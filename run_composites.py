"""
run_composites.py
=================
Generates ONLY the missing composites for all 22 scenes.
Does NOT touch or delete existing channel PNGs/GeoTIFFs.
Reads existing output_v2/<ts>/manifest.json to know which composites
are missing, then generates only those.
"""
import gc, glob, hashlib, json, logging, os, sys, traceback
from datetime import datetime

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "composites_log.txt")
log = logging.getLogger("composites")
log.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%Y-%m-%dT%H:%M:%S")
fh = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w"); fh.setFormatter(fmt); log.addHandler(fh)
ch = logging.StreamHandler(sys.stdout);                          ch.setFormatter(fmt); log.addHandler(ch)

BASE     = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_ROOT = os.path.join(BASE, "output_v2")
sys.path.insert(0, BASE)

import numpy as np
from satpy import Scene
from satpy.enhancements.enhancer import get_enhanced_image
from visualize_msg_v2 import (
    standard_nat_name, make_hardlink, remove_if_exists, ensure_dir,
    save_rgb_png, save_geotiff, get_composite_array,
    COMPOSITES, _title, slug, _EQC_PROJ, _EQC_EXTENT_M
)

def md5(p):
    h = hashlib.md5()
    with open(p,"rb") as f:
        for c in iter(lambda: f.read(65536), b""): h.update(c)
    return h.hexdigest()

def find_nat_for_ts(ts_str):
    """Find the .nat file whose sensing time matches ts_str (YYYYMMDDTHHmmZ)."""
    nat_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.nat")))
    nat_files = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]
    seen, unique = {}, []
    for f in nat_files:
        h = md5(f)
        if h not in seen: seen[h]=f; unique.append(f)
    # Match by reading sensing time from summary.json
    summary_path = os.path.join(OUT_ROOT, "summary.json")
    with open(summary_path, encoding="utf-8") as f:
        summary = json.load(f)
    for m in summary["manifests"]:
        t = str(m.get("sensing_start",""))
        # Convert to ts format
        try:
            dt = datetime.strptime(t[:16], "%Y-%m-%d %H:%M")
            if dt.strftime("%Y%m%dT%H%MZ") == ts_str:
                fname = m["file"]
                for fp in unique:
                    if os.path.basename(fp) == fname:
                        return fp
        except: pass
    return None

def main():
    log.info("="*70)
    log.info("run_composites.py  --  composites-only pass")
    log.info("Started: %s", datetime.utcnow().isoformat())

    # Find all timestamp folders
    ts_folders = sorted([
        d for d in glob.glob(os.path.join(OUT_ROOT, "*"))
        if os.path.isdir(d) and not d.endswith("summary.json")
    ])
    log.info("Found %d scene folders", len(ts_folders))

    total_ok = total_err = 0

    for folder_path in ts_folders:
        ts_str = os.path.basename(folder_path)
        log.info("")
        log.info("-"*70)
        log.info("Scene: %s", ts_str)

        # Find matching .nat file
        nat_path = find_nat_for_ts(ts_str)
        if nat_path is None:
            log.error("  Cannot find .nat file for %s", ts_str)
            continue

        out_comp = ensure_dir(os.path.join(folder_path, "composites"))
        out_tif  = ensure_dir(os.path.join(folder_path, "geotiff"))

        tmp = standard_nat_name(DATA_DIR, abs(hash(ts_str)) % 9000 + 1000)
        make_hardlink(nat_path, tmp)

        # Get metadata
        try:
            sc0 = Scene(reader="seviri_l1b_native", filenames=[tmp])
            sc0.load(["IR_108"])
            t0   = sc0["IR_108"].attrs.get("start_time", datetime.utcnow())
            t1   = sc0["IR_108"].attrs.get("end_time",   datetime.utcnow())
            plat = sc0["IR_108"].attrs.get("platform_name", "Meteosat-9")
            del sc0; gc.collect()
        except Exception as exc:
            log.error("  Cannot read metadata: %s", exc)
            remove_if_exists(tmp)
            continue

        scene_ok = scene_err = 0

        for comp_key, comp_label in COMPOSITES.items():
            sl       = slug(comp_label)
            png_path = os.path.join(out_comp, "{}.png".format(sl))
            tif_path = os.path.join(out_tif,  "{}.tif".format(sl))

            # Skip if already exists
            if os.path.exists(png_path) and os.path.getsize(png_path) > 10000:
                log.info("  SKIP (exists): %s", comp_label)
                scene_ok += 1
                continue

            try:
                sc = Scene(reader="seviri_l1b_native", filenames=[tmp])
                sc.load([comp_key])

                # Access on native scene, force compute while file open
                try:
                    _ = sc[comp_key]
                except KeyError:
                    raise RuntimeError("'{}' not in native scene".format(comp_key))

                arr = get_composite_array(sc, comp_key)
                del sc; gc.collect()

                if arr is None:
                    raise RuntimeError("get_composite_array returned None")

                # Resize from native 3712x3712 to 2048x2048
                if arr.shape[0] != 2048:
                    from scipy.ndimage import zoom as _zoom
                    factor = 2048 / arr.shape[0]
                    arr = np.stack([
                        _zoom(arr[:,:,b], factor, order=1) for b in range(3)
                    ], axis=2).astype(np.float32)
                    arr = np.clip(arr, 0, 1)

                title = _title(comp_label, plat, t0, t1)
                ok_p  = save_rgb_png(arr, png_path, title)
                ok_t  = save_geotiff(arr, tif_path)

                if ok_p:
                    scene_ok += 1
                    log.info("  OK  %s", comp_label)
                else:
                    scene_err += 1
                    log.error("  FAIL  %s (PNG save error)", comp_label)

                del arr; gc.collect()

            except Exception as exc:
                scene_err += 1
                log.error("  FAIL  %s: %s", comp_label, str(exc)[:120])
                traceback.print_exc()
                gc.collect()

        remove_if_exists(tmp)
        total_ok  += scene_ok
        total_err += scene_err
        log.info("  Scene done: %d OK | %d errors", scene_ok, scene_err)

    log.info("")
    log.info("="*70)
    log.info("COMPOSITES COMPLETE: %d OK | %d errors", total_ok, total_err)
    log.info("="*70)

if __name__ == "__main__":
    main()
