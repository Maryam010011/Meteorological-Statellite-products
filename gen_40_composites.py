"""
gen_40_composites.py
====================
Generates 40 color-variant versions of each composite for each scene.
Structure:
  output_v2/<ts>/composites_40/<comp_name>/<slug>_<colorname>.png

Run with --sample to generate only scene 1, dust composite as preview.
Run without args to generate all 22 scenes x 9 composites x 40 colors.
"""
import warnings; warnings.filterwarnings("ignore")
import gc, glob, hashlib, json, os, sys, traceback
from datetime import datetime

import numpy as np
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import cartopy.crs as ccrs
import cartopy.feature as cfeature
from scipy.ndimage import zoom as _zoom

BASE     = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_ROOT = os.path.join(BASE, "output_v2")
sys.path.insert(0, BASE)

from satpy import Scene
from satpy.enhancements.enhancer import get_enhanced_image
from visualize_msg_v2 import (
    standard_nat_name, make_hardlink, remove_if_exists, ensure_dir, slug
)

# ── Cartopy setup ─────────────────────────────────────────────────────────────
# Use Geostationary projection — naturally clips to circular disk boundary
# so coastlines CANNOT go outside the disk. This matches the reference image.
_CRS      = ccrs.Geostationary(central_longitude=45.5, satellite_height=35_785_831)
_GEOS_EXT = 5_570_248.4773   # native SEVIRI full-disk half-extent in metres

# ── 40 color palettes ─────────────────────────────────────────────────────────
PALETTES = [
    # id,  cmap_name,            display_label
    ("01", "gray_r",             "IR Gray (cold bright)"),
    ("02", "gray",               "IR Gray (warm bright)"),
    ("03", "jet",                "Rainbow Jet"),
    ("04", "hsv",                "HSV Spectrum"),
    ("05", "inferno",            "Thermal Inferno"),
    ("06", "magma",              "Magma"),
    ("07", "plasma",             "Plasma"),
    ("08", "viridis",            "Viridis"),
    ("09", "cividis",            "Cividis"),
    ("10", "Spectral",           "Spectral"),
    ("11", "Spectral_r",         "Spectral Reversed"),
    ("12", "RdYlBu",             "Red-Yellow-Blue"),
    ("13", "RdYlBu_r",           "Red-Yellow-Blue Inv"),
    ("14", "coolwarm",           "Cool-Warm"),
    ("15", "bwr",                "Blue-White-Red"),
    ("16", "seismic",            "Seismic"),
    ("17", "YlOrRd_r",           "Yellow-Orange-Red"),
    ("18", "YlOrRd",             "Yellow-Orange-Red Inv"),
    ("19", "hot",                "Hot"),
    ("20", "hot_r",              "Hot Reversed"),
    ("21", "afmhot",             "AFM Hot"),
    ("22", "gist_heat",          "Gist Heat"),
    ("23", "copper",             "Copper"),
    ("24", "autumn",             "Autumn"),
    ("25", "summer",             "Summer"),
    ("26", "spring",             "Spring"),
    ("27", "winter",             "Winter"),
    ("28", "ocean",              "Ocean"),
    ("29", "terrain",            "Terrain"),
    ("30", "gist_earth",         "Gist Earth"),
    ("31", "gist_ncar",          "NCAR"),
    ("32", "gist_rainbow",       "Gist Rainbow"),
    ("33", "nipy_spectral",      "NIPY Spectral"),
    ("34", "turbo",              "Turbo"),
    ("35", "gnuplot",            "GNUPlot"),
    ("36", "gnuplot2",           "GNUPlot2"),
    ("37", "CMRmap",             "CMR Map"),
    ("38", "cubehelix",          "Cube Helix"),
    ("39", "brg",                "BRG"),
    ("40", "tab20b",             "Tab 20B"),
]

COMPOSITES = [
    "natural_color",
    "airmass",
    "dust",
    "ash",
    "convection",
    "fog",
    "night_microphysics",
    "day_severe_storms",
    "colorized_ir_clouds",
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for c in iter(lambda: f.read(65536), b""): h.update(c)
    return h.hexdigest()

def get_composite_rgb(nat_tmp, comp_key):
    """Load composite from native scene, force compute, return (H,W,3) 0-1."""
    sc = Scene(reader="seviri_l1b_native", filenames=[nat_tmp])
    sc.load([comp_key])
    try:
        _ = sc[comp_key]
    except KeyError:
        del sc; gc.collect()
        return None
    ximg = get_enhanced_image(sc[comp_key])
    data = ximg.data.compute().values      # (3, H, W)
    arr  = np.transpose(data[:3], (1,2,0)).astype(np.float32)
    arr  = np.clip(arr, 0, 1)
    del sc; gc.collect()
    # Resize to 2048x2048 if needed
    if arr.shape[0] != 2048:
        factor = 2048 / arr.shape[0]
        arr = np.stack(
            [_zoom(arr[:,:,b], factor, order=1) for b in range(3)], axis=2
        ).astype(np.float32)
        arr = np.clip(arr, 0, 1)
    return arr

def rgb_to_luminance(arr_hwc):
    """Convert (H,W,3) RGB to (H,W) luminance for colormap rendering."""
    return 0.299*arr_hwc[:,:,0] + 0.587*arr_hwc[:,:,1] + 0.114*arr_hwc[:,:,2]

def render_colormap(lum, cmap_name, out_path, title_str):
    """Render luminance array in Geostationary projection.
    The geos CRS naturally clips everything to the circular disk —
    coastlines physically cannot appear outside the disk boundary."""
    import matplotlib.patches as mpatches

    fig, ax = plt.subplots(1, 1, figsize=(10, 10),
                           subplot_kw={"projection": _CRS},
                           facecolor="white")
    ax.set_facecolor("black")
    ax.set_global()   # show full disk

    # Place the data using the geostationary metre-based extent
    img = ax.imshow(lum, origin="upper",
                    cmap=cmap_name, vmin=0.0, vmax=1.0,
                    extent=[-_GEOS_EXT, _GEOS_EXT, -_GEOS_EXT, _GEOS_EXT],
                    transform=_CRS,
                    interpolation="nearest", zorder=3)

    # Coastlines and borders — automatically clipped to disk by geos CRS
    ax.add_feature(cfeature.COASTLINE.with_scale("50m"),
                   linewidth=0.6, edgecolor="white", zorder=5)
    ax.add_feature(cfeature.BORDERS.with_scale("50m"),
                   linewidth=0.3, edgecolor="white", zorder=5)

    # Gridlines with lat/lon labels
    try:
        gl = ax.gridlines(crs=ccrs.PlateCarree(), draw_labels=True,
                          linewidth=0.3, color="white", alpha=0.4,
                          x_inline=False, y_inline=False)
        gl.top_labels   = False
        gl.right_labels = False
        gl.xlabel_style = {"size": 7, "color": "black"}
        gl.ylabel_style = {"size": 7, "color": "black"}
        gl.xlocator = mticker.FixedLocator(range(-180, 181, 30))
        gl.ylocator = mticker.FixedLocator(range(-90, 91, 30))
    except Exception:
        pass

    # Colorbar and title
    cb = fig.colorbar(img, ax=ax, orientation="horizontal",
                      pad=0.04, fraction=0.03, shrink=0.7)
    cb.set_label("Composite intensity", color="black", fontsize=7)
    cb.ax.xaxis.set_tick_params(color="black", labelcolor="black", labelsize=6)
    ax.set_title(title_str, color="black", fontsize=8, pad=8, loc="left")

    fig.savefig(out_path, dpi=150, facecolor="white")
    plt.close(fig)
    gc.collect()

# ── Per-scene processor ───────────────────────────────────────────────────────

def process_scene(nat_path, ts_str, sensing_start, sensing_end, platform,
                  comp_list, idx):
    """Generate 40-color variants for each composite in comp_list."""
    tmp = standard_nat_name(DATA_DIR, idx + 2000)
    make_hardlink(nat_path, tmp)

    ts_label = (sensing_start.strftime("%Y-%m-%d %H:%M UTC")
                if hasattr(sensing_start,"strftime") else str(sensing_start)[:16]+" UTC")
    te_label = (sensing_end.strftime("%H:%M UTC")
                if hasattr(sensing_end,"strftime") else str(sensing_end)[11:16]+" UTC")

    scene_ok = scene_err = 0

    try:
        for comp_key in comp_list:
            # Output folder: output_v2/<ts>/composites_40/<comp_key>/
            out_dir = ensure_dir(os.path.join(
                OUT_ROOT, ts_str, "composites_40", comp_key))

            # Check if already fully done (40 files exist)
            existing = [f for f in os.listdir(out_dir) if f.endswith(".png")]
            if len(existing) >= 40:
                print(f"    SKIP {comp_key} (already {len(existing)} files)")
                scene_ok += 40
                continue

            # Load composite RGB
            try:
                arr = get_composite_rgb(tmp, comp_key)
            except Exception as exc:
                print(f"    FAIL load {comp_key}: {exc}")
                scene_err += 40
                continue

            if arr is None:
                print(f"    FAIL {comp_key}: returned None")
                scene_err += 40
                continue

            # Convert to luminance for colormap rendering
            lum = rgb_to_luminance(arr)
            del arr; gc.collect()

            # Generate 40 colormap versions
            comp_label = comp_key.replace("_", " ").title()
            for pal_id, cmap_name, pal_label in PALETTES:
                fname = "{}_{}_{}_{}.png".format(
                    ts_str, comp_key, pal_id,
                    cmap_name.replace("_","").replace(".","").lower()[:12])
                out_path = os.path.join(out_dir, fname)

                if os.path.exists(out_path) and os.path.getsize(out_path) > 5000:
                    scene_ok += 1
                    continue

                title = ("{} — {}\n{} IODC 45.5°E  |  {} - {}".format(
                    comp_label, pal_label, platform, ts_label, te_label))

                try:
                    render_colormap(lum, cmap_name, out_path, title)
                    scene_ok += 1
                except Exception as exc:
                    print(f"      FAIL {comp_key} {pal_id}: {exc}")
                    scene_err += 1

            del lum; gc.collect()
            print(f"    OK  {comp_key}  ({len(PALETTES)} colors)")

    finally:
        remove_if_exists(tmp)

    return scene_ok, scene_err


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    sample_mode = "--sample" in sys.argv

    # Load scene list from summary.json
    summary_path = os.path.join(OUT_ROOT, "summary.json")
    with open(summary_path, encoding="utf-8") as f:
        summary = json.load(f)

    # Build nat file lookup
    nat_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.nat")))
    nat_files = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]
    seen, unique = {}, []
    for fp in nat_files:
        h = md5(fp)
        if h not in seen: seen[h]=fp; unique.append(fp)

    fname_map = {os.path.basename(fp): fp for fp in unique}

    manifests = summary.get("manifests", [])

    if sample_mode:
        manifests = manifests[:1]
        comp_list = ["dust"]
        print("SAMPLE MODE: 1 scene, dust composite, 40 colors")
    else:
        comp_list = COMPOSITES
        print(f"FULL MODE: {len(manifests)} scenes x {len(comp_list)} composites x {len(PALETTES)} colors")
        print(f"Total images to generate: {len(manifests)*len(comp_list)*len(PALETTES):,}")

    total_ok = total_err = 0

    for i, m in enumerate(manifests, 1):
        fname = m.get("file","")
        nat_path = fname_map.get(fname)
        if not nat_path:
            print(f"  SKIP scene {i}: cannot find .nat for {fname}")
            continue

        ts_str = m.get("sensing_start","")[:16].replace("-","").replace(" ","T").replace(":","")[:13]+"Z"
        # Fix format: 20260622T0000Z
        try:
            from datetime import datetime as _dt
            dt = _dt.strptime(str(m.get("sensing_start",""))[:16], "%Y-%m-%d %H:%M")
            ts_str = dt.strftime("%Y%m%dT%H%MZ")
            sensing_start = dt
            sensing_end_str = str(m.get("sensing_end",""))[:16]
            sensing_end = _dt.strptime(sensing_end_str, "%Y-%m-%d %H:%M")
        except Exception:
            sensing_start = sensing_end = None

        platform = m.get("platform","Meteosat-9")

        print(f"\nScene {i}/{len(manifests)}  --  {fname}  [{ts_str}]")

        ok, err = process_scene(
            nat_path, ts_str, sensing_start, sensing_end,
            platform, comp_list, i)

        total_ok  += ok
        total_err += err
        print(f"  Scene result: {ok} OK | {err} errors")

    print(f"\n{'='*60}")
    print(f"DONE: {total_ok} images OK | {total_err} errors")
    print(f"Output: {OUT_ROOT}/<ts>/composites_40/<comp>/")


if __name__ == "__main__":
    main()
