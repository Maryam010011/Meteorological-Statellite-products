"""
batch_run.py  -  Process all 22 unique MSG SEVIRI scenes sequentially.
Writes output_v2/<timestamp>/ folders + output_v2/summary.json
"""
import gc, glob, hashlib, json, logging, os, sys, traceback
from datetime import datetime

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "batch_log.txt")
log = logging.getLogger("batch")
log.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s", datefmt="%Y-%m-%dT%H:%M:%S")
fh = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w"); fh.setFormatter(fmt); log.addHandler(fh)
ch = logging.StreamHandler(sys.stdout);                          ch.setFormatter(fmt); log.addHandler(ch)

BASE     = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_ROOT = os.path.join(BASE, "output_v2")
sys.path.insert(0, BASE)

from visualize_msg_v2 import process_file, ensure_dir

def md5(p):
    h = hashlib.md5()
    with open(p, "rb") as f:
        for c in iter(lambda: f.read(65536), b""): h.update(c)
    return h.hexdigest()

def sensing_time(fp, idx):
    from satpy import Scene
    from visualize_msg_v2 import standard_nat_name, make_hardlink, remove_if_exists
    tmp = standard_nat_name(DATA_DIR, idx + 1000)
    make_hardlink(fp, tmp)
    try:
        sc = Scene(reader="seviri_l1b_native", filenames=[tmp])
        sc.load(["IR_108"])
        t = sc["IR_108"].attrs.get("start_time")
        del sc; gc.collect()
        return t
    except:
        return None
    finally:
        remove_if_exists(tmp)

def main():
    log.info("="*70)
    log.info("batch_run.py  --  MSG SEVIRI full batch")
    log.info("Started: %s", datetime.utcnow().isoformat())

    nat_files = sorted(glob.glob(os.path.join(DATA_DIR, "*.nat")))
    nat_files = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]
    log.info("Found %d .nat files", len(nat_files))

    seen, unique = {}, []
    for f in nat_files:
        h = md5(f)
        if h in seen:
            log.warning("Skipping duplicate: %s", os.path.basename(f))
        else:
            seen[h] = f; unique.append(f)
    log.info("%d unique files", len(unique))

    log.info("Sorting by sensing time ...")
    timed = []
    for i, fp in enumerate(unique):
        t = sensing_time(fp, i)
        timed.append((t or datetime.max, fp))
    timed.sort(key=lambda x: x[0])
    ordered = [fp for _, fp in timed]

    log.info("Chronological order:")
    for t, fp in timed:
        log.info("  %s  |  %s", str(t)[:19], os.path.basename(fp))

    ensure_dir(OUT_ROOT)
    manifests, ok_count = [], 0
    total_ok = total_err = 0
    failures = []

    for idx, fp in enumerate(ordered, 1):
        fname = os.path.basename(fp)
        log.info("")
        log.info("-"*70)
        log.info("SCENE %d/%d  --  %s", idx, len(ordered), fname)
        log.info("-"*70)
        try:
            m = process_file(fp, DATA_DIR, OUT_ROOT, idx)
            manifests.append(m)
            prods = m.get("products", [])
            ok  = sum(1 for p in prods if p["status"] == "ok")
            err = sum(1 for p in prods if p["status"] == "error")
            total_ok += ok; total_err += err
            ts = str(m.get("sensing_start","?"))[:19]
            if err:
                log.warning("! SCENE %d  %s  --  %d OK | %d errors", idx, ts, ok, err)
                for p in prods:
                    if p["status"] == "error":
                        log.warning("    FAIL  %s: %s", p["product"], str(p.get("reason",""))[:100])
                        failures.append({"scene":idx,"file":fname,"product":p["product"],"reason":str(p.get("reason",""))[:200]})
            else:
                log.info("OK SCENE %d  %s  --  %d/%d images", idx, ts, ok, ok+err)
            ok_count += 1
        except Exception as exc:
            log.error("Unhandled error %s: %s", fname, exc)
            traceback.print_exc()
            failures.append({"scene":idx,"file":fname,"reason":str(exc)})
        gc.collect()

    log.info("")
    log.info("="*70)
    log.info("RUN COMPLETE")
    log.info("  Scenes : %d / %d succeeded", ok_count, len(ordered))
    log.info("  Images : %d OK | %d errors", total_ok, total_err)
    log.info("  Output : %s", OUT_ROOT)
    log.info("="*70)

    with open(os.path.join(OUT_ROOT, "summary.json"), "w", encoding="utf-8") as f:
        json.dump({"run_time":datetime.utcnow().isoformat(),
                   "scenes_succeeded":ok_count, "total_ok_images":total_ok,
                   "total_errors":total_err, "failures":failures,
                   "manifests":manifests}, f, indent=2, default=str)
    log.info("Summary: %s", os.path.join(OUT_ROOT, "summary.json"))
    log.info("Log:     %s", LOG_FILE)

if __name__ == "__main__":
    main()
