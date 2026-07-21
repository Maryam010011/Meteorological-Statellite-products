"""
inspect_msg.py
==============
Part 1 Inspection Script — MSG SEVIRI Level 1.5 native (.nat) files.

Calculates MD5 checksums, reads internal metadata (platform, sensing time,
projection) via Satpy, detects duplicates, and prints a clean summary.

Usage:
    python inspect_msg.py <data_dir>

Example:
    python inspect_msg.py data/
"""

import gc
import glob
import hashlib
import os
import sys
import traceback

PYTHON = sys.executable

def calculate_md5(filepath):
    """Return the MD5 hex digest of *filepath*."""
    h = hashlib.md5()
    with open(filepath, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def standard_nat_name(data_dir, index):
    """Return a temp filename that Satpy's seviri_l1b_native reader accepts."""
    return os.path.join(
        data_dir,
        f"MSG3-SEVI-MSG15-0100-NA-20000101000000.000000000Z-INSPECT{index:04d}.nat"
    )


def make_hardlink(src, dst):
    import shutil
    if os.path.exists(dst):
        os.remove(dst)
    try:
        os.link(src, dst)
    except OSError:
        shutil.copy2(src, dst)


def inspect_files(data_dir):
    nat_files = sorted(glob.glob(os.path.join(data_dir, "*.nat")))
    nat_files = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]

    print("=" * 100)
    print("  MSG SEVIRI Level 1.5 Native Data — Inspection Report")
    print("=" * 100)
    print(f"  Directory : {os.path.abspath(data_dir)}")
    print(f"  Files found: {len(nat_files)}")
    print()

    # -----------------------------------------------------------------------
    # Pass 1: MD5 checksums (no Satpy needed)
    # -----------------------------------------------------------------------
    print(f"{'#':<4} {'Filename':<22} {'Size (MB)':<10} {'MD5 Checksum':<34}")
    print("-" * 75)
    checksums = {}
    file_md5 = {}
    for i, fp in enumerate(nat_files, 1):
        fname = os.path.basename(fp)
        size_mb = os.path.getsize(fp) / 1_048_576
        md5 = calculate_md5(fp)
        file_md5[fp] = md5
        checksums.setdefault(md5, []).append(fname)
        print(f"{i:<4} {fname:<22} {size_mb:<10.1f} {md5:<34}")

    print()
    duplicates = {md5: files for md5, files in checksums.items() if len(files) > 1}
    if duplicates:
        print("[!] DUPLICATE FILES DETECTED (identical MD5):")
        for md5, files in duplicates.items():
            print(f"    MD5: {md5}")
            for f in files:
                print(f"      -> {f}")
        print()
    else:
        print("[OK] All files have unique MD5 checksums.\n")

    # -----------------------------------------------------------------------
    # Pass 2: Satpy metadata (sensing time, platform, projection, channels)
    # -----------------------------------------------------------------------
    print("Loading metadata with Satpy (seviri_l1b_native reader)...")
    print("-" * 100)
    header = (f"{'Filename':<22} | {'Platform':<12} | {'Start Time (UTC)':<20} "
              f"| {'End Time (UTC)':<20} | {'lon_0':>6} | {'Status'}")
    print(header)
    print("-" * 100)

    from satpy import Scene

    seen_md5 = set()
    results = []

    for idx, fp in enumerate(nat_files):
        fname = os.path.basename(fp)
        md5 = file_md5[fp]
        is_dup = md5 in seen_md5
        seen_md5.add(md5)

        temp_path = standard_nat_name(data_dir, idx)
        try:
            make_hardlink(fp, temp_path)
            scene = Scene(reader="seviri_l1b_native", filenames=[temp_path])
            datasets = scene.available_dataset_names()
            ref = "IR_108" if "IR_108" in datasets else (datasets[0] if datasets else None)
            if ref is None:
                raise ValueError("No datasets available")
            scene.load([ref])
            meta = scene[ref].attrs
            start = meta.get("start_time")
            end   = meta.get("end_time")
            plat  = meta.get("platform_name", "Unknown")
            area  = meta.get("area")
            lon0  = area.crs.to_dict().get("lon_0", "?") if area else "?"
            status = "DUPLICATE" if is_dup else "OK"
            row = dict(fname=fname, md5=md5, platform=plat,
                       start=str(start), end=str(end),
                       lon0=lon0, channels=sorted(datasets),
                       status=status)
            print(f"{fname:<22} | {plat:<12} | {str(start):<20} | {str(end):<20} "
                  f"| {str(lon0):>6} | {status}")
            results.append(row)
            del scene
            gc.collect()
        except Exception as exc:
            print(f"{fname:<22} | {'?':<12} | {'?':<20} | {'?':<20} | {'?':>6} | ERROR: {exc}")
            results.append(dict(fname=fname, md5=md5, status=f"ERROR: {exc}"))
            traceback.print_exc()
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    print("\n" + "=" * 100)
    unique_scenes = [r for r in results if r.get("status") == "OK"]
    dup_scenes    = [r for r in results if r.get("status") == "DUPLICATE"]
    error_scenes  = [r for r in results if r.get("status", "").startswith("ERROR")]

    print(f"  Total files     : {len(nat_files)}")
    print(f"  Unique scenes   : {len(unique_scenes)}")
    print(f"  Duplicates      : {len(dup_scenes)}")
    print(f"  Errors          : {len(error_scenes)}")

    if unique_scenes:
        starts = [r["start"] for r in unique_scenes if "start" in r]
        print(f"  Earliest scene  : {min(starts)}")
        print(f"  Latest scene    : {max(starts)}")

    # Print channel list from last successfully loaded scene
    for r in reversed(results):
        if "channels" in r:
            print(f"\n  Available SEVIRI channels (from last loaded scene):")
            print("  " + ", ".join(r["channels"]))
            break

    # SatDump status
    print("\n" + "=" * 100)
    print("  SatDump Status:")
    import subprocess
    result = subprocess.run(["where", "satdump"], capture_output=True, text=True, shell=True)
    if result.returncode == 0:
        print(f"  [OK] SatDump found at: {result.stdout.strip()}")
    else:
        print("  [!!] SatDump is NOT installed on this system.")
        print("       For MSG Level 1.5 native files, SatDump has PARTIAL support")
        print("       (experimental pipeline). Since it is absent, we use Satpy exclusively.")
    print("=" * 100)


if __name__ == "__main__":
    data_directory = sys.argv[1] if len(sys.argv) > 1 else "data"
    inspect_files(data_directory)
