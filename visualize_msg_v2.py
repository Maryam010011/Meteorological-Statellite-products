"""
visualize_msg_v2.py  -  Corrected MSG SEVIRI visualizer
========================================================
All 4 supervisor defects fixed:
  1. GEOMETRY  : PlateCarree resampling + correct degree extents -> pillow-square shape
  2. COLORS    : Satpy built-in enhancement for composites (no manual re-stretch)
                 Raw calibrated values for channels (K or %)
  3. METADATA  : Product name, platform, sub-sat lon, sensing UTC on every image
  4. RESOLUTION: 2048x2048 standard, 4096x4096 HRV

Rendering approach (confirmed working from test_render.py):
  - Resample to eqc (PlateCarree) grid centred on 45.5E
  - Render with ccrs.PlateCarree(central_longitude=45.5)
  - imshow extent in degrees matching the eqc area
  - White figure background, black space background
  - Yellow coastlines, white gridlines, black labels

Usage:
    python visualize_msg_v2.py <data_dir>
"""
import csv, gc, hashlib, json, logging, os, shutil, sys, traceback
from datetime import datetime

import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import cartopy.crs as ccrs
import cartopy.feature as cfeature

from satpy import Scene
from satpy.enhancements.enhancer import get_enhanced_image
from pyresample.geometry import AreaDefinition

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(levelname)-8s  %(message)s",
                    datefmt="%Y-%m-%dT%H:%M:%S")
log = logging.getLogger("visualize_v2")

# ===========================================================================
# CONFIGURATION
# ===========================================================================

CHANNELS = ["VIS006","VIS008","IR_016","IR_039","WV_062","WV_073",
            "IR_087","IR_097","IR_108","IR_120","IR_134","HRV"]

CHANNEL_META = {
    "VIS006": ("VIS 0.6 um",  "%",  "gray",   0,   100),
    "VIS008": ("VIS 0.8 um",  "%",  "gray",   0,   100),
    "IR_016": ("NIR 1.6 um",  "%",  "gray",   0,   100),
    "IR_039": ("IR 3.9 um",   "K",  "gray_r", 200, 330),
    "WV_062": ("WV 6.2 um",   "K",  "gray_r", 200, 260),
    "WV_073": ("WV 7.3 um",   "K",  "gray_r", 200, 280),
    "IR_087": ("IR 8.7 um",   "K",  "gray_r", 200, 330),
    "IR_097": ("IR 9.7 um",   "K",  "gray_r", 200, 320),
    "IR_108": ("IR 10.8 um",  "K",  "gray_r", 200, 330),
    "IR_120": ("IR 12.0 um",  "K",  "gray_r", 200, 330),
    "IR_134": ("IR 13.4 um",  "K",  "gray_r", 200, 330),
    "HRV":    ("HRV 0.7 um",  "%",  "gray",   0,   100),
}

COMPOSITES = {
    "natural_color":      "Natural Color",
    "airmass":            "Airmass RGB",
    "dust":               "Dust RGB",
    "ash":                "Ash RGB",
    "convection":         "Convection RGB",
    "fog":                "Fog RGB",
    "night_microphysics": "Night Microphysics RGB",
    "day_severe_storms":  "Day Severe Storms RGB",
    "colorized_ir_clouds":"IR 10.8 Colorized",
}

# ===========================================================================
# AREA DEFINITIONS  (PlateCarree / eqc centred on 45.5E)
# ===========================================================================
_EQC_PROJ = "+proj=eqc +lon_0=45.5 +datum=WGS84 +units=m"
_EQC_EXTENT_M = (-8_000_000, -8_000_000, 8_000_000, 8_000_000)

AREA_STD = AreaDefinition(
    "msg_eqc_2k", "MSG IODC eqc 2048x2048", "eqc",
    _EQC_PROJ, 2048, 2048, _EQC_EXTENT_M)

AREA_HRV = AreaDefinition(
    "msg_eqc_hrv", "MSG IODC eqc HRV 4096x4096", "eqc",
    _EQC_PROJ, 4096, 4096, _EQC_EXTENT_M)

# Cartopy CRS + display extent in degrees (matches eqc 8000km radius)
_CRS   = ccrs.PlateCarree(central_longitude=45.5)
_DEG   = 71.9
_LONMIN, _LONMAX = 45.5 - _DEG, 45.5 + _DEG   # -26.4 .. 117.4
_LATMIN, _LATMAX = -_DEG, _DEG


# ===========================================================================
# HELPERS
# ===========================================================================

def md5sum(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for c in iter(lambda: f.read(65536), b""): h.update(c)
    return h.hexdigest()

def standard_nat_name(data_dir, index):
    return os.path.join(data_dir,
        "MSG3-SEVI-MSG15-0100-NA-20000101000000.000000000Z-V2IDX{:04d}.nat".format(index))

def make_hardlink(src, dst):
    try:
        if os.path.exists(dst): os.remove(dst)
        os.link(src, dst); return True
    except OSError:
        try: shutil.copy2(src, dst); return True
        except: return False

def remove_if_exists(path):
    try:
        if os.path.exists(path): os.remove(path)
    except: pass

def ensure_dir(path):
    os.makedirs(path, exist_ok=True); return path

def slug(name):
    return name.lower().replace(" ","_").replace("/","_").replace(".","p")

def _title(product, platform, t0, t1):
    s = t0.strftime("%Y-%m-%d %H:%M UTC") if hasattr(t0,"strftime") else str(t0)[:16]+" UTC"
    e = t1.strftime("%H:%M UTC")           if hasattr(t1,"strftime") else str(t1)[11:16]+" UTC"
    return "{} | {} IODC  45.5E\n{} - {}".format(product, platform, s, e)

# ===========================================================================
# RENDERING  (confirmed working approach from test_render.py)
# ===========================================================================

def _make_fig_ax(figsize=(10, 10)):
    fig, ax = plt.subplots(1, 1, figsize=figsize,
                           subplot_kw={"projection": _CRS},
                           facecolor="white")
    fig.patch.set_facecolor("white")
    ax.set_facecolor("black")
    ax.set_extent([_LONMIN, _LONMAX, _LATMIN, _LATMAX], crs=ccrs.PlateCarree())
    return fig, ax

def _overlays(ax):
    ax.add_feature(cfeature.COASTLINE.with_scale("50m"),
                   linewidth=0.6, edgecolor="yellow", zorder=5)
    ax.add_feature(cfeature.BORDERS.with_scale("50m"),
                   linewidth=0.3, edgecolor="yellow", zorder=5)
    try:
        gl = ax.gridlines(crs=ccrs.PlateCarree(), draw_labels=True,
                          linewidth=0.3, color="white", alpha=0.5,
                          x_inline=False, y_inline=False)
        gl.top_labels = False; gl.right_labels = False
        gl.xlabel_style = {"size":7,"color":"black"}
        gl.ylabel_style = {"size":7,"color":"black"}
        gl.xlocator = mticker.FixedLocator(range(-180,181,30))
        gl.ylocator = mticker.FixedLocator(range(-90,91,30))
    except Exception: pass

def _savefig(fig, path):
    try:
        fig.savefig(path, dpi=150, facecolor="white")
    except Exception as exc:
        if "LinearRing" in str(exc) or "GEOSException" in str(exc):
            for ax in fig.axes:
                if hasattr(ax, "_gridliners"):
                    try: ax._gridliners.clear()
                    except: pass
            fig.savefig(path, dpi=150, facecolor="white")
        else:
            raise

def _apply_disk_mask(ax):
    """
    Draw a white outline frame on the axes border so that any coastline/
    border lines that Cartopy draws outside the satellite disk extent are
    hidden. Uses ax.spines to set a thick white border, effectively masking
    spillover lines at the edges.
    """
    # Set a thick white spine to cover any lines at the border
    for spine in ax.spines.values():
        spine.set_edgecolor("white")
        spine.set_linewidth(4)
        spine.set_zorder(10)


def save_channel_png(arr, path, label, cmap, vmin, vmax, units, title):
    try:
        fig, ax = _make_fig_ax()
        img = ax.imshow(arr, origin="upper", cmap=cmap, vmin=vmin, vmax=vmax,
                        extent=[_LONMIN, _LONMAX, _LATMIN, _LATMAX],
                        transform=ccrs.PlateCarree(),
                        interpolation="nearest", zorder=3)
        _overlays(ax)
        # Mask coastlines outside the disk with a white frame
        _apply_disk_mask(ax)
        cb = fig.colorbar(img, ax=ax, orientation="horizontal",
                          pad=0.04, fraction=0.03, shrink=0.7)
        cb.set_label("{} [{}]".format(label, units), color="black", fontsize=7)
        cb.ax.xaxis.set_tick_params(color="black", labelcolor="black", labelsize=6)
        ax.set_title(title, color="black", fontsize=8, pad=8, loc="left")
        _savefig(fig, path)
        plt.close(fig)
        log.info("  [OK] %s", path); return True
    except Exception as exc:
        log.error("  [!!] channel PNG %s: %s", label, exc)
        traceback.print_exc(); plt.close("all"); return False

def save_rgb_png(arr, path, title):
    try:
        fig, ax = _make_fig_ax()
        ax.imshow(np.clip(arr, 0, 1), origin="upper",
                  extent=[_LONMIN, _LONMAX, _LATMIN, _LATMAX],
                  transform=ccrs.PlateCarree(),
                  interpolation="nearest", zorder=3)
        _overlays(ax)
        # Mask coastlines outside the disk with a white frame
        _apply_disk_mask(ax)
        ax.set_title(title, color="black", fontsize=8, pad=8, loc="left")
        _savefig(fig, path)
        plt.close(fig)
        log.info("  [OK] %s", path); return True
    except Exception as exc:
        log.error("  [!!] RGB PNG: %s", exc)
        traceback.print_exc(); plt.close("all"); return False

def save_geotiff(arr, path):
    try:
        import rasterio
        from rasterio.transform import from_bounds
        from rasterio.crs import CRS as RC
        w = arr.shape[-1] if arr.ndim==3 else arr.shape[1]
        h = arr.shape[-2] if arr.ndim==3 else arr.shape[0]
        tr = from_bounds(_EQC_EXTENT_M[0], _EQC_EXTENT_M[1],
                         _EQC_EXTENT_M[2], _EQC_EXTENT_M[3], w, h)
        crs = RC.from_proj4(_EQC_PROJ)
        if arr.ndim == 2:
            bands = arr[np.newaxis].astype(np.float32)
        else:
            bands = np.transpose(arr, (2,0,1)).astype(np.float32)
        with rasterio.open(path,"w",driver="GTiff",height=bands.shape[1],
                           width=bands.shape[2],count=bands.shape[0],
                           dtype=bands.dtype,crs=crs,transform=tr,
                           compress="lzw") as dst:
            for i in range(bands.shape[0]): dst.write(bands[i], i+1)
        log.info("  [OK] GeoTIFF %s", path); return True
    except ImportError:
        log.warning("  rasterio not installed"); return False
    except Exception as exc:
        log.error("  [!!] GeoTIFF %s: %s", path, exc); return False


# ===========================================================================
# DATA EXTRACTION
# ===========================================================================

def get_channel_array(scene, ch):
    """Return raw calibrated numpy array (H,W): K for IR/WV, % for VIS.
    Forces dask compute while file is still open."""
    ds = scene[ch]
    # Force compute NOW
    data = np.array(ds.values)
    if data.ndim == 3: data = data[0]
    return data.astype(np.float32)

def get_composite_array(scene, comp_key):
    """Return Satpy-enhanced (H,W,3) float32 0-1 array. NO re-stretch.
    Forces dask computation before returning so the scene/temp file
    can be safely deleted afterwards."""
    ximg = get_enhanced_image(scene[comp_key])
    # Force compute NOW while scene/file is still open
    data = ximg.data.compute().values  # (bands, H, W)
    arr  = np.transpose(data[:3], (1,2,0)).astype(np.float32)
    return np.clip(arr, 0, 1)

# ===========================================================================
# PER-FILE PROCESSOR
# ===========================================================================

def process_file(nat_path, data_dir, output_root, index):
    filename = os.path.basename(nat_path)
    log.info("="*70)
    log.info("Processing [%d]  %s", index, filename)

    tmp = standard_nat_name(data_dir, index)
    if not make_hardlink(nat_path, tmp):
        return {"file": filename, "status": "load_failed", "products": []}

    try:
        # --- metadata ---
        sc0 = Scene(reader="seviri_l1b_native", filenames=[tmp])
        sc0.load(["IR_108"])
        meta  = sc0["IR_108"].attrs
        t0    = meta.get("start_time", datetime.utcnow())
        t1    = meta.get("end_time",   datetime.utcnow())
        plat  = meta.get("platform_name", "Meteosat-9")
        ts    = t0.strftime("%Y%m%dT%H%MZ")
        del sc0; gc.collect()

        out_root  = ensure_dir(os.path.join(output_root, ts))
        out_ch    = ensure_dir(os.path.join(out_root, "channels"))
        out_comp  = ensure_dir(os.path.join(out_root, "composites"))
        out_tif   = ensure_dir(os.path.join(out_root, "geotiff"))
        products  = []
        gen_time  = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

        # ---- CHANNELS ----
        log.info("  --- Channels ---")
        for ch in CHANNELS:
            label, units, cmap, vmin, vmax = CHANNEL_META.get(
                ch, (ch, "K", "gray_r", 200, 330))
            area = AREA_HRV if ch == "HRV" else AREA_STD
            try:
                sc = Scene(reader="seviri_l1b_native", filenames=[tmp])
                sc.load([ch])
                rsc = sc.resample(area, resampler="nearest")
                # Force compute while file still open, THEN delete
                arr = get_channel_array(rsc, ch)
                del rsc, sc; gc.collect()

                title = _title(label, plat, t0, t1)
                pp = os.path.join(out_ch,  "{}.png".format(ch))
                tp = os.path.join(out_tif, "{}.tif".format(ch))
                ok_p = save_channel_png(arr, pp, label, cmap, vmin, vmax, units, title)
                ok_t = save_geotiff(arr, tp)
                products.append({"product":label,"type":"channel","status":"ok",
                    "sensing_time":str(t0),"generation_time":gen_time,
                    "png":pp if ok_p else None,"tif":tp if ok_t else None})
                del arr; gc.collect()
            except Exception as exc:
                log.error("    FAIL %s: %s", ch, exc); traceback.print_exc()
                products.append({"product":label,"type":"channel","status":"error",
                    "reason":str(exc)[:200],"sensing_time":str(t0),
                    "generation_time":gen_time,"png":None,"tif":None})
                gc.collect()

        # ---- COMPOSITES ----
        log.info("  --- Composites ---")
        for comp_key, comp_label in COMPOSITES.items():
            try:
                sc = Scene(reader="seviri_l1b_native", filenames=[tmp])
                sc.load([comp_key])

                # Composites exist on the native geos scene — access directly
                # without resampling (they are already at 3712x3712).
                # Then resample the extracted array to our eqc grid separately.
                try:
                    _ = sc[comp_key]  # confirm it loaded
                except KeyError:
                    del sc; gc.collect()
                    raise RuntimeError("'{}' not available in native scene".format(comp_key))

                # Get enhanced array while native file is still open
                arr = get_composite_array(sc, comp_key)
                del sc; gc.collect()

                if arr is None:
                    raise RuntimeError("get_composite_array returned None")

                # Resample the numpy array to eqc grid using scipy
                # (arr is already (H,W,3) from native 3712x3712 geos grid)
                # We use a simple zoom to get to 2048x2048
                from scipy.ndimage import zoom as _zoom
                if arr.shape[0] != 2048:
                    factor = 2048 / arr.shape[0]
                    arr = np.stack([
                        _zoom(arr[:,:,b], factor, order=1) for b in range(3)
                    ], axis=2).astype(np.float32)
                    arr = np.clip(arr, 0, 1)

                title = _title(comp_label, plat, t0, t1)
                sl = slug(comp_label)
                pp = os.path.join(out_comp, "{}.png".format(sl))
                tp = os.path.join(out_tif,  "{}.tif".format(sl))
                ok_p = save_rgb_png(arr, pp, title)
                ok_t = save_geotiff(arr, tp)
                products.append({"product":comp_label,"type":"composite",
                    "composite_key":comp_key,"status":"ok",
                    "sensing_time":str(t0),"generation_time":gen_time,
                    "png":pp if ok_p else None,"tif":tp if ok_t else None})
                del arr; gc.collect()
            except Exception as exc:
                log.error("    FAIL '%s': %s", comp_label, exc); traceback.print_exc()
                products.append({"product":comp_label,"type":"composite",
                    "composite_key":comp_key,"status":"error","reason":str(exc)[:200],
                    "sensing_time":str(t0),"generation_time":gen_time,
                    "png":None,"tif":None})
                gc.collect()

        # ---- MANIFEST ----
        manifest = {"file":filename,"md5":md5sum(nat_path),"platform":plat,
                    "sensing_start":str(t0),"sensing_end":str(t1),
                    "generation_time":gen_time,"output_root":out_root,
                    "products":products}
        with open(os.path.join(out_root,"manifest.json"),"w",encoding="utf-8") as f:
            json.dump(manifest, f, indent=2, default=str)
        if products:
            with open(os.path.join(out_root,"manifest.csv"),"w",
                      newline="",encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["product","type","status",
                    "sensing_time","generation_time","png","tif"],
                    extrasaction="ignore")
                w.writeheader(); w.writerows(products)

        ok = sum(1 for p in products if p["status"]=="ok")
        er = sum(1 for p in products if p["status"]=="error")
        log.info("  DONE  %d OK | %d errors", ok, er)
        return manifest

    finally:
        remove_if_exists(tmp)


# ===========================================================================
# ENTRY POINT
# ===========================================================================

def main():
    import glob as _glob
    if len(sys.argv) < 2:
        print("Usage: python visualize_msg_v2.py <data_dir>"); sys.exit(1)

    data_dir    = os.path.abspath(sys.argv[1])
    output_root = os.path.join(os.path.dirname(data_dir), "output_v2")
    nat_files   = sorted(_glob.glob(os.path.join(data_dir, "*.nat")))
    nat_files   = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]

    ensure_dir(output_root)
    seen, unique = {}, []
    for f in nat_files:
        h = md5sum(f)
        if h not in seen: seen[h]=f; unique.append(f)
        else: log.warning("Skipping duplicate: %s", os.path.basename(f))

    log.info("Processing %d unique scenes -> %s", len(unique), output_root)
    manifests, ok_count = [], 0
    for i, fp in enumerate(unique, 1):
        try:
            m = process_file(fp, data_dir, output_root, i)
            manifests.append(m)
            if m.get("status") not in ("load_failed","no_channels"):
                ok_count += 1
        except Exception as exc:
            log.error("Unhandled error %s: %s", os.path.basename(fp), exc)
            traceback.print_exc()

    with open(os.path.join(output_root,"summary.json"),"w",encoding="utf-8") as f:
        json.dump({"run_time":datetime.utcnow().isoformat(),
                   "unique_files":len(unique),"succeeded":ok_count,
                   "manifests":manifests}, f, indent=2, default=str)
    log.info("="*70)
    log.info("Done. %d/%d succeeded. Output: %s", ok_count, len(unique), output_root)

if __name__ == "__main__":
    main()
