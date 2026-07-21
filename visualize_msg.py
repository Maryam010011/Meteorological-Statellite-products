"""
visualize_msg.py
================
Visualize EUMETSAT Meteosat Second Generation (MSG) SEVIRI Level 1.5 native (.nat)
data using Satpy, Cartopy, and matplotlib.

Usage:
    python visualize_msg.py <data_dir>

Example:
    python visualize_msg.py data/

For every .nat file found the script generates:
  A) Grayscale / colorized images for all 12 SEVIRI channels.
  B) Standard EUMETSAT RGB composites.
  C) Each image is reprojected and overlaid with Cartopy coastlines / borders.
  D) Output is saved as PNG and GeoTIFF in an organised folder tree:
       output/<timestamp>/channels/<name>.png
       output/<timestamp>/composites/<name>.png
       output/<timestamp>/geotiff/<name>.tif
  E) A manifest JSON + CSV is written per timestamp.

Author: Antigravity (generated 2026-07-15)
"""

# ---------------------------------------------------------------------------
# Standard library
# ---------------------------------------------------------------------------
import csv
import gc
import hashlib
import json
import logging
import os
import shutil
import sys
import tempfile
import traceback
from datetime import datetime, timezone
from pathlib import Path

# ---------------------------------------------------------------------------
# Third-party
# ---------------------------------------------------------------------------
import matplotlib
matplotlib.use("Agg")                          # non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import cartopy.crs as ccrs
import cartopy.feature as cfeature
from cartopy.mpl.gridliner import LONGITUDE_FORMATTER, LATITUDE_FORMATTER  # noqa: F401 (kept for compat)

# ---------------------------------------------------------------------------
# Satpy / pyresample
# ---------------------------------------------------------------------------
from satpy import Scene
from satpy.writers import get_enhanced_image
from pyresample.geometry import AreaDefinition

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%S",
)
log = logging.getLogger("visualize_msg")

# ===========================================================================
# CONFIGURATION
# ===========================================================================

# Channels to process individually
CHANNELS = [
    "VIS006", "VIS008", "IR_016", "IR_039",
    "WV_062", "WV_073", "IR_087", "IR_097",
    "IR_108", "IR_120", "IR_134", "HRV",
]

# Colourmap configuration per channel (cmap, vmin, vmax, invert)
CHANNEL_CMAPS = {
    # Visible / near-IR: bright = white
    "VIS006": ("gray",           0,   100, False),
    "VIS008": ("gray",           0,   100, False),
    "IR_016": ("gray",           0,   100, False),
    "HRV":    ("gray",           0,   100, False),
    # Solar / mixed channels
    "IR_039": ("gray_r",       200,   330, True),
    # Water vapour – cold = bright
    "WV_062": ("gray_r",       200,   260, True),
    "WV_073": ("gray_r",       200,   280, True),
    # Thermal IR – cold cloud tops = bright
    "IR_087": ("gray_r",       200,   330, True),
    "IR_097": ("gray_r",       200,   320, True),
    "IR_108": ("gray_r",       200,   330, True),
    "IR_120": ("gray_r",       200,   330, True),
    "IR_134": ("gray_r",       200,   330, True),
}

# Composites to generate (Satpy name → display name)
COMPOSITES = {
    "natural_color":     "Natural Color",
    "airmass":           "Airmass RGB",
    "dust":              "Dust RGB",
    "ash":               "Ash RGB",
    "convection":        "Convection RGB",
    "fog":               "Fog RGB",
    "night_microphysics":"Night Microphysics RGB",
    "day_severe_storms": "Severe Storms RGB",
    "hrv_clouds":        "HRV Clouds (HRV Clean)",
    "hrv_fog":           "HRV Fog",
    "colorized_ir_clouds":"IR10.8 Colorized Clouds",
}

# Resampling target – Plate Carrée, centred on Meteosat-9 IODC sub-satellite point
# We build a custom AreaDefinition so we don't need an external area config file.
TARGET_PROJ = "+proj=eqc +lon_0=45.5 +lat_0=0 +datum=WGS84 +units=m"
# 2048 x 2048 pixels covering ± 80° in lat/lon around 45.5°E
TARGET_AREA = AreaDefinition(
    area_id="msg_platecarree_2k",
    description="Plate Carrée 2048×2048 centred 45.5E for MSG-9 IODC",
    proj_id="eqc",
    projection=TARGET_PROJ,
    width=2048,
    height=2048,
    area_extent=(-8_900_000, -8_900_000, 8_900_000, 8_900_000),
)

# Cartopy projection matching the target area
CARTOPY_CRS = ccrs.PlateCarree(central_longitude=45.5)

# ===========================================================================
# HELPERS
# ===========================================================================

def md5sum(path: str) -> str:
    """Return MD5 hex digest of *path*."""
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def standard_nat_name(index: int) -> str:
    """
    Return a valid EUMETSAT-style filename that Satpy's seviri_l1b_native
    reader will recognise.  We embed a fixed placeholder timestamp; the
    real sensing time is read from the file header by Satpy internally.
    """
    return f"MSG3-SEVI-MSG15-0100-NA-20000101000000.000000000Z-TMPIDX{index:04d}.nat"


def make_hardlink(src: str, dst: str) -> bool:
    """Create *dst* as a hard-link to *src*.  Return True on success."""
    try:
        if os.path.exists(dst):
            os.remove(dst)
        os.link(src, dst)
        return True
    except OSError as exc:
        log.warning("Cannot hard-link %s → %s (%s); trying copy …", src, dst, exc)
        try:
            shutil.copy2(src, dst)
            return True
        except Exception as exc2:
            log.error("Copy also failed: %s", exc2)
            return False


def remove_if_exists(path: str) -> None:
    """Silently remove *path* if it exists."""
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def slug(name: str) -> str:
    """Convert a product name to a filesystem-safe slug."""
    return name.lower().replace(" ", "_").replace("/", "_").replace(".", "p")


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


# ===========================================================================
# IMAGE RENDERING
# ===========================================================================

def _apply_cartopy_overlays(ax, extent_lonlat=(-180, 180, -85, 85)):
    """Add coastlines, borders, gridlines to a Cartopy GeoAxes."""
    ax.add_feature(cfeature.COASTLINE.with_scale("50m"), linewidth=0.6,
                   edgecolor="yellow", zorder=5)
    ax.add_feature(cfeature.BORDERS.with_scale("50m"), linewidth=0.4,
                   edgecolor="white", linestyle="--", zorder=5)
    ax.add_feature(cfeature.OCEAN.with_scale("50m"),
                   facecolor="#0a1929", zorder=0)
    ax.add_feature(cfeature.LAND.with_scale("50m"),
                   facecolor="#1a1a1a", zorder=1)
    # draw_labels=False avoids the Cartopy/Shapely LinearRing crash (Shapely 2.x
    # raises GEOSException when the map-boundary path has too few vertices).
    try:
        gl = ax.gridlines(crs=ccrs.PlateCarree(), draw_labels=False,
                          linewidth=0.3, color="white", alpha=0.4,
                          x_inline=False, y_inline=False)
        gl.xlocator = mticker.FixedLocator(range(-180, 181, 30))
        gl.ylocator = mticker.FixedLocator(range(-90, 91, 30))
    except Exception:
        pass   # gridlines are decorative; skip silently if they fail
    return ax


def _safe_savefig(fig, out_path: str) -> None:
    """
    Save figure, retrying without gridlines if Shapely/Cartopy gridliner
    raises a GEOSException (known Cartopy 0.25 + Shapely 2.x incompatibility).
    """
    try:
        plt.savefig(out_path, dpi=150, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
    except Exception as exc:
        if "LinearRing" in str(exc) or "GEOSException" in str(exc) or "linestring" in str(exc).lower():
            log.warning("  Gridliner error on first save; retrying without gridlines: %s", exc)
            # Remove all gridliner artists and retry
            for ax in fig.axes:
                try:
                    if hasattr(ax, "_gridliners"):
                        for gl in list(ax._gridliners):
                            try:
                                gl.remove()
                            except Exception:
                                pass
                        ax._gridliners.clear()
                except Exception:
                    pass
            plt.savefig(out_path, dpi=150, bbox_inches="tight",
                        facecolor=fig.get_facecolor())
        else:
            raise



def save_channel_png(
    data_np: np.ndarray,
    out_path: str,
    title: str,
    cmap: str,
    vmin: float,
    vmax: float,
    sensing_time: str,
    units: str = "K",
) -> bool:
    """
    Render a single-band array with Cartopy overlays and save as PNG.
    *data_np* must already be on the TARGET_AREA grid (2048 × 2048).
    """
    try:
        fig = plt.figure(figsize=(10, 10), facecolor="#0d0d0d")
        ax = plt.axes(projection=CARTOPY_CRS)
        ax.set_extent([-79.5 + 45.5, 79.5 + 45.5, -80, 80], crs=ccrs.PlateCarree())
        ax.set_facecolor("#0a1929")

        # Build extent in target projection coordinates
        extent_m = (-8_900_000, 8_900_000, -8_900_000, 8_900_000)   # (left, right, bottom, top)
        img = ax.imshow(
            data_np,
            origin="upper",
            extent=extent_m,
            transform=CARTOPY_CRS,
            cmap=cmap,
            vmin=vmin,
            vmax=vmax,
            interpolation="nearest",
            zorder=3,
        )
        _apply_cartopy_overlays(ax)

        cbar = plt.colorbar(img, ax=ax, orientation="horizontal",
                            pad=0.03, fraction=0.046, shrink=0.8)
        cbar.set_label(f"{title} [{units}]", color="white", fontsize=8)
        cbar.ax.xaxis.set_tick_params(color="white", labelcolor="white", labelsize=7)

        ax.set_title(f"{title}\n{sensing_time} UTC  |  Meteosat-9 IODC",
                     color="white", fontsize=10, pad=8)
        fig.patch.set_facecolor("#0d0d0d")

        plt.tight_layout()
        _safe_savefig(fig, out_path)
        plt.close(fig)
        log.info("  [OK] PNG saved: %s", out_path)
        return True
    except Exception as exc:
        log.error("  [!!] Failed PNG for %s: %s", title, exc)
        traceback.print_exc()
        plt.close("all")
        return False


def save_rgb_png(
    rgb_np: np.ndarray,
    out_path: str,
    title: str,
    sensing_time: str,
) -> bool:
    """
    Render an RGB array (H × W × 3, float 0–1) with Cartopy overlays.
    """
    try:
        fig = plt.figure(figsize=(10, 10), facecolor="#0d0d0d")
        ax = plt.axes(projection=CARTOPY_CRS)
        ax.set_extent([-79.5 + 45.5, 79.5 + 45.5, -80, 80], crs=ccrs.PlateCarree())
        ax.set_facecolor("#0a1929")

        extent_m = (-8_900_000, 8_900_000, -8_900_000, 8_900_000)
        ax.imshow(
            np.clip(rgb_np, 0, 1),
            origin="upper",
            extent=extent_m,
            transform=CARTOPY_CRS,
            interpolation="nearest",
            zorder=3,
        )
        _apply_cartopy_overlays(ax)

        ax.set_title(f"{title}\n{sensing_time} UTC  |  Meteosat-9 IODC",
                     color="white", fontsize=10, pad=8)
        fig.patch.set_facecolor("#0d0d0d")

        plt.tight_layout()
        _safe_savefig(fig, out_path)
        plt.close(fig)
        log.info("  [OK] PNG saved: %s", out_path)
        return True
    except Exception as exc:
        log.error("  [!!] Failed RGB PNG for %s: %s", title, exc)
        traceback.print_exc()
        plt.close("all")
        return False


def save_geotiff(data_np: np.ndarray, out_path: str, bands: int = 1) -> bool:
    """
    Save array as a GeoTIFF using rasterio (preferred) or a minimal stub.
    data_np shape: (H, W) for single band, (H, W, 3) for RGB.
    """
    try:
        import rasterio
        from rasterio.transform import from_bounds
        from rasterio.crs import CRS as RioCRS

        extent = (-8_900_000, -8_900_000, 8_900_000, 8_900_000)  # (w, s, e, n)
        transform = from_bounds(extent[0], extent[1], extent[2], extent[3],
                                data_np.shape[-1] if data_np.ndim == 3 else data_np.shape[1],
                                data_np.shape[-2] if data_np.ndim == 3 else data_np.shape[0])
        crs_wkt = RioCRS.from_proj4(TARGET_PROJ)

        if data_np.ndim == 2:
            arr = data_np[np.newaxis, ...].astype(np.float32)
            band_count = 1
        else:
            arr = np.transpose(data_np, (2, 0, 1)).astype(np.float32)
            band_count = arr.shape[0]

        with rasterio.open(
            out_path, "w",
            driver="GTiff",
            height=arr.shape[1],
            width=arr.shape[2],
            count=band_count,
            dtype=arr.dtype,
            crs=crs_wkt,
            transform=transform,
            compress="lzw",
        ) as dst:
            for b in range(band_count):
                dst.write(arr[b], b + 1)

        log.info("  [OK] GeoTIFF saved: %s", out_path)
        return True

    except ImportError:
        # rasterio not available – save a minimal world-file PNG-as-TIF
        log.warning("  [!] rasterio not installed; skipping GeoTIFF for %s", out_path)
        return False
    except Exception as exc:
        log.error("  [!!] GeoTIFF failed for %s: %s", out_path, exc)
        return False


# ===========================================================================
# SCENE LOADING & RESAMPLING
# ===========================================================================

def load_scene(nat_path: str, data_dir: str, index: int):
    """
    Return an open Satpy Scene for *nat_path*.
    Creates a temporary hard-link with a standard EUMETSAT filename so that
    the seviri_l1b_native reader can recognise the file.
    Returns (scene, temp_path) or (None, None) on failure.
    """
    temp_name = standard_nat_name(index)
    temp_path = os.path.join(data_dir, temp_name)

    if not make_hardlink(nat_path, temp_path):
        return None, None

    try:
        scene = Scene(reader="seviri_l1b_native", filenames=[temp_path])
        return scene, temp_path
    except Exception as exc:
        log.error("Cannot create Scene for %s: %s", nat_path, exc)
        remove_if_exists(temp_path)
        return None, None


def resample_dataset(scene, ds_name: str):
    """
    Return the resampled numpy array for *ds_name*, or None.
    For RGB composites the returned array is (H, W, 3) float32 0-1.
    For single channels the returned array is (H, W) float32 in physical units.

    HRV-based composites (hrv_clouds, hrv_fog) span two resolution grids and
    must be resampled without a dataset filter so Satpy can reconcile both
    grids internally.
    """
    try:
        # First attempt: targeted resample
        resampled = scene.resample(TARGET_AREA, datasets=[ds_name],
                                   resampler="nearest")
        # Check if ds_name is actually present in the resampled scene
        available = resampled.available_dataset_names()
        if ds_name not in available:
            # Fallback: resample the whole scene (handles HRV composites)
            log.warning("    '%s' not in targeted resample; retrying full-scene resample ...",
                        ds_name)
            resampled = scene.resample(TARGET_AREA, resampler="nearest")
            available = resampled.available_dataset_names()
            if ds_name not in available:
                log.warning("    '%s' still not available after full-scene resample.", ds_name)
                return None

        data = resampled[ds_name]
        arr = np.array(data.values, dtype=np.float32)

        # Handle RGB (3, H, W) -> (H, W, 3) and normalise to 0-1
        if arr.ndim == 3 and arr.shape[0] in (3, 4):
            arr = np.transpose(arr, (1, 2, 0))
            arr = arr[..., :3]                 # drop alpha if present
            arr_min, arr_max = np.nanpercentile(arr, 2), np.nanpercentile(arr, 98)
            if arr_max > arr_min:
                arr = (arr - arr_min) / (arr_max - arr_min)
            arr = np.clip(arr, 0, 1)
        return arr
    except Exception as exc:
        log.warning("    resample failed for '%s': %s", ds_name, exc)
        return None


# ===========================================================================
# MAIN PROCESSING LOOP
# ===========================================================================

def process_file(nat_path: str, data_dir: str, output_root: str, index: int) -> dict:
    """
    Process one .nat file → all channels + composites.
    Returns a manifest dict.
    """
    filename = os.path.basename(nat_path)
    log.info("=" * 70)
    log.info("Processing [%d] %s", index, filename)

    # ---- 1. Load scene ---------------------------------------------------
    scene, temp_path = load_scene(nat_path, data_dir, index)
    if scene is None:
        return {"file": filename, "status": "load_failed", "products": []}

    try:
        # ---- 2. Inspect metadata -----------------------------------------
        available_channels = [c for c in CHANNELS if c in scene.available_dataset_names()]
        log.info("  Available channels: %s", available_channels)

        # Load a reference channel for metadata
        if not available_channels:
            log.error("  No known channels available – skipping.")
            return {"file": filename, "status": "no_channels", "products": []}

        ref_ch = "IR_108" if "IR_108" in available_channels else available_channels[0]
        scene.load([ref_ch])
        meta = scene[ref_ch].attrs
        start_time = meta.get("start_time", datetime.utcnow())
        end_time   = meta.get("end_time",   datetime.utcnow())
        platform   = meta.get("platform_name", "Meteosat-9")
        ts_str     = start_time.strftime("%Y%m%dT%H%MZ")
        sensing_str = start_time.strftime("%Y-%m-%d %H:%M")

        log.info("  Platform: %s  |  Start: %s  |  End: %s",
                 platform, start_time, end_time)

        # ---- 3. Output directories ----------------------------------------
        out_root     = ensure_dir(os.path.join(output_root, ts_str))
        out_channels = ensure_dir(os.path.join(out_root, "channels"))
        out_composites = ensure_dir(os.path.join(out_root, "composites"))
        out_geotiff  = ensure_dir(os.path.join(out_root, "geotiff"))

        products = []          # manifest entries
        gen_time = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

        # ==================================================================
        # A) INDIVIDUAL CHANNELS
        # ==================================================================
        log.info("  --- Individual channels ---")
        for ch in CHANNELS:
            ch_label = ch   # e.g. "IR_108"
            try:
                if ch not in scene.available_dataset_names():
                    log.warning("    Channel %s not available – skipping.", ch)
                    products.append({
                        "product": ch_label, "type": "channel",
                        "status": "skipped", "reason": "not available",
                        "channels_used": [ch],
                        "sensing_time": sensing_str,
                        "generation_time": gen_time,
                        "png": None, "tif": None,
                    })
                    continue

                # Reload the scene to avoid memory buildup
                scene_ch = Scene(reader="seviri_l1b_native",
                                 filenames=[temp_path])
                scene_ch.load([ch])
                arr = resample_dataset(scene_ch, ch)
                del scene_ch
                gc.collect()

                if arr is None:
                    raise RuntimeError("resample returned None")

                cmap, vmin, vmax, _ = CHANNEL_CMAPS.get(ch, ("gray_r", None, None, True))

                # auto-stretch if vmin/vmax not set
                if vmin is None:
                    vmin = float(np.nanpercentile(arr, 2))
                if vmax is None:
                    vmax = float(np.nanpercentile(arr, 98))

                units = "K" if ch.startswith(("IR", "WV")) else "%"

                png_path = os.path.join(out_channels, f"{ch}.png")
                tif_path = os.path.join(out_geotiff, f"{ch}.tif")

                ok_png = save_channel_png(arr, png_path, ch_label,
                                          cmap, vmin, vmax, sensing_str, units)
                ok_tif = save_geotiff(arr, tif_path)

                products.append({
                    "product": ch_label, "type": "channel", "status": "ok",
                    "channels_used": [ch],
                    "sensing_time": sensing_str,
                    "generation_time": gen_time,
                    "png": png_path if ok_png else None,
                    "tif": tif_path if ok_tif else None,
                })
                del arr
                gc.collect()

            except Exception as exc:
                log.error("    ✗  Channel %s failed: %s", ch, exc)
                traceback.print_exc()
                products.append({
                    "product": ch_label, "type": "channel",
                    "status": "error", "reason": str(exc),
                    "channels_used": [ch],
                    "sensing_time": sensing_str,
                    "generation_time": gen_time,
                    "png": None, "tif": None,
                })
                gc.collect()
                continue

        # ==================================================================
        # B) RGB COMPOSITES
        # ==================================================================
        log.info("  --- RGB Composites ---")
        for comp_key, comp_label in COMPOSITES.items():
            try:
                scene_comp = Scene(reader="seviri_l1b_native",
                                   filenames=[temp_path])
                scene_comp.load([comp_key])

                arr = resample_dataset(scene_comp, comp_key)
                del scene_comp
                gc.collect()

                if arr is None:
                    raise RuntimeError("resample returned None")

                # Determine output path and channels used (approximate)
                comp_slug = slug(comp_label)
                png_path  = os.path.join(out_composites, f"{comp_slug}.png")
                tif_path  = os.path.join(out_geotiff,    f"{comp_slug}.tif")

                if arr.ndim == 3:                  # RGB
                    ok_png = save_rgb_png(arr, png_path, comp_label, sensing_str)
                    ok_tif = save_geotiff(arr, tif_path)
                else:                              # single-band (e.g. colorized_ir)
                    ok_png = save_channel_png(arr, png_path, comp_label,
                                              "RdYlBu_r", None, None, sensing_str, "K")
                    ok_tif = save_geotiff(arr, tif_path)

                products.append({
                    "product": comp_label, "type": "composite",
                    "composite_key": comp_key,
                    "status": "ok",
                    "sensing_time": sensing_str,
                    "generation_time": gen_time,
                    "png": png_path if ok_png else None,
                    "tif": tif_path if ok_tif else None,
                })
                del arr
                gc.collect()

            except Exception as exc:
                log.error("    ✗  Composite '%s' failed: %s", comp_label, exc)
                traceback.print_exc()
                products.append({
                    "product": comp_label, "type": "composite",
                    "composite_key": comp_key,
                    "status": "error", "reason": str(exc),
                    "sensing_time": sensing_str,
                    "generation_time": gen_time,
                    "png": None, "tif": None,
                })
                gc.collect()
                continue

        # ==================================================================
        # C) MANIFEST (JSON + CSV)
        # ==================================================================
        manifest = {
            "file":           filename,
            "md5":            md5sum(nat_path),
            "platform":       platform,
            "sensing_start":  str(start_time),
            "sensing_end":    str(end_time),
            "generation_time": gen_time,
            "output_root":    out_root,
            "products":       products,
        }

        json_path = os.path.join(out_root, "manifest.json")
        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(manifest, fh, indent=2, default=str)
        log.info("  Manifest JSON: %s", json_path)

        csv_path = os.path.join(out_root, "manifest.csv")
        if products:
            fieldnames = ["product", "type", "status", "sensing_time",
                          "generation_time", "png", "tif"]
            with open(csv_path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(products)
        log.info("  Manifest CSV:  %s", csv_path)

        ok_count  = sum(1 for p in products if p["status"] == "ok")
        err_count = sum(1 for p in products if p["status"] == "error")
        skip_count = sum(1 for p in products if p["status"] == "skipped")
        log.info("  Done: %d OK | %d errors | %d skipped", ok_count, err_count, skip_count)

        return manifest

    finally:
        remove_if_exists(temp_path)


# ===========================================================================
# ENTRY POINT
# ===========================================================================

def main():
    import glob as _glob

    if len(sys.argv) < 2:
        print(f"Usage: python {sys.argv[0]} <data_dir>")
        sys.exit(1)

    data_dir    = os.path.abspath(sys.argv[1])
    output_root = os.path.join(os.path.dirname(data_dir), "output")

    nat_files = sorted(_glob.glob(os.path.join(data_dir, "*.nat")))
    # Exclude any leftover temporary files from previous runs
    nat_files = [f for f in nat_files if not os.path.basename(f).startswith("MSG")]

    if not nat_files:
        log.error("No .nat files found in '%s'", data_dir)
        sys.exit(1)

    log.info("Found %d .nat files in %s", len(nat_files), data_dir)
    log.info("Output root: %s", output_root)
    ensure_dir(output_root)

    # Detect and report duplicates (by MD5)
    md5_map: dict[str, list[str]] = {}
    for f in nat_files:
        h = md5sum(f)
        md5_map.setdefault(h, []).append(os.path.basename(f))

    seen_md5: set[str] = set()
    unique_files = []
    for f in nat_files:
        h = md5sum(f)
        if h in seen_md5:
            log.warning("Skipping duplicate file %s (same MD5 as %s)",
                        os.path.basename(f), md5_map[h][0])
            continue
        seen_md5.add(h)
        unique_files.append(f)

    log.info("Processing %d unique files (%d duplicates skipped)",
             len(unique_files), len(nat_files) - len(unique_files))

    all_manifests = []
    success_count = 0

    for idx, nat_path in enumerate(unique_files):
        try:
            manifest = process_file(nat_path, data_dir, output_root, idx)
            all_manifests.append(manifest)
            if manifest.get("status") not in ("load_failed", "no_channels"):
                success_count += 1
        except Exception as exc:
            log.error("Unhandled error for %s: %s", nat_path, exc)
            traceback.print_exc()

    # ---- Write global summary manifest -----------------------------------
    summary_path = os.path.join(output_root, "summary.json")
    with open(summary_path, "w", encoding="utf-8") as fh:
        json.dump({
            "run_time":       datetime.utcnow().isoformat(),
            "data_dir":       data_dir,
            "total_files":    len(nat_files),
            "unique_files":   len(unique_files),
            "succeeded":      success_count,
            "manifests":      all_manifests,
        }, fh, indent=2, default=str)

    log.info("=" * 70)
    log.info("Run complete.")
    log.info("  Total files    : %d", len(nat_files))
    log.info("  Unique scenes  : %d", len(unique_files))
    log.info("  Succeeded      : %d", success_count)
    log.info("  Summary JSON   : %s", summary_path)
    log.info("  Output root    : %s", output_root)


if __name__ == "__main__":
    main()
