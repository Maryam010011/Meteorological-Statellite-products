"""
test_colormaps.py
=================
Generates IR_108 for one scene using 10 different color palettes
as a sample/preview before full batch.
"""
import warnings; warnings.filterwarnings("ignore")
import os, gc, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import cartopy.crs as ccrs
import cartopy.feature as cfeature
from pyresample.geometry import AreaDefinition
from satpy import Scene
from visualize_msg_v2 import (
    standard_nat_name, make_hardlink, remove_if_exists, ensure_dir,
    _apply_disk_mask, _overlays
)

OUT = r"e:\NASTAP\Assignment 4\sample_colormaps"
ensure_dir(OUT)
DATA_DIR = r"e:\NASTAP\Assignment 4\data"
NAT = r"e:\NASTAP\Assignment 4\data\msg15 (1).nat"

# 10 sample color palettes
PALETTES = [
    ("01_Gray",          "gray_r",      "Standard IR Gray (cold=bright)"),
    ("02_Rainbow",       "jet",         "Rainbow / Jet"),
    ("03_Thermal",       "inferno",     "Thermal Inferno"),
    ("04_Spectral",      "Spectral",    "Spectral"),
    ("05_RdYlBu",        "RdYlBu",      "Red-Yellow-Blue"),
    ("06_Plasma",        "plasma",      "Plasma"),
    ("07_Viridis",       "viridis",     "Viridis"),
    ("08_Coolwarm",      "coolwarm",    "Cool-Warm"),
    ("09_YlOrRd",        "YlOrRd_r",   "Yellow-Orange-Red (inverted)"),
    ("10_Cyan_Purple",   "PRGn",        "Purple-Green"),
]

eqc_area = AreaDefinition(
    "eqc","eqc","eqc",
    "+proj=eqc +lon_0=45.5 +datum=WGS84 +units=m",
    2048, 2048, (-8000000,-8000000,8000000,8000000)
)
_CRS = ccrs.PlateCarree(central_longitude=45.5)
lon_min = 45.5 - 72; lon_max = 45.5 + 72

tmp = standard_nat_name(DATA_DIR, 95)
make_hardlink(NAT, tmp)

try:
    sc = Scene(reader="seviri_l1b_native", filenames=[tmp])
    sc.load(["IR_108"])
    t0   = sc["IR_108"].attrs.get("start_time")
    t1   = sc["IR_108"].attrs.get("end_time")
    plat = sc["IR_108"].attrs.get("platform_name","Meteosat-9")
    rsc  = sc.resample(eqc_area, resampler="nearest")
    arr  = rsc["IR_108"].values.astype("float32")
    del sc, rsc; gc.collect()

    print(f"Array: {arr.shape}, min={arr[~np.isnan(arr)].min():.1f}K, max={arr[~np.isnan(arr)].max():.1f}K")

    ts = t0.strftime("%Y-%m-%d %H:%M UTC") if t0 else "2026-06-22 00:00 UTC"
    te = t1.strftime("%H:%M UTC")           if t1 else "00:15 UTC"

    for slug, cmap, label in PALETTES:
        fig, ax = plt.subplots(1,1,figsize=(10,10),
            subplot_kw={"projection": _CRS}, facecolor="white")
        ax.set_facecolor("black")
        ax.set_extent([lon_min, lon_max, -72, 72], crs=ccrs.PlateCarree())

        img = ax.imshow(arr, origin="upper", cmap=cmap,
                        vmin=200, vmax=330,
                        extent=[lon_min, lon_max, -72, 72],
                        transform=ccrs.PlateCarree(),
                        interpolation="nearest", zorder=3)

        ax.add_feature(cfeature.COASTLINE.with_scale("50m"),
                       linewidth=0.6, edgecolor="yellow", zorder=5)
        ax.add_feature(cfeature.BORDERS.with_scale("50m"),
                       linewidth=0.3, edgecolor="yellow", zorder=5)
        try:
            gl = ax.gridlines(crs=ccrs.PlateCarree(), draw_labels=True,
                              linewidth=0.3, color="white", alpha=0.5,
                              x_inline=False, y_inline=False)
            gl.top_labels=False; gl.right_labels=False
            gl.xlabel_style={"size":7,"color":"black"}
            gl.ylabel_style={"size":7,"color":"black"}
            gl.xlocator=mticker.FixedLocator(range(-180,181,30))
            gl.ylocator=mticker.FixedLocator(range(-90,91,30))
        except: pass

        cb = fig.colorbar(img, ax=ax, orientation="horizontal",
                          pad=0.04, fraction=0.03, shrink=0.7)
        cb.set_label(f"IR 10.8 µm [K]  —  {label}", color="black", fontsize=7)
        cb.ax.xaxis.set_tick_params(color="black",labelcolor="black",labelsize=6)

        ax.set_title(
            f"IR 10.8 µm  |  {label}  |  {plat} IODC 45.5°E\n{ts} - {te}",
            color="black", fontsize=8, pad=8, loc="left")

        out_path = os.path.join(OUT, f"IR_108_{slug}.png")
        fig.savefig(out_path, dpi=150, facecolor="white")
        plt.close(fig)
        print(f"  Saved: IR_108_{slug}.png  ({os.path.getsize(out_path)//1024} KB)")
        gc.collect()

    print(f"\nDone. Open {OUT} to review the 10 color samples.")

finally:
    remove_if_exists(tmp)
