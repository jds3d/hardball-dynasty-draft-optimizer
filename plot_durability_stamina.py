#!/usr/bin/env python3
"""
Plot pitcher durability-group score (z) vs Stamina (x) and Durability (y)
using the proportional penalty from algorithm.json.
Run: python plot_durability_stamina.py        # saves PNG and opens window
     python plot_durability_stamina.py --save-only   # only saves PNG
Requires: matplotlib, numpy (pip install matplotlib numpy)
"""
import json
import sys
from pathlib import Path

import numpy as np

try:
    import matplotlib
    if "--save-only" in sys.argv:
        matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from mpl_toolkits.mplot3d import Axes3D
except ImportError:
    print("Install matplotlib and numpy: pip install matplotlib numpy")
    raise

# Paths
SCRIPT_DIR = Path(__file__).resolve().parent
ALGORITHM_FILE = SCRIPT_DIR / "algorithm.json"
try:
    from app_dir import get_algorithm_file
    ALGORITHM_FILE = get_algorithm_file()
except Exception:
    pass


def load_algorithm():
    with open(ALGORITHM_FILE, encoding="utf-8") as f:
        return json.load(f)


def poly(x, a3, a2, a1, a0):
    """Polynomial a3*x^3 + a2*x^2 + a1*x + a0 (vectorized)."""
    return a3 * x**3 + a2 * x**2 + a1 * x + a0


def penalty_proportional(durability, stamina, floor=0.05):
    """Proportional penalty: scale * floor, with scale from algorithm."""
    combined = np.minimum(1, (2 * stamina + durability) / 120)
    dura_scale = np.minimum(1, durability / 10)
    stam_scale = np.minimum(1, stamina / 10)
    scale = combined * dura_scale * stam_scale
    return np.maximum(floor, scale)


def main():
    algo = load_algorithm()
    coeff = algo.get("polynomial", {})
    a3 = coeff.get("a3", 0)
    a2 = coeff.get("a2", 0)
    a1 = coeff.get("a1", 0)
    a0 = coeff.get("a0", 0)

    pitchers = algo.get("pitchers", {}).get("groups", {}).get("durability", {})
    ratings = pitchers.get("ratings", {"Durability": 1.0, "Stamina": 4.0})
    w_dura = float(ratings.get("Durability", 1))
    w_stam = float(ratings.get("Stamina", 4))
    pen_cfg = pitchers.get("penalty", {})
    floor = float(pen_cfg.get("floor", 0.05))

    # Grid: stamina = x, durability = y
    n = 101
    stamina = np.linspace(0, 100, n)
    durability = np.linspace(0, 100, n)
    S, D = np.meshgrid(stamina, durability)

    base = w_dura * poly(D, a3, a2, a1, a0) + w_stam * poly(S, a3, a2, a1, a0)
    pen = penalty_proportional(D, S, floor=floor)
    z = base * pen

    fig = plt.figure(figsize=(12, 5))

    # 3D surface
    ax1 = fig.add_subplot(121, projection="3d")
    surf = ax1.plot_surface(S, D, z, cmap="viridis", edgecolor="none", antialiased=True)
    ax1.set_xlabel("Stamina")
    ax1.set_ylabel("Durability")
    ax1.set_zlabel("Durability group score")
    ax1.set_title("Score (z) vs Stamina (x) & Durability (y)")
    fig.colorbar(surf, ax=ax1, shrink=0.6)

    # 2D heatmap
    ax2 = fig.add_subplot(122)
    im = ax2.pcolormesh(stamina, durability, z, shading="auto", cmap="viridis")
    ax2.set_xlabel("Stamina")
    ax2.set_ylabel("Durability")
    ax2.set_title("Durability group score (heatmap)")
    fig.colorbar(im, ax=ax2, label="Score")
    ax2.set_aspect("equal")

    plt.tight_layout()
    out = SCRIPT_DIR / "durability_stamina_plot.png"
    plt.savefig(out, dpi=150, bbox_inches="tight")
    print(f"Saved {out}")
    if "--save-only" not in sys.argv:
        plt.show()


if __name__ == "__main__":
    main()
