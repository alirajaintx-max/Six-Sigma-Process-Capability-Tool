# ============================================================
#  SIX SIGMA PROCESS CAPABILITY TOOL
#  Author: [Your Name]
#  Description: Calculates Cp, Cpk, runs control charts
#               (X-bar & R-chart), and exports a formatted
#               quality report to Excel.
# ============================================================

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False
    print("Install openpyxl: pip install openpyxl")


# ============================================================
# CONFIGURATION
# ============================================================

PROCESS_NAME  = "Shaft Diameter"
UNIT          = "mm"
TARGET        = 50.00          # Target (nominal) value
LSL           = 49.70          # Lower Specification Limit
USL           = 50.30          # Upper Specification Limit
SUBGROUP_SIZE = 5              # Samples per subgroup
N_SUBGROUPS   = 30             # Number of subgroups collected
RANDOM_SEED   = 7


# ============================================================
# STEP 1: Generate process data
# (Replace with your real measurements CSV)
# ============================================================

def generate_process_data():
    np.random.seed(RANDOM_SEED)
    mean  = TARGET + 0.02       # Slight mean shift to make it realistic
    std   = 0.08                # Process standard deviation

    data  = np.random.normal(mean, std, (N_SUBGROUPS, SUBGROUP_SIZE))

    # Inject a few out-of-control points for realism
    data[14, 2] += 0.28
    data[22, 0] -= 0.26

    df = pd.DataFrame(data, columns=[f"Sample {i+1}" for i in range(SUBGROUP_SIZE)])
    df.index = [f"SG {i+1}" for i in range(N_SUBGROUPS)]
    print(f"✅ Generated {N_SUBGROUPS} subgroups × {SUBGROUP_SIZE} samples for '{PROCESS_NAME}'")
    return df


# ============================================================
# STEP 2: Process capability indices (Cp, Cpk, Pp, Ppk)
# ============================================================

def calculate_capability(data):
    all_values = data.values.flatten()
    mean       = all_values.mean()
    std        = all_values.std(ddof=1)

    cp   = (USL - LSL) / (6 * std)
    cpu  = (USL - mean) / (3 * std)
    cpl  = (mean - LSL) / (3 * std)
    cpk  = min(cpu, cpl)
    sigma_level = cpk * 3

    # Estimated defect rate (parts per million)
    ppm_upper = stats.norm.sf(USL, loc=mean, scale=std) * 1_000_000
    ppm_lower = stats.norm.cdf(LSL, loc=mean, scale=std) * 1_000_000
    ppm_total = ppm_upper + ppm_lower

    results = {
        "mean": mean, "std": std, "cp": cp, "cpu": cpu,
        "cpl": cpl, "cpk": cpk, "sigma_level": sigma_level,
        "ppm_total": ppm_total,
    }

    print(f"\n📊 Process Capability Results:")
    print(f"   Mean:        {mean:.4f} {UNIT}   (Target: {TARGET} {UNIT})")
    print(f"   Std Dev:     {std:.4f} {UNIT}")
    print(f"   Cp:          {cp:.3f}   {'✅ Capable' if cp >= 1.33 else '⚠️  Marginal' if cp >= 1.0 else '❌ Not Capable'}")
    print(f"   Cpk:         {cpk:.3f}   {'✅ Centred' if cpk >= 1.33 else '⚠️  Off-Centre' if cpk >= 1.0 else '❌ Not Centred'}")
    print(f"   Sigma Level: {sigma_level:.2f}σ")
    print(f"   Est. Defects:{ppm_total:.0f} PPM")

    return results


# ============================================================
# STEP 3: Control chart constants & limits
# ============================================================

# Standard control chart constants for subgroup size n
CONTROL_CONSTANTS = {
    2:  {"A2": 1.880, "D3": 0,     "D4": 3.267, "d2": 1.128},
    3:  {"A2": 1.023, "D3": 0,     "D4": 2.575, "d2": 1.693},
    4:  {"A2": 0.729, "D3": 0,     "D4": 2.282, "d2": 2.059},
    5:  {"A2": 0.577, "D3": 0,     "D4": 2.115, "d2": 2.326},
    6:  {"A2": 0.483, "D3": 0,     "D4": 2.004, "d2": 2.534},
    7:  {"A2": 0.419, "D3": 0.076, "D4": 1.924, "d2": 2.704},
    8:  {"A2": 0.373, "D3": 0.136, "D4": 1.864, "d2": 2.847},
    9:  {"A2": 0.337, "D3": 0.184, "D4": 1.816, "d2": 2.970},
    10: {"A2": 0.308, "D3": 0.223, "D4": 1.777, "d2": 3.078},
}

def calculate_control_limits(data):
    n   = SUBGROUP_SIZE
    cc  = CONTROL_CONSTANTS.get(n, CONTROL_CONSTANTS[5])

    xbars = data.mean(axis=1)
    ranges = data.max(axis=1) - data.min(axis=1)

    xbar_bar = xbars.mean()
    r_bar    = ranges.mean()

    xbar_ucl = xbar_bar + cc["A2"] * r_bar
    xbar_lcl = xbar_bar - cc["A2"] * r_bar
    r_ucl    = cc["D4"] * r_bar
    r_lcl    = cc["D3"] * r_bar

    return {
        "xbars": xbars, "ranges": ranges,
        "xbar_bar": xbar_bar, "r_bar": r_bar,
        "xbar_ucl": xbar_ucl, "xbar_lcl": xbar_lcl,
        "r_ucl": r_ucl, "r_lcl": r_lcl,
    }


def flag_ooc(values, ucl, lcl):
    """Returns boolean array — True where point is out of control."""
    return (values > ucl) | (values < lcl)


# ============================================================
# STEP 4: Plot everything
# ============================================================

def plot_results(data, capability, limits):
    fig = plt.figure(figsize=(16, 10))
    fig.patch.set_facecolor("#0d1117")
    gs  = gridspec.GridSpec(2, 3, figure=fig, hspace=0.5, wspace=0.35)

    ax_xbar  = fig.add_subplot(gs[0, :2])   # X-bar chart (wide)
    ax_r     = fig.add_subplot(gs[1, :2])   # R chart (wide)
    ax_hist  = fig.add_subplot(gs[0, 2])    # Histogram
    ax_kpi   = fig.add_subplot(gs[1, 2])    # KPI scorecard

    for ax in [ax_xbar, ax_r, ax_hist, ax_kpi]:
        ax.set_facecolor("#0d1117")
        ax.tick_params(colors="white")
        for spine in ax.spines.values():
            spine.set_edgecolor("#333")
        ax.grid(color="#1e1e2e", linewidth=0.5, linestyle="--")

    sg_labels = data.index
    x_idx     = np.arange(len(sg_labels))

    xbars  = limits["xbars"]
    ranges = limits["ranges"]
    ooc_x  = flag_ooc(xbars, limits["xbar_ucl"], limits["xbar_lcl"])
    ooc_r  = flag_ooc(ranges, limits["r_ucl"], limits["r_lcl"])

    # --- X-bar Chart ---
    ax_xbar.plot(x_idx, xbars, color="#4A90D9", linewidth=1.5, marker="o", markersize=4, label="X-bar")
    ax_xbar.scatter(x_idx[ooc_x], xbars[ooc_x], color="#FF4444", s=80, zorder=5, label="Out of Control")
    ax_xbar.axhline(limits["xbar_bar"], color="#7ED321", linewidth=1.2, linestyle="--", label="CL")
    ax_xbar.axhline(limits["xbar_ucl"], color="#F5A623", linewidth=1.2, linestyle="--", label=f"UCL={limits['xbar_ucl']:.3f}")
    ax_xbar.axhline(limits["xbar_lcl"], color="#F5A623", linewidth=1.2, linestyle="--", label=f"LCL={limits['xbar_lcl']:.3f}")
    ax_xbar.axhline(USL, color="#FF4444", linewidth=0.8, linestyle=":", alpha=0.6)
    ax_xbar.axhline(LSL, color="#FF4444", linewidth=0.8, linestyle=":", alpha=0.6)
    ax_xbar.set_title(f"X-bar Control Chart — {PROCESS_NAME}", color="white", fontsize=11, fontweight="bold")
    ax_xbar.set_ylabel(f"Subgroup Mean ({UNIT})", color="white")
    ax_xbar.legend(fontsize=8, facecolor="#1a1a2e", labelcolor="white", edgecolor="#333", ncol=4)
    ax_xbar.set_xticks(x_idx[::3])
    ax_xbar.set_xticklabels(sg_labels[::3], rotation=45, fontsize=8)

    # --- R Chart ---
    ax_r.plot(x_idx, ranges, color="#BD10E0", linewidth=1.5, marker="s", markersize=4, label="Range")
    ax_r.scatter(x_idx[ooc_r], ranges[ooc_r], color="#FF4444", s=80, zorder=5, label="Out of Control")
    ax_r.axhline(limits["r_bar"], color="#7ED321", linewidth=1.2, linestyle="--", label="R-bar")
    ax_r.axhline(limits["r_ucl"], color="#F5A623", linewidth=1.2, linestyle="--", label=f"UCL={limits['r_ucl']:.3f}")
    ax_r.axhline(limits["r_lcl"], color="#F5A623", linewidth=1.2, linestyle="--", label=f"LCL={limits['r_lcl']:.3f}")
    ax_r.set_title("R Chart (Range / Variability)", color="white", fontsize=11, fontweight="bold")
    ax_r.set_ylabel(f"Subgroup Range ({UNIT})", color="white")
    ax_r.set_xlabel("Subgroup", color="white")
    ax_r.legend(fontsize=8, facecolor="#1a1a2e", labelcolor="white", edgecolor="#333", ncol=4)
    ax_r.set_xticks(x_idx[::3])
    ax_r.set_xticklabels(sg_labels[::3], rotation=45, fontsize=8)

    # --- Histogram with spec limits ---
    all_vals = data.values.flatten()
    ax_hist.hist(all_vals, bins=20, color="#4A90D9", edgecolor="#0d1117", density=True, alpha=0.8)
    xmin, xmax = all_vals.min() - 0.05, all_vals.max() + 0.05
    xs = np.linspace(xmin, xmax, 200)
    ax_hist.plot(xs, stats.norm.pdf(xs, capability["mean"], capability["std"]),
                 color="#7ED321", linewidth=2, label="Normal fit")
    ax_hist.axvline(LSL, color="#FF4444", linewidth=2, linestyle="--", label=f"LSL={LSL}")
    ax_hist.axvline(USL, color="#FF4444", linewidth=2, linestyle="--", label=f"USL={USL}")
    ax_hist.axvline(TARGET, color="#F5A623", linewidth=1.5, linestyle=":", label=f"Target={TARGET}")
    ax_hist.set_title("Process Distribution", color="white", fontsize=11, fontweight="bold")
    ax_hist.set_xlabel(UNIT, color="white")
    ax_hist.set_ylabel("Density", color="white")
    ax_hist.legend(fontsize=8, facecolor="#1a1a2e", labelcolor="white", edgecolor="#333")

    # --- KPI Scorecard ---
    ax_kpi.axis("off")
    cp, cpk = capability["cp"], capability["cpk"]
    def grade(v): return "✅" if v >= 1.33 else ("⚠️" if v >= 1.0 else "❌")
    kpis = [
        ("Cp",          f"{cp:.3f}  {grade(cp)}"),
        ("Cpk",         f"{cpk:.3f}  {grade(cpk)}"),
        ("Sigma Level", f"{capability['sigma_level']:.2f}σ"),
        ("Mean",        f"{capability['mean']:.4f} {UNIT}"),
        ("Std Dev",     f"{capability['std']:.4f} {UNIT}"),
        ("Est. PPM",    f"{capability['ppm_total']:.0f}"),
        ("OOC Points",  f"{ooc_x.sum() + ooc_r.sum()}"),
    ]
    ax_kpi.set_title("KPI Scorecard", color="white", fontsize=11, fontweight="bold", pad=10)
    for i, (label, val) in enumerate(kpis):
        y = 0.9 - i * 0.13
        ax_kpi.text(0.05, y, label, color="#AAAAAA", fontsize=10, transform=ax_kpi.transAxes)
        ax_kpi.text(0.95, y, val, color="white", fontsize=10, fontweight="bold",
                    ha="right", transform=ax_kpi.transAxes)
        ax_kpi.axhline(y - 0.04, xmin=0.05, xmax=0.95, color="#333", linewidth=0.5,
                       transform=ax_kpi.transAxes)

    fig.suptitle(f"Six Sigma Process Capability Report — {PROCESS_NAME}",
                 color="white", fontsize=14, fontweight="bold", y=1.01)
    plt.savefig("sixsigma_report.png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.show()
    print("\n💾 Chart saved as sixsigma_report.png")


# ============================================================
# STEP 5: Export to Excel
# ============================================================

def export_excel(data, capability, limits, ooc_x, ooc_r):
    if not EXCEL_OK:
        print("⚠️  Skipping Excel export.")
        return

    fname = "sixsigma_report.xlsx"
    with pd.ExcelWriter(fname, engine="openpyxl") as writer:

        # KPI summary
        cp, cpk = capability["cp"], capability["cpk"]
        kpi_df = pd.DataFrame({
            "KPI":   ["Cp","Cpk","Sigma Level","Mean","Std Dev","Est. PPM","OOC (X-bar)","OOC (R)"],
            "Value": [round(cp,3), round(cpk,3), round(capability["sigma_level"],2),
                      round(capability["mean"],4), round(capability["std"],4),
                      round(capability["ppm_total"],0), int(ooc_x.sum()), int(ooc_r.sum())],
            "Status":["Capable ✅" if cp>=1.33 else "Marginal ⚠️" if cp>=1.0 else "Not Capable ❌",
                      "Centred ✅" if cpk>=1.33 else "Off-Centre ⚠️" if cpk>=1.0 else "Not Centred ❌",
                      "","","","","",""]
        })
        kpi_df.to_excel(writer, sheet_name="Capability Summary", index=False)

        # Raw data
        data.to_excel(writer, sheet_name="Raw Measurements")

        # Control chart data
        cc_df = pd.DataFrame({
            "Subgroup":   data.index,
            "X-bar":      limits["xbars"].round(4),
            "Range":      limits["ranges"].round(4),
            "X-bar UCL":  round(limits["xbar_ucl"], 4),
            "X-bar LCL":  round(limits["xbar_lcl"], 4),
            "R UCL":      round(limits["r_ucl"], 4),
            "OOC (Xbar)": ooc_x.astype(int),
            "OOC (R)":    ooc_r.astype(int),
        })
        cc_df.to_excel(writer, sheet_name="Control Chart Data", index=False)

    print(f"📁 Excel report saved as {fname}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 55)
    print("   SIX SIGMA PROCESS CAPABILITY TOOL")
    print(f"   Process: {PROCESS_NAME}  |  LSL: {LSL}  USL: {USL}")
    print("=" * 55)

    data        = generate_process_data()
    capability  = calculate_capability(data)
    limits      = calculate_control_limits(data)

    ooc_x = flag_ooc(limits["xbars"], limits["xbar_ucl"], limits["xbar_lcl"])
    ooc_r = flag_ooc(limits["ranges"], limits["r_ucl"],   limits["r_lcl"])
    print(f"\n🔎 Out-of-control points:  X-bar: {ooc_x.sum()}  |  R-chart: {ooc_r.sum()}")

    plot_results(data, capability, limits)
    export_excel(data, capability, limits, ooc_x, ooc_r)

    print("\n✅ All done!")
    print("=" * 55)


if __name__ == "__main__":
    main()
