"""
TLM Settlement Analysis – Interactive Streamlit Application
Author: Research-use template
Run:
    pip install -r requirements.txt
    streamlit run tlm_settlement_app.py

Model scope:
- Axisymmetric cylindrical geometry is used for area/volume calculations.
- Surcharge is assumed uniformly distributed over the column top.
- Layer-wise surcharge-induced compression is calculated with a logarithmic
  compression-index model, partitioned into instantaneous and primary
  consolidation components to avoid counting the same compression twice.
- Secondary compression and first-order biological settlement are optional.
- This is an engineering screening/calibration tool, not a validated
  hydro-mechanical finite-element solver. Hydraulic conductivity and pipe
  geometry are reported and used to define drainage assumptions; they do not
  independently determine Cv or settlement.
"""

from __future__ import annotations

import io
import math
import zipfile
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(page_title="TLM Settlement Analysis", page_icon="🏗️", layout="wide")

st.title("Triple-Layered Landfill (TLM) – Settlement Analysis")
st.caption("Cylindrical, three-layer, time-dependent engineering calculation tool")

st.warning(
    "Research caution: predictions are only as reliable as the measured input parameters. "
    "The default material values are placeholders for sensitivity analysis, not measured "
    "properties of your MSW. Do not report placeholder-based predictions as experimental results."
)

G = 9.80665
GAMMA_W = 9.80665  # kN/m3, approximately

def consolidation_degree(Tv: np.ndarray, terms: int = 80) -> np.ndarray:
    """Average 1-D Terzaghi degree of consolidation for double/single drainage path."""
    Tv = np.maximum(np.asarray(Tv, dtype=float), 0.0)
    U = np.zeros_like(Tv)
    for n in range(terms):
        m = 2 * n + 1
        U += (8.0 / (math.pi**2 * m**2)) * np.exp(
            -(m**2 * math.pi**2 * Tv / 4.0)
        )
    return np.clip(1.0 - U, 0.0, 1.0)

def make_time_grid(hours: float, points: int) -> np.ndarray:
    """Include t=0 and logarithmically spaced positive times in hours."""
    positive = np.geomspace(max(hours / (points * 1000), 1e-5), hours, points - 1)
    return np.unique(np.concatenate(([0.0], positive)))

def layer_results(layer, q_kpa, t_hours, immediate_fraction):
    """
    Calculate surcharge-induced compression for one layer.
    Total ultimate primary-mechanical compression from Cc is partitioned:
      immediate = f * S_mech,ultimate
      primary(t) = (1-f) * U(t) * S_mech,ultimate
    Secondary and biological components are additional calibrated mechanisms.
    """
    H = layer["H"]
    Cc = layer["Cc"]
    e0 = layer["e0"]
    sigma0 = layer["sigma0"]
    Cv = layer["Cv"]
    Hd = layer["Hd"]
    C_alpha = layer["Calpha"]
    ep = layer["ep"]
    t_primary = layer["t_primary"]
    eps_b = layer["eps_b"]
    kb = layer["kb"]

    if sigma0 <= 0:
        raise ValueError("Initial effective stress must be greater than zero.")
    if H <= 0:
        raise ValueError("Waste thickness must be greater than zero.")
    if e0 <= -1 or ep <= -1:
        raise ValueError("Void ratios must be greater than -1.")
    if Cv < 0 or Hd <= 0:
        raise ValueError("Cv must be non-negative and drainage path must be positive.")

    # Surcharge-induced ultimate compression from the logarithmic Cc relation.
    S_mech_ultimate = H * Cc / (1.0 + e0) * math.log10((sigma0 + q_kpa) / sigma0)
    S_im = immediate_fraction * S_mech_ultimate

    t_seconds = t_hours * 3600.0
    Tv = Cv * t_seconds / (Hd**2)
    U = consolidation_degree(Tv)
    S_primary = (1.0 - immediate_fraction) * S_mech_ultimate * U

    # Secondary compression is zero before its specified reference time.
    t_years = np.maximum(t_hours / 8760.0, 1e-12)
    tp_years = max(t_primary / 8760.0, 1e-12)
    S_secondary = np.zeros_like(t_hours)
    mask = t_hours >= t_primary
    if C_alpha > 0:
        S_secondary[mask] = (
            H * C_alpha / (1.0 + ep)
            * np.log10(np.maximum(t_years[mask] / tp_years, 1.0))
        )

    # First-order biological strain; time is expressed in days for kb [1/day].
    t_days = t_hours / 24.0
    S_bio = H * eps_b * (1.0 - np.exp(-kb * t_days))

    total = S_im + S_primary + S_secondary + S_bio
    return {
        "Surcharge ultimate mechanical (mm)": S_mech_ultimate * 1000,
        "Immediate (mm)": np.full_like(t_hours, S_im * 1000),
        "Primary consolidation (mm)": S_primary * 1000,
        "Secondary compression (mm)": S_secondary * 1000,
        "Biological (mm)": S_bio * 1000,
        "Total (mm)": total * 1000,
        "Degree consolidation (%)": U * 100,
        "Time factor Tv": Tv,
    }

def build_excel(summary_df, time_df, input_df, plot_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["TLM Settlement Analysis"])
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    ws.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(["Layer-wise summary at final analysis time"])
    for row in [summary_df.columns.tolist()] + summary_df.values.tolist():
        ws.append(list(row))

    ws2 = wb.create_sheet("Time Series")
    for row in [time_df.columns.tolist()] + time_df.values.tolist():
        ws2.append(list(row))

    ws3 = wb.create_sheet("Inputs")
    for row in [input_df.columns.tolist()] + input_df.values.tolist():
        ws3.append(list(row))

    ws4 = wb.create_sheet("Plot Data")
    for row in [plot_df.columns.tolist()] + plot_df.values.tolist():
        ws4.append(list(row))

    # Add a native Excel line chart for total settlement.
    chart = LineChart()
    chart.title = "Total TLM Settlement vs Time"
    chart.y_axis.title = "Settlement (mm)"
    chart.x_axis.title = "Time (hours)"
    data = Reference(ws2, min_col=ws2.max_column, min_row=1, max_row=ws2.max_row)
    cats = Reference(ws2, min_col=1, min_row=2, max_row=ws2.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 9
    chart.width = 18
    ws2.add_chart(chart, "L2")

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        for col in sheet.columns:
            max_len = 0
            letter = get_column_letter(col[0].column)
            for cell in col[: min(len(col), 300)]:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            sheet.column_dimensions[letter].width = min(max(max_len + 2, 12), 42)
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()

with st.sidebar:
    st.header("1. Column geometry")
    diameter_cm = st.number_input("Internal column diameter (cm)", min_value=1.0, value=30.0, step=0.5)
    total_height_m = st.number_input("Total column height (m)", min_value=0.1, value=1.0, step=0.05)
    n_layers = 3
    st.caption("This version models exactly three waste layers.")
    layer_thickness = []
    gravel_thickness = []
    for i, default in enumerate([0.2833, 0.2833, 0.2833], start=1):
        layer_thickness.append(st.number_input(f"Layer {i} compressible waste thickness (m)", min_value=0.01, value=default, step=0.005, key=f"h{i}"))
    for i in range(1, 4):
        gravel_thickness.append(st.number_input(f"Drainage gravel thickness below layer {i} (m)", min_value=0.0, value=0.05, step=0.005, key=f"g{i}"))

    st.header("2. Drainage pipes and outlets")
    pipe_diameter_cm = st.number_input("Collection pipe internal diameter (cm)", min_value=0.1, value=2.0, step=0.1)
    pipe_count = st.number_input("Number of collection pipes per outlet", min_value=1, value=1, step=1)
    outlet_elevations = [
        st.number_input("Outlet 1 depth from top (m)", min_value=0.0, value=0.33, step=0.01),
        st.number_input("Outlet 2 depth from top (m)", min_value=0.0, value=0.66, step=0.01),
        st.number_input("Outlet 3 depth from top (m)", min_value=0.0, value=1.00, step=0.01),
    ]
    pipe_length_m = st.number_input("Approximate pipe length per outlet (m)", min_value=0.0, value=0.15, step=0.01)

    st.header("3. Hydraulic properties")
    Kr = st.number_input("Radial hydraulic conductivity Kr (m/s)", min_value=0.0, value=0.0028, format="%.6f")
    Kz = st.number_input("Vertical hydraulic conductivity Kz (m/s)", min_value=0.0, value=0.0032, format="%.6f")
    porosity = st.number_input("Effective porosity (-)", min_value=0.01, max_value=0.99, value=0.45, step=0.01)
    leachate_litres = st.number_input("Initial leachate application (L)", min_value=0.0, value=1.0, step=0.1)
    drainage_mode = st.selectbox("Drainage condition for 1-D consolidation approximation",
                                 ["Double drainage (Hd = H/2)", "Single drainage (Hd = H)", "User-defined drainage path"])

    st.header("4. Surcharge")
    surcharge_mode = st.selectbox("Loading cases", ["All three: 0.5, 1.0, 1.5 kg", "Single user-defined load"])
    custom_mass = st.number_input("Custom surcharge mass (kg)", min_value=0.0, value=1.0, step=0.1)
    load_plate_diameter_cm = st.number_input("Load plate diameter (cm; 0 = full column)", min_value=0.0, value=30.0, step=0.5)
    load_uniform = st.checkbox("Assume uniform pressure under the plate", value=True)

    st.header("5. Analysis time")
    horizon_hours = st.number_input("Prediction horizon (hours)", min_value=1.0, value=24.0*365.0*10, step=24.0)
    n_time = st.slider("Number of time points", min_value=50, max_value=1000, value=240, step=10)
    st.caption("Time horizon controls the plotted prediction period; it is not a prediction of when settlement physically stops.")

st.header("Material and settlement parameters")
st.info(
    "Enter measured values where available. Placeholder defaults are provided only to let you test "
    "the interface and perform sensitivity analysis. Replace them before using numerical predictions in a thesis."
)

cols = st.columns(3)
layer_inputs = []
defaults = [
    {"rho": 600.0, "e0": 1.0, "M": 100.0, "Cc": 0.20, "sigma0": 5.0, "Cv": 1e-7, "Hd": 0.14165, "Ca": 0.02, "ep": 0.8, "tp": 24.0, "epsb": 0.05, "kb": 0.01},
    {"rho": 600.0, "e0": 1.0, "M": 100.0, "Cc": 0.20, "sigma0": 8.0, "Cv": 1e-7, "Hd": 0.14165, "Ca": 0.02, "ep": 0.8, "tp": 24.0, "epsb": 0.05, "kb": 0.01},
    {"rho": 600.0, "e0": 1.0, "M": 100.0, "Cc": 0.20, "sigma0": 11.0, "Cv": 1e-7, "Hd": 0.14165, "Ca": 0.02, "ep": 0.8, "tp": 24.0, "epsb": 0.05, "kb": 0.01},
]
for i in range(3):
    with cols[i]:
        st.subheader(f"Layer {i+1}")
        d = defaults[i]
        rho = st.number_input("Bulk density (kg/m³)", min_value=1.0, value=d["rho"], step=10.0, key=f"rho{i}")
        e0 = st.number_input("Initial void ratio e0", min_value=0.0, value=d["e0"], step=0.05, key=f"e0{i}")
        M = st.number_input("Constrained modulus M (kPa; reference only)", min_value=0.01, value=d["M"], step=5.0, key=f"M{i}")
        Cc = st.number_input("Compression index Cc", min_value=0.0, value=d["Cc"], step=0.01, key=f"Cc{i}")
        sigma0 = st.number_input("Initial effective stress σ'0 (kPa)", min_value=0.01, value=d["sigma0"], step=0.5, key=f"sigma{i}")
        Cv = st.number_input("Vertical consolidation coefficient Cv (m²/s)", min_value=0.0, value=d["Cv"], format="%.2e", key=f"Cv{i}")
        if drainage_mode == "Double drainage (Hd = H/2)":
            Hd = layer_thickness[i] / 2.0
            st.caption(f"Computed Hd = {Hd:.4f} m")
        elif drainage_mode == "Single drainage (Hd = H)":
            Hd = layer_thickness[i]
            st.caption(f"Computed Hd = {Hd:.4f} m")
        else:
            Hd = st.number_input("User drainage path Hd (m)", min_value=0.001, value=d["Hd"], step=0.005, key=f"Hd{i}")
        Calpha = st.number_input("Secondary compression index Cα", min_value=0.0, value=d["Ca"], step=0.005, key=f"Ca{i}")
        ep = st.number_input("Void ratio at primary end ep", min_value=0.0, value=d["ep"], step=0.05, key=f"ep{i}")
        t_primary = st.number_input("Primary reference time (hours)", min_value=0.01, value=d["tp"], step=1.0, key=f"tp{i}")
        eps_b = st.number_input("Ultimate biological strain (-)", min_value=0.0, value=d["epsb"], step=0.01, key=f"epsb{i}")
        kb = st.number_input("Biodegradation rate kb (1/day)", min_value=0.0, value=d["kb"], format="%.4f", key=f"kb{i}")
        layer_inputs.append({
            "H": layer_thickness[i], "rho": rho, "e0": e0, "M": M, "Cc": Cc,
            "sigma0": sigma0, "Cv": Cv, "Hd": Hd, "Calpha": Calpha,
            "ep": ep, "t_primary": t_primary, "eps_b": eps_b, "kb": kb
        })

st.header("6. Model assumptions")
c1, c2, c3 = st.columns(3)
with c1:
    immediate_fraction = st.slider("Instantaneous fraction of Cc-based mechanical compression", 0.0, 1.0, 0.20, 0.05)
with c2:
    use_secondary = st.checkbox("Include secondary compression", value=True)
with c3:
    use_bio = st.checkbox("Include biological settlement", value=False)

st.caption(
    "The Cc-based ultimate compression is partitioned between immediate and primary consolidation. "
    "The constrained modulus M is shown for comparison/sensitivity but is not added as a second "
    "compression calculation, avoiding double-counting the same surcharge response."
)

# Geometry validation
D = diameter_cm / 100.0
R = D / 2.0
A = math.pi * R**2
H_w_total = sum(layer_thickness)
H_g_total = sum(gravel_thickness)
H_occupied = H_w_total + H_g_total
pipe_d = pipe_diameter_cm / 100.0
pipe_area = math.pi * (pipe_d / 2.0)**2
plate_d = load_plate_diameter_cm / 100.0 if load_plate_diameter_cm > 0 else D
plate_area = math.pi * (plate_d / 2.0)**2

geometry_errors = []
if H_occupied > total_height_m + 1e-9:
    geometry_errors.append(f"Total waste + gravel thickness ({H_occupied:.4f} m) exceeds column height ({total_height_m:.4f} m).")
if plate_d > D + 1e-9:
    geometry_errors.append("Load plate diameter cannot exceed column internal diameter in this simple model.")
if any(x < 0 or x > total_height_m for x in outlet_elevations):
    geometry_errors.append("All outlet depths must be within the column height.")
if pipe_d >= D:
    geometry_errors.append("Pipe diameter must be smaller than column diameter.")
if abs(H_occupied - total_height_m) > 0.02:
    st.warning(f"Geometry check: waste + gravel = {H_occupied:.4f} m, while column height = {total_height_m:.4f} m. Check any gaps or unmodelled layers.")

if geometry_errors:
    for e in geometry_errors:
        st.error(e)
    st.stop()

st.subheader("Geometry and loading checks")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Column cross-sectional area", f"{A:.5f} m²")
m2.metric("Total waste volume", f"{A*H_w_total*1000:.2f} L")
m3.metric("Total gravel volume", f"{A*H_g_total*1000:.2f} L")
m4.metric("Pipe cross-sectional area", f"{pipe_area*1e6:.1f} mm²")

if not load_uniform and plate_d != D:
    st.warning("Non-uniform plate pressure distribution is not solved by this 1-D model. The calculation still uses average pressure over the plate area; use a 3-D/axisymmetric mechanical model for edge effects.")

if surcharge_mode == "All three: 0.5, 1.0, 1.5 kg":
    loads = [0.5, 1.0, 1.5]
else:
    loads = [custom_mass]

times = make_time_grid(float(horizon_hours), int(n_time))
all_summaries = []
all_curves = {}
all_time_tables = []

for mass in loads:
    load_area = plate_area
    q_kpa = mass * G / load_area / 1000.0
    layer_curves = []
    for i, layer in enumerate(layer_inputs):
        result = layer_results(layer, q_kpa, times, immediate_fraction)
        if not use_secondary:
            result["Secondary compression (mm)"][:] = 0.0
        if not use_bio:
            result["Biological (mm)"][:] = 0.0
        layer_curves.append(result)

        final = {k: float(v[-1]) for k, v in result.items() if isinstance(v, np.ndarray)}
        all_summaries.append({
            "Load (kg)": mass, "Surcharge (kPa)": q_kpa, "Layer": f"Layer {i+1}",
            "Waste thickness (m)": layer["H"], "Bulk density (kg/m³)": layer["rho"],
            "Initial void ratio": layer["e0"], "Constrained modulus (kPa; reference)": layer["M"],
            "Cc": layer["Cc"], "Initial effective stress (kPa)": layer["sigma0"],
            "Cv (m²/s)": layer["Cv"], "Drainage path (m)": layer["Hd"],
            "Immediate (mm)": final["Immediate (mm)"],
            "Primary consolidation (mm)": final["Primary consolidation (mm)"],
            "Secondary compression (mm)": final["Secondary compression (mm)"],
            "Biological (mm)": final["Biological (mm)"],
            "Total layer settlement (mm)": final["Total (mm)"],
            "Degree consolidation (%)": final["Degree consolidation (%)"],
        })

    total_curve = sum(x["Total (mm)"] for x in layer_curves)
    all_curves[mass] = {
        "time": times,
        "layers": layer_curves,
        "total": total_curve,
        "q": q_kpa,
    }
    time_dict = {"Time (hours)": times, "Time (days)": times / 24.0, "Time (years)": times / 8760.0}
    for i, res in enumerate(layer_curves, start=1):
        time_dict[f"Layer {i} total (mm)"] = res["Total (mm)"]
        time_dict[f"Layer {i} primary (mm)"] = res["Primary consolidation (mm)"]
        time_dict[f"Layer {i} secondary (mm)"] = res["Secondary compression (mm)"]
        time_dict[f"Layer {i} biological (mm)"] = res["Biological (mm)"]
    time_dict["Total TLM settlement (mm)"] = total_curve
    time_dict["Surcharge (kg)"] = np.full_like(times, mass)
    all_time_tables.append(pd.DataFrame(time_dict))

summary_df = pd.DataFrame(all_summaries)
time_df = pd.concat(all_time_tables, ignore_index=True)

# Input register
input_rows = [
    ("Column diameter (m)", D), ("Column height (m)", total_height_m),
    ("Total waste thickness (m)", H_w_total), ("Total gravel thickness (m)", H_g_total),
    ("Pipe diameter (m)", pipe_d), ("Pipe count per outlet", pipe_count),
    ("Pipe length per outlet (m)", pipe_length_m),
    ("Outlet 1 depth (m)", outlet_elevations[0]), ("Outlet 2 depth (m)", outlet_elevations[1]),
    ("Outlet 3 depth (m)", outlet_elevations[2]), ("Radial hydraulic conductivity (m/s)", Kr),
    ("Vertical hydraulic conductivity (m/s)", Kz), ("Effective porosity (-)", porosity),
    ("Leachate application (L)", leachate_litres), ("Prediction horizon (hours)", horizon_hours),
    ("Instantaneous fraction", immediate_fraction), ("Secondary included", use_secondary),
    ("Biological included", use_bio), ("Assumed uniform top pressure", load_uniform),
]
for i, layer in enumerate(layer_inputs, start=1):
    for k, v in layer.items():
        input_rows.append((f"Layer {i} – {k}", v))
input_df = pd.DataFrame(input_rows, columns=["Input parameter", "Value"])

st.subheader("7. Numerical results")
final_rows = []
for mass, data in all_curves.items():
    final_rows.append({
        "Surcharge (kg)": mass,
        "Pressure (kPa)": data["q"],
        "Layer 1 (mm)": data["layers"][0]["Total (mm)"][-1],
        "Layer 2 (mm)": data["layers"][1]["Total (mm)"][-1],
        "Layer 3 (mm)": data["layers"][2]["Total (mm)"][-1],
        "Total TLM (mm)": data["total"][-1],
    })
final_df = pd.DataFrame(final_rows)
st.dataframe(final_df, use_container_width=True)

st.caption(
    f"Values shown at the selected prediction horizon of {horizon_hours:g} hours "
    f"({horizon_hours/24:.2f} days; {horizon_hours/8760:.2f} years). "
    "They are model outputs, not measured settlement."
)

# Charts
st.subheader("8. Settlement graphs")
tab1, tab2, tab3 = st.tabs(["Total settlement vs time", "Layer-wise settlement", "Settlement components"])
with tab1:
    fig, ax = plt.subplots()
    for mass, data in all_curves.items():
        ax.plot(data["time"], data["total"], label=f"{mass:g} kg")
    ax.set_xlabel("Time (hours)")
    ax.set_ylabel("Total settlement (mm)")
    ax.set_title("Predicted total TLM settlement")
    ax.grid(True, alpha=0.3)
    ax.legend(title="Surcharge")
    st.pyplot(fig)
    plt.close(fig)

with tab2:
    selected_mass = st.selectbox("Select surcharge case for layer-wise plot", loads, format_func=lambda x: f"{x:g} kg")
    data = all_curves[selected_mass]
    fig, ax = plt.subplots()
    for i, layer_curve in enumerate(data["layers"], start=1):
        ax.plot(data["time"], layer_curve["Total (mm)"], label=f"Layer {i}")
    ax.plot(data["time"], data["total"], linestyle="--", linewidth=2, label="Total TLM")
    ax.set_xlabel("Time (hours)")
    ax.set_ylabel("Settlement (mm)")
    ax.set_title(f"Layer-wise settlement – {selected_mass:g} kg surcharge")
    ax.grid(True, alpha=0.3)
    ax.legend()
    st.pyplot(fig)
    plt.close(fig)

with tab3:
    data = all_curves[selected_mass]
    fig, ax = plt.subplots()
    for i, layer_curve in enumerate(data["layers"], start=1):
        ax.plot(data["time"], layer_curve["Immediate (mm)"], label=f"L{i} immediate")
        ax.plot(data["time"], layer_curve["Primary consolidation (mm)"], label=f"L{i} primary")
        ax.plot(data["time"], layer_curve["Secondary compression (mm)"], label=f"L{i} secondary")
        ax.plot(data["time"], layer_curve["Biological (mm)"], label=f"L{i} biological")
    ax.set_xlabel("Time (hours)")
    ax.set_ylabel("Settlement component (mm)")
    ax.set_title("Settlement components by layer")
    ax.grid(True, alpha=0.3)
    ax.legend(ncol=2, fontsize=8)
    st.pyplot(fig)
    plt.close(fig)

# Geometry table and hydraulic reporting
st.subheader("9. Outlet and drainage configuration")
outlet_df = pd.DataFrame({
    "Outlet": ["Outlet 1", "Outlet 2", "Outlet 3"],
    "Depth from top (m)": outlet_elevations,
    "Pipe diameter (mm)": [pipe_diameter_cm * 10] * 3,
    "Pipe count": [pipe_count] * 3,
    "Pipe length (m)": [pipe_length_m] * 3,
})
st.dataframe(outlet_df, use_container_width=True)
st.write(
    f"Hydraulic anisotropy Kr/Kz = {Kr/Kz:.4f}" if Kz > 0 else "Kz is zero; anisotropy ratio is undefined."
)
st.write(
    f"Applied leachate volume = {leachate_litres:.2f} L. "
    f"Nominal pore volume of waste = {porosity * A * H_w_total * 1000:.2f} L "
    "(assuming uniform effective porosity)."
)

# Validation / warnings
st.subheader("10. Model checks and validation status")
warnings = []
if abs(H_occupied - total_height_m) > 0.005:
    warnings.append("Waste plus gravel thickness does not exactly match total column height.")
if any(outlet_elevations[i] >= outlet_elevations[i+1] for i in range(2)):
    warnings.append("Outlet elevations should increase from top to bottom.")
if any(layer["sigma0"] <= 0 for layer in layer_inputs):
    warnings.append("Initial effective stress must be positive.")
if any(layer["Cv"] == 0 for layer in layer_inputs):
    warnings.append("A zero Cv produces no primary consolidation progression.")
if not warnings:
    st.success("Basic geometry and input checks passed. This is not experimental validation.")
else:
    for w in warnings:
        st.warning(w)

st.info(
    "Experimental validation is pending because no settlement observations have been entered. "
    "To validate, enter measured layer-interface displacements at matching times and compare "
    "measured versus predicted settlement using RMSE, MAE and R². This application currently "
    "does not infer measured settlement from leachate volumes."
)

# Export
st.subheader("11. Export results")
plot_rows = []
for mass, data in all_curves.items():
    for j, t in enumerate(data["time"]):
        plot_rows.append({
            "Surcharge (kg)": mass,
            "Time (hours)": t,
            "Total settlement (mm)": data["total"][j],
            "Layer 1 (mm)": data["layers"][0]["Total (mm)"][j],
            "Layer 2 (mm)": data["layers"][1]["Total (mm)"][j],
            "Layer 3 (mm)": data["layers"][2]["Total (mm)"][j],
        })
plot_df = pd.DataFrame(plot_rows)
excel_bytes = build_excel(final_df, time_df, input_df, plot_df)
st.download_button(
    "Download complete Excel workbook",
    data=excel_bytes,
    file_name="TLM_settlement_analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

csv_bytes = time_df.to_csv(index=False).encode("utf-8")
st.download_button("Download time-series CSV", data=csv_bytes, file_name="TLM_settlement_time_series.csv", mime="text/csv")

st.subheader("12. Interpretation and limitations")
st.markdown(
    """
- The surcharge pressure is calculated from the selected load and load-plate area.
- The compression-index model predicts surcharge-induced mechanical compression using the entered initial effective stress and Cc.
- The instantaneous fraction partitions that Cc-based compression between immediate and primary components; it is a calibration parameter.
- The constrained modulus is recorded for comparison and parameter checks; it is not added to the Cc-based compression to avoid double-counting.
- The 1-D Terzaghi consolidation approximation is not a full axisymmetric coupled hydro-mechanical solution.
- Pipe diameter, outlet elevation, gravel thickness, Kr and Kz describe the drainage configuration but do not automatically determine pore pressure or Cv.
- The secondary compression term is active only after the user-entered primary reference time.
- The biological term is an optional first-order long-term approximation. It should normally be switched off for short 24-hour laboratory settlement predictions unless independently justified.
- Wall friction, waste-gravel interface slip, nonuniform stress transfer, gas pressure, and unsaturated suction are not explicitly solved.
"""
)
