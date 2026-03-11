import streamlit as st
import openpyxl
import networkx as nx
import matplotlib.pyplot as plt
import re
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Excel Model Auditor", layout="wide")
st.title("Excel Model Auditor")
st.caption("Upload a workbook to visualize its dependency graph and detect structural issues.")

# --- Sidebar ---
with st.sidebar:
    st.header("Settings")
    uploaded_file = st.file_uploader("Upload Excel workbook", type=["xlsx"])
    input_sheets = st.text_input(
        "Input sheet names (comma-separated)",
        value="Inputs",
        help="Cells in these sheets are expected to hold constants — they won't be flagged as hardcoded.",
    )
    input_sheet_list = [s.strip() for s in input_sheets.split(",") if s.strip()]

if uploaded_file is None:
    st.info("Upload an Excel file in the sidebar to get started.")
    st.stop()

# --- Helpers ---
def get_formula_str(value):
    """Return the formula as a plain string.
    openpyxl returns dynamic array formulas (FILTER, UNIQUE, SORT, XLOOKUP, etc.)
    as ArrayFormula objects rather than strings — unwrap them here."""
    if hasattr(value, "text"):   # ArrayFormula namedtuple
        return value.text or ""
    return value or ""

def expand_range(sheet, start, end):
    col1, row1 = coordinate_from_string(start)
    col2, row2 = coordinate_from_string(end)
    col1_i = column_index_from_string(col1)
    col2_i = column_index_from_string(col2)
    cells = []
    for col in range(col1_i, col2_i + 1):
        for row in range(row1, row2 + 1):
            cells.append(f"{sheet}!{get_column_letter(col)}{row}")
    return cells

def node_color(node, hardcoded_cells):
    if node in hardcoded_cells:
        return "gold"
    sheet = node.split("!")[0]
    if sheet in input_sheet_list:
        return "lightgreen"
    colors = ["lightskyblue", "salmon", "plum", "peachpuff", "lightcyan"]
    idx = sheet_names.index(sheet) if sheet in sheet_names else -1
    return colors[idx % len(colors)]

# --- Load workbook ---
wb = openpyxl.load_workbook(uploaded_file, data_only=False)
sheet_names = wb.sheetnames

st.subheader("Sheets detected")
st.write(", ".join(sheet_names))

# --- Preprocess: expand named ranges in formulas ---
named_ranges_map = {}
for name in wb.defined_names:
    defn = wb.defined_names[name]
    try:
        parts = []
        for sheet_name, cell_range in defn.destinations:
            clean = cell_range.replace("$", "")
            sheet_ref = f"'{sheet_name}'" if " " in sheet_name else sheet_name
            parts.append(f"{sheet_ref}!{clean}")
        named_ranges_map[name] = ",".join(parts)
    except Exception:
        pass

if named_ranges_map:
    sorted_names = sorted(named_ranges_map, key=len, reverse=True)
    for sheet in sheet_names:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f" and cell.value:
                    formula = get_formula_str(cell.value)
                    for name in sorted_names:
                        pattern = re.compile(r"\b" + re.escape(name) + r"\b", re.IGNORECASE)
                        formula = pattern.sub(named_ranges_map[name], formula)
                    cell.value = formula

# --- Detect hardcoded constants ---
hardcoded_cells = set()
for sheet in sheet_names:
    if sheet in input_sheet_list:
        continue
    ws = wb[sheet]
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "n" and cell.value is not None:
                hardcoded_cells.add(f"{sheet}!{cell.coordinate}")

# --- Build dependency graph ---
G = nx.DiGraph()
cell_pattern = r"[A-Za-z_']+!\$?[A-Za-z]+\$?\d+"
range_pattern = r"([A-Za-z_']+)!([A-Za-z]+\d+):([A-Za-z]+\d+)"

for sheet in sheet_names:
    ws = wb[sheet]
    for row in ws.iter_rows():
        for cell in row:
            location = f"{sheet}!{cell.coordinate}"
            if cell.data_type == "f":
                formula = get_formula_str(cell.value)
                for r in re.findall(range_pattern, formula):
                    rng_sheet, start, end = r
                    rng_sheet = rng_sheet.strip("'")
                    for ref in expand_range(rng_sheet, start, end):
                        G.add_edge(ref, location)
                for ref in re.findall(cell_pattern, formula):
                    ref = ref.replace("$", "").strip("'")
                    G.add_edge(ref, location)

for cell in hardcoded_cells:
    if cell not in G:
        G.add_node(cell)

# --- Layout: assign x position by sheet order ---
layers = {name: i for i, name in enumerate(sheet_names)}
pos = {}
y_positions = {}
for node in sorted(G.nodes):
    sheet = node.split("!")[0].strip("'")
    layer = layers.get(sheet, len(sheet_names))
    y_positions.setdefault(sheet, 0)
    pos[node] = (layer, -y_positions[sheet])
    y_positions[sheet] += 1

# --- Draw graph ---
colors = [node_color(n, hardcoded_cells) for n in G.nodes]
sizes = [400 + len(nx.descendants(G, n)) * 100 for n in G.nodes]

fig, ax = plt.subplots(figsize=(14, 8))
nx.draw(
    G, pos,
    ax=ax,
    with_labels=True,
    node_color=colors,
    node_size=sizes,
    font_size=8,
    arrows=True,
)
ax.set_title("Excel Model Dependency Graph")

# Legend
from matplotlib.patches import Patch
legend_entries = [Patch(color="gold", label="Hardcoded constant")]
sheet_colors = ["lightgreen", "lightskyblue", "salmon", "plum", "peachpuff", "lightcyan"]
for i, name in enumerate(sheet_names):
    legend_entries.append(Patch(color=sheet_colors[i % len(sheet_colors)], label=name))
ax.legend(handles=legend_entries, loc="upper left")

st.subheader("Dependency graph")
st.pyplot(fig)
plt.close(fig)

# --- Results panels ---
col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Nodes", len(G.nodes))

with col2:
    st.metric("Edges", len(G.edges))

with col3:
    st.metric("Hardcoded constants", len(hardcoded_cells))

col_a, col_b = st.columns(2)

with col_a:
    st.subheader("Hardcoded constants")
    st.caption("Numeric values outside input sheets — should these be in Inputs?")
    if hardcoded_cells:
        st.dataframe({"Cell": sorted(hardcoded_cells)}, use_container_width=True)
    else:
        st.success("None found.")

with col_b:
    st.subheader("Orphan cells")
    st.caption("Cells with no connections — overwritten formulas, unused assumptions, or abandoned calculations.")
    orphans = [n for n in G.nodes if G.degree(n) == 0]
    if orphans:
        st.dataframe({"Cell": sorted(orphans)}, use_container_width=True)
    else:
        st.success("None found.")

if named_ranges_map:
    with st.expander("Named ranges resolved"):
        st.dataframe(
            {"Name": list(named_ranges_map.keys()), "Resolves to": list(named_ranges_map.values())},
            use_container_width=True,
        )
