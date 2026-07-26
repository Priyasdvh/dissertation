"""
Generate Fig 5: DQ dimensions, issue themes, and affected data elements.

Input
-----
DQD_Theme_Element_Matrix.xlsx
Sheet: Matrix

Outputs
-------
Fig5.tif   PLOS submission file
Fig5.eps   Vector alternative
Fig5.png   Preview file

Figure specifications
---------------------
- 300 dpi
- 2250 × 2610 pixels
- Flattened RGB TIFF
- LZW compression
- Maximum file size: 10 MB
- Figure text: 8–10 pt
- Full data-element labels rotated diagonally
- No figure number, title, or caption inside the figure
"""

from pathlib import Path
import warnings

import matplotlib

# Allows figure generation on a Linux server without a desktop.
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Rectangle
import pandas as pd
from PIL import Image


# ============================================================
# 1. File settings
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILENAME = "DQD_Theme_Element_Matrix.xlsx"
SHEET_NAME = "Matrix"

OUTPUT_TIFF = BASE_DIR / "Fig5.tif"
OUTPUT_EPS = BASE_DIR / "Fig5.eps"
OUTPUT_PNG = BASE_DIR / "Fig5.png"


# ============================================================
# 2. PLOS figure settings
# ============================================================

DPI = 300

# 7.5 × 8.7 inches at 300 dpi:
# 2250 × 2610 pixels
FIGURE_WIDTH_INCHES = 7.5
FIGURE_HEIGHT_INCHES = 8.7

FIGURE_SIZE = (
    FIGURE_WIDTH_INCHES,
    FIGURE_HEIGHT_INCHES,
)

EXPECTED_WIDTH_PIXELS = round(
    FIGURE_WIDTH_INCHES * DPI
)

EXPECTED_HEIGHT_PIXELS = round(
    FIGURE_HEIGHT_INCHES * DPI
)

MIN_WIDTH_PIXELS = 789
MAX_WIDTH_PIXELS = 2250
MAX_HEIGHT_PIXELS = 2625
MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024

STAR_SIZE = 180

# Set to True after Arial or Times is installed.
# False permits DejaVu Sans temporarily on the server.
STRICT_PLOS_FONT = False


# ============================================================
# 3. Font selection
# ============================================================

def select_figure_font() -> tuple[str, bool]:
    """
    Select an installed PLOS-compatible font.

    Returns
    -------
    tuple[str, bool]
        Font name and whether it is formally PLOS-approved.
    """

    installed_fonts = {
        font.name
        for font in font_manager.fontManager.ttflist
    }

    approved_fonts = [
        "Arial",
        "Times New Roman",
        "Times",
    ]

    for font_name in approved_fonts:
        if font_name in installed_fonts:
            return font_name, True

    if STRICT_PLOS_FONT:
        raise RuntimeError(
            "Arial or Times is not installed.\n"
            "Install one of these fonts before generating the "
            "final publication figure."
        )

    fallback_font = "DejaVu Sans"

    if fallback_font not in installed_fonts:
        raise RuntimeError(
            "No suitable fallback font was found."
        )

    warnings.warn(
        "Arial or Times is not installed. "
        "DejaVu Sans will be used temporarily. "
        "Regenerate the final publication figure with Arial or Times.",
        stacklevel=2,
    )

    return fallback_font, False


FIGURE_FONT, FONT_IS_APPROVED = select_figure_font()

plt.rcParams.update(
    {
        "font.family": FIGURE_FONT,
        "font.size": 8,
        "axes.labelsize": 8,
        "xtick.labelsize": 8,
        "ytick.labelsize": 8,
        "figure.facecolor": "white",
        "axes.facecolor": "white",

        # Embed TrueType fonts in EPS.
        "ps.fonttype": 42,
        "pdf.fonttype": 42,
    }
)

print(f"Using font: {FIGURE_FONT}")

if FONT_IS_APPROVED:
    print("Font status: PLOS-approved")
else:
    print(
        "Font status: temporary drafting font; "
        "regenerate the final version with Arial or Times."
    )


# ============================================================
# 4. Excel columns
# ============================================================

DIMENSION_COLUMN = "DQ Dimension"
THEME_COLUMN = "Issue theme"

ELEMENT_COLUMNS = [
    "Diagnostic codes",
    "Demographic fields",
    "Temporal / date fields",
    "Procedure & service codes",
    "Pharmacy / drug records",
    "Patient & record linkage fields",
]


# ============================================================
# 5. Full display labels
# ============================================================

# These labels appear diagonally above the matrix.
ELEMENT_DISPLAY_LABELS = {
    "Diagnostic codes":
        "Diagnostic codes",

    "Demographic fields":
        "Demographic fields",

    "Temporal / date fields":
        "Temporal / date fields",

    "Procedure & service codes":
        "Procedure & service codes",

    "Pharmacy / drug records":
        "Pharmacy / drug records",

    "Patient & record linkage fields":
        "Patient & record linkage fields",
}


# ============================================================
# 6. DQ dimensions and colours
# ============================================================

DQD_ORDER = [
    "Completeness",
    "Conformance",
    "Plausibility",
    "Consistency",
    "Accuracy / Correctness / Validity",
    "Concordance",
    "Currency",
    "Uniqueness",
    "Temporal Relationships",
]

DQD_COLORS = {
    "Completeness": "#1B9E8A",
    "Conformance": "#2E7FB8",
    "Plausibility": "#8E44AD",
    "Consistency": "#B8860B",
    "Accuracy / Correctness / Validity": "#D62728",
    "Concordance": "#C2185B",
    "Currency": "#7F7F7F",
    "Uniqueness": "#2E8B57",
    "Temporal Relationships": "#555555",
}

DQD_DISPLAY_LABELS = {
    "Completeness":
        "Completeness",

    "Conformance":
        "Conformance",

    "Plausibility":
        "Plausibility",

    "Consistency":
        "Consistency",

    "Accuracy / Correctness / Validity":
        "Accuracy /\ncorrectness /\nvalidity",

    "Concordance":
        "Concordance",

    "Currency":
        "Currency",

    "Uniqueness":
        "Uniqueness",

    "Temporal Relationships":
        "Temporal\nrelationships",
}


# ============================================================
# 7. Theme display labels
# ============================================================

THEME_DISPLAY_LABELS = {
    "Atemporal / temporal plausibility":
        "Atemporal / temporal\nplausibility",

    "Accuracy / Correctness":
        "Accuracy /\ncorrectness",

    "Relevance / fit-for-purpose validity":
        "Relevance / fit-for-purpose\nvalidity",

    "Temporal / system-transition inconsistency":
        "Temporal / system-transition\ninconsistency",

    "Cross-site / linkage consistency":
        "Cross-site / linkage\nconsistency",
}


# ============================================================
# 8. Locate the Excel workbook
# ============================================================

def locate_input_file(filename: str) -> Path:
    """
    Find the Excel workbook in the script directory
    or its parent directory.
    """

    direct_path = BASE_DIR / filename

    if direct_path.exists():
        return direct_path

    matches: list[Path] = []

    for search_root in [BASE_DIR, BASE_DIR.parent]:
        matches.extend(
            search_root.rglob(filename)
        )

    unique_matches = list(
        dict.fromkeys(
            path.resolve()
            for path in matches
        )
    )

    if not unique_matches:
        raise FileNotFoundError(
            f"Could not find {filename!r}.\n"
            "Place the Excel file in the same folder as this script."
        )

    if len(unique_matches) > 1:
        print(
            "Several matching Excel files were found. "
            "Using the first one:"
        )

        for match in unique_matches:
            print(f"  {match}")

    return unique_matches[0]


EXCEL_PATH = locate_input_file(
    EXCEL_FILENAME
)

print(f"Using workbook: {EXCEL_PATH}")


# ============================================================
# 9. Load and validate the workbook
# ============================================================

df = pd.read_excel(
    EXCEL_PATH,
    sheet_name=SHEET_NAME,
)

df.columns = [
    str(column).strip()
    for column in df.columns
]

required_columns = [
    DIMENSION_COLUMN,
    THEME_COLUMN,
    *ELEMENT_COLUMNS,
]

missing_columns = [
    column
    for column in required_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        "The following required columns are missing:\n- "
        + "\n- ".join(missing_columns)
    )

# Remove totals rows and incomplete rows.
df = df[
    df[DIMENSION_COLUMN].notna()
    & df[THEME_COLUMN].notna()
].copy()

df[DIMENSION_COLUMN] = (
    df[DIMENSION_COLUMN]
    .astype(str)
    .str.strip()
)

df[THEME_COLUMN] = (
    df[THEME_COLUMN]
    .astype(str)
    .str.strip()
)

unknown_dimensions = sorted(
    set(df[DIMENSION_COLUMN])
    - set(DQD_ORDER)
)

if unknown_dimensions:
    print(
        "\nWarning: rows with the following unrecognised "
        "DQ dimensions will be excluded:"
    )

    for dimension in unknown_dimensions:
        print(f"  {dimension}")

df = df[
    df[DIMENSION_COLUMN].isin(DQD_ORDER)
].copy()

if df.empty:
    raise ValueError(
        "No valid issue-theme rows remained after filtering."
    )


# ============================================================
# 10. Convert Excel marks to relationships
# ============================================================

def is_marked(value) -> bool:
    """
    Return True when a cell represents a relationship.

    Accepted positive values:
    X, x, ✓, ✔, 1, yes, true
    """

    if pd.isna(value):
        return False

    normalised_value = (
        str(value)
        .strip()
        .casefold()
    )

    positive_values = {
        "x",
        "✓",
        "✔",
        "1",
        "yes",
        "true",
    }

    return normalised_value in positive_values


dimension_rank = {
    dimension: index
    for index, dimension in enumerate(DQD_ORDER)
}

df["_dimension_rank"] = (
    df[DIMENSION_COLUMN]
    .map(dimension_rank)
)

# Preserve the original issue-theme order within each dimension.
df = (
    df.sort_values(
        "_dimension_rank",
        kind="stable",
    )
    .reset_index(drop=True)
)

records = []

for _, row in df.iterrows():
    linked_elements = {
        column
        for column in ELEMENT_COLUMNS
        if is_marked(row[column])
    }

    records.append(
        {
            "dimension": row[DIMENSION_COLUMN],
            "theme": row[THEME_COLUMN],
            "elements": linked_elements,
        }
    )

number_of_rows = len(records)
number_of_columns = len(ELEMENT_COLUMNS)

print(
    f"\nLoaded {number_of_rows} issue themes and "
    f"{number_of_columns} data-element categories."
)

themes_without_links = [
    record["theme"]
    for record in records
    if not record["elements"]
]

if themes_without_links:
    print(
        "\nWarning: the following themes have no marked links:"
    )

    for theme in themes_without_links:
        print(f"  {theme}")


# ============================================================
# 11. Determine DQ-dimension boundaries
# ============================================================

group_boundaries: dict[str, list[int]] = {}

for row_index, record in enumerate(records):
    dimension = record["dimension"]

    if dimension not in group_boundaries:
        group_boundaries[dimension] = [
            row_index,
            row_index,
        ]
    else:
        group_boundaries[dimension][1] = row_index


# ============================================================
# 12. Create the figure
# ============================================================

fig = plt.figure(
    figsize=FIGURE_SIZE,
    dpi=DPI,
    facecolor="white",
)

# Manual axis placement preserves the exact output dimensions.
ax = fig.add_axes(
    [0.025, 0.025, 0.95, 0.95]
)

ax.set_facecolor("white")

# First record appears at the top.
y_positions = {
    row_index: number_of_rows - row_index
    for row_index in range(number_of_rows)
}


# ============================================================
# 13. Full diagonal matrix column headings
# ============================================================

HEADER_Y_POSITION = number_of_rows + 0.72
HEADER_ROTATION = 22

for column_index, column_name in enumerate(
    ELEMENT_COLUMNS
):
    display_label = ELEMENT_DISPLAY_LABELS[
        column_name
    ]

    ax.text(
        column_index - 0.08,
        HEADER_Y_POSITION,
        display_label,
        ha="left",
        va="bottom",
        fontsize=8,
        color="#333333",
        rotation=HEADER_ROTATION,
        rotation_mode="anchor",
        clip_on=False,
    )


# ============================================================
# 14. Matrix grid lines
# ============================================================

# Vertical grid lines extend slightly above the matrix,
# matching the reference design.
for column_index in range(number_of_columns):
    ax.plot(
        [
            column_index,
            column_index,
        ],
        [
            0.5,
            number_of_rows + 1.75,
        ],
        color="#E0E0E0",
        linewidth=0.55,
        zorder=0,
    )

# Horizontal grid lines.
for row_index in range(number_of_rows):
    y_position = y_positions[row_index]

    ax.plot(
        [
            -0.5,
            number_of_columns - 0.5,
        ],
        [
            y_position,
            y_position,
        ],
        color="#ECECEC",
        linewidth=0.55,
        zorder=0,
    )


# ============================================================
# 15. Matrix border
# ============================================================

matrix_border = Rectangle(
    (-0.5, 0.5),
    width=number_of_columns,
    height=number_of_rows,
    fill=False,
    edgecolor="#808080",
    linewidth=0.75,
    zorder=2,
)

ax.add_patch(
    matrix_border
)


# ============================================================
# 16. Theme labels and stars
# ============================================================

for row_index, record in enumerate(records):
    y_position = y_positions[row_index]

    dimension = record["dimension"]
    theme = record["theme"]

    colour = DQD_COLORS.get(
        dimension,
        "#444444",
    )

    theme_label = THEME_DISPLAY_LABELS.get(
        theme,
        theme,
    )

    ax.text(
        -0.72,
        y_position,
        theme_label,
        ha="right",
        va="center",
        fontsize=8,
        color="#222222",
        linespacing=0.95,
    )

    for column_index, column_name in enumerate(
        ELEMENT_COLUMNS
    ):
        if column_name in record["elements"]:
            ax.scatter(
                column_index,
                y_position,
                s=STAR_SIZE,
                marker="*",
                color=colour,
                edgecolors="white",
                linewidths=0.35,
                zorder=3,
            )


# ============================================================
# 17. DQ-dimension brackets and labels
# ============================================================

dimension_bracket_x = -4.05
dimension_label_x = -4.27

for dimension in DQD_ORDER:
    if dimension not in group_boundaries:
        continue

    first_index, last_index = (
        group_boundaries[dimension]
    )

    y_top = (
        y_positions[first_index]
        + 0.44
    )

    y_bottom = (
        y_positions[last_index]
        - 0.44
    )

    colour = DQD_COLORS.get(
        dimension,
        "#444444",
    )

    # Vertical bracket.
    ax.plot(
        [
            dimension_bracket_x,
            dimension_bracket_x,
        ],
        [
            y_bottom,
            y_top,
        ],
        color=colour,
        linewidth=2.6,
        solid_capstyle="round",
        zorder=2,
    )

    # Upper bracket end.
    ax.plot(
        [
            dimension_bracket_x,
            dimension_bracket_x + 0.13,
        ],
        [
            y_top,
            y_top,
        ],
        color=colour,
        linewidth=1.1,
        zorder=2,
    )

    # Lower bracket end.
    ax.plot(
        [
            dimension_bracket_x,
            dimension_bracket_x + 0.13,
        ],
        [
            y_bottom,
            y_bottom,
        ],
        color=colour,
        linewidth=1.1,
        zorder=2,
    )

    display_dimension = DQD_DISPLAY_LABELS.get(
        dimension,
        dimension,
    )

    ax.text(
        dimension_label_x,
        (y_top + y_bottom) / 2,
        display_dimension,
        ha="right",
        va="center",
        fontsize=8,
        fontweight="bold",
        color=colour,
        linespacing=0.95,
    )


# ============================================================
# 18. Group headings
# ============================================================

GROUP_HEADING_Y = number_of_rows + 2.35

ax.text(
    -5.35,
    GROUP_HEADING_Y,
    "DQ Dimension",
    fontsize=8,
    fontweight="bold",
    fontstyle="italic",
    color="#555555",
    ha="left",
    va="center",
)

ax.text(
    -0.72,
    GROUP_HEADING_Y,
    "Issue Theme",
    fontsize=8,
    fontweight="bold",
    fontstyle="italic",
    color="#555555",
    ha="right",
    va="center",
)


# ============================================================
# 19. Axis limits and appearance
# ============================================================

# Additional space on the right ensures that the final
# diagonal heading is not clipped.
ax.set_xlim(
    -5.85,
    number_of_columns + 1.50,
)

ax.set_ylim(
    0.15,
    number_of_rows + 2.85,
)

ax.set_xticks([])
ax.set_yticks([])

for spine in ax.spines.values():
    spine.set_visible(False)


# ============================================================
# 20. TIFF validation
# ============================================================

def validate_tiff(output_path: Path) -> None:
    """
    Validate the TIFF against the configured PLOS limits.
    """

    with Image.open(output_path) as image:
        width_pixels, height_pixels = image.size
        image_mode = image.mode
        recorded_dpi = image.info.get("dpi")
        compression = image.info.get("compression")
        number_of_frames = getattr(
            image,
            "n_frames",
            1,
        )

    file_size_bytes = output_path.stat().st_size

    errors = []

    if width_pixels != EXPECTED_WIDTH_PIXELS:
        errors.append(
            f"Expected width {EXPECTED_WIDTH_PIXELS} pixels, "
            f"but found {width_pixels}."
        )

    if height_pixels != EXPECTED_HEIGHT_PIXELS:
        errors.append(
            f"Expected height {EXPECTED_HEIGHT_PIXELS} pixels, "
            f"but found {height_pixels}."
        )

    if not (
        MIN_WIDTH_PIXELS
        <= width_pixels
        <= MAX_WIDTH_PIXELS
    ):
        errors.append(
            f"Width {width_pixels} is outside the permitted "
            f"range of {MIN_WIDTH_PIXELS}–"
            f"{MAX_WIDTH_PIXELS} pixels."
        )

    if height_pixels > MAX_HEIGHT_PIXELS:
        errors.append(
            f"Height {height_pixels} exceeds the maximum "
            f"of {MAX_HEIGHT_PIXELS} pixels."
        )

    if image_mode != "RGB":
        errors.append(
            f"Image mode is {image_mode}; expected RGB."
        )

    if number_of_frames != 1:
        errors.append(
            f"The TIFF contains {number_of_frames} pages; "
            "only one page is permitted."
        )

    if compression != "tiff_lzw":
        errors.append(
            f"Compression is {compression}; expected tiff_lzw."
        )

    if file_size_bytes > MAX_FILE_SIZE_BYTES:
        errors.append(
            f"File size is "
            f"{file_size_bytes / (1024 * 1024):.2f} MB; "
            "the maximum is 10 MB."
        )

    if recorded_dpi:
        horizontal_dpi = float(recorded_dpi[0])
        vertical_dpi = float(recorded_dpi[1])

        if not (
            299 <= horizontal_dpi <= 301
            and 299 <= vertical_dpi <= 301
        ):
            errors.append(
                f"Recorded resolution is {recorded_dpi}; "
                "expected approximately 300 dpi."
            )

    if errors:
        raise ValueError(
            f"\n{output_path.name} failed validation:\n- "
            + "\n- ".join(errors)
        )

    print(
        f"\nValidated {output_path.name}"
        f"\n  Dimensions: "
        f"{width_pixels} × {height_pixels} pixels"
        f"\n  Resolution: {recorded_dpi}"
        f"\n  Colour mode: {image_mode}"
        f"\n  Compression: {compression}"
        f"\n  Pages: {number_of_frames}"
        f"\n  File size: "
        f"{file_size_bytes / (1024 * 1024):.2f} MB"
    )


# ============================================================
# 21. Export figure
# ============================================================

def export_figure() -> None:
    """
    Export TIFF, EPS, and PNG versions.
    """

    temporary_png = (
        BASE_DIR
        / "_Fig5_temporary_render.png"
    )

    # Do not use bbox_inches="tight".
    # It would change the final pixel dimensions.
    fig.savefig(
        temporary_png,
        format="png",
        dpi=DPI,
        facecolor="white",
        transparent=False,
    )

    # Create a flattened RGB TIFF with LZW compression.
    with Image.open(temporary_png) as image:
        flattened_image = image.convert("RGB")

        flattened_image.save(
            OUTPUT_TIFF,
            format="TIFF",
            compression="tiff_lzw",
            dpi=(DPI, DPI),
        )

    # Preserve the exact PNG rendering as a preview.
    temporary_png.replace(
        OUTPUT_PNG
    )

    # Create an EPS vector alternative.
    fig.savefig(
        OUTPUT_EPS,
        format="eps",
        facecolor="white",
        transparent=False,
    )

    validate_tiff(
        OUTPUT_TIFF
    )

    eps_size_bytes = OUTPUT_EPS.stat().st_size

    if eps_size_bytes > MAX_FILE_SIZE_BYTES:
        print(
            f"\nWarning: {OUTPUT_EPS.name} is "
            f"{eps_size_bytes / (1024 * 1024):.2f} MB."
        )

    print(f"\nSaved TIFF: {OUTPUT_TIFF}")
    print(f"Saved EPS:  {OUTPUT_EPS}")
    print(f"Saved PNG:  {OUTPUT_PNG}")


export_figure()
plt.close(fig)

print("\nFigure 5 was generated successfully.")