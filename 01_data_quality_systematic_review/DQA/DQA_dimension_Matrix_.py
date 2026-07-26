from pathlib import Path
import re

import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np
import pandas as pd
from PIL import Image


# ============================================================
# 1. File settings
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "DQ_Dimensions_Matrix.xlsx"
SHEET_NAME = "DQ_Dimensions_Matrix"

OUT_DIM_PATH = BASE_DIR / "Fig3.tif"
OUT_METH_PATH = BASE_DIR / "Fig4.tif"


# ============================================================
# 2. PLOS figure settings
# ============================================================

DPI = 300

# PLOS full-page width:
# 7.5 inches × 300 dpi = 2250 pixels
#
# Height:
# 8.5 inches × 300 dpi = 2550 pixels
#
# PLOS maximum height at 300 dpi = 2625 pixels
FIGURE_SIZE = (7.5, 8.5)

MIN_WIDTH_PIXELS = 789
MAX_WIDTH_PIXELS = 2250
MAX_HEIGHT_PIXELS = 2625
MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024

SHOW_FIGURES = False
SORT_BY_YEAR = True

EXPECTED_NUMBER_OF_STUDIES = 19

EXPECTED_DIM_COUNTS = np.array(
    [12, 9, 8, 6, 5, 3, 2, 1, 1],
    dtype=int,
)

EXPECTED_METH_COUNTS = np.array(
    [15, 11, 11, 3],
    dtype=int,
)


# ============================================================
# 3. Font configuration
# ============================================================

def select_plos_font() -> str:
    """
    Select an installed font accepted by PLOS.

    PLOS permits Arial, Times, or Symbol within figures.
    """

    installed_fonts = {
        font.name for font in font_manager.fontManager.ttflist
    }

    accepted_fonts = [
        "Arial",
        "Times New Roman",
        "Times",
        "DejaVu Sans",  # temporary server fallback
    ]

    for font_name in accepted_fonts:
        if font_name in installed_fonts:
            return font_name

    raise RuntimeError(
        "No PLOS-compatible font was found.\n"
        "Install Arial, Times New Roman, or Times and rerun the script."
    )


PLOS_FONT = select_plos_font()

plt.rcParams.update(
    {
        "font.family": PLOS_FONT,
        "font.size": 9,
        "axes.labelsize": 10,
        "axes.titlesize": 11,
        "xtick.labelsize": 8,
        "ytick.labelsize": 8,
        "legend.fontsize": 8,
        "figure.facecolor": "white",
        "axes.facecolor": "white",

        # Preserve fonts when exporting EPS.
        "ps.fonttype": 42,
        "pdf.fonttype": 42,
    }
)

print(f"Using figure font: {PLOS_FONT}")


# ============================================================
# 4. Excel layout
# ============================================================

# Zero-based column index containing the study label.
AUTHOR_COL = 1

# Excel rows 4–22 contain the 19 included studies.
# Python iloc excludes the final row.
DATA_START_ROW = 3
DATA_END_ROW = 22

# Excel columns D–L contain DQ dimensions.
DIM_COLS = [3, 4, 5, 6, 7, 8, 9, 10, 11]

# Excel columns M–P contain DQA methods.
METH_COLS = [12, 13, 14, 15]


# ============================================================
# 5. Labels
# ============================================================

DIM_LABELS = [
    "Completeness",
    "Conformance",
    "Plausibility",
    "Consistency",
    "Accuracy /\ncorrectness /\nvalidity",
    "Concordance",
    "Currency",
    "Uniqueness",
    "Temporal\nrelationships",
]

METH_LABELS = [
    "Element-level\nassessment",
    "Cross-database or\nreference-based\ncomparison",
    "Rule-based\nchecks",
    "Structured DQA\nframework or\nprogramme",
]


# ============================================================
# 6. Colours
# ============================================================

DIM_COLORS = [
    "#2A9D8F",  # Completeness
    "#2F80C1",  # Conformance
    "#8E44AD",  # Plausibility
    "#C58B00",  # Consistency
    "#E03131",  # Accuracy / correctness / validity
    "#D81B60",  # Concordance
    "#7A7A7A",  # Currency
    "#2E8B57",  # Uniqueness
    "#4D4D4D",  # Temporal relationships
]

METH_COLORS = [
    "#3AAFA9",  # Element-level assessment
    "#5DADE2",  # Cross-database/reference-based comparison
    "#7E57C2",  # Rule-based checks
    "#E67E22",  # Structured framework/programme
]


# ============================================================
# 7. Load and validate the workbook
# ============================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{EXCEL_PATH}\n\n"
        "Place DQ_Dimensions_Matrix.xlsx in the same directory "
        "as this Python script."
    )

df_raw = pd.read_excel(
    EXCEL_PATH,
    sheet_name=SHEET_NAME,
    header=None,
)

required_last_column = max(DIM_COLS + METH_COLS)

if df_raw.shape[1] <= required_last_column:
    raise ValueError(
        "The workbook structure does not match the expected layout.\n"
        f"Expected at least {required_last_column + 1} columns, "
        f"but found {df_raw.shape[1]}."
    )

data = df_raw.iloc[DATA_START_ROW:DATA_END_ROW].copy()
data.columns = range(data.shape[1])

authors = (
    data[AUTHOR_COL]
    .fillna("")
    .astype(str)
    .str.strip()
    .tolist()
)

if len(authors) != EXPECTED_NUMBER_OF_STUDIES:
    raise ValueError(
        f"Expected {EXPECTED_NUMBER_OF_STUDIES} studies, "
        f"but found {len(authors)}.\n"
        "Check DATA_START_ROW and DATA_END_ROW."
    )

empty_author_rows = [
    index + DATA_START_ROW + 1
    for index, author in enumerate(authors)
    if not author
]

if empty_author_rows:
    raise ValueError(
        "The following Excel rows have no study label: "
        + ", ".join(map(str, empty_author_rows))
    )


# ============================================================
# 8. Convert Excel marks to binary values
# ============================================================

def to_binary(value) -> int:
    """
    Convert common positive indicators to 1.

    Accepted positive values:
    X, x, ✓, ✔, 1, yes, true

    Blank cells and all other values become 0.
    """

    if pd.isna(value):
        return 0

    text = str(value).strip().casefold()

    positive_values = {
        "x",
        "✓",
        "✔",
        "1",
        "yes",
        "true",
    }

    return int(text in positive_values)


dim_matrix = (
    data[DIM_COLS]
    .apply(lambda column: column.map(to_binary))
    .to_numpy(dtype=int)
)

meth_matrix = (
    data[METH_COLS]
    .apply(lambda column: column.map(to_binary))
    .to_numpy(dtype=int)
)


# ============================================================
# 9. Extract publication years and sort studies
# ============================================================

def extract_year(label: str) -> int:
    """Extract the final four-digit publication year from a label."""

    matches = re.findall(
        r"\b(?:19|20)\d{2}\b",
        str(label),
    )

    if not matches:
        raise ValueError(
            "Could not identify a four-digit publication year in "
            f"the study label: {label!r}"
        )

    return int(matches[-1])


if SORT_BY_YEAR:
    years = np.array(
        [extract_year(author) for author in authors],
        dtype=int,
    )

    # Stable ascending order: oldest study first.
    order = np.argsort(years, kind="stable")

    authors = [authors[index] for index in order]
    dim_matrix = dim_matrix[order, :]
    meth_matrix = meth_matrix[order, :]


# ============================================================
# 10. Validate the study totals
# ============================================================

dimension_totals = dim_matrix.sum(axis=0)
method_totals = meth_matrix.sum(axis=0)

print(f"\nStudies loaded: {len(authors)}")

print("\nData quality dimensions")
for label, count in zip(DIM_LABELS, dimension_totals):
    clean_label = label.replace("\n", " ")
    print(f"{clean_label:45s} {int(count):>2}")

print("\nData quality assessment methods")
for label, count in zip(METH_LABELS, method_totals):
    clean_label = label.replace("\n", " ")
    print(f"{clean_label:45s} {int(count):>2}")

if not np.array_equal(
    dimension_totals,
    EXPECTED_DIM_COUNTS,
):
    raise ValueError(
        "\nThe DQ-dimension totals do not match the validated counts.\n"
        f"Expected: {EXPECTED_DIM_COUNTS.tolist()}\n"
        f"Found:    {dimension_totals.tolist()}\n\n"
        "Check the DQ-dimension marks in the Excel matrix."
    )

if not np.array_equal(
    method_totals,
    EXPECTED_METH_COUNTS,
):
    raise ValueError(
        "\nThe DQA-method totals do not match the validated counts.\n"
        f"Expected: {EXPECTED_METH_COUNTS.tolist()}\n"
        f"Found:    {method_totals.tolist()}\n\n"
        "Check the assessment-method marks in the Excel matrix."
    )


# ============================================================
# 11. Figure export and validation
# ============================================================

def validate_tiff(output_path: Path) -> None:
    """Validate a TIFF file against the main PLOS specifications."""

    with Image.open(output_path) as image:
        width_pixels, height_pixels = image.size
        image_mode = image.mode
        recorded_dpi = image.info.get("dpi")
        compression = image.info.get("compression")

    file_size_bytes = output_path.stat().st_size

    errors = []

    if not MIN_WIDTH_PIXELS <= width_pixels <= MAX_WIDTH_PIXELS:
        errors.append(
            f"Width is {width_pixels} pixels; "
            f"required range is {MIN_WIDTH_PIXELS}–"
            f"{MAX_WIDTH_PIXELS} pixels."
        )

    if height_pixels > MAX_HEIGHT_PIXELS:
        errors.append(
            f"Height is {height_pixels} pixels; "
            f"maximum permitted height is {MAX_HEIGHT_PIXELS} pixels."
        )

    if image_mode != "RGB":
        errors.append(
            f"Colour mode is {image_mode}; expected RGB."
        )

    if file_size_bytes > MAX_FILE_SIZE_BYTES:
        errors.append(
            f"File size is "
            f"{file_size_bytes / (1024 * 1024):.2f} MB; "
            "the maximum is 10 MB."
        )

    if errors:
        raise ValueError(
            f"\n{output_path.name} does not meet the configured "
            "PLOS requirements:\n- "
            + "\n- ".join(errors)
        )

    print(
        f"\nValidated: {output_path.name}\n"
        f"  Dimensions: {width_pixels} × {height_pixels} pixels\n"
        f"  Resolution: {recorded_dpi}\n"
        f"  Colour mode: {image_mode}\n"
        f"  Compression: {compression}\n"
        f"  File size: "
        f"{file_size_bytes / (1024 * 1024):.2f} MB"
    )


def export_figure(
    fig: plt.Figure,
    output_path: Path,
) -> None:
    """
    Export a figure as TIFF, EPS, and PNG.

    TIFF is flattened, RGB, 300 dpi, and LZW-compressed.
    EPS is generated as an accepted vector alternative.
    PNG is created only as a preview.
    """

    temporary_png = output_path.with_name(
        f"{output_path.stem}_temporary.png"
    )

    # Render at the exact figure size and resolution.
    # Do not use bbox_inches="tight", as it changes dimensions.
    fig.savefig(
        temporary_png,
        format="png",
        dpi=DPI,
        facecolor="white",
        transparent=False,
    )

    # Convert the rendered image to flattened RGB TIFF
    # using LZW compression.
    with Image.open(temporary_png) as image:
        rgb_image = image.convert("RGB")

        rgb_image.save(
            output_path,
            format="TIFF",
            compression="tiff_lzw",
            dpi=(DPI, DPI),
        )

    temporary_png.unlink(missing_ok=True)

    # Accepted vector output.
    eps_path = output_path.with_suffix(".eps")

    fig.savefig(
        eps_path,
        format="eps",
        facecolor="white",
        transparent=False,
    )

    # Preview output for GitHub or local inspection.
    preview_path = output_path.with_suffix(".png")

    fig.savefig(
        preview_path,
        format="png",
        dpi=DPI,
        facecolor="white",
        transparent=False,
    )

    validate_tiff(output_path)

    if eps_path.stat().st_size > MAX_FILE_SIZE_BYTES:
        print(
            f"Warning: {eps_path.name} is larger than 10 MB."
        )

    print(f"Saved TIFF: {output_path}")
    print(f"Saved EPS:  {eps_path}")
    print(f"Saved PNG:  {preview_path}")


# ============================================================
# 12. Shared figure-drawing function
# ============================================================

def draw_figure(
    matrix: np.ndarray,
    labels: list[str],
    colors: list[str],
    output_path: Path,
    star_size: float,
    x_label_fontsize: float,
    x_label_rotation: float,
    x_label_pad: float,
    y_label_fontsize: float = 8,
) -> None:
    """
    Draw a bar chart above a study-by-category star matrix.
    """

    number_studies, number_columns = matrix.shape

    if number_columns != len(labels):
        raise ValueError(
            f"The matrix contains {number_columns} columns, "
            f"but {len(labels)} labels were supplied."
        )

    if number_columns != len(colors):
        raise ValueError(
            f"The matrix contains {number_columns} columns, "
            f"but {len(colors)} colours were supplied."
        )

    x_positions = np.arange(number_columns)
    column_totals = matrix.sum(axis=0)

    fig, (bar_axis, matrix_axis) = plt.subplots(
        nrows=2,
        ncols=1,
        sharex=True,
        figsize=FIGURE_SIZE,
        gridspec_kw={
            "height_ratios": [2.1, 7.0],
            "hspace": 0.03,
        },
    )

    fig.patch.set_facecolor("white")

    # Explicit internal margins preserve the exact exported dimensions.
    fig.subplots_adjust(
        left=0.35,
        right=0.98,
        top=0.97,
        bottom=0.22,
        hspace=0.03,
    )

    # --------------------------------------------------------
    # Upper bar chart
    # --------------------------------------------------------

    bars = bar_axis.bar(
        x_positions,
        column_totals,
        width=0.60,
        color=colors,
        edgecolor="white",
        linewidth=0.8,
        zorder=2,
    )

    for bar, total, color in zip(
        bars,
        column_totals,
        colors,
    ):
        bar_axis.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.22,
            str(int(total)),
            ha="center",
            va="bottom",
            fontsize=9,
            fontweight="bold",
            color=color,
        )

    bar_axis.set_ylabel(
        "Studies (n)",
        fontsize=10,
        labelpad=8,
    )

    bar_axis.set_ylim(
        0,
        max(column_totals) + 3.2,
    )

    bar_axis.tick_params(
        axis="y",
        labelsize=8,
        length=3,
    )

    bar_axis.tick_params(
        axis="x",
        bottom=False,
        labelbottom=False,
    )

    bar_axis.spines["top"].set_visible(False)
    bar_axis.spines["right"].set_visible(False)
    bar_axis.spines["left"].set_color("#BEBEBE")
    bar_axis.spines["bottom"].set_color("#BEBEBE")

    # --------------------------------------------------------
    # Lower star matrix
    # --------------------------------------------------------

    for row_index in range(number_studies):
        for column_index in range(number_columns):
            if matrix[row_index, column_index] == 1:
                matrix_axis.scatter(
                    x_positions[column_index],
                    row_index + 1,
                    marker="*",
                    s=star_size,
                    color=colors[column_index],
                    edgecolors="white",
                    linewidths=0.3,
                    zorder=3,
                )

    for x_position in x_positions:
        matrix_axis.axvline(
            x_position,
            color="#D8D8D8",
            linestyle=":",
            linewidth=0.7,
            zorder=0,
        )

    for y_position in range(1, number_studies + 1):
        matrix_axis.axhline(
            y_position,
            color="#E3E3E3",
            linestyle=":",
            linewidth=0.7,
            zorder=0,
        )

    matrix_axis.set_xlim(
        -0.75,
        number_columns - 0.25,
    )

    matrix_axis.set_ylim(
        0.3,
        number_studies + 0.7,
    )

    matrix_axis.set_yticks(
        np.arange(1, number_studies + 1)
    )

    matrix_axis.set_yticklabels(
        authors,
        fontsize=y_label_fontsize,
    )

    # The first, oldest publication appears at the top.
    matrix_axis.invert_yaxis()

    matrix_axis.set_ylabel(
        "Publications (first author and year)",
        fontsize=9,
        labelpad=9,
    )

    matrix_axis.set_xticks(x_positions)

    matrix_axis.set_xticklabels(
        labels,
        fontsize=x_label_fontsize,
        rotation=x_label_rotation,
        ha="right",
        rotation_mode="anchor",
    )

    matrix_axis.tick_params(
        axis="x",
        pad=x_label_pad,
        length=0,
    )

    matrix_axis.tick_params(
        axis="y",
        length=0,
    )

    for spine in matrix_axis.spines.values():
        spine.set_visible(True)
        spine.set_color("#808080")
        spine.set_linewidth(0.7)

    export_figure(
        fig=fig,
        output_path=output_path,
    )

    if SHOW_FIGURES:
        plt.show()

    plt.close(fig)


# ============================================================
# 13. Fig 3: Data quality dimensions
# ============================================================

draw_figure(
    matrix=dim_matrix,
    labels=DIM_LABELS,
    colors=DIM_COLORS,
    output_path=OUT_DIM_PATH,
    star_size=115,
    x_label_fontsize=8,
    x_label_rotation=34,
    x_label_pad=6,
    y_label_fontsize=8,
)


# ============================================================
# 14. Fig 4: Data quality assessment methods
# ============================================================

draw_figure(
    matrix=meth_matrix,
    labels=METH_LABELS,
    colors=METH_COLORS,
    output_path=OUT_METH_PATH,
    star_size=125,
    x_label_fontsize=8,
    x_label_rotation=30,
    x_label_pad=6,
    y_label_fontsize=8,
)


print("\nAll figures were generated successfully.")