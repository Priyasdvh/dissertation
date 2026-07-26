from pathlib import Path
import re
import sys

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Rectangle


# ============================================================
# 1. File settings
# ============================================================

BASE_DIR = Path(__file__).resolve().parent


DEFAULT_EXCEL_PATH = BASE_DIR / "Additional_file_1_study_characteristics.xlsx"

# Ignore Jupyter/IPython arguments such as --f=...
valid_args = [
    arg for arg in sys.argv[1:]
    if not arg.startswith("-")
]

EXCEL_PATH = (
    Path(valid_args[0]).expanduser().resolve()
    if valid_args
    else DEFAULT_EXCEL_PATH
)

SHEET_NAME = "study_characteristics"

OUTPUT_PATH = BASE_DIR / "Figure_2_Study_Characteristics.png"


# ============================================================
# 2. Load and validate the Excel sheet
# ============================================================

if not EXCEL_PATH.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{EXCEL_PATH}\n\n"
        "Place the Excel file in the same folder as this script, "
        "or provide its path when running the script."
    )

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

# Remove accidental spaces from column names and empty rows.
df.columns = df.columns.astype(str).str.strip()
df = df.dropna(how="all").copy()

required_columns = [
    "Year",
    "Country",
    "Healthcare domain",
    "Study design",
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        f"Missing required columns: {missing_columns}\n"
        f"Available columns: {df.columns.tolist()}"
    )

number_of_studies = len(df)
print(f"Studies loaded: {number_of_studies}")


# ============================================================
# 3. Helper functions
# ============================================================

def normalise_text(value):
    """Standardise spaces, dashes and capitalisation for matching."""
    text = str(value).strip().casefold()
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text


def study_design_group(value):
    """
    Convert the study-design text to one of the five figure categories.

    Mixed methods is checked before Methodological because the word
    'methods' contains 'method'.
    """
    text = normalise_text(value)

    if "mixed" in text:
        return "Mixed methods"

    if "cross-database" in text or "cross database" in text or "comparison" in text:
        return "Cross-database\ncomparison"

    if "empirical" in text or "validation" in text or "rule-based" in text:
        return "Empirical\nvalidation"

    if (
        "observational" in text
        or "retrospective" in text
        or "longitudinal" in text
        or "cross-sectional" in text
        or "cross sectional" in text
    ):
        return "Observational/\nretrospective"

    if "methodological" in text or "method" in text:
        return "Methodological"

    return "Unmapped"


def year_group(value):
    """Assign publication year to the five prespecified periods."""
    try:
        year = int(float(value))
    except (TypeError, ValueError):
        return "Unmapped"

    if year <= 2010:
        return "≤2010"

    if 2015 <= year <= 2017:
        return "2015–2017"

    if 2018 <= year <= 2020:
        return "2018–2020"

    if 2021 <= year <= 2023:
        return "2021–2023"

    if 2024 <= year <= 2025:
        return "2024–2025"

    return "Unmapped"


def country_group(value):
    """Group Belgium and Poland as other European countries."""
    text = normalise_text(value)

    country_mapping = {
        "usa": "USA",
        "us": "USA",
        "united states": "USA",
        "united states of america": "USA",
        "germany": "Germany",
        "canada": "Canada",
        "taiwan": "Taiwan",
        "australia": "Australia",
        "japan": "Japan",
        "belgium": "Other European\ncountries",
        "poland": "Other European\ncountries",
    }

    return country_mapping.get(text, "Unmapped")


def healthcare_domain_group(value):
    """
    Convert detailed healthcare domains to the five display categories.
    Maternal-infant and substance-use/OUD studies are grouped as Others.
    """
    text = normalise_text(value)

    if "general" in text or "multi-domain" in text or "multi domain" in text:
        return "General /\nMulti-domain"

    if "cardiovascular" in text or "heart" in text:
        return "Cardiovascular"

    if "chronic" in text or "diabetes" in text or "metabolic" in text:
        return "Chronic disease"

    if "paediatric" in text or "pediatric" in text:
        return "Paediatric"

    return "Others"


def count_in_order(series, categories):
    """Return counts in the same order as the displayed labels."""
    counts = series.value_counts()

    return [
        int(counts.get(category, 0))
        for category in categories
    ]


def check_unmapped(original_column, grouped_column, variable_name):
    """Stop the script rather than silently omitting a new category."""
    mask = grouped_column.eq("Unmapped")

    if mask.any():
        values = sorted(
            original_column.loc[mask]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        raise ValueError(
            f"Unmapped {variable_name} value(s): {values}\n"
            "Update the corresponding grouping function before producing "
            "the final figure."
        )


# ============================================================
# 4. Create grouped variables
# ============================================================

df["Study design group"] = df["Study design"].apply(study_design_group)
df["Year group"] = df["Year"].apply(year_group)
df["Country group"] = df["Country"].apply(country_group)
df["Healthcare domain group"] = df["Healthcare domain"].apply(
    healthcare_domain_group
)

check_unmapped(
    df["Study design"],
    df["Study design group"],
    "study-design",
)

check_unmapped(
    df["Year"],
    df["Year group"],
    "publication-year",
)

check_unmapped(
    df["Country"],
    df["Country group"],
    "country",
)

check_unmapped(
    df["Healthcare domain"],
    df["Healthcare domain group"],
    "healthcare-domain",
)


# ============================================================
# 5. Define labels and calculate counts
# ============================================================

study_design_labels = [
    "Empirical\nvalidation",
    "Cross-database\ncomparison",
    "Observational/\nretrospective",
    "Methodological",
    "Mixed methods",
]

year_labels = [
    "≤2010",
    "2015–2017",
    "2018–2020",
    "2021–2023",
    "2024–2025",
]

country_labels = [
    "USA",
    "Germany",
    "Other European\ncountries",
    "Canada",
    "Taiwan",
    "Australia",
    "Japan",
]

domain_labels = [
    "General /\nMulti-domain",
    "Cardiovascular",
    "Chronic disease",
    "Paediatric",
    "Others",
]


parts = [
    {
        "part": "Part (a)",
        "title": "Study design",
        "labels": study_design_labels,
        "values": count_in_order(
            df["Study design group"],
            study_design_labels,
        ),
        "color": "#4C78A8",
    },
    {
        "part": "Part (b)",
        "title": "Published year",
        "labels": year_labels,
        "values": count_in_order(
            df["Year group"],
            year_labels,
        ),
        "color": "#8E6BBE",
    },
    {
        "part": "Part (c)",
        "title": "Country of study",
        "labels": country_labels,
        "values": count_in_order(
            df["Country group"],
            country_labels,
        ),
        "color": "#2F80C1",
    },
    {
        "part": "Part (d)",
        "title": "Healthcare domain",
        "labels": domain_labels,
        "values": count_in_order(
            df["Healthcare domain group"],
            domain_labels,
        ),
        "color": "#D97757",
    },
]


# ============================================================
# 6. Verify all calculated counts
# ============================================================

for part in parts:
    panel_total = sum(part["values"])

    print(f"\n{part['part']} — {part['title']}")

    for label, value in zip(part["labels"], part["values"]):
        clean_label = label.replace("\n", " ")
        print(f"{clean_label}: {value}")

    print(f"Total: {panel_total}")

    if panel_total != number_of_studies:
        raise ValueError(
            f"{part['title']} totals {panel_total}, but the Excel sheet "
            f"contains {number_of_studies} studies."
        )


# ============================================================
# 7. Figure layout
# ============================================================

column_width = 1.0
total_columns = sum(len(part["labels"]) for part in parts)

header_height = 1.35
label_height = 2.10
plot_height = 3.00

total_height = header_height + label_height + plot_height

figure_width = 18
figure_height = 7

fig, ax = plt.subplots(
    figsize=(figure_width, figure_height)
)

fig.patch.set_facecolor("white")
ax.set_facecolor("white")

ax.set_xlim(-0.85, total_columns)
ax.set_ylim(0, total_height)
ax.axis("off")


# ============================================================
# 8. Figure style
# ============================================================

header_background = "white"
cell_background = "white"

border_color = "#5F5F5F"
text_color = "#1F1F1F"
count_color = "#333333"

# A shared scale allows direct comparison across all four panels.
maximum_value = max(
    1,
    max(
        max(part["values"])
        for part in parts
    ),
)


# ============================================================
# 9. Draw the figure
# ============================================================

x_start = 0

for part in parts:
    number_categories = len(part["labels"])
    panel_width = number_categories * column_width
    panel_color = part["color"]

    # Header block
    ax.add_patch(
        Rectangle(
            (
                x_start,
                plot_height + label_height,
            ),
            panel_width,
            header_height,
            facecolor=header_background,
            edgecolor=border_color,
            linewidth=1.2,
        )
    )

    # Part title
    ax.text(
        x_start + panel_width / 2,
        plot_height
        + label_height
        + header_height * 0.72,
        part["part"],
        ha="center",
        va="center",
        fontsize=13,
        fontweight="bold",
        family="serif",
        color=text_color,
    )

    # Panel title
    ax.text(
        x_start + panel_width / 2,
        plot_height
        + label_height
        + header_height * 0.34,
        part["title"],
        ha="center",
        va="center",
        fontsize=15,
        fontweight="bold",
        family="serif",
        color=text_color,
    )

    # Category columns and bars
    for index, (label, value) in enumerate(
        zip(
            part["labels"],
            part["values"],
        )
    ):
        x_position = x_start + index * column_width

        # Label cell
        ax.add_patch(
            Rectangle(
                (
                    x_position,
                    plot_height,
                ),
                column_width,
                label_height,
                facecolor=cell_background,
                edgecolor=border_color,
                linewidth=0.8,
            )
        )

        # Bar cell
        ax.add_patch(
            Rectangle(
                (
                    x_position,
                    0,
                ),
                column_width,
                plot_height,
                facecolor=cell_background,
                edgecolor=border_color,
                linewidth=0.8,
            )
        )

        # Rotated category label
        ax.text(
            x_position + column_width / 2,
            plot_height + label_height / 2,
            label,
            ha="center",
            va="center",
            rotation=90,
            fontsize=12,
            fontweight="bold",
            family="serif",
            color=text_color,
        )

        # Bar
        bar_width = 0.44
        bar_height = (
            value / maximum_value
        ) * (plot_height - 0.50)

        bar_x = (
            x_position
            + (column_width - bar_width) / 2
        )

        ax.add_patch(
            Rectangle(
                (
                    bar_x,
                    0.08,
                ),
                bar_width,
                bar_height,
                facecolor=panel_color,
                edgecolor=panel_color,
                linewidth=0.5,
            )
        )

        # Neutral count label above the bar
        label_y = 0.08 + bar_height + 0.08

        ax.text(
            x_position + column_width / 2,
            label_y,
            str(value),
            ha="center",
            va="bottom",
            fontsize=11,
            fontweight="bold",
            family="serif",
            color=count_color,
        )

    # Outer border around each panel
    ax.add_patch(
        Rectangle(
            (
                x_start,
                0,
            ),
            panel_width,
            total_height,
            fill=False,
            edgecolor=border_color,
            linewidth=1.2,
        )
    )

    x_start += panel_width


# ============================================================
# 10. Save and display
# ============================================================

plt.tight_layout()

# Save high-resolution PNG
fig.savefig(
    OUTPUT_PATH,
    dpi=600,
    bbox_inches="tight",
    pad_inches=0.05,
    facecolor="white",
)

# Save vector PDF
PDF_OUTPUT_PATH = OUTPUT_PATH.with_suffix(".pdf")

fig.savefig(
    PDF_OUTPUT_PATH,
    format="pdf",
    bbox_inches="tight",
    pad_inches=0.05,
    facecolor="white",
)

# Save vector EPS for journal submission
EPS_OUTPUT_PATH = OUTPUT_PATH.with_suffix(".eps")

fig.savefig(
    EPS_OUTPUT_PATH,
    format="eps",
    bbox_inches="tight",
    pad_inches=0.05,
    facecolor="white",
)

plt.show()
plt.close(fig)

print(f"\nSaved PNG:\n{OUTPUT_PATH}")
print(f"\nSaved PDF:\n{PDF_OUTPUT_PATH}")
print(f"\nSaved EPS:\n{EPS_OUTPUT_PATH}")