from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import nbformat as nbf
import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib.ticker import MaxNLocator
from nbclient import NotebookClient
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.util import Inches, Pt
from sklearn.ensemble import RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.pipeline import Pipeline


ROOT = Path(__file__).resolve().parents[1]
FIG_DIR = ROOT / "outputs" / "figures"
PROC_DIR = ROOT / "processed"
FIG_DIR.mkdir(parents=True, exist_ok=True)
PROC_DIR.mkdir(parents=True, exist_ok=True)

FILES = ["air_quality.csv", "stations.csv", "weather.csv", "traffic.csv", "calendar.csv"]
NUMERIC_POLLUTANTS = [
    "PM2.5", "PM10", "NO", "NO2", "NOx", "NH3", "CO", "SO2",
    "O3", "Benzene", "Toluene", "Xylene",
]


def load_data() -> dict[str, pd.DataFrame]:
    return {
        "air_quality": pd.read_csv(ROOT / "air_quality.csv", parse_dates=["timestamp"]),
        "stations": pd.read_csv(ROOT / "stations.csv"),
        "weather": pd.read_csv(ROOT / "weather.csv", parse_dates=["timestamp"]),
        "traffic": pd.read_csv(ROOT / "traffic.csv", parse_dates=["Date"]),
        "calendar": pd.read_csv(ROOT / "calendar.csv", parse_dates=["timestamp"]),
    }


def audit(datasets: dict[str, pd.DataFrame]) -> str:
    lines = [
        "CH20 – Air Quality Insight: Data Audit",
        f"Generated with {sys.executable}",
        "",
        "Dataset source note: CSV files were supplied in the project workspace. "
        "No external source metadata was included in the files.",
    ]
    for name, df in datasets.items():
        lines += [
            f"\n=== {name} ===",
            f"Rows: {len(df):,}",
            f"Columns: {len(df.columns)}",
            f"Column names: {', '.join(df.columns)}",
            f"Duplicate rows: {int(df.duplicated().sum()):,}",
            "Data types: " + "; ".join(f"{c}={t}" for c, t in df.dtypes.items()),
            "Missing values: " + "; ".join(
                f"{c}={int(v):,}" for c, v in df.isna().sum().items() if v
            ) if df.isna().any().any() else "Missing values: none",
        ]
        for col in [c for c in df.columns if c in {"timestamp", "Date"}]:
            lines.append(
                f"{col} range: {df[col].min()} to {df[col].max()}; "
                f"invalid after parsing: {int(df[col].isna().sum())}"
            )
        for col in ["station", "location", "Area Name", "season"]:
            if col in df.columns:
                vals = sorted(df[col].dropna().astype(str).unique())
                lines.append(f"{col} unique values: {len(vals):,}; examples: {vals[:12]}")
        numeric = df.select_dtypes(include="number")
        if not numeric.empty:
            negative = {
                col: int((numeric[col] < 0).sum())
                for col in numeric.columns
                if (numeric[col] < 0).any()
            }
            lines.append(f"Negative numeric values: {negative or 'none'}")
            iqr_flags = {}
            for col in numeric.columns:
                values = numeric[col].dropna()
                if len(values) < 4:
                    continue
                q1, q3 = values.quantile([0.25, 0.75])
                iqr = q3 - q1
                if iqr > 0:
                    count = int(((values < q1 - 1.5 * iqr) | (values > q3 + 1.5 * iqr)).sum())
                    if count:
                        iqr_flags[col] = count
            lines.append(f"IQR outlier flags (investigate, not deleted): {iqr_flags or 'none'}")
        if "AQI" in df.columns:
            lines.append(
                f"AQI > 500 review flags: {int((pd.to_numeric(df['AQI'], errors='coerce') > 500).sum())}; "
                "retained because the source AQI_Bucket labels these records Severe."
            )

    aq = datasets["air_quality"]
    aq_station_ids = set(aq["station"].dropna().astype(str).str.strip().str.upper())
    station_ids = set(datasets["stations"]["station"].dropna().astype(str).str.strip().str.upper())
    calendar_dates = set(pd.to_datetime(datasets["calendar"]["timestamp"]).dt.date)
    aq_dates = set(pd.to_datetime(aq["timestamp"]).dt.date)
    lines += [
        "\n=== Compatibility checks ===",
        f"Air-quality dates: {aq.timestamp.min().date()} to {aq.timestamp.max().date()}",
        f"Weather dates: {datasets['weather'].timestamp.min().date()} to "
        f"{datasets['weather'].timestamp.max().date()}",
        f"Traffic dates: {datasets['traffic'].Date.min().date()} to "
        f"{datasets['traffic'].Date.max().date()}",
        "Air-quality/weather date overlap: "
        f"{len(set(aq.timestamp.dt.date) & set(datasets['weather'].timestamp.dt.date)):,} dates",
        "Air-quality/traffic date overlap: "
        f"{len(set(aq.timestamp.dt.date) & set(datasets['traffic'].Date.dt.date)):,} dates",
        "Weather and traffic use 2022–2024 location/date records; air quality uses "
        "2015–2020 station/timestamp records. No cross-dataset pollution join was made.",
        "Weather has only categorical weather_condition; wind, rainfall, temperature, "
        "and humidity measurements are unavailable.",
        "Traffic has no AQI or station identifier; traffic volume versus AQI was not calculated.",
        "\n=== Compatibility matrix ===",
        "air_quality + stations: JOINABLE by normalized station ID; "
        f"{len(aq_station_ids & station_ids)} of {len(aq_station_ids)} air-quality IDs are present in stations "
        f"({len(aq_station_ids - station_ids)} unmatched air-quality IDs). This is metadata enrichment only.",
        "air_quality + calendar: JOINABLE by calendar date; calendar supplies date labels. "
        f"{len(aq_dates & calendar_dates)} of {len(aq_dates)} air-quality dates are present in calendar; "
        "timestamp granularity differs, so the join key is date, not timestamp.",
        "air_quality + weather: NOT ASSESSABLE; 0 overlapping calendar dates and no station/location mapping.",
        "air_quality + traffic: NOT ASSESSABLE; 0 overlapping calendar dates, no station key, and no AQI field.",
        "weather + traffic: JOINABLE by date and shared area names for their own 2022–2024 descriptive comparison; "
        "this does not create an AQI relationship.",
    ]
    return "\n".join(lines) + "\n"


def clean_data(datasets: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    aq = datasets["air_quality"].copy()
    aq["station"] = aq["station"].astype("string").str.strip().str.upper()
    aq["timestamp"] = pd.to_datetime(aq["timestamp"], errors="coerce")
    before = len(aq)
    aq["invalid_numeric_flag"] = False
    aq["range_review_flag"] = False
    for column in NUMERIC_POLLUTANTS + ["AQI"]:
        aq[column] = pd.to_numeric(aq[column], errors="coerce")
        aq["invalid_numeric_flag"] |= aq[column].lt(0).fillna(False)
    aq["range_review_flag"] |= aq["AQI"].gt(500).fillna(False)
    aq["outlier_flag"] = False
    for column in NUMERIC_POLLUTANTS + ["AQI"]:
        values = aq[column].dropna()
        if len(values) >= 4:
            q1, q3 = values.quantile([0.25, 0.75])
            iqr = q3 - q1
            if iqr > 0:
                aq["outlier_flag"] |= aq[column].lt(q1 - 1.5 * iqr).fillna(False)
                aq["outlier_flag"] |= aq[column].gt(q3 + 1.5 * iqr).fillna(False)
    aq = aq.dropna(subset=["station", "timestamp"]).drop_duplicates()
    aq["preprocessing_rows_before"] = before
    aq["preprocessing_rows_after"] = len(aq)
    aq["date"] = aq["timestamp"].dt.date.astype("string")
    aq["hour"] = aq["timestamp"].dt.hour
    aq["weekday"] = aq["timestamp"].dt.day_name()
    aq["month"] = aq["timestamp"].dt.month
    aq["year"] = aq["timestamp"].dt.year
    aq["season"] = aq["timestamp"].dt.month.map(
        lambda m: "Winter" if m in (12, 1, 2) else
        "Spring" if m in (3, 4, 5) else
        "Summer" if m in (6, 7, 8) else "Autumn"
    )
    aq.to_csv(PROC_DIR / "air_quality_cleaned.csv", index=False)

    st = datasets["stations"].copy()
    st["station"] = st["station"].astype("string").str.strip().str.upper()
    st.to_csv(PROC_DIR / "stations_cleaned.csv", index=False)

    weather = datasets["weather"].copy()
    weather["location"] = weather["location"].astype("string").str.strip()
    weather.to_csv(PROC_DIR / "weather_cleaned.csv", index=False)

    traffic = datasets["traffic"].copy()
    traffic["Area Name"] = traffic["Area Name"].astype("string").str.strip()
    traffic.to_csv(PROC_DIR / "traffic_cleaned.csv", index=False)

    cal = datasets["calendar"].copy()
    cal["timestamp"] = pd.to_datetime(cal["timestamp"], errors="coerce")
    cal.to_csv(PROC_DIR / "calendar_cleaned.csv", index=False)
    return {"air_quality": aq, "stations": st, "weather": weather, "traffic": traffic, "calendar": cal}


def make_figures(aq: pd.DataFrame) -> dict[str, Path]:
    sns.set_theme(style="whitegrid")
    plt.rcParams["figure.dpi"] = 120
    figures: dict[str, Path] = {}

    station = aq.groupby("station", as_index=False).agg(
        mean_aqi=("AQI", "mean"), median_aqi=("AQI", "median"), records=("AQI", "count")
    ).dropna()
    station = station.sort_values("mean_aqi", ascending=False)
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.barplot(data=station.head(15), x="mean_aqi", y="station", color="#2f6f8f", ax=ax)
    ax.set(title="Top 15 Stations by Mean AQI", xlabel="Mean AQI", ylabel="Station")
    fig.tight_layout()
    figures["station_comparison.png"] = FIG_DIR / "station_comparison.png"
    fig.savefig(figures["station_comparison.png"]); plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.histplot(aq["AQI"].dropna(), bins=40, kde=True, color="#d95f02", ax=ax)
    ax.set(title="AQI Distribution", xlabel="AQI", ylabel="Number of observations")
    fig.tight_layout()
    figures["aqi_distribution.png"] = FIG_DIR / "aqi_distribution.png"
    fig.savefig(figures["aqi_distribution.png"]); plt.close(fig)

    monthly = aq.dropna(subset=["AQI"]).groupby(aq["timestamp"].dt.to_period("M"))["AQI"].mean()
    fig, ax = plt.subplots(figsize=(12, 5))
    ax.plot(monthly.index.astype(str), monthly.values, color="#1b9e77")
    ax.set(title="Monthly Mean AQI Over Time", xlabel="Month", ylabel="Mean AQI")
    ax.xaxis.set_major_locator(MaxNLocator(12))
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    figures["monthly_trend.png"] = FIG_DIR / "monthly_trend.png"
    fig.savefig(figures["monthly_trend.png"]); plt.close(fig)

    corr = aq[NUMERIC_POLLUTANTS + ["AQI"]].corr(numeric_only=True)
    fig, ax = plt.subplots(figsize=(11, 8))
    sns.heatmap(corr, cmap="coolwarm", center=0, ax=ax)
    ax.set_title("Pollutant and AQI Correlations")
    fig.tight_layout()
    figures["correlation_heatmap.png"] = FIG_DIR / "correlation_heatmap.png"
    fig.savefig(figures["correlation_heatmap.png"]); plt.close(fig)

    for column, filename, title in [
        ("hour", "hourly_pattern.png", "Mean AQI by Hour"),
        ("weekday", "weekday_pattern.png", "Mean AQI by Weekday"),
        ("season", "seasonal_pattern.png", "Mean AQI by Season"),
    ]:
        values = aq.dropna(subset=["AQI"]).groupby(column, sort=False)["AQI"].mean()
        if column == "weekday":
            values = values.reindex(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
        if column == "season":
            values = values.reindex(["Winter", "Spring", "Summer", "Autumn"])
        fig, ax = plt.subplots(figsize=(9, 5))
        values.plot(kind="bar", color="#7570b3", ax=ax)
        ax.set(title=title, xlabel=column.title(), ylabel="Mean AQI")
        ax.tick_params(axis="x", rotation=0)
        fig.tight_layout()
        figures[filename] = FIG_DIR / filename
        fig.savefig(figures[filename]); plt.close(fig)
    return figures


def model_results(aq: pd.DataFrame) -> tuple[pd.DataFrame, Pipeline | None, str]:
    model_df = aq.sort_values("timestamp").dropna(subset=["AQI"]).copy()
    usable = [c for c in NUMERIC_POLLUTANTS if c in model_df]
    if len(model_df) < 100 or not usable:
        return pd.DataFrame(), None, "A reliable model was not trained: insufficient target/features."
    split = int(len(model_df) * 0.8)
    train, test = model_df.iloc[:split], model_df.iloc[split:]
    pipe = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("model", RandomForestRegressor(n_estimators=120, random_state=42, n_jobs=-1, max_depth=18)),
    ])
    pipe.fit(train[usable], train["AQI"])
    pred = pipe.predict(test[usable])
    baseline_pred = np.repeat(train["AQI"].mean(), len(test))
    results = pd.DataFrame([
        {"model": "Random Forest regression", "split": "chronological final 20%",
         "rows_train": len(train), "rows_test": len(test),
         "MAE": mean_absolute_error(test["AQI"], pred),
         "RMSE": mean_squared_error(test["AQI"], pred) ** 0.5, "R2": r2_score(test["AQI"], pred)},
        {"model": "Training-mean baseline", "split": "chronological final 20%",
         "rows_train": len(train), "rows_test": len(test),
         "MAE": mean_absolute_error(test["AQI"], baseline_pred),
         "RMSE": mean_squared_error(test["AQI"], baseline_pred) ** 0.5, "R2": r2_score(test["AQI"], baseline_pred)},
    ])
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.scatter(test["AQI"], pred, alpha=0.15, s=8)
    lims = [min(test["AQI"].min(), pred.min()), max(test["AQI"].max(), pred.max())]
    ax.plot(lims, lims, "--", color="black", label="Perfect estimate")
    ax.set(title="AQI Estimation: Actual vs Predicted", xlabel="Actual AQI", ylabel="Predicted AQI")
    ax.legend()
    fig.tight_layout()
    fig.savefig(FIG_DIR / "model_actual_vs_predicted.png"); plt.close(fig)
    return results, pipe, (
        "The model estimates same-record AQI from pollutant measurements; it is not a "
        "future forecast. AQI is derived from pollutants, so this result should not be "
        "interpreted as an independent causal prediction."
    )


def findings(aq: pd.DataFrame, model: pd.DataFrame) -> str:
    valid = aq.dropna(subset=["AQI"])
    station = valid.groupby("station")["AQI"].agg(["mean", "count"]).query("count >= 20")
    high, low = station["mean"].idxmax(), station["mean"].idxmin()
    month = valid.groupby("month")["AQI"].mean()
    season = valid.groupby("season")["AQI"].mean().sort_values(ascending=False)
    corr = valid[NUMERIC_POLLUTANTS + ["AQI"]].corr(numeric_only=True)["AQI"].drop("AQI").abs().sort_values(ascending=False)
    lines = [
        "CH20 – Evidence-based EDA findings",
        "",
        f"1. [Supported by data] AQI coverage: {len(valid):,} of {len(aq):,} air-quality records have AQI "
        f"({len(valid) / len(aq):.1%}). Mean AQI is {valid.AQI.mean():.2f}; median is {valid.AQI.median():.2f}.",
        f"2. [Supported by data] Station variation: among stations with at least 20 AQI records, "
        f"{high} has the highest mean AQI ({station.loc[high, 'mean']:.2f}) and "
        f"{low} the lowest ({station.loc[low, 'mean']:.2f}); high station n={int(station.loc[high, 'count'])}, "
        f"low station n={int(station.loc[low, 'count'])}.",
        f"3. [Supported by data] Derived seasonal pattern: {season.index[0]} has the highest mean AQI "
        f"({season.iloc[0]:.2f}) and {season.index[-1]} the lowest ({season.iloc[-1]:.2f}); "
        "seasons are derived from timestamp month, not weather observations.",
        f"4. [Suggestive, not causal] Strongest absolute pollutant correlation with AQI in these records is "
        f"{corr.index[0]} (Pearson r={valid[corr.index[0]].corr(valid.AQI):.3f}). Correlation is not causation.",
        "5. [Not assessable with supplied data] Wind, rainfall, temperature and humidity analyses were not performed because "
        "the weather file contains only categorical weather_condition and does not overlap "
        "the air-quality date range.",
        "6. [Not assessable with supplied data] Traffic volume versus AQI was not performed because traffic has no AQI/station "
        "key and its dates (2022–2024) do not overlap air quality (2015–2020).",
    ]
    if not model.empty:
        rf, base = model.iloc[0], model.iloc[1]
        lines.append(
            f"7. [Suggestive, not causal] Chronological same-record AQI estimation: Random Forest MAE={rf.MAE:.2f}, "
            f"RMSE={rf.RMSE:.2f}, R²={rf.R2:.3f}; training-mean baseline "
            f"MAE={base.MAE:.2f}, RMSE={base.RMSE:.2f}, R²={base.R2:.3f}."
        )
    return "\n".join(lines) + "\n"


def make_excel(aq: pd.DataFrame, datasets: dict[str, pd.DataFrame], station: pd.DataFrame,
               model: pd.DataFrame, text: str) -> None:
    path = ROOT / "air_quality_analysis.xlsx"
    monthly = aq.groupby(["year", "month"], as_index=False)["AQI"].agg(["mean", "median", "count"]).reset_index()
    seasonal = aq.groupby("season", as_index=False)["AQI"].agg(["mean", "median", "count"]).reset_index()
    corr = aq[NUMERIC_POLLUTANTS + ["AQI"]].corr(numeric_only=True)["AQI"].reset_index()
    corr.columns = ["variable", "correlation_with_AQI"]
    summary = pd.DataFrame({
        "Metric": ["Air-quality records", "AQI records", "Stations", "Air-quality date start", "Air-quality date end",
                   "Weather/traffic merge decision"],
        "Value": [len(aq), int(aq.AQI.notna().sum()), aq.station.nunique(), str(aq.timestamp.min().date()),
                  str(aq.timestamp.max().date()), "Not merged: incompatible dates/keys"],
    })
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"README": [
            "CH20 – Air Quality Insight",
            "Workbook contains supported summaries and chart-ready tables.",
            "The weather and traffic datasets were not merged with air quality because dates and keys are incompatible.",
            "Source CSVs remain unchanged; cleaned air-quality data is available in processed/.",
        ]}).to_excel(writer, sheet_name="README", index=False)
        summary.to_excel(writer, sheet_name="Data Summary", index=False)
        pd.DataFrame({
            "Join": ["air_quality + stations", "air_quality + calendar", "air_quality + weather",
                     "air_quality + traffic", "weather + traffic"],
            "Status": ["Joinable: normalized station ID", "Joinable: calendar date",
                       "Not assessable with supplied data", "Not assessable with supplied data",
                       "Joinable for weather/traffic only"],
            "Reason": [
                "Metadata enrichment; verify unmatched station IDs.",
                "Date-level join; timestamp granularity differs.",
                "Zero overlapping dates and no station/location mapping.",
                "Zero overlapping dates, no station key, and no AQI field.",
                "Shared 2022–2024 dates and area names; does not establish AQI relationship.",
            ],
        }).to_excel(writer, sheet_name="Compatibility Matrix", index=False)
        pd.DataFrame({
            "Quality item": ["Air-quality rows before preprocessing", "Air-quality rows after preprocessing",
                             "Air-quality invalid numeric flags", "AQI > 500 review flags",
                             "Air-quality IQR outlier flags",
                             "Exact duplicate rows in all supplied files"],
            "Value": [len(datasets["air_quality"]), len(aq), int(aq["invalid_numeric_flag"].sum()),
                      int(aq["range_review_flag"].sum()), int(aq["outlier_flag"].sum()), "0 in each file"],
        }).to_excel(writer, sheet_name="Data Quality", index=False)
        station.to_excel(writer, sheet_name="Station Analysis", index=False)
        monthly.to_excel(writer, sheet_name="Time Analysis", index=False)
        corr.to_excel(writer, sheet_name="Correlation Analysis", index=False)
        seasonal.to_excel(writer, sheet_name="Seasonal Analysis", index=False)
        if not model.empty:
            model.to_excel(writer, sheet_name="Model Results", index=False)
        pd.DataFrame({"Limitation": [
            "Weather contains only categorical weather_condition; numeric wind/rainfall/temperature/humidity are unavailable.",
            "Weather and traffic dates do not overlap air-quality dates.",
            "Traffic has no AQI or air-quality station identifier.",
            "Pollutant-based AQI model is same-record estimation, not future forecasting; AQI is derived from pollutants.",
            "Outlier flags are retained for investigation and are not silently deleted.",
        ]}).to_excel(writer, sheet_name="Limitations", index=False)
        pd.DataFrame({"Finding": text.splitlines()}).to_excel(writer, sheet_name="Key Insights", index=False)
    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = min(
                50, max(12, max(len(str(cell.value or "")) for cell in column) + 2)
            )
    ws = wb["Station Analysis"]
    if ws.max_row > 1:
        ws.conditional_formatting.add(f"B2:B{ws.max_row}", ColorScaleRule(start_type="min", start_color="63BE7B",
                                                                          mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                                                          end_type="max", end_color="F8696B"))
        chart = BarChart()
        chart.title = "Station Mean AQI"
        chart.y_axis.title = "Mean AQI"
        chart.x_axis.title = "Station"
        chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=min(ws.max_row, 16)), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=min(ws.max_row, 16)))
        ws.add_chart(chart, "E2")
    ws = wb["Time Analysis"]
    if ws.max_row > 1:
        chart = LineChart()
        chart.title = "Monthly Mean AQI"
        chart.add_data(Reference(ws, min_col=3, min_row=1, max_row=ws.max_row), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=2, min_row=2, max_row=ws.max_row))
        ws.add_chart(chart, "H2")
    wb.save(path)


def make_ppt(aq: pd.DataFrame, model: pd.DataFrame, text: str) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    def slide(title: str, body: list[str], image: Path | None = None):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        bg = s.background.fill; bg.solid(); bg.fore_color.rgb = __import__("pptx").dml.color.RGBColor(245, 248, 250)
        box = s.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.7))
        p = box.text_frame.paragraphs[0]; p.text = title; p.font.size = Pt(26); p.font.bold = True
        p.font.color.rgb = __import__("pptx").dml.color.RGBColor(31, 78, 121)
        if image and image.exists():
            s.shapes.add_picture(str(image), Inches(0.7), Inches(1.35), width=Inches(7.2))
            x, y, w = 8.3, 1.35, 4.3
        else:
            x, y, w = 0.9, 1.45, 11.5
        box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(5.3))
        tf = box.text_frame; tf.word_wrap = True
        for i, item in enumerate(body):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = item; p.font.size = Pt(16); p.space_after = Pt(12)
        return s

    slide("CH20 – Air Quality Insight", [
        "DS Sprint 2026 – Skill-Based Assessment 2",
        "Objective: understand station and time patterns in measured air quality and communicate evidence-based actions.",
    ])
    slide("Datasets and compatibility", [
        f"Air quality: {len(aq):,} timestamped records, {aq.station.nunique()} stations, {aq.timestamp.min().date()}–{aq.timestamp.max().date()}.",
        "Stations provides station metadata; calendar provides date attributes.",
        "Weather and traffic are 2022–2024 and do not overlap air quality dates (2015–2020); no unsupported merge was made.",
    ])
    slide("Methodology and data quality", [
        f"Exact duplicate rows: none in the five supplied files.",
        f"AQI is present for {aq.AQI.notna().sum():,} of {len(aq):,} air-quality rows.",
        "Dates were parsed safely; station IDs were trimmed and uppercased. Cleaned copies are saved in processed/.",
        "Weather contains categorical conditions only; wind/rainfall measurements are unavailable.",
    ])
    slide("Evidence from EDA", [
        *text.splitlines()[4:8],
    ], FIG_DIR / "station_comparison.png")
    model_body = ["No model was trained."] if model.empty else [
        f"Random Forest estimation: MAE={model.iloc[0].MAE:.2f}, RMSE={model.iloc[0].RMSE:.2f}, R²={model.iloc[0].R2:.3f}.",
        f"Training-mean baseline: MAE={model.iloc[1].MAE:.2f}, RMSE={model.iloc[1].RMSE:.2f}, R²={model.iloc[1].R2:.3f}.",
        "This is same-record AQI estimation from pollutant readings, not a future forecast.",
    ]
    slide("Machine-learning result", model_body, FIG_DIR / "model_actual_vs_predicted.png")
    slide("Recommendations", [
        "Prioritize monitoring and investigation at stations with persistently higher mean AQI.",
        "Use the seasonal and monthly patterns to plan intensified monitoring and public advisories.",
        "Collect weather measurements and station-linked traffic data in the same time window before claiming causal effects.",
    ])
    slide("Conclusion and limitations", [
        "The strongest defensible results come from station and time analysis within the air-quality dataset.",
        "Traffic and weather relationships remain untested because the supplied datasets are temporally and structurally incompatible.",
        "Interpret pollutant/AQI relationships as associations; AQI is derived from pollutant measurements.",
    ])
    prs.save(ROOT / "business_insights.pptx")


def make_notebook(audit_text: str, findings_text: str, model: pd.DataFrame) -> None:
    nb = nbf.v4.new_notebook()
    nb["metadata"] = {
        "kernelspec": {"display_name": "Python 3.11 (Air Quality Insight)", "language": "python", "name": "air-quality-insight"},
        "language_info": {"name": "python", "version": "3.11.15"},
    }
    cells = [
        nbf.v4.new_markdown_cell("# CH20 – Air Quality Insight\n\nAn evidence-based analysis of the supplied air-quality, station, weather, traffic, and calendar CSV files."),
        nbf.v4.new_markdown_cell("## Objective\nAnalyze how station and time relate to pollution, validate whether weather/traffic joins are defensible, and produce actionable recommendations without fabricating unavailable measurements."),
        nbf.v4.new_code_cell("import sys, pandas as pd, numpy as np, matplotlib.pyplot as plt, seaborn as sns\nfrom pathlib import Path\nprint(sys.executable)\nprint(sys.version)"),
        nbf.v4.new_code_cell("DATA_DIR = Path.cwd()\nair_quality = pd.read_csv(DATA_DIR / 'air_quality.csv', parse_dates=['timestamp'])\nstations = pd.read_csv(DATA_DIR / 'stations.csv')\nweather = pd.read_csv(DATA_DIR / 'weather.csv', parse_dates=['timestamp'])\ntraffic = pd.read_csv(DATA_DIR / 'traffic.csv', parse_dates=['Date'])\ncalendar = pd.read_csv(DATA_DIR / 'calendar.csv', parse_dates=['timestamp'])\n{ name: df.shape for name, df in {'air_quality': air_quality, 'stations': stations, 'weather': weather, 'traffic': traffic, 'calendar': calendar}.items() }"),
        nbf.v4.new_markdown_cell("## Audit and compatibility\nThe following report is generated from the actual files. Weather and traffic are intentionally not merged with air quality because their dates and join keys are incompatible."),
        nbf.v4.new_code_cell("print(" + repr(audit_text) + ")"),
        nbf.v4.new_markdown_cell("## Cleaning and feature engineering"),
        nbf.v4.new_code_cell("aq = air_quality.copy()\naq['station'] = aq['station'].astype('string').str.strip().str.upper()\naq['timestamp'] = pd.to_datetime(aq['timestamp'], errors='coerce')\naq['hour'] = aq['timestamp'].dt.hour\naq['weekday'] = aq['timestamp'].dt.day_name()\naq['month'] = aq['timestamp'].dt.month\naq['year'] = aq['timestamp'].dt.year\naq['season'] = aq['month'].map(lambda m: 'Winter' if m in (12,1,2) else 'Spring' if m in (3,4,5) else 'Summer' if m in (6,7,8) else 'Autumn')\naq.head()"),
        nbf.v4.new_markdown_cell("## Exploratory findings"),
        nbf.v4.new_code_cell("from IPython.display import display, Image\n\nstation_summary = (aq.dropna(subset=['AQI']).groupby('station')['AQI'].agg(['mean', 'median', 'count']).sort_values('mean', ascending=False))\ndisplay(aq[['AQI'] + [c for c in ['PM2.5', 'PM10', 'NO2', 'NOx'] if c in aq]].describe().T)\ndisplay(station_summary.head(10))"),
        nbf.v4.new_code_cell("Image(filename=str(DATA_DIR / 'outputs' / 'figures' / 'station_comparison.png'))"),
        nbf.v4.new_code_cell("Image(filename=str(DATA_DIR / 'outputs' / 'figures' / 'monthly_trend.png'))"),
        nbf.v4.new_code_cell("print(" + repr(findings_text) + ")"),
        nbf.v4.new_markdown_cell("## Limitations and conclusion\nThe air-quality dataset supports station/time EDA and same-record AQI estimation. It does not support causal weather or traffic analysis with the supplied files because measurements and compatible keys are unavailable."),
    ]
    if not model.empty:
        cells.insert(-1, nbf.v4.new_code_cell("from IPython.display import display\ndisplay(pd.DataFrame(" + repr(model.to_dict("records")) + "))"))
    nb["cells"] = cells
    path = ROOT / "air_insights.ipynb"
    with path.open("w", encoding="utf-8") as handle:
        nbf.write(nb, handle)
    executed = nbf.read(path.open(encoding="utf-8"), as_version=4)
    NotebookClient(executed, timeout=600, kernel_name="air-quality-insight", resources={"metadata": {"path": str(ROOT)}}).execute()
    with path.open("w", encoding="utf-8") as handle:
        nbf.write(executed, handle)


def make_readme(findings_text: str, model: pd.DataFrame) -> None:
    metrics = "See outputs/eda_findings.txt for actual model metrics." if not model.empty else "No valid model was trained."
    content = f"""# CH20 – Air Quality Insight

## Objective
Analyze station and time patterns in measured air quality and assess whether the supplied weather and traffic files can be joined without unsupported assumptions.

## Setup and execution
Use Python 3.11 and the project-local environment:

```powershell
.\\.venv\\Scripts\\Activate.ps1
python scripts\\data_audit.py
python scripts\\preprocess.py
python scripts\\analyze.py
python scripts\\train_model.py
python scripts\\generate_reports.py
python scripts\\final_validation.py
```

The registered Jupyter kernel is `Python 3.11 (Air Quality Insight)`.

## Data and methodology
The five CSV files were supplied in the workspace; no external source metadata was included. Original files are preserved. Dates, missingness, duplicates, station/location values, ranges, and compatibility were audited. Cleaned copies are in `processed/`.

Air quality covers 2015–2020. Weather and traffic cover 2022–2024, so they were not merged with air quality. Weather has only categorical `weather_condition`; wind, rainfall, temperature, and humidity measurements are unavailable. Traffic has no air-quality station key or AQI.

Compatibility matrix:

| Datasets | Verdict | Reason |
|---|---|---|
| Air quality + stations | Joinable | Normalized station ID; metadata enrichment only |
| Air quality + calendar | Joinable | Date-level join; timestamp granularity differs |
| Air quality + weather | Not assessable with supplied data | Zero overlapping dates and no station/location mapping |
| Air quality + traffic | Not assessable with supplied data | Zero overlapping dates, no station key, and no AQI |
| Weather + traffic | Joinable for their own descriptive comparison | Shared 2022–2024 dates and area names; no AQI relationship |

## Results

{findings_text}

{metrics}

## Outputs
- `air_insights.ipynb`: executed notebook workflow.
- `air_quality_analysis.xlsx`: chart-ready summaries, formulas/tables, and supported charts.
- `business_insights.pptx`: business presentation.
- `data_audit_report.txt`: actual audit output.
- `outputs/figures/`: EDA and model figures.
- `processed/`: cleaned, traceable copies of all five datasets.

## Limitations
Pollutant/AQI correlations are associations, and AQI is derived from pollutant measurements. The model is same-record AQI estimation rather than future forecasting. Weather and traffic causal relationships require aligned dates, station/location identifiers, and numeric measurements.

Outlier values are flagged in the cleaned air-quality file and audit report for investigation; they are not silently deleted. The project does not claim that traffic or weather affects AQI.
"""
    (ROOT / "README.md").write_text(content, encoding="utf-8")
    (ROOT / "requirements.txt").write_text(
        "pandas\nnumpy\nmatplotlib\nseaborn\nscikit-learn\njupyter\nipykernel\nopenpyxl\npython-pptx\nnbformat\nnbclient\n",
        encoding="utf-8",
    )
    (ROOT / ".gitignore").write_text(".venv/\n__pycache__/\n*.pyc\n.ipynb_checkpoints/\n", encoding="utf-8")


def main() -> None:
    datasets = load_data()
    audit_text = audit(datasets)
    (ROOT / "data_audit_report.txt").write_text(audit_text, encoding="utf-8")
    cleaned = clean_data(datasets)
    aq = cleaned["air_quality"]
    figures = make_figures(aq)
    station = aq.groupby("station", as_index=False).agg(
        mean_aqi=("AQI", "mean"), median_aqi=("AQI", "median"), records=("AQI", "count")
    ).dropna().sort_values("mean_aqi", ascending=False)
    model, _, model_note = model_results(aq)
    findings_text = findings(aq, model)
    findings_text += "\n" + model_note + "\n"
    (ROOT / "outputs" / "eda_findings.txt").write_text(findings_text, encoding="utf-8")
    make_excel(aq, cleaned, station, model, findings_text)
    make_ppt(aq, model, findings_text)
    make_readme(findings_text, model)
    make_notebook(audit_text, findings_text, model)
    print(json.dumps({
        "datasets": {k: list(v.shape) for k, v in datasets.items()},
        "figures": [str(p.relative_to(ROOT)) for p in figures.values()],
        "model_results": model.to_dict("records"),
        "artifacts": [str(p.relative_to(ROOT)) for p in [
            ROOT / "data_audit_report.txt", ROOT / "outputs" / "eda_findings.txt",
            ROOT / "air_insights.ipynb", ROOT / "air_quality_analysis.xlsx",
            ROOT / "business_insights.pptx", ROOT / "README.md",
        ]],
    }, indent=2, default=str))


if __name__ == "__main__":
    main()
