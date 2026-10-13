# CH20 – Air Quality Insight

## Objective
Analyze station and time patterns in measured air quality and assess whether the supplied weather and traffic files can be joined without unsupported assumptions.

## Setup and execution
Use Python 3.11 and the project-local environment:

```powershell
.\.venv\Scripts\Activate.ps1
python scripts\data_audit.py
python scripts\preprocess.py
python scripts\analyze.py
python scripts\train_model.py
python scripts\generate_reports.py
python scripts\final_validation.py
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

CH20 – Evidence-based EDA findings

1. [Supported by data] AQI coverage: 87,025 of 108,035 air-quality records have AQI (80.6%). Mean AQI is 179.75; median is 132.00.
2. [Supported by data] Station variation: among stations with at least 20 AQI records, GJ001 has the highest mean AQI (452.12) and MZ001 the lowest (34.77); high station n=1334, low station n=111.
3. [Supported by data] Derived seasonal pattern: Winter has the highest mean AQI (243.48) and Summer the lowest (117.39); seasons are derived from timestamp month, not weather observations.
4. [Suggestive, not causal] Strongest absolute pollutant correlation with AQI in these records is PM10 (Pearson r=0.885). Correlation is not causation.
5. [Not assessable with supplied data] Wind, rainfall, temperature and humidity analyses were not performed because the weather file contains only categorical weather_condition and does not overlap the air-quality date range.
6. [Not assessable with supplied data] Traffic volume versus AQI was not performed because traffic has no AQI/station key and its dates (2022–2024) do not overlap air quality (2015–2020).
7. [Suggestive, not causal] Chronological same-record AQI estimation: Random Forest MAE=17.46, RMSE=26.97, R²=0.902; training-mean baseline MAE=93.95, RMSE=105.32, R²=-0.492.

The model estimates same-record AQI from pollutant measurements; it is not a future forecast. AQI is derived from pollutants, so this result should not be interpreted as an independent causal prediction.


See outputs/eda_findings.txt for actual model metrics.

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
