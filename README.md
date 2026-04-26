# FX PARITIES OVERVIEW REPORT / HTML DASHBOARD / AUTMATION
## Overview
The FX Parities Reporting Engine is a Python-based treasury reporting solution designed to automate the generation of standardized FX parity reports. It supports both live API-based data retrieval and user-provided Excel uploads, transforming daily USD-based FX rates into a structured reporting package that can be used for analysis, review, and presentation.

The engine was built to reduce manual reporting effort, improve consistency across FX datasets, and provide a repeatable process for producing treasury-friendly FX reporting outputs. In a single run, it creates formatted Excel deliverables and an interactive HTML dashboard for executive-level review.

## What the Project Does
This project automates the full reporting flow from FX source data to presentation-ready outputs. It:

- retrieves daily FX rates through API mode or accepts existing user-uploaded FX data
- standardizes the dataset into a common reporting structure
- fills missing calendar dates using forward-fill logic
- enriches the dataset with reporting-ready masterdata fields
- calculates rate movement and volatility metrics
- exports structured Excel files and a multi-page HTML dashboard

## Key Features
- **Two operating modes**
  - **API Mode**: automatically fetches daily USD-based FX rates
  - **User-Upload Mode**: accepts an Excel file and reshapes it into the same standardized format

- **Automated data preparation**
  - column normalization
  - currency catalog construction
  - missing-day calendar completion
  - forward-fill logic for non-trading days
  - masterdata enrichment for reporting use

- **Reporting-ready outputs**
  - main Excel workbook with raw and masterdata sheets
  - standalone masterdata workbook
  - interactive HTML FX dashboard

- **Executive reporting focus**
  - rate trend views
  - daily / weekly / month-end views
  - latest movement monitoring
  - prior-month comparison
  - rolling 30-day volatility analysis

## Outputs
Each execution generates the following deliverables:

1. **Main Excel Workbook**  
   Includes the raw FX dataset and the enriched masterdata sheet.

2. **Standalone Masterdata Workbook**  
   Contains only the reporting-ready masterdata output.

3. **Interactive HTML Dashboard**  
   A multi-page dashboard for reviewing FX trends, movements, volatility, and summary insights.

## Why This Project Was Built
In many treasury environments, FX reporting is still partially manual, fragmented across spreadsheets, and highly dependent on repetitive data cleaning steps. This project was designed to solve that problem by creating a more controlled and repeatable workflow.

The goal is not only to generate data outputs, but also to produce a reporting package that is already structured for business review, treasury analysis, and dashboard-style presentation.

## High-Level Workflow
1. Start the engine
2. Choose the data source mode
3. Load FX data from API or Excel file
4. Standardize and validate the input structure
5. Build a complete daily calendar across currencies
6. Apply forward-fill for missing dates
7. Enrich the dataset with masterdata attributes and analytics
8. Export Excel outputs
9. Generate the HTML dashboard
10. Save all deliverables into the output folder

## Technology Stack
- **Python**
- **pandas**
- **requests**
- **openpyxl**
- **HTML / JavaScript dashboard output**

## Typical Use Case
This project is suitable for treasury, finance, and reporting teams that need a controlled way to generate FX overview reports from either automated market data or internal prepared datasets.

It is especially useful when the reporting objective is to create:
- standardized FX parity tables
- reusable masterdata for downstream reporting
- dashboard-ready FX trend analysis
- presentation-friendly reporting outputs with minimal manual intervention

## Repository Contents
This repository may include:
- the main reporting engine script
- technical documentation
- user runbook / execution guide
- generated sample outputs or reference materials

## Notes
This project is designed as a practical reporting engine rather than a generic FX analytics library. Its structure, outputs, and workflow are aligned with treasury reporting needs and presentation-ready deliverables.
