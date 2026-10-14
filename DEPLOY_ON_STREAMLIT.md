# Update citykart-gm-visibility.streamlit.app

This is the **online edition**. It generates the report on Streamlit Community Cloud without Windows, Microsoft Excel, pywin32 or COM. Downloads are **XLSX**, including a report with live business formulas and cached calculated values, plus a completely value-pasted summary.

## Update your existing app

1. Extract `GM_Visibility_Report_Portal_ONLINE.zip` on your computer.
2. Open the GitHub repository and branch currently connected to **citykart-gm-visibility.streamlit.app**. Use the app's Manage app panel to identify them if needed.
3. Replace the old portal files with the extracted files. Upload **all files and folders**, including `.streamlit`, `template`, `input_templates`, `buy_plan_templates` and `master_templates`. Keep `app.py` and `requirements.txt` at the repository root. If your existing entrypoint is in a subfolder, put the complete package in that folder and keep the configured entrypoint path consistent.
4. Commit the changes. The existing app normally rebuilds after the commit. If necessary, use **Manage app → Reboot app**.
5. If you deploy as a new app, select your repository/branch, main file **app.py**, and **Python 3.12** in Advanced settings. Use the `citykart-gm-visibility` subdomain if it belongs to your app and is available.
6. Open your site, upload fresh source data and both latest buy plans, then click **Generate visibility report**. Download the newly generated XLSX report and summary.

Do not retain the old Windows-only requirements or old `app.py`. Do not upload the ZIP itself as the app: Streamlit runs the extracted code from GitHub.

## Files and inputs

- Four source inputs: GRC_DATA, PPO_DATA, GM_TNA, MNL_PPO. Separate files or one combined workbook are supported.
- Both latest buy plans: Regular and Winter, with department/article sheets.
- Optional metadata workbook: Division Merchant, Article Master and Department Master.
- XLSX, XLSB and legacy XLS source/buy-plan/metadata uploads are supported. XLSX is recommended. Uploaded formula fields need saved calculated values: recalculate/save the workbook in Excel first if quantities are reported blank.
- Custom **formatting/report templates must be XLSX**. To reuse an XLSB report as a custom template, save it as XLSX in Excel first.
- Original uploaded sheets are **not copied unchanged** into the report. Normalized data is retained in hidden sheets for audit.

## Latest report rules retained

- Sort: Merchant A–Z → Division A–Z → Section A–Z → Department A–Z (Article reports only) → period Buy Plan largest to smallest within the group.
- Starting month controls the three regular months, combined caption (OND, NDJ, etc.), GRC date window and PPO buckets, including year rollover.
- First PPO bucket includes the first month and overdue pending quantities from all prior months. SCHEDULE DT means Delivery Date.
- Source-only articles are included with Buy Plan 0; listing comes from sources and buy plans.
- Source quantities, metadata, individual Buy Plans, combined Buy Plan and total system PPO are values. Capped quantities, fill rates, visibility, coverage, adjusted monthly plans and grand totals remain formulas.
- Article sheets have no Attribute column; Winter sheets have no historical STH columns.
- Each output sheet selects one Merchant cell. The report opens on REGULAR_DEPT.
- Summary contains exactly REGULAR, WINTER and BOTTOM, with all cells as values, font size 8, quantities without decimals and whole-percent display.
- Bottom 15 uses lowest first-month Visibility % among departments with Sale Rank 1–100 and a positive first-month Buy Plan.
- The GRC cutoff defaults to yesterday in India time. Report filename: `GM_Visibility_Report_Till_DD-Mmm-YY.xlsx`.

For a new period, upload plans with matching `MON_YYYY_BUY_PLAN` headers (e.g. NOV_2026, DEC_2026, JAN_2027). Winter uses the selected first month and WINTER_BUY_PLAN. Generate again after changing uploads or settings; Download alone does not recalculate.

## Validation and limits

Percentage cells use 0%; coverage quantities use #,##0. Colour rules retain their conditions, fills and fonts but do not override cell number formats with General.

The cloud engine calculates the bundled report's arithmetic, MIN, MAX, IF, IFERROR, SUM and SUBTOTAL formulas, writes cached results, and retains the formulas for Excel to recalculate when report input cells change. Custom templates using other formula functions are rejected with a clear message rather than returning unchecked results.

The standard XLSX styles, widths, row heights, merged headers, outline groups and conditional-format rules come from the converted reference template. Binary XLSB-only features and macros are not supported. Exact rendering can differ between Excel and other spreadsheet viewers.

The pipeline was tested on the reference's actual source data. Its cached business formulas were independently compared with LibreOffice recalculation. Month rollover, metadata joins, source-only zero plans, empty Winter populations, sorting and summary snapshots were also checked. The live Streamlit deployment must be updated through your connected GitHub repository; this ZIP does not itself deploy the site.

## Run locally, if needed

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

No Excel installation is needed for this edition.
