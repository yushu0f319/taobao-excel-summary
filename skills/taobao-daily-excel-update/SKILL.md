---
name: taobao-daily-excel-update
description: Use when the user wants to update a fixed daily Taobao scale-data Excel workbook from local Taobao export files. Handles the three Taobao source tables: TB门店, TB流量, and TB履约.
---

# Taobao Daily Excel Update

Use this skill to generate a new daily scale-data workbook from:

1. The previous completed workbook, such as `6月4日度-规模数据.xlsx`.
2. Three local Taobao export workbooks for the new date.

The output is a new Excel workbook with the Taobao source data appended to the fixed template sheets.

## Source Table Recognition

Do not rely on file names. Identify each source workbook from its headers:

- `TB门店`: contains `总成交额（元）`, `预计收入（元）`, `有效订单`, `顾客实付金额（元）`.
- `TB流量`: contains `总曝光人数`, `总进店人数`, `总下单人数`, `进店转化率`, `下单转化率`.
- `TB履约`: contains `履约渗透率`, `轨迹渗透率`, `及时送达率`, `流失订单`, `拣货及时订单率`.

If normal spreadsheet reading appears to show only `A1` or an empty table, still run the bundled script. Taobao exports may have an incorrect XLSX `dimension` value while the real rows are present in worksheet XML.

## Update Rules

For each recognized source table:

1. Locate the matching template sheet: `TB门店`, `TB流量`, or `TB履约`.
2. Append a new block at the end of the sheet:
   - one blank row
   - one source header row
   - all source data rows
3. Copy formatting from the previous block in the template sheet.
4. Convert values to template-compatible types:
   - `日期` -> number
   - `门店id` -> number
   - amounts, counts, ratings, durations, and rate decimals -> number
   - percent text such as `11.6%` -> `0.116` with the template percentage format
   - `总管id`, `供应商id`, and name fields -> text
5. Preserve formulas, workbook structure, images, hidden sheets, and WPS-specific content.
6. Set the workbook to recalculate formulas when opened in Excel or WPS.

## Run

Ask the user for:

- previous completed workbook path
- the three Taobao source workbook paths, or the folder containing them
- output workbook path

Then run:

```bash
python3 scripts/update_taobao_daily.py \
  --base "/path/to/previous.xlsx" \
  --input "/path/to/source-folder-or-file" \
  --output "/path/to/output.xlsx"
```

Resolve `scripts/update_taobao_daily.py` relative to this `SKILL.md` file before running it.
Use multiple `--input` arguments if files are in different locations.

## Validation

The bundled script must fail instead of guessing when:

- any of the three required source table types is missing
- multiple files match the same source table type
- the three source tables contain different dates
- source headers do not match the fixed template sheet headers
- written values do not match the source data after required type conversion

After a successful run, tell the user:

- output file path
- processed date
- which source file mapped to each sheet
- appended row range for each sheet
- whether value validation passed
