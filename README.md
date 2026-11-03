# CartCheck

**Food Delivery Fees and Discount Analysis in Microsoft Excel**

CartCheck checks whether a food-delivery discount still saves money after delivery, packaging, platform and other charges are added. The project combines an Excel dashboard, editable bill formulas, three native PivotTables, six charts, a fee planner and data quality checks.

The study uses **100 simulated orders** and three fictional platforms: QuickBite, MealCart and DineDrop. Findings describe this demonstration dataset.

[Open the Excel workbook](CartCheck_Excel_Project.xlsx) · [Read the project report](docs/CartCheck_Project_Report.pdf) · [Download the input CSV](data/CartCheck_Input_Data.csv)

## Dashboard preview

![CartCheck dashboard with headline measures, average platform fees and monthly final bills](images/dashboard-overview.png)

## Main findings

| Measure | Result |
| --- | ---: |
| Orders analysed | 100 |
| Listed food subtotals | INR 34,020 |
| Discounts | INR 6,750 |
| Extra charges | INR 7,634 |
| Final bills | INR 34,904 |
| Net saving against listed subtotals | **INR -884** |
| Positive / break-even / negative orders | 40 / 2 / 58 |
| Orders above the upper fee fence | 4 |

Extra charges exceed discounts by INR 884. This explains why total final bills are INR 884 above the listed food subtotals. The comparison baseline is the food subtotal recorded for each order; the data does not contain matched restaurant or competitor prices.

## Open the project

1. Download this repository using **Code > Download ZIP**, then extract it.
2. Open **CartCheck_Excel_Project.xlsx** in Microsoft Excel.
3. Start with **Dashboard**, then inspect **Orders** and **Analysis**.
4. Use **Fee Planner** for a hypothetical input demonstration.
5. After changing study records in **Orders**, choose **Data > Refresh All** to refresh the PivotTables.

The delivered analysis runs inside the workbook. Its source logic is the Excel formulas; using it does not require a Python program, VBA macro, database or web server. Rehearse the live demonstration in Microsoft Excel before presenting.

## Workbook design

| Worksheet | Purpose |
| --- | --- |
| Dashboard | Main measures and six editable charts |
| Fee Planner | Independent hypothetical bill with editable inputs |
| Analysis | Formula summaries, statistics and three native PivotTables |
| Orders | 100 input records and row calculations |
| Data Quality | Nine summary checks and row-level reviews |
| Project Guide | Formula examples, refresh instructions and maintenance notes |

Study inputs are in **Orders A8:K107**. The **OrdersData** table occupies **A7:T107**. Calculated measures are in L:T, with additional review formulas in U:V. Amber cells identify editable inputs.

## Calculations

| Measure | Calculation |
| --- | --- |
| Extra charges | Delivery + packaging + platform + other charges |
| Final bill | Food subtotal - discount + extra charges |
| Net saving | Discount - extra charges |
| Net saving rate | Net saving / food subtotal |

For the first order, CC-001:

- Food subtotal: INR 200
- Discount: INR 60
- Fees: INR 20 + 10 + 5 + 10 = INR 45
- Final bill: INR 185
- Net saving: INR 15, or 7.5% of the subtotal

Actual formulas in **Orders row 8**:

```excel
M8 =SUM(H8:K8)
N8 =F8-G8+M8
O8 =G8-M8
P8 =IFERROR(O8/F8,"n.a.")
S8 =IF(O8>0,"Positive",IF(O8<0,"Negative","Break-even"))
L8 =VLOOKUP(C8,$W$8:$X$10,2,FALSE)
```

Click a calculated cell to inspect the formula bar, or select **Formulas > Show Formulas**. The [formula reference](docs/CartCheck_Formula_Reference.txt) lists representative formulas and their locations. The workbook is the complete source.

## Excel techniques

- SUM, IF, nested IF, AND, OR, NOT and IFERROR
- VLOOKUP, SUMIF, AVERAGEIF, COUNTIF and COUNTIFS
- CONCATENATE and four named ranges
- Dropdowns, numeric validation and conditional formatting
- Sorting, filtering and typed date/currency/percentage formats
- PivotTables, charts, central tendency and IQR analysis

## PivotTables and charts

The three native PivotTables share **OrdersData** as their source and are placed on **Analysis**.

| PivotTable | Result range | Group | Aggregation |
| --- | --- | --- | --- |
| FeesByPlatform | B32:C36 | Platform | Average extra charges |
| MonthlyFinalBills | B42:C47 | Month | Sum of final bill |
| SavingOutcomeCounts | B53:C57 | Saving outcome | Count of order ID |

Monthly grouping uses a month helper generated from the numeric order date. Three dashboard charts read PivotTable result cells; the other three read formula summaries. Refresh the PivotTables after editing Orders.

![CartCheck comparisons of saving outcomes, average net saving, fee components and fee distribution](images/dashboard-comparisons.png)

## Statistical review

- Mean extra charges: **INR 76.34**
- Median extra charges: **INR 69.00**
- Mode of delivery fee: **INR 20**, occurring in 15 orders
- Q1 of extra charges: **INR 55.75**
- Q3 of extra charges: **INR 85.00**
- IQR: **INR 29.25**
- Upper fee fence: **85 + 1.5 × 29.25 = INR 128.875**

Four orders exceed the unrounded upper fence. They remain included in the totals. A high fee is a review flag; it is not automatically a data error or evidence of fraud. The delivery-fee mode describes a different variable from the extra-charge mean and median.

## Live demonstration

![CartCheck Fee Planner showing editable inputs and calculated results](images/fee-planner.png)

1. In **Fee Planner**, leave C8:C13 at **200, 60, 20, 10, 5 and 10**.
2. Observe extra charges of **INR 45**, final bill of **INR 185** and net saving of **INR 15**.
3. Change the delivery fee in **C10 from 20 to 30**.
4. Observe extra charges of **INR 55**, final bill of **INR 195** and net saving of **INR 5**.
5. Restore C10 to **20**.

The Fee Planner is separate from the 100 study orders, so this demonstration changes only the hypothetical bill. With starting fees of INR 45 and target saving C15 of INR 20, the required discount is INR 65.

## Data quality and limitations

The saved workbook has zero issues or reconciliation differences for its nine checks: missing input cells, duplicate-affected rows, unmapped codes, invalid bill inputs, out-of-period dates, bill reconciliation, outcome counts, fee-band counts and delivery-fee frequency coverage.

The study covers **1 July to 23 October 2026**, with 25 simulated orders in each of the four months. These records do not establish real platform performance or customer demand. Order baskets differ and the data does not control for restaurant, distance or service quality.

The current analysis uses fixed ranges for 100 records. Adding orders, months or fee denominations requires the range and list updates described in the report and Project Guide, followed by a PivotTable refresh. Do not remove unusual fees merely because they are large.

## Project files

| File | Description |
| --- | --- |
| [Excel workbook](CartCheck_Excel_Project.xlsx) | Complete editable analysis and source formulas |
| [Input CSV](data/CartCheck_Input_Data.csv) | Raw simulated order inputs |
| [Project report PDF](docs/CartCheck_Project_Report.pdf) | 26-page college report with screenshots, methods and results |
| [Project report Word](docs/CartCheck_Project_Report.docx) | Editable report |
| [Demo and viva guide PDF](docs/CartCheck_Demo_and_Viva_Guide.pdf) | Presentation preparation and viva questions |
| [Demo and viva guide Word](docs/CartCheck_Demo_and_Viva_Guide.docx) | Editable demonstration guide |
| [Formula reference](docs/CartCheck_Formula_Reference.txt) | Representative Excel formulas with cell locations |

## Author

**Prathamesh Prabhakar Bodhale**  
MCA Artificial Intelligence and Machine Learning, Chandigarh University  
Subject: Data Analytics  
Project guide: Mrs. Deepali Saini

## Method references

The report includes the full references. Key method sources:

- [Microsoft Excel functions by category](https://support.microsoft.com/en-us/excel/excel-functions-by-category)
- [Create a PivotTable](https://support.microsoft.com/en-us/excel/get-started/create-a-pivottable-to-analyze-worksheet-data)
- [Refresh PivotTable data](https://support.microsoft.com/en-us/excel/refresh-pivottable-data)
- [NIST box plot and IQR method](https://itl.nist.gov/div898/handbook/eda/section3/boxplot.htm)
