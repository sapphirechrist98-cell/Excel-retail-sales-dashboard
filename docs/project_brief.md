# Project Brief

## Project Name
Retail Sales Performance Analysis Using Excel

## Scenario
You are working as a junior data analyst for a retail company. The business wants to understand sales performance across regions, product categories, customer segments, and months.

Your task is to clean the sales data, analyze it using Excel, and create a dashboard that helps business users quickly understand performance.

## Dataset Columns
- `Order_ID` - unique order identifier
- `Order_Date` - date of purchase
- `Customer_Name` - customer name
- `Segment` - customer type
- `Region` - sales region
- `Category` - product category
- `Product` - product name
- `Quantity` - units sold
- `Unit_Price` - price per unit
- `Sales` - total sales amount
- `Cost` - total cost amount
- `Profit` - sales minus cost
- `Payment_Mode` - payment method used

## Excel Work Steps
1. Open `data/sales_data.csv` in Excel.
2. Convert the dataset into an Excel Table using `Ctrl + T`.
3. Check column data types, especially date and number columns.
4. Add a new column called `Month`.
5. Add a new column called `Profit_Margin`.
6. Create pivot tables for region, category, segment, and month analysis.
7. Create charts from the pivot tables.
8. Build one dashboard sheet using KPIs, charts, and slicers.
9. Write 4-6 business insights.

## Suggested Formulas

Month:
```excel
=TEXT([@Order_Date],"mmm-yyyy")
```

Profit Margin:
```excel
=IFERROR([@Profit]/[@Sales],0)
```

Average Order Value:
```excel
=Total Sales / Total Orders
```

## Skills Demonstrated
- Data cleaning
- Excel tables
- Pivot tables
- Pivot charts
- KPI reporting
- Dashboard design
- Business analysis
- Data storytelling
