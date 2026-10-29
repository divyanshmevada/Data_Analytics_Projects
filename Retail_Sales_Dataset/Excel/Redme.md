# Retail Sales Performance Analysis (Excel)

An end-to-end Excel analysis of two years of retail orders: data preparation with lookup formulas, 14 pivot tables, 9 charts and an interactive dashboard with slicers.

> **Dataset note:** The data is synthetic practice data created for this project. It does not represent a real company.

---

## 1. Objective

- Measure overall sales, profit and profit margin
- Find the best and worst performing categories, products, regions and states
- Understand seasonality and year-over-year growth
- Check the impact of discounts on profit
- Measure how many orders are lost to returns and cancellations

## 2. Dataset

| Sheet | Rows | Description |
|---|---|---|
| Orders | 1,200 | Orders from Jan 2024 to Dec 2025 |
| Products | 22 | Product name, category, cost price, unit price |
| Customers | 200 | Customer name, city, state, region, segment |
| Monthly_Targets_2025 | 5 | Monthly sales targets by region (wide format) |

**Order status:** Delivered 1,001 | Returned 83 | Cancelled 67 | Pending 49

## 3. Tools & Skills Used

Excel: `VLOOKUP`, `TEXT`, Excel Tables, Pivot Tables, Pivot Charts, Slicers, Dashboard design

## 4. What I Did (Before vs After)

### Before
The Orders sheet had only 9 raw columns: Order ID, Order Date, Customer ID, Product ID, Quantity, Discount, Ship Mode, Payment Mode, Status. There was no price, sales, cost or profit, and no product or customer details.

### After
The Orders sheet now has 23 columns (14 added), converted to an Excel Table:

| Added column | How it was created |
|---|---|
| `unit_price`, `cost_price` | `VLOOKUP` from the Products sheet |
| `Product Name`, `Category` | `VLOOKUP` from the Products sheet |
| `State`, `Region`, `segment` | `VLOOKUP` from the Customers sheet |
| `discount_amount` | `Quantity × unit_price × Discount` |
| `net_sales` | `Quantity × unit_price − discount_amount` |
| `total_cost` | `Quantity × cost_price` |
| `profit` | `net_sales − total_cost` |
| `Month`, `Months`, `year` | `TEXT()` on Order Date |

**Definition used in this project: Sales = `net_sales` (after discount).**

### Analysis steps
1. Built lookup columns and calculated sales, cost and profit
2. Verified results on sample rows (for example, Row 2: discount ₹859.80, net sales ₹16,336.20, profit ₹5,136.20)
3. Built 14 pivot tables (month, category, product, region, state, segment, status, payment mode, orders by state)
4. Built 9 charts from the pivots
5. Created a dashboard with 4 KPI cards and 6 slicers (Month, Year, Region, Status, Ship Mode, Category)

## 5. Dashboard

![Dashboard](images/dashboard.png)

- **KPI cards:** Total Sales, Total Profit, Total Orders, Average Order Value
- **Charts:** Month-wise sales and profit (line), category-wise profit, product-wise sales and profit, region-wise profit, state-wise sales and profit, state-wise orders, order status (pie), segment-wise profit (doughnut)
- **Slicers:** Month, Year, Region, Status, Ship Mode, Category

> By default the dashboard includes all order statuses. For business analysis, set the **Status slicer to Delivered + Pending**. All insights below use Delivered + Pending orders.

## 6. Key Insights

**Overall (Delivered + Pending, 1,050 orders)**

| Metric | Value |
|---|---|
| Total Sales | ₹69.27 lakh |
| Total Profit | ₹24.34 lakh |
| Profit Margin | 35.1% |
| Average Order Value | ₹6,597 |

1. **Sales grew 18.4% year over year.** 2025 sales were ₹37.56 lakh versus ₹31.71 lakh in 2024, and profit grew 15.9%.
2. **Q4 is the peak season.** October to December contributes 34% of annual sales. December 2025 was the best month (₹4.78 lakh).
3. **Furniture drives the business.** It gives 58.5% of sales and 52.5% of profit, led by Office Chair (₹3.55 lakh profit), Study Table and Filing Cabinet. Its margin (32%) is lower than the average.
4. **Clothing has the best margin (46%)** but only 16.7% of sales. Denim Jeans (45% margin) is the strongest clothing product. This is a growth opportunity.
5. **Office Supplies is small.** It contributes only 5.1% of sales. A4 Paper Ream has the lowest margin (27%).
6. **West is the top region** with ₹19.97 lakh sales (28.8%), and **Gujarat is the top state** with ₹14.09 lakh sales and 207 orders. East has the lowest sales (₹9.83 lakh) but the best margin (37%).
7. **Discounts reduce profit.** Margin falls from 39.1% with no discount to 25.7% at 20% discount. Total discount given was about ₹4.75 lakh.
8. **Returns and cancellations hurt revenue.** 6.9% of orders were returned and 5.6% cancelled, which is ₹7.98 lakh of sales (₹2.84 lakh profit) not realised.
9. **Consumer is the largest segment** with 60.8% of sales.

## 7. Recommendations

- Stock up and run campaigns before Q4, especially for Furniture
- Push Clothing, since it has the highest margin
- Limit discounts above 10% on Furniture, where margins are already lower
- Investigate the reasons behind returns and cancellations
- Look for ways to grow East and Central, the two weakest regions

## 8. Repository Structure

```
01_Retail_Sales_Performance_Analysis/
├── README.md
├── data/
│   └── Retail_Sales_Dataset.xlsx      (raw data)
├── excel/
│   └── Retail_Sales_Dashboard.xlsx    (analysis + dashboard)
└── images/
    └── dashboard.png
```

## 9. Next Steps

- Compare 2025 actual sales against `Monthly_Targets_2025` using Power Query
- Add the SQL version of this analysis (`sql/queries.sql`)
- Rebuild the dashboard in Power BI

## Author

**Divyansh Mevada** | [LinkedIn](https://linkedin.com/in/divyanshmevada) | [GitHub](https://github.com/divyanshmevada)
