# Excel Project 
#### Coffee Order Sales 

## Table of Contents
- [Overview](#overview)
- [Project Structure](#project-structure)
- [Data Source](#data-source)
- [Exploratory Data Analysis (EDA)](#exploratory-data-analysis-eda)
- [Data Analysis](#data-analysis)
- [Interactive Dashboard](#interactive-dashboard)
- [Conclusion](#conclusion)

## Overview
This Excel project involves the analysis and visualization of sales data for a coffee company. The project utilizes three main sheets: `order`, `customer`, and `products`. The `order` sheet contains information about individual orders, including order details, customer details, and product details. The `customer` sheet holds customer information, while the `products` sheet stores data related to various coffee products.


## Project Structure
The project consists of three primary sheets:
- **Order Sheet:** Contains columns for Order ID, Order Date, Customer ID, Product ID, Quantity, Customer Name, Email, Country, Coffee Type, Roast Type, Size, Unit Price, Sales, Coffee Type Name, Roast Type Name, and Loyalty Card.
- **Customer Sheet:** Includes columns for Customer ID, Customer Name, Email, Phone Number, Address Line 1, City, Country, Postcode, and Loyalty Card.
- **Products Sheet:** Contains columns for Product ID, Coffee Type, Roast Type, Size, Unit Price, Price per 100g, and Profit.

## Data Source
The data for this project was sourced from GitHub. XLOOKUP function was used to fill in missing rows from the `customer` and `products` sheets into the `order` sheet, ensuring comprehensive data for analysis.

## Exploratory Data Analysis (EDA)
The following steps were taken during the EDA process:
- Checking for missing values and handling them appropriately, using XLOOKUP to fill in missing rows from `customer` and `products` sheets into the `order` sheet.
- Exploring the distribution of key variables such as sales, quantity, and unit price.
- Analyzing the relationship between variables, such as sales and product type, customer demographics, and loyalty card status.
- Identifying any outliers or anomalies in the data and addressing them as necessary.

## Data Analysis
The following steps were taken to analyze the data:
1. **Data Cleaning:** Checked for and corrected any inconsistencies, missing values, or errors in the dataset.
2. **Data Integration:** Utilized XLOOKUP function to integrate missing customer and product data into the order sheet for comprehensive analysis.
3. **Calculations:** Calculated total sales, profit margins, and other relevant metrics based on the available data.
4. **Sales Trend Analysis:** Analyzed sales trends over time to identify seasonal patterns or growth trends.
5. **Customer Analysis:** Identified top customers based on sales volume and loyalty card usage.
6. **Geographical Analysis:** Examined sales distribution by country to identify target markets or areas for improvement.

## Interactive Dashboard
An interactive dashboard was created to provide a user-friendly interface for exploring the data. The dashboard features:
- Visualization of order dates, roast type names, loyalty cards, and size.
- Total sales trend over time.
- Sales breakdown by country.
- Top 5 customers based on sales.

![coffeeSales_dashboard](https://github.com/Remmy69440/Excel/assets/159604919/2e4809f0-156d-4800-9535-98a4f396e374)

## Conclusion
This Excel project offers valuable insights into the sales performance of the coffee company. The interactive dashboard provides a convenient way to analyze different aspects of the data and identify key trends and patterns.
