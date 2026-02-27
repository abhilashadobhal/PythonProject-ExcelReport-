Project Title

Automated Excel Report Generator from CSV Data (Python-Based Reporting System)

Objective

To develop a robust Python application capable of ingesting raw sales data from a CSV file, performing complex data analysis (aggregation, pivoting), generating data visualizations, and exporting the complete, styled analysis into a single, professional Microsoft Excel report (.xlsx format).

Report Structure and Analysis

The generated Excel report is structured across two sheets for comprehensive analysis:

1. Summary Sheet

Overall Sales Statistics: Provides high-level metrics for the SALES column, including Count, Mean, Standard Deviation, Min, Max, and Total Sum.

Sales by Product Line (Pivot Table): Ranks product lines by their total sales revenue to identify top performers.

Visualization: An embedded Bar Chart visually representing the Total Sales by Product Line.

. Country_Analysis Sheet

Average Sales by Country and Deal Size (Pivot Table): Offers a multi-dimensional view of sales performance, showing the average transaction value segmented by the customer's COUNTRY and the size of the DEALSIZE (Small, Medium, Large).

Technical Achievements

End-to-End Automation: Successful integration of data handling (pandas), visualization (matplotlib), and advanced file formatting (openpyxl).

Robust Data Formatting: Implementation of explicit string casting to handle currency formatting ($X,XXX.XX) and prevent future pandas data type warnings.

Professional Styling: Application of borders, bold titles, customized column widths, and right-aligned numeric data within the Excel sheets for enhanced readability and professionalism.
