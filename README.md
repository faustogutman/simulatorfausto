Mortgage and Property Comparison Calculator (Tkinter & Pandas)
This Python application provides a robust and user-friendly tool for comparing different properties and their associated mortgage scenarios. Built with Tkinter for the graphical user interface and pandas for data handling and calculations, it allows users to input property details, explore various loan scenarios, and export detailed reports to Excel.

Features
Multi-Property Comparison: Manage and analyze up to three different properties simultaneously using an intuitive tabbed interface.

Detailed Property Inputs: Input essential data for each property, including:

Alias & Link

Property Price

Area (Square Meters)

Loan-to-Value (LTV) Percentage

Expected Monthly Rent

Options to skip Purchase Tax and Broker Fees.

Multiple Loan Scenarios: For each property, define up to three distinct loan scenarios by specifying:

Annual Interest Rate

Loan Term (in years)

Comprehensive Calculations: The application automatically calculates:

Estimated Purchase Tax

Required Down Payment

Estimated Lawyer Fees

Estimated Broker Fees

Total Capital Required

Price per Square Meter

Loan Amount

Monthly Mortgage Payments for each scenario

Total Interest and Total Payment over the loan term for each scenario

Amortization Tables & Graphs: Visualize the breakdown of principal and interest payments over time for each loan scenario with detailed amortization tables and interactive plots.

Rent-to-Mortgage Comparison: See a clear ratio of expected monthly rent to the calculated monthly mortgage payment for each scenario.

Data Persistence (CSV):

Save Inputs: Export all property input data to a CSV file for easy saving and sharing.

Load Inputs: Import previously saved CSV data to quickly populate your property tabs.

Export to Excel (Comprehensive Report): Generate a detailed Excel report (.xlsx) that includes:

Separate Sheet per Property: Each property gets its own dedicated sheet in the Excel workbook.

Full Details per Property: Within each property's sheet, you'll find:

A summary of all input fields and general property cost calculations.

All three loan scenarios, each with its:

Detailed amortization table.

Rent comparison string.

Embedded amortization plot.

Technologies Used
Python 3.x

Tkinter: For the graphical user interface.

Pandas: For efficient data manipulation and calculations.

Matplotlib: For generating amortization plots.

Pillow (PIL Fork): For handling image processing to embed plots in Excel.

Openpyxl: The backend engine used by Pandas for writing to .xlsx files and for directly embedding images.

<img width="1039" height="901" alt="image" src="https://github.com/user-attachments/assets/8f9d1a9d-3325-4fa4-813f-8bfc08e373e4" />


<img width="926" height="1009" alt="image" src="https://github.com/user-attachments/assets/01b054a8-61e8-4532-8273-bedf398adba1" />


<img width="836" height="1030" alt="image" src="https://github.com/user-attachments/assets/6ac25a03-b453-4258-a763-24834063578f" />
