Excell-file
Repository for the creation of the Excel file with the e-commerce exercise 

I have created an Excel file as the e-commerce exercise for the higher education courses shop that includes interconnected sheets for financial analysis. The file contains three main worksheets with structured data and relationships between key business assumptions.

File Structure

Key Assumptions in the Model:

Revenue Calculation

  Revenue is computed by multiplying sales volumes (ProductMix) by unit prices (Pricing).
  This is done via formulas in the Cash Flow sheet:
  Example: =SUM(ProductMix!B2:B10*Pricing!B2:B10) (January).

Expenses

  Fixed monthly expenses are assumed (15,000 in January, 16,000 in February, 17,000 in March).
  These are hardcoded but can be adjusted manually.

Product Mix (Sales Volumes)

  Courses A–I each have monthly sales values for Jan–Mar.
  Example: Course A sells 100, 120, 130 units in Jan–Mar.

Pricing

  Each course has a fixed price per unit ranging from 100 to 500.
  Example: Course A = 500, Course I = 100.

Cash Flow Structure

  Net Cash Flow = Revenue – Expenses.
  Cumulative Cash Flow adds each month’s net flow to the prior balance.


Key Features of the Code & Excel File

1. Three Linked Tabs

  Cash Flow: Consolidates revenue, expenses, net cash flow, and cumulative position.
  ProductMix: Monthly course sales volumes.
  Pricing: Price per course unit.

2. Interconnected with Formulas

  The Cash Flow sheet dynamically references both the ProductMix and Pricing sheets.
  If you update sales volumes or pricing, revenues and cash flows recalculate automatically.

3. Scalable Structure

  You can add more courses to ProductMix and Pricing, extend months, or modify expenses.
  The formulas will adapt if you extend ranges.

4. Excel Output

  The final Excel file (HigherEd_Courses_Business_Model.xlsx) is created with openpyxl.
  It contains ready-to-use links for scenario testing (e.g., change prices or volumes to see impacts).

5. Python Script Automation

  The .py file (create_highered_excel.py) automates the creation of the model.
  Running the script regenerates the Excel file, ensuring a repeatable process.

This model allows exploring various scenarios by adjusting key assumptions such as course volumes, pricing strategies, or cost structures.
