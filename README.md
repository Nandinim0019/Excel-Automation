# Excel-Automation
Excel Automation using Python (openpyxl)

This project automates the process of calculating sales tax and total amount for each order in an Excel sheet using Python. It demonstrates how to manipulate Excel files programmatically with openpyxl, apply business logic with user-defined functions, and efficiently process tabular data stored in dictionary structures.

📂 Project Overview

The script reads an existing Excel sheet containing order details such as:

Order ID
Customer ID
Order Date
Subtotal
Location
Items Ordered

It then:
Calculates the tax for each order using a custom function.
Computes the total amount (Subtotal + Tax).
Writes the computed values back into the Excel sheet.
This automation eliminates manual calculations and improves efficiency and accuracy in data entry.

⚙️ Technologies Used
-Python 3
-Openpyxl package (for Excel manipulation)
-Dictionary & Loop Structures (for data processing)
-User-defined functions

🧠 Key Features

-Reads and writes data directly from Excel workbooks
-Automates tax and total column calculations
-Uses dictionaries for structured data handling
-Error-free and repeatable data processing logic

📘 Usage Instructions
Install dependencies:
pip install openpyxl
Place your input Excel file (e.g., sales_orders.xlsx) in the same directory.
Run the Python script or Jupyter notebook:
(or open the .ipynb notebook and execute the cells)

The updated Excel file will include the Tax and Total columns filled automatically.

🧾 Example Output
Order_ID	Customer_ID	Subtotal	Tax	Total
100001	C00007	899.97	62.99	962.96
100002	C00015	799.97	55.99	855.96
📈 Learning Outcomes

Automating Excel operations using openpyxl
Applying functions, loops, and dictionaries to manage structured data
Improving productivity by replacing manual Excel updates with Python scripts

👩‍💻 Author

Nandini Naik
Mini Project – Excel Data Automation using Python
(For learning and demonstration purposes)
