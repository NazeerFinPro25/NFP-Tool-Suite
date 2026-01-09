📊 NFP Tool Suite | NazeerFinPro

Advanced Financial Solutions & Automation

Welcome to the official repository for the NFP Tool Suite. This application is a comprehensive financial automation portal designed to streamline repetitive HR and Finance tasks for small and medium businesses.

👇 Live App: https://nfp-tools.streamlit.app

🚀 Key Features

1. 🏢 Auto-Attendance Generator

One-Click Automation: Converts raw biometric Excel data into formatted, payroll-ready attendance sheets instantly.

Smart Time Generation: Automatically fills missing times with natural-looking variations (e.g., 09:03, 18:10) to save hours of manual entry.

Holiday Management: Built-in gazetted holiday handling that marks holidays correctly on the sheets.

Customizable: Set your own Company Name and Month/Year for the report.

2. 🧾 Bulk GST Invoice Maker

Bulk Processing: Upload a single Sales Register file containing multiple Delivery Challans (DCs).

Auto-Generation: Generates a professional, print-ready PDF/HTML invoice for every single DC automatically.

Tax Calculation: Automatically calculates 18% GST (customizable rate) and Grand Totals.

Amount in Words: Automatically converts the final amount into words (e.g., "Fifty Thousand Rupees Only").

Fully Customizable Header: Set your own Company Name, Address, NTN, Logo, and Contact details directly from the app interface.

3. 🇵🇰 Pakistan Salary Tax Calculator (2025-26)

Updated FBR Slabs: Accurate tax calculation based on the latest FBR Tax Slabs for Tax Year 2025-2026.

Instant Breakdown: Enter a monthly gross salary to see the Annual Tax, Monthly Deduction, and Net Salary instantly.

🛠️ Installation & Usage

Clone the repository:

git clone [https://github.com/NazeerFinPro25/NFP-Tool-Suite.git](https://github.com/NazeerFinPro25/NFP-Tool-Suite.git)


Install Requirements:
Make sure you have Python installed, then run:

pip install -r requirements.txt


Run the Application:
To launch the app, use the following command:

python -m streamlit run app.py


(Alternative command if Streamlit is in your PATH: streamlit run app.py)

📂 Input File Formats (Templates)

The app requires specific Excel formats to work correctly. You can download sample templates directly from the app interface or use the structure below:

1. Attendance Data (data.xlsx):

Columns: S#, CODE, NAME, ABSENT DAYS, Overtime Hours

2. Sales Register (sales_register.xlsx):

Columns: DC No., Invoice No., Invoice Date, Customer Name, Bill To Address, Customer NTN, Credit Terms, Item Description, H.S Code, UOM, Qty, Unit Price (PKR), Total Value (PKR)

👨‍💼 Author

Nazeer Ahmed Khan Founder, NazeerFinPro

Email: nazeerfinpro@gmail.com

WhatsApp: +92 333 3126614

LinkedIn: Nazeer Fin Pro

© 2025 NazeerFinPro. All Rights Reserved.