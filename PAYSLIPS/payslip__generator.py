import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables (EMAIL_USER and EMAIL_PASS)
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Ensure payslips folder exists
os.makedirs("payslips", exist_ok=True)

# Read Excel file
try:
    df = pd.read_excel("employees.xlsx")
except FileNotFoundError:
    print("Error: employees.xlsx not found.")
    exit()

# Initialize email client
try:
    yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
except Exception as e:
    print("Failed to initialize email client:", e)
    exit()

# Generate and send payslips
for _, row in df.iterrows():
    try:
        emp_id = str(row['Employee ID'])
        name = row['Name']
        email = row['Email']
        basic = float(row['Basic Salary'])
        allowances = float(row['Allowances'])
        deductions = float(row['Deductions'])
        net_salary = basic + allowances - deductions

        # Create PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, f"Payslip for {name}", ln=True, align="C")

        pdf.set_font("Arial", size=12)
        pdf.ln(10)
        pdf.cell(100, 10, f"Employee ID: {emp_id}", ln=True)
        pdf.cell(100, 10, f"Basic Salary: ${basic:.2f}", ln=True)
        pdf.cell(100, 10, f"Allowances: ${allowances:.2f}", ln=True)
        pdf.cell(100, 10, f"Deductions: ${deductions:.2f}", ln=True)
        pdf.cell(100, 10, f"Net Salary: ${net_salary:.2f}", ln=True)

        payslip_path = f"payslips/{emp_id}.pdf"
        pdf.output(payslip_path)

        # Send email
        subject = "Your Payslip for This Month"
        body = f"Dear {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Department"
        yag.send(to=email, subject=subject, contents=body, attachments=payslip_path)

        print(f"Payslip sent to {name} at {email}")

    except Exception as e:
        print(f"Error processing {row.get('Name', 'Unknown')}: {e}")
