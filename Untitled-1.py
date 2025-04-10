import os
import pandas as pd
from fpdf import FPDF
import yagmail
from datetime import datetime
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Email configuration
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS", "tapiwamwedzi7@gmail.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "xqmmexvkpqsrijyr")

# Create payslip directory if it doesn't exist
os.makedirs("payslips", exist_ok=True)

def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Header with blue background
    pdf.set_fill_color(0, 102, 204)
    pdf.rect(0, 0, 210, 20, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Uncommon", ln=True, align='C')
    pdf.ln(10)

    # Reference Number
    today = datetime.now().strftime("%Y-%m-%d")
    ref = f"REF-{employee['Employee ID']}-{today.replace('-', '')}"
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(50)
    pdf.cell(0, 10, f"Reference: {ref}", ln=True, align='R')
    pdf.ln(5)

    # Employee Info
    pdf.set_text_color(0)
    pdf.set_font("Arial", '', 12)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(190, 10, "Employee Information", ln=True, fill=True)
    pdf.cell(190, 8, f"Employee ID: {employee['Employee ID']}", ln=True)
    pdf.cell(190, 8, f"Name: {employee['Name']}", ln=True)
    pdf.cell(190, 8, f"Email: {employee['Email']}", ln=True)
    pdf.ln(5)

    # Salary Breakdown
    pdf.set_font("Arial", 'B', 12)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(190, 10, "Salary Breakdown", ln=True, fill=True)
    pdf.set_font("Arial", '', 12)
    pdf.cell(95, 8, "Basic Salary", 1)
    pdf.cell(95, 8, f"${employee['Basic Salary']:.2f}", 1, ln=True)
    pdf.cell(95, 8, "Allowances", 1)
    pdf.cell(95, 8, f"${employee['Allowances']:.2f}", 1, ln=True)
    pdf.cell(95, 8, "Deductions", 1)
    pdf.cell(95, 8, f"${employee['Deductions']:.2f}", 1, ln=True)

    # Net Salary
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 13)
    pdf.set_fill_color(204, 255, 204)
    pdf.cell(190, 10, f"Net Salary: ${employee['Net Salary']:.2f}", ln=True, fill=True, align='C')

    # Signature
    pdf.ln(15)
    try:
        # Change to "signature.jpg" here if the image is in JPG format
        pdf.image(r"C:\Users\Tapiwa Mwedzi\Desktop\New folder\New folder\signature.jpg", x=10, y=pdf.get_y(), w=40)
        pdf.ln(25)
    except RuntimeError:
        pdf.cell(0, 10, "[Signature image missing]", ln=True)
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 6, "Authorized Signature", ln=True)

    # Note
    pdf.ln(10)
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(80)
    pdf.multi_cell(0, 8, "Note: Thank you for your continuous contribution to Uncommon. "
                         "If you have any questions regarding your payslip, please contact HR.")

    # Footer
    pdf.set_y(-15)
    pdf.set_font("Arial", "I", 8)
    pdf.set_text_color(100)
    pdf.cell(0, 10, f"Generated on {today} | HR Department - Uncommon", 0, 0, 'C')

    # Save file
    filename = f"payslips/{employee['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

def send_email(receiver_email, file_path, name):
    try:
        yag = yagmail.SMTP(EMAIL_ADDRESS, EMAIL_PASSWORD)
        subject = "Your Payslip for This Month"
        body = f"Hi {name},\n\nPlease find attached your payslip for this month.\n\nRegards,\nHR Department"
        yag.send(to=receiver_email, subject=subject, contents=body, attachments=file_path)
        print(f"Email sent to {name} ({receiver_email})")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}: {e}")

def main():
    try:
        df = pd.read_excel(r"C:/Users/Tapiwa Mwedzi/Desktop/New folder/New folder/employees.xlsx")

        # Ensure required columns exist
        required_columns = {"Employee ID", "Name", "Email", "Basic Salary", "Allowances", "Deductions"}
        if not required_columns.issubset(df.columns):
            raise ValueError("Excel file is missing required columns.")

        # Process each employee
        for _, row in df.iterrows():
            try:
                row['Net Salary'] = row['Basic Salary'] + row['Allowances'] - row['Deductions']
                pdf_path = generate_payslip(row)
                send_email(row['Email'], pdf_path, row['Name'])
            except Exception as e:
                print(f"Error processing employee {row['Employee ID']}: {e}")

    except FileNotFoundError:
        print("employees.xlsx not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()