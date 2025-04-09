import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv
import logging

# Load environment variables
load_dotenv()

# Email configuration
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS", "tapiwamwedzi7@mail.com")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "xqmmexvkpqsrijyr")

# Create payslip directory if it doesn't exist
os.makedirs("payslips", exist_ok=True)

# Set up logging to log errors and successes
logging.basicConfig(filename="payslip_generator.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

def generate_payslip(employee):
    """Generate a payslip PDF for a single employee."""
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
        pdf.ln(10)

        pdf.cell(200, 10, txt=f"Employee ID: {employee['Employee ID']}", ln=True)
        pdf.cell(200, 10, txt=f"Name: {employee['Name']}", ln=True)
        pdf.cell(200, 10, txt=f"Basic Salary: ${employee['Basic Salary']:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Allowances: ${employee['Allowances']:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Deductions: ${employee['Deductions']:.2f}", ln=True)
        pdf.cell(200, 10, txt=f"Net Salary: ${employee['Net Salary']:.2f}", ln=True)

        filename = f"payslips/{employee['Employee ID']}.pdf"
        pdf.output(filename)
        logging.info(f"Payslip generated for {employee['Name']} (ID: {employee['Employee ID']})")
        return filename
    except Exception as e:
        logging.error(f"Error generating payslip for {employee['Name']} (ID: {employee['Employee ID']}): {e}")
        raise

def send_email(receiver_email, file_path, name):
    """Send email with payslip attached."""
    try:
        yag = yagmail.SMTP(EMAIL_ADDRESS, EMAIL_PASSWORD)
        subject = "Your Payslip for This Month"
        body = f"Hi {name},\n\nPlease find attached your payslip for this month.\n\nRegards,\nHR Department"
        yag.send(to=receiver_email, subject=subject, contents=body, attachments=file_path)
        logging.info(f"Email sent to {name} ({receiver_email})")
    except Exception as e:
        logging.error(f"Failed to send email to {receiver_email}: {e}")
        raise

def main():
    try:
        # Prompt for Excel file path to make it dynamic
        excel_file = input("Enter the path to the employees Excel file: ")

        if not os.path.isfile(excel_file):
            logging.error(f"Excel file not found at {excel_file}")
            print("The specified Excel file does not exist. Exiting.")
            return

        df = pd.read_excel(excel_file)

        # Ensure required columns exist
        required_columns = {"Employee ID", "Name", "Email", "Basic Salary", "Allowances", "Deductions"}
        if not required_columns.issubset(df.columns):
            logging.error("Excel file is missing required columns.")
            print("The Excel file is missing required columns. Exiting.")
            return

        # Process each employee
        for _, row in df.iterrows():
            try:
                row['Net Salary'] = row['Basic Salary'] + row['Allowances'] - row['Deductions']
                pdf_path = generate_payslip(row)
                send_email(row['Email'], pdf_path, row['Name'])
            except Exception as e:
                logging.error(f"Error processing employee {row['Employee ID']}: {e}")

        logging.info("All payslips processed successfully.")

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        print("An unexpected error occurred. Please check the log for details.")

if __name__ == "__main__":
    main()
