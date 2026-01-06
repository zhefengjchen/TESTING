#!/usr/bin/env python3
"""
Generate sample Excel files for testing the Invoice Management System.
"""

import os
from datetime import datetime, timedelta
import random
from openpyxl import Workbook


# Sample data for generating test invoices
SAMPLE_ASINS = [
    "B08N5WRWNW", "B09V3KXJPB", "B07ZPKN6YR", "B08HR5SXPS",
    "B0B8BJDL9H", "B09JQMJHXY", "B07S829LBX", "B08L5VN68Y"
]

SAMPLE_BRANDS = [
    "AmazonBasics", "Anker", "JBL", "Sony", "Samsung",
    "Apple", "LG", "Bose", "Philips", "Logitech"
]

SAMPLE_PRODUCTS = [
    "Wireless Earbuds", "USB-C Cable", "Bluetooth Speaker",
    "Power Bank 10000mAh", "Smart Watch", "Laptop Stand",
    "Webcam HD 1080p", "Mechanical Keyboard", "Gaming Mouse",
    "USB Hub 4-Port", "Phone Case", "Screen Protector"
]

TEST_TYPES = [
    "Benchmark Test", "Usability Test", "Performance Protocol",
    "PVT", "PT", "ET", "VT", "Packaging", "ECR", "PT-Initial"
]

DEVELOPMENT_CENTERS = ["SEA", "LUX/EU", "JP"]

LOCATIONS = ["Shenzhen", "Shanghai", "Hong Kong", "Taiwan", "Vietnam"]


def generate_tracking_number():
    """Generate a random tracking number."""
    return f"TBT-{random.randint(100000, 999999)}"


def generate_invoice_number(lab_name):
    """Generate a random invoice number."""
    return f"{lab_name}-INV-{random.randint(1000, 9999)}"


def generate_report_number(lab_name):
    """Generate a random report number."""
    return f"{lab_name}-RPT-{random.randint(10000, 99999)}"


def random_date(start_date, end_date):
    """Generate a random date between start and end."""
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    return start_date + timedelta(days=random_days)


def generate_sample_excel(lab_name: str, num_records: int = 20, output_dir: str = "tests"):
    """Generate a sample Excel file for testing."""

    wb = Workbook()

    # Create TESTING+SERVICES sheet
    ws = wb.active
    ws.title = "TESTING+SERVICES"

    # Headers
    headers = [
        "Tajan Bidding Tracking Number",
        "Amazon Test Request #",
        "Amazon Tracker Number",
        "ASIN#",
        "Development Center (SEA, LUX/EU, JP)",
        "Category (New HCC)",
        "Product Brand",
        "Product Description",
        "Testing SLA (Working days)",
        "Test / Service Type",
        "Quotation/Order Number",
        "Request Date",
        "Test Start Date",
        "Report Delivered Date",
        "Report number",
        "Test / inspection Location",
        "Product line",
        "AMAZON QUALITY MANAGER",
        "AMAZON SOURCING MANAGER",
        "Invoice #",
        "Lab contact",
        "COMMENT (IF ANY)",
        "AMOUNT (Currency=USD)"
    ]

    ws.append(headers)

    # Generate sample records
    base_date = datetime(2024, 1, 1)
    end_date = datetime(2024, 12, 31)

    for i in range(num_records):
        request_date = random_date(base_date, end_date)
        start_date = request_date + timedelta(days=random.randint(1, 5))
        delivery_date = start_date + timedelta(days=random.randint(5, 30))

        # Occasionally create records with validation issues
        tracking_number = generate_tracking_number()
        test_type = random.choice(TEST_TYPES)

        if i % 5 == 0:  # Every 5th record has missing tracking number
            tracking_number = ""
        if i % 7 == 0:  # Every 7th record has invalid test type
            test_type = "Invalid Test Type XYZ"

        row = [
            tracking_number,
            f"ATR-{random.randint(10000, 99999)}",
            f"ATN-{random.randint(100000, 999999)}",
            random.choice(SAMPLE_ASINS),
            random.choice(DEVELOPMENT_CENTERS),
            f"HCC-{random.randint(100, 999)}",
            random.choice(SAMPLE_BRANDS),
            random.choice(SAMPLE_PRODUCTS),
            random.choice([5, 10, 15, 20, 30]),
            test_type,
            f"QO-{random.randint(1000, 9999)}",
            request_date.strftime("%Y-%m-%d"),
            start_date.strftime("%Y-%m-%d"),
            delivery_date.strftime("%Y-%m-%d"),
            generate_report_number(lab_name),
            random.choice(LOCATIONS),
            f"Product Line {random.choice(['A', 'B', 'C', 'D'])}",
            f"QM_{random.randint(1, 10)}@amazon.com",
            f"SM_{random.randint(1, 10)}@amazon.com",
            generate_invoice_number(lab_name),
            f"{lab_name.lower()}_contact@{lab_name.lower()}.com",
            random.choice(["", "Rush order", "Priority", "Standard"]),
            round(random.uniform(100, 5000), 2)
        ]

        ws.append(row)

    # Add some footer rows that should be skipped
    ws.append([])
    ws.append(["Subtotal", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", sum(random.uniform(100, 5000) for _ in range(num_records))])
    ws.append(["Tax Rate", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "5%"])
    ws.append(["Sales Tax", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 250.00])
    ws.append(["Total", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5250.00])

    # Create another sheet (should be ignored)
    ws2 = wb.create_sheet("Summary")
    ws2.append(["This sheet should be ignored"])

    # Save the file
    filename = f"{lab_name}_Invoice_Sample_{datetime.now().strftime('%Y%m%d')}.xlsx"
    filepath = os.path.join(output_dir, filename)

    # Ensure directory exists
    os.makedirs(output_dir, exist_ok=True)

    wb.save(filepath)
    print(f"Generated: {filepath}")
    return filepath


def main():
    """Generate sample files for all labs."""
    labs = ["BV", "ITS", "TUV", "SGS"]

    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data")

    for lab in labs:
        generate_sample_excel(lab, num_records=25, output_dir=output_dir)

    print(f"\nSample files generated in: {output_dir}")
    print("You can use these files to test the Invoice Management System.")


if __name__ == "__main__":
    main()
