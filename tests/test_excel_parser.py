#!/usr/bin/env python3
"""
Unit tests for the Excel parser module.
"""

import os
import sys
import unittest
import tempfile
from datetime import datetime

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.excel_parser import ExcelParser, DataValidator


class TestExcelParser(unittest.TestCase):
    """Test cases for the ExcelParser class."""

    def setUp(self):
        """Set up test fixtures."""
        self.parser = ExcelParser()

    def test_detect_lab_name_bv(self):
        """Test lab name detection for BV."""
        self.assertEqual(self.parser.detect_lab_name("BV_Invoice_2024.xlsx"), "BV")
        self.assertEqual(self.parser.detect_lab_name("invoice_bv_q1.xlsx"), "BV")

    def test_detect_lab_name_its(self):
        """Test lab name detection for ITS."""
        self.assertEqual(self.parser.detect_lab_name("ITS_Invoice_2024.xlsx"), "ITS")

    def test_detect_lab_name_tuv(self):
        """Test lab name detection for TUV."""
        self.assertEqual(self.parser.detect_lab_name("TUV_Report.xlsx"), "TUV")

    def test_detect_lab_name_sgs(self):
        """Test lab name detection for SGS."""
        self.assertEqual(self.parser.detect_lab_name("SGS_data.xlsx"), "SGS")

    def test_detect_lab_name_not_found(self):
        """Test lab name detection when not found."""
        self.assertIsNone(self.parser.detect_lab_name("unknown_lab.xlsx"))
        self.assertIsNone(self.parser.detect_lab_name("invoice.xlsx"))

    def test_custom_lab_names(self):
        """Test with custom lab names."""
        parser = ExcelParser(lab_names=["CUSTOM", "LAB"])
        self.assertEqual(parser.detect_lab_name("CUSTOM_file.xlsx"), "CUSTOM")
        self.assertIsNone(parser.detect_lab_name("BV_file.xlsx"))


class TestDataValidator(unittest.TestCase):
    """Test cases for the DataValidator class."""

    def setUp(self):
        """Set up test fixtures."""
        self.allowed_types = [
            "Benchmark Test", "PVT", "PT", "ET", "VT", "ECR"
        ]
        self.validator = DataValidator(self.allowed_types)

    def test_valid_record(self):
        """Test validation of a valid record."""
        record = {
            "tajan_bidding_tracking_number": "TBT-123456",
            "test_service_type": "PVT"
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertTrue(is_valid)
        self.assertEqual(len(errors), 0)

    def test_missing_tracking_number(self):
        """Test validation with missing tracking number."""
        record = {
            "tajan_bidding_tracking_number": "",
            "test_service_type": "PVT"
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertFalse(is_valid)
        self.assertIn("Missing Tracking Number", errors)

    def test_invalid_test_type(self):
        """Test validation with invalid test type."""
        record = {
            "tajan_bidding_tracking_number": "TBT-123456",
            "test_service_type": "Invalid Type"
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertFalse(is_valid)
        self.assertTrue(any("Invalid Test Type" in e for e in errors))

    def test_missing_test_type(self):
        """Test validation with missing test type."""
        record = {
            "tajan_bidding_tracking_number": "TBT-123456",
            "test_service_type": ""
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertFalse(is_valid)
        self.assertIn("Missing Test/Service Type", errors)

    def test_multiple_errors(self):
        """Test validation with multiple errors."""
        record = {
            "tajan_bidding_tracking_number": "",
            "test_service_type": "Invalid Type"
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertFalse(is_valid)
        self.assertEqual(len(errors), 2)

    def test_case_insensitive_test_type(self):
        """Test that test type validation is case insensitive."""
        record = {
            "tajan_bidding_tracking_number": "TBT-123456",
            "test_service_type": "pvt"  # lowercase
        }
        is_valid, errors = self.validator.validate_record(record)
        self.assertTrue(is_valid)

    def test_validate_records_batch(self):
        """Test batch validation of records."""
        records = [
            {"tajan_bidding_tracking_number": "TBT-1", "test_service_type": "PVT"},
            {"tajan_bidding_tracking_number": "", "test_service_type": "PVT"},
            {"tajan_bidding_tracking_number": "TBT-3", "test_service_type": "Invalid"},
            {"tajan_bidding_tracking_number": "TBT-4", "test_service_type": "ET"},
        ]

        valid, invalid = self.validator.validate_records(records)

        self.assertEqual(len(valid), 2)
        self.assertEqual(len(invalid), 2)


if __name__ == "__main__":
    unittest.main()
