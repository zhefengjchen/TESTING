#!/usr/bin/env python3
"""
Unit tests for the database module.
"""

import os
import sys
import unittest
import tempfile

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.database import Database


class TestDatabase(unittest.TestCase):
    """Test cases for the Database class."""

    def setUp(self):
        """Set up test fixtures."""
        # Create a temporary database file
        self.temp_dir = tempfile.mkdtemp()
        self.db_path = os.path.join(self.temp_dir, "test_invoices.db")
        self.db = Database(self.db_path)

    def tearDown(self):
        """Tear down test fixtures."""
        self.db.close()
        if os.path.exists(self.db_path):
            os.remove(self.db_path)
        os.rmdir(self.temp_dir)

    def test_database_creation(self):
        """Test that database and tables are created."""
        self.assertTrue(os.path.exists(self.db_path))

    def test_default_lab_names(self):
        """Test that default lab names are initialized."""
        lab_names = self.db.get_lab_names()
        self.assertIn("BV", lab_names)
        self.assertIn("ITS", lab_names)
        self.assertIn("TUV", lab_names)
        self.assertIn("SGS", lab_names)

    def test_default_test_types(self):
        """Test that default test types are initialized."""
        test_types = self.db.get_test_types()
        self.assertIn("Benchmark Test", test_types)
        self.assertIn("PVT", test_types)
        self.assertIn("ECR", test_types)
        self.assertEqual(len(test_types), 22)

    def test_add_lab_name(self):
        """Test adding a new lab name."""
        result = self.db.add_lab_name("NEW_LAB")
        self.assertTrue(result)
        self.assertIn("NEW_LAB", self.db.get_lab_names())

    def test_add_duplicate_lab_name(self):
        """Test that duplicate lab names are rejected."""
        result = self.db.add_lab_name("BV")
        self.assertFalse(result)

    def test_delete_lab_name(self):
        """Test deleting a lab name."""
        self.db.add_lab_name("TEST_LAB")
        result = self.db.delete_lab_name("TEST_LAB")
        self.assertTrue(result)
        self.assertNotIn("TEST_LAB", self.db.get_lab_names())

    def test_insert_invoice(self):
        """Test inserting an invoice record."""
        invoice_data = {
            "tajan_bidding_tracking_number": "TBT-123456",
            "amazon_test_request": "ATR-789",
            "asin": "B08N5WRWNW",
            "test_service_type": "PVT",
            "amount_usd": 1500.00,
            "lab_name": "BV",
            "invoice_date": "2024-01-15"
        }

        invoice_id = self.db.insert_invoice(invoice_data)
        self.assertIsNotNone(invoice_id)
        self.assertGreater(invoice_id, 0)

    def test_get_all_invoices(self):
        """Test retrieving all invoices."""
        # Insert test data
        for i in range(3):
            self.db.insert_invoice({
                "tajan_bidding_tracking_number": f"TBT-{i}",
                "lab_name": "BV"
            })

        invoices = self.db.get_all_invoices()
        self.assertEqual(len(invoices), 3)

    def test_update_invoice(self):
        """Test updating an invoice record."""
        invoice_id = self.db.insert_invoice({
            "tajan_bidding_tracking_number": "TBT-UPDATE",
            "amount_usd": 100.00
        })

        result = self.db.update_invoice(invoice_id, {"amount_usd": 200.00})
        self.assertTrue(result)

        invoices = self.db.get_all_invoices()
        updated = next(i for i in invoices if i['id'] == invoice_id)
        self.assertEqual(updated['amount_usd'], 200.00)

    def test_delete_invoice(self):
        """Test deleting an invoice record."""
        invoice_id = self.db.insert_invoice({
            "tajan_bidding_tracking_number": "TBT-DELETE"
        })

        result = self.db.delete_invoice(invoice_id)
        self.assertTrue(result)

        invoices = self.db.get_all_invoices()
        self.assertEqual(len([i for i in invoices if i['id'] == invoice_id]), 0)

    def test_insert_abnormal_invoice(self):
        """Test inserting an abnormal invoice record."""
        invoice_data = {
            "tajan_bidding_tracking_number": "",
            "test_service_type": "Invalid Type"
        }

        invoice_id = self.db.insert_abnormal_invoice(invoice_data, "Missing Tracking Number")
        self.assertIsNotNone(invoice_id)

        abnormal = self.db.get_all_abnormal_invoices()
        self.assertEqual(len(abnormal), 1)
        self.assertEqual(abnormal[0]['validation_error'], "Missing Tracking Number")

    def test_search_invoices(self):
        """Test searching invoices with filters."""
        # Insert test data
        self.db.insert_invoice({
            "tajan_bidding_tracking_number": "TBT-SEARCH1",
            "lab_name": "BV",
            "invoice_date": "2024-01-15"
        })
        self.db.insert_invoice({
            "tajan_bidding_tracking_number": "TBT-SEARCH2",
            "lab_name": "ITS",
            "invoice_date": "2024-02-15"
        })

        # Search by lab name
        results = self.db.search_invoices({"lab_name": "BV"})
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]['lab_name'], "BV")

        # Search by date range
        results = self.db.search_invoices({
            "date_from": "2024-01-01",
            "date_to": "2024-01-31"
        })
        self.assertEqual(len(results), 1)

        # Search by text
        results = self.db.search_invoices({"search_text": "SEARCH1"})
        self.assertEqual(len(results), 1)


if __name__ == "__main__":
    unittest.main()
