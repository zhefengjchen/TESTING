"""
Database module for Invoice Management System.
Handles SQLite database operations for invoices, abnormal items, and settings.
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple


class Database:
    """SQLite database handler for invoice management."""

    # Invoice table columns matching Excel headers
    INVOICE_COLUMNS = [
        "id",
        "tajan_bidding_tracking_number",
        "amazon_test_request",
        "amazon_tracker_number",
        "asin",
        "development_center",
        "category",
        "product_brand",
        "product_description",
        "testing_sla",
        "test_service_type",
        "quotation_order_number",
        "request_date",
        "test_start_date",
        "report_delivered_date",
        "report_number",
        "test_inspection_location",
        "product_line",
        "amazon_quality_manager",
        "amazon_sourcing_manager",
        "invoice_number",
        "lab_contact",
        "comment",
        "amount_usd",
        "lab_name",
        "invoice_date",
        "created_at",
        "updated_at"
    ]

    # Default allowed test/service types
    DEFAULT_TEST_TYPES = [
        "Benchmark Test",
        "Benchmark Sample Purchase",
        "Usability Test",
        "Usability Protocol",
        "Comparison",
        "DDC Design Document Collection",
        "Delorean DDC",
        "DP Review",
        "PRD & Protocol Upgrade",
        "Performance Protocol",
        "PVT",
        "PT",
        "PPT",
        "PT-Initial",
        "PT-Final",
        "ET",
        "VT",
        "Delorean PT",
        "Packaging",
        "Failure & Defect Analysis",
        "Declaration of Conformity",
        "ECR"
    ]

    # Default lab names
    DEFAULT_LAB_NAMES = ["BV", "ITS", "TUV", "SGS"]

    # Default database directory for Windows
    DEFAULT_DB_DIR = r"C:\Amazon\AI\CODING PROJECT\INVOICE MANAGEMENT SYSTEM\invoice database"
    DEFAULT_DB_PATH = os.path.join(DEFAULT_DB_DIR, "invoices.db")
    IPC_DB_PATH = os.path.join(DEFAULT_DB_DIR, "ipc.db")
    FQA_DB_PATH = os.path.join(DEFAULT_DB_DIR, "fqa.db")
    PSI_DB_PATH = os.path.join(DEFAULT_DB_DIR, "psi.db")

    def __init__(self, db_path: str = None):
        """Initialize database connection."""
        if db_path is None:
            db_path = self.DEFAULT_DB_PATH
        self.db_path = db_path
        self._ensure_db_directory()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()
        self._initialize_settings()

    def _ensure_db_directory(self):
        """Ensure the database directory exists."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)

    def _create_tables(self):
        """Create all required database tables."""
        cursor = self.conn.cursor()

        # Valid invoices table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tajan_bidding_tracking_number TEXT,
                amazon_test_request TEXT,
                amazon_tracker_number TEXT,
                asin TEXT,
                development_center TEXT,
                category TEXT,
                product_brand TEXT,
                product_description TEXT,
                testing_sla TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Abnormal invoice items table (same structure + validation_error field)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS abnormal_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tajan_bidding_tracking_number TEXT,
                amazon_test_request TEXT,
                amazon_tracker_number TEXT,
                asin TEXT,
                development_center TEXT,
                category TEXT,
                product_brand TEXT,
                product_description TEXT,
                testing_sla TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                validation_error TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Lab names settings table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS lab_names (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Test/Service types settings table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS test_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        self.conn.commit()

    def _initialize_settings(self):
        """Initialize default settings if not present."""
        cursor = self.conn.cursor()

        # Check if lab names exist
        cursor.execute("SELECT COUNT(*) FROM lab_names")
        if cursor.fetchone()[0] == 0:
            for lab in self.DEFAULT_LAB_NAMES:
                cursor.execute("INSERT OR IGNORE INTO lab_names (name) VALUES (?)", (lab,))

        # Check if test types exist
        cursor.execute("SELECT COUNT(*) FROM test_types")
        if cursor.fetchone()[0] == 0:
            for test_type in self.DEFAULT_TEST_TYPES:
                cursor.execute("INSERT OR IGNORE INTO test_types (name) VALUES (?)", (test_type,))

        self.conn.commit()

    # Invoice CRUD operations
    def insert_invoice(self, data: Dict[str, Any]) -> int:
        """Insert a new invoice record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.INVOICE_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns]

        cursor.execute(f"""
            INSERT INTO invoices ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def insert_abnormal_invoice(self, data: Dict[str, Any], validation_error: str) -> int:
        """Insert an abnormal invoice record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.INVOICE_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        columns.append('validation_error')

        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns[:-1]]
        values.append(validation_error)

        cursor.execute(f"""
            INSERT INTO abnormal_invoices ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def get_all_invoices(self) -> List[Dict[str, Any]]:
        """Get all valid invoice records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM invoices ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def get_all_abnormal_invoices(self) -> List[Dict[str, Any]]:
        """Get all abnormal invoice records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM abnormal_invoices ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def update_invoice(self, invoice_id: int, data: Dict[str, Any]) -> bool:
        """Update an invoice record."""
        cursor = self.conn.cursor()

        # Build update query dynamically
        set_clauses = []
        values = []
        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(invoice_id)

        query = f"UPDATE invoices SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()

        return cursor.rowcount > 0

    def update_abnormal_invoice(self, invoice_id: int, data: Dict[str, Any]) -> bool:
        """Update an abnormal invoice record."""
        cursor = self.conn.cursor()

        set_clauses = []
        values = []
        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(invoice_id)

        query = f"UPDATE abnormal_invoices SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()

        return cursor.rowcount > 0

    def delete_invoice(self, invoice_id: int) -> bool:
        """Delete an invoice record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM invoices WHERE id = ?", (invoice_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_abnormal_invoice(self, invoice_id: int) -> bool:
        """Delete an abnormal invoice record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM abnormal_invoices WHERE id = ?", (invoice_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def search_invoices(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search invoices with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM invoices WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('test_service_type'):
            query += " AND test_service_type = ?"
            params.append(filters['test_service_type'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                tajan_bidding_tracking_number LIKE ? OR
                amazon_test_request LIKE ? OR
                asin LIKE ? OR
                product_brand LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ?
            )"""
            params.extend([search_term] * 6)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def search_abnormal_invoices(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search abnormal invoices with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM abnormal_invoices WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('test_service_type'):
            query += " AND test_service_type = ?"
            params.append(filters['test_service_type'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                tajan_bidding_tracking_number LIKE ? OR
                amazon_test_request LIKE ? OR
                asin LIKE ? OR
                product_brand LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ?
            )"""
            params.extend([search_term] * 6)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    # Lab names management
    def get_lab_names(self) -> List[str]:
        """Get all lab names."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM lab_names ORDER BY name")
        return [row[0] for row in cursor.fetchall()]

    def add_lab_name(self, name: str) -> bool:
        """Add a new lab name."""
        cursor = self.conn.cursor()
        try:
            cursor.execute("INSERT INTO lab_names (name) VALUES (?)", (name.upper(),))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    def delete_lab_name(self, name: str) -> bool:
        """Delete a lab name."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM lab_names WHERE name = ?", (name,))
        self.conn.commit()
        return cursor.rowcount > 0

    # Test types management
    def get_test_types(self) -> List[str]:
        """Get all test/service types."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM test_types ORDER BY name")
        return [row[0] for row in cursor.fetchall()]

    def add_test_type(self, name: str) -> bool:
        """Add a new test/service type."""
        cursor = self.conn.cursor()
        try:
            cursor.execute("INSERT INTO test_types (name) VALUES (?)", (name,))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

    def delete_test_type(self, name: str) -> bool:
        """Delete a test/service type."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM test_types WHERE name = ?", (name,))
        self.conn.commit()
        return cursor.rowcount > 0

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()


class PsiDatabase:
    """SQLite database handler for PSI inspection data."""

    PSI_COLUMNS = [
        "id",
        "inspection_id",
        "factory_id",
        "amazon_tracker_number",
        "factory_name",
        "product_description",
        "manday",
        "test_service_type",
        "quotation_order_number",
        "request_date",
        "test_start_date",
        "report_delivered_date",
        "report_number",
        "test_inspection_location",
        "product_line",
        "amazon_quality_manager",
        "amazon_sourcing_manager",
        "invoice_number",
        "lab_contact",
        "comment",
        "amount_usd",
        "lab_name",
        "invoice_date",
        "created_at",
        "updated_at"
    ]

    def __init__(self, db_path: str = None):
        """Initialize database connection."""
        if db_path is None:
            db_path = Database.IPC_DB_PATH
        self.db_path = db_path
        self._ensure_db_directory()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _ensure_db_directory(self):
        """Ensure the database directory exists."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)

    def _create_tables(self):
        """Create PSI inspection tables."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS psi_inspections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inspection_id TEXT,
                factory_id TEXT,
                amazon_tracker_number TEXT,
                factory_name TEXT,
                product_description TEXT,
                manday TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()

    def insert_inspection(self, data: Dict[str, Any]) -> int:
        """Insert a new inspection record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.PSI_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns]

        cursor.execute(f"""
            INSERT INTO psi_inspections ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def get_all_inspections(self) -> List[Dict[str, Any]]:
        """Get all inspection records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM psi_inspections ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def update_inspection(self, inspection_id: int, data: Dict[str, Any]) -> bool:
        """Update an inspection record."""
        cursor = self.conn.cursor()
        set_clauses = []
        values = []

        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(inspection_id)

        query = f"UPDATE psi_inspections SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_inspection(self, inspection_id: int) -> bool:
        """Delete an inspection record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM psi_inspections WHERE id = ?", (inspection_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def search_inspections(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search inspection records with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM psi_inspections WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('test_service_type'):
            query += " AND test_service_type = ?"
            params.append(filters['test_service_type'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                inspection_id LIKE ? OR
                factory_id LIKE ? OR
                amazon_tracker_number LIKE ? OR
                factory_name LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ? OR
                report_number LIKE ?
            )"""
            params.extend([search_term] * 7)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()


class FqaDatabase:
    """SQLite database handler for FQA data."""

    FQA_COLUMNS = Database.INVOICE_COLUMNS

    def __init__(self, db_path: str = None):
        """Initialize database connection."""
        if db_path is None:
            db_path = Database.FQA_DB_PATH
        self.db_path = db_path
        self._ensure_db_directory()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _ensure_db_directory(self):
        """Ensure the database directory exists."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)

    def _create_tables(self):
        """Create FQA tables."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS fqa_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tajan_bidding_tracking_number TEXT,
                amazon_test_request TEXT,
                amazon_tracker_number TEXT,
                asin TEXT,
                development_center TEXT,
                category TEXT,
                product_brand TEXT,
                product_description TEXT,
                testing_sla TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()

    def insert_invoice(self, data: Dict[str, Any]) -> int:
        """Insert a new FQA record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.FQA_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns]

        cursor.execute(f"""
            INSERT INTO fqa_invoices ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def get_all_invoices(self) -> List[Dict[str, Any]]:
        """Get all FQA records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM fqa_invoices ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def update_invoice(self, invoice_id: int, data: Dict[str, Any]) -> bool:
        """Update an FQA record."""
        cursor = self.conn.cursor()
        set_clauses = []
        values = []

        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(invoice_id)

        query = f"UPDATE fqa_invoices SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_invoice(self, invoice_id: int) -> bool:
        """Delete an FQA record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM fqa_invoices WHERE id = ?", (invoice_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def search_invoices(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search FQA records with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM fqa_invoices WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('test_service_type'):
            query += " AND test_service_type = ?"
            params.append(filters['test_service_type'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                tajan_bidding_tracking_number LIKE ? OR
                amazon_test_request LIKE ? OR
                amazon_tracker_number LIKE ? OR
                product_brand LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ? OR
                report_number LIKE ?
            )"""
            params.extend([search_term] * 7)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()


class PsiInspectionDatabase:
    """SQLite database handler for PSI inspection data."""

    PSI_COLUMNS = Database.INVOICE_COLUMNS

    def __init__(self, db_path: str = None):
        """Initialize database connection."""
        if db_path is None:
            db_path = Database.PSI_DB_PATH
        self.db_path = db_path
        self._ensure_db_directory()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _ensure_db_directory(self):
        """Ensure the database directory exists."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)

    def _create_tables(self):
        """Create PSI inspection tables."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS psi_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tajan_bidding_tracking_number TEXT,
                amazon_test_request TEXT,
                amazon_tracker_number TEXT,
                asin TEXT,
                development_center TEXT,
                category TEXT,
                product_brand TEXT,
                product_description TEXT,
                testing_sla TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()

    def insert_invoice(self, data: Dict[str, Any]) -> int:
        """Insert a new PSI record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.PSI_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns]

        cursor.execute(f"""
            INSERT INTO psi_invoices ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def get_all_invoices(self) -> List[Dict[str, Any]]:
        """Get all PSI records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM psi_invoices ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def update_invoice(self, invoice_id: int, data: Dict[str, Any]) -> bool:
        """Update a PSI record."""
        cursor = self.conn.cursor()
        set_clauses = []
        values = []

        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(invoice_id)

        query = f"UPDATE psi_invoices SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_invoice(self, invoice_id: int) -> bool:
        """Delete a PSI record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM psi_invoices WHERE id = ?", (invoice_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def search_invoices(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search PSI records with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM psi_invoices WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('test_service_type'):
            query += " AND test_service_type = ?"
            params.append(filters['test_service_type'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                tajan_bidding_tracking_number LIKE ? OR
                amazon_test_request LIKE ? OR
                amazon_tracker_number LIKE ? OR
                product_brand LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ? OR
                report_number LIKE ?
            )"""
            params.extend([search_term] * 7)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()


class PsiInspectionDatabase:
    """SQLite database handler for PSI inspection data."""

    PSI_COLUMNS = Database.INVOICE_COLUMNS

    def __init__(self, db_path: str = None):
        """Initialize database connection."""
        if db_path is None:
            db_path = Database.PSI_DB_PATH
        self.db_path = db_path
        self._ensure_db_directory()
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _ensure_db_directory(self):
        """Ensure the database directory exists."""
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir)

    def _create_tables(self):
        """Create PSI inspection tables."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS psi_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tajan_bidding_tracking_number TEXT,
                amazon_test_request TEXT,
                amazon_tracker_number TEXT,
                asin TEXT,
                development_center TEXT,
                category TEXT,
                product_brand TEXT,
                product_description TEXT,
                testing_sla TEXT,
                test_service_type TEXT,
                quotation_order_number TEXT,
                request_date TEXT,
                test_start_date TEXT,
                report_delivered_date TEXT,
                report_number TEXT,
                test_inspection_location TEXT,
                product_line TEXT,
                amazon_quality_manager TEXT,
                amazon_sourcing_manager TEXT,
                invoice_number TEXT,
                lab_contact TEXT,
                comment TEXT,
                amount_usd REAL,
                lab_name TEXT,
                invoice_date TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        self.conn.commit()

    def insert_invoice(self, data: Dict[str, Any]) -> int:
        """Insert a new PSI record."""
        cursor = self.conn.cursor()
        columns = [col for col in self.PSI_COLUMNS if col not in ('id', 'created_at', 'updated_at')]
        placeholders = ', '.join(['?' for _ in columns])
        column_names = ', '.join(columns)

        values = [data.get(col, '') for col in columns]

        cursor.execute(f"""
            INSERT INTO psi_invoices ({column_names})
            VALUES ({placeholders})
        """, values)

        self.conn.commit()
        return cursor.lastrowid

    def get_all_invoices(self) -> List[Dict[str, Any]]:
        """Get all PSI records."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM psi_invoices ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def update_invoice(self, invoice_id: int, data: Dict[str, Any]) -> bool:
        """Update a PSI record."""
        cursor = self.conn.cursor()
        set_clauses = []
        values = []

        for key, value in data.items():
            if key not in ('id', 'created_at'):
                set_clauses.append(f"{key} = ?")
                values.append(value)

        set_clauses.append("updated_at = ?")
        values.append(datetime.now().isoformat())
        values.append(invoice_id)

        query = f"UPDATE psi_invoices SET {', '.join(set_clauses)} WHERE id = ?"
        cursor.execute(query, values)
        self.conn.commit()
        return cursor.rowcount > 0

    def delete_invoice(self, invoice_id: int) -> bool:
        """Delete a PSI record."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM psi_invoices WHERE id = ?", (invoice_id,))
        self.conn.commit()
        return cursor.rowcount > 0

    def search_invoices(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Search PSI records with filters."""
        cursor = self.conn.cursor()

        query = "SELECT * FROM psi_invoices WHERE 1=1"
        params = []

        if filters.get('lab_name'):
            query += " AND lab_name = ?"
            params.append(filters['lab_name'])

        if filters.get('date_from'):
            query += " AND invoice_date >= ?"
            params.append(filters['date_from'])

        if filters.get('date_to'):
            query += " AND invoice_date <= ?"
            params.append(filters['date_to'])

        if filters.get('search_text'):
            search_term = f"%{filters['search_text']}%"
            query += """ AND (
                tajan_bidding_tracking_number LIKE ? OR
                amazon_test_request LIKE ? OR
                amazon_tracker_number LIKE ? OR
                product_brand LIKE ? OR
                product_description LIKE ? OR
                invoice_number LIKE ? OR
                report_number LIKE ?
            )"""
            params.extend([search_term] * 7)

        query += " ORDER BY id DESC"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows]

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()
