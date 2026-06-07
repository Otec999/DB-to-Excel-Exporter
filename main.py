import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QTextEdit, QFileDialog,
    QMessageBox, QProgressBar, QGroupBox, QLineEdit, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
import sqlalchemy as sa
from sqlalchemy import inspect, text
from datetime import datetime


class DatabaseWorker(QThread):
    """Worker thread for database operations"""
    finished = pyqtSignal(object)
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, connection_string, action='test'):
        super().__init__()
        self.connection_string = connection_string
        self.action = action
        self.table_name = None

    def run(self):
        try:
            if self.action == 'list_tables':
                self._list_tables()
            elif self.action == 'export':
                self._export_table()
        except Exception as e:
            self.error.emit(str(e))

    def _get_engine(self):
        return sa.create_engine(self.connection_string)

    def _list_tables(self):
        engine = self._get_engine()
        inspector = inspect(engine)
        tables = inspector.get_table_names()
        self.finished.emit(tables)

    def _export_table(self):
        engine = self._get_engine()
        self.progress.emit(f"Exporting table: {self.table_name}...")
        query = f"SELECT * FROM [{self.table_name}]"
        df = pd.read_sql_query(query, engine)
        self.progress.emit(f"Read {len(df)} rows from {self.table_name}")
        self.finished.emit([df, self.table_name])


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.current_table = None
        self.worker = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("SQL to Excel Exporter")
        self.setGeometry(100, 100, 750, 600)
        self.setMinimumSize(600, 500)
        self.setStyleSheet("""
            QMainWindow { background-color: #f5f5f5; }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #bdbdbd;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QPushButton {
                background-color: #4CAF50; color: white; border: none;
                padding: 8px 16px; border-radius: 4px;
                font-size: 13px; min-height: 30px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:disabled { background-color: #bdbdbd; color: #757575; }
            QPushButton#btnConnect { background-color: #1976D2; }
            QPushButton#btnConnect:hover { background-color: #1565C0; }
            QPushButton#btnExport { background-color: #F57C00; }
            QPushButton#btnExport:hover { background-color: #E65100; }
            QPushButton#btnClear { background-color: #757575; }
            QPushButton#btnClear:hover { background-color: #616161; }
            QTextEdit {
                background-color: white; border: 1px solid #bdbdbd;
                border-radius: 4px; font-family: Consolas, monospace; font-size: 12px;
            }
            QComboBox, QLineEdit {
                padding: 5px; border: 1px solid #bdbdbd;
                border-radius: 4px; background-color: white; min-height: 25px;
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)

        # Title
        title = QLabel("SQL Database to Excel Exporter")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Connection group
        conn_grp = QGroupBox("Database Connection")
        conn_lay = QVBoxLayout()

        self.conn_input = QLineEdit()
        self.conn_input.setPlaceholderText(
            "Example: mssql+pyodbc://server/db?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server"
        )
        conn_lay.addWidget(QLabel("Connection String:"))
        conn_lay.addWidget(self.conn_input)

        # Quick templates
        quick = QHBoxLayout()
        btn_sql = QPushButton("SQL Server (local)")
        btn_sql.setStyleSheet("background-color: #607D8B;")
        btn_sql.clicked.connect(
            lambda: self.conn_input.setText(
                "mssql+pyodbc://./YourDatabase?trusted_connection=yes&driver=ODBC+Driver+17+for+SQL+Server"
            )
        )

        btn_mysql = QPushButton("MySQL")
        btn_mysql.setStyleSheet("background-color: #607D8B;")
        btn_mysql.clicked.connect(
            lambda: self.conn_input.setText(
                "mysql+pymysql://root:password@localhost/your_database"
            )
        )

        btn_pg = QPushButton("PostgreSQL")
        btn_pg.setStyleSheet("background-color: #607D8B;")
        btn_pg.clicked.connect(
            lambda: self.conn_input.setText(
                "postgresql://postgres:password@localhost:5432/your_database"
            )
        )

        quick.addWidget(btn_sql)
        quick.addWidget(btn_mysql)
        quick.addWidget(btn_pg)
        conn_lay.addLayout(quick)

        # Connect button
        self.btn_connect = QPushButton("Connect & Load Tables")
        self.btn_connect.setObjectName("btnConnect")
        self.btn_connect.clicked.connect(self.connect_database)
        conn_lay.addWidget(self.btn_connect)

        conn_grp.setLayout(conn_lay)
        layout.addWidget(conn_grp)

        # Table selection group
        tbl_grp = QGroupBox("Table Selection & Export")
        tbl_lay = QVBoxLayout()

        sel = QHBoxLayout()
        sel.addWidget(QLabel("Select Table:"))
        self.table_combo = QComboBox()
        self.table_combo.setMinimumHeight(30)
        sel.addWidget(self.table_combo, 1)
        tbl_lay.addLayout(sel)

        # Options
        opts = QHBoxLayout()
        self.chk_index = QCheckBox("Include row index")
        self.chk_index.setChecked(False)
        self.chk_open = QCheckBox("Open file after export")
        self.chk_open.setChecked(True)
        opts.addWidget(self.chk_index)
        opts.addWidget(self.chk_open)
        opts.addStretch()
        tbl_lay.addLayout(opts)

        # Buttons
        btns = QHBoxLayout()
        self.btn_export = QPushButton("Export Selected Table to Excel")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.clicked.connect(self.export_to_excel)
        self.btn_export.setEnabled(False)

        self.btn_export_all = QPushButton("Export ALL Tables")
        self.btn_export_all.setObjectName("btnExport")
        self.btn_export_all.clicked.connect(self.export_all_tables)
        self.btn_export_all.setEnabled(False)

        btns.addWidget(self.btn_export)
        btns.addWidget(self.btn_export_all)
        tbl_lay.addLayout(btns)

        tbl_grp.setLayout(tbl_lay)
        layout.addWidget(tbl_grp)

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Log
        log_grp = QGroupBox("Operation Log")
        log_lay = QVBoxLayout()
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(160)
        log_lay.addWidget(self.log_output)

        clear_btn = QPushButton("Clear Log")
        clear_btn.setObjectName("btnClear")
        clear_btn.clicked.connect(lambda: self.log_output.clear())
        log_lay.addWidget(clear_btn)

        log_grp.setLayout(log_lay)
        layout.addWidget(log_grp)

        self.statusBar().showMessage("Ready")
        self.log("Application started. Enter connection string and click Connect.")

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_output.append(f"[{ts}] {msg}")

    def set_controls_enabled(self, enabled):
        self.btn_connect.setEnabled(enabled)
        self.conn_input.setEnabled(enabled)
        self.btn_export.setEnabled(enabled and self.table_combo.count() > 0)
        self.btn_export_all.setEnabled(enabled and self.table_combo.count() > 0)

    def connect_database(self):
        conn = self.conn_input.text().strip()
        if not conn:
            QMessageBox.warning(self, "Warning", "Please enter a connection string!")
            return

        self.log("Connecting to database...")
        self.set_controls_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = DatabaseWorker(conn, 'list_tables')
        self.worker.finished.connect(self.on_tables_loaded)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_tables_loaded(self, tables):
        self.progress_bar.setVisible(False)
        self.set_controls_enabled(True)

        self.table_combo.clear()
        if tables:
            self.table_combo.addItems(tables)
            self.log(f"Loaded {len(tables)} tables from database")
            self.btn_export.setEnabled(True)
            self.btn_export_all.setEnabled(True)
        else:
            self.log("No tables found in database")

    def on_error(self, err):
        self.progress_bar.setVisible(False)
        self.set_controls_enabled(True)
        self.log(f"ERROR: {err}")
        QMessageBox.critical(self, "Error", str(err))

    def export_to_excel(self):
        table = self.table_combo.currentText()
        if not table:
            QMessageBox.warning(self, "Warning", "Please select a table!")
            return

        conn = self.conn_input.text().strip()
        if not conn:
            QMessageBox.warning(self, "Warning", "Please enter connection string!")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, f"Save '{table}' as Excel", f"{table}.xlsx", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        self.log(f"Exporting '{table}'...")
        self.set_controls_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = DatabaseWorker(conn, 'export')
        self.worker.table_name = table
        self.worker.finished.connect(lambda res: self._finish_export(res, path))
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def _finish_export(self, result, path):
        df, table = result
        try:
            df.to_excel(path, index=self.chk_index.isChecked(), engine='openpyxl')
            self.log(f"Exported '{table}' ({len(df)} rows) -> {path}")

            self.progress_bar.setVisible(False)
            self.set_controls_enabled(True)

            if self.chk_open.isChecked():
                reply = QMessageBox.question(
                    self, "Export Complete",
                    f"Table '{table}' exported!\nRows: {len(df)}\n\nOpen file?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    os.startfile(path)
            else:
                QMessageBox.information(
                    self, "Export Complete",
                    f"Table '{table}' exported!\nRows: {len(df)}\nFile: {path}"
                )
        except Exception as e:
            self.log(f"ERROR saving Excel: {e}")
            QMessageBox.critical(self, "Error", f"Failed to save Excel:\n{e}")
            self.progress_bar.setVisible(False)
            self.set_controls_enabled(True)

    def export_all_tables(self):
        conn = self.conn_input.text().strip()
        if not conn:
            QMessageBox.warning(self, "Warning", "Please enter connection string!")
            return

        folder = QFileDialog.getExistingDirectory(self, "Select folder to save all tables")
        if not folder:
            return

        self.log("Exporting ALL tables...")
        self.set_controls_enabled(False)
        self.progress_bar.setVisible(True)

        engine = sa.create_engine(conn)
        inspector = inspect(engine)
        tables = inspector.get_table_names()

        self.progress_bar.setRange(0, len(tables))
        ok_count = 0

        for i, t in enumerate(tables):
            try:
                self.log(f"  Exporting {t}...")
                df = pd.read_sql_query(f"SELECT * FROM [{t}]", engine)
                fp = os.path.join(folder, f"{t}.xlsx")
                df.to_excel(fp, index=self.chk_index.isChecked(), engine='openpyxl')
                self.log(f"    -> {len(df)} rows saved")
                ok_count += 1
            except Exception as e:
                self.log(f"    ERROR: {e}")
            self.progress_bar.setValue(i + 1)

        self.progress_bar.setVisible(False)
        self.set_controls_enabled(True)

        QMessageBox.information(
            self, "Export Complete",
            f"Exported {ok_count} of {len(tables)} tables\nLocation: {folder}"
        )


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
