import re
import sys
import time
import queue
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.sync_api import Browser, BrowserContext, Page, Playwright, TimeoutError, sync_playwright
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QComboBox,
)


@dataclass
class MarkRecord:
    seat_number: str
    marks: str
    status: str = "PENDING"


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("E-Samarth Marks Entry Assistant (PyQt6)")
        self.resize(1200, 760)

        self.playwright: Optional[Playwright] = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

        self._selection_queue: queue.Queue[dict] = queue.Queue()
        self.table_selector: str = ""
        self.seat_column_index: Optional[int] = None

        self.df: Optional[pd.DataFrame] = None
        self.records: list[MarkRecord] = []

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)

        source_box = QGroupBox("1) Source File")
        source_layout = QGridLayout(source_box)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Select Excel/CSV file")

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self._on_browse_file)

        load_btn = QPushButton("Load Data")
        load_btn.clicked.connect(self._on_load_data)

        source_layout.addWidget(QLabel("File:"), 0, 0)
        source_layout.addWidget(self.file_path_edit, 0, 1)
        source_layout.addWidget(browse_btn, 0, 2)
        source_layout.addWidget(load_btn, 0, 3)

        self.seat_col_combo = QComboBox()
        self.marks_col_combo = QComboBox()

        source_layout.addWidget(QLabel("Seat Number Column:"), 1, 0)
        source_layout.addWidget(self.seat_col_combo, 1, 1)
        source_layout.addWidget(QLabel("Marks Column:"), 1, 2)
        source_layout.addWidget(self.marks_col_combo, 1, 3)

        layout.addWidget(source_box)

        browser_box = QGroupBox("2) Browser / Portal")
        browser_layout = QGridLayout(browser_box)

        launch_btn = QPushButton("Open Chromium (Debug)")
        launch_btn.clicked.connect(self._on_open_browser)

        select_table_btn = QPushButton("Select Table")
        select_table_btn.clicked.connect(self._on_select_table)

        select_seat_btn = QPushButton("Select Seat Column")
        select_seat_btn.clicked.connect(self._on_select_seat_column)

        self.marks_col_index_spin = QSpinBox()
        self.marks_col_index_spin.setMinimum(1)
        self.marks_col_index_spin.setMaximum(100)
        self.marks_col_index_spin.setValue(5)

        self.delay_spin = QSpinBox()
        self.delay_spin.setMinimum(100)
        self.delay_spin.setMaximum(5000)
        self.delay_spin.setValue(450)
        self.delay_spin.setSuffix(" ms")

        browser_layout.addWidget(launch_btn, 0, 0)
        browser_layout.addWidget(select_table_btn, 0, 1)
        browser_layout.addWidget(select_seat_btn, 0, 2)
        browser_layout.addWidget(QLabel("Marks Column Index in Portal:"), 1, 0)
        browser_layout.addWidget(self.marks_col_index_spin, 1, 1)
        browser_layout.addWidget(QLabel("Delay between rows:"), 1, 2)
        browser_layout.addWidget(self.delay_spin, 1, 3)

        layout.addWidget(browser_box)

        run_box = QGroupBox("3) Execute")
        run_layout = QHBoxLayout(run_box)

        start_btn = QPushButton("Start Entry")
        start_btn.clicked.connect(self._on_start_entry)

        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)

        self.progress_label = QLabel("Progress: 0/0")

        run_layout.addWidget(start_btn)
        run_layout.addWidget(self.progress, 1)
        run_layout.addWidget(self.progress_label)

        layout.addWidget(run_box)

        self.table_widget = QTableWidget(0, 3)
        self.table_widget.setHorizontalHeaderLabels(["Seat Number", "Marks", "Status"])
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table_widget, 1)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Activity log...")
        layout.addWidget(self.log_box, 1)

    def _log(self, text: str) -> None:
        self.log_box.append(text)
        self.log_box.ensureCursorVisible()

    def _on_browse_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select data file",
            str(Path.cwd()),
            "Data Files (*.xlsx *.xls *.csv)",
        )
        if file_path:
            self.file_path_edit.setText(file_path)

    def _on_load_data(self) -> None:
        file_path = self.file_path_edit.text().strip()
        if not file_path:
            self._error("Please select a file first.")
            return

        path = Path(file_path)
        if not path.exists():
            self._error("Selected file does not exist.")
            return

        try:
            if path.suffix.lower() == ".csv":
                self.df = pd.read_csv(path)
            else:
                self.df = pd.read_excel(path)
        except Exception as exc:
            self._error(f"Failed to read file: {exc}")
            return

        if self.df is None or self.df.empty:
            self._error("Loaded file is empty.")
            return

        cols = [str(c) for c in self.df.columns]
        self.seat_col_combo.clear()
        self.marks_col_combo.clear()
        self.seat_col_combo.addItems(cols)
        self.marks_col_combo.addItems(cols)

        auto_seat = self._find_col_index(cols, ["seat", "roll", "exam roll"])
        auto_marks = self._find_col_index(cols, ["mark", "score", "obtained"])
        if auto_seat is not None:
            self.seat_col_combo.setCurrentIndex(auto_seat)
        if auto_marks is not None:
            self.marks_col_combo.setCurrentIndex(auto_marks)

        self._build_records_from_selection()
        self._log(f"Loaded {len(self.records)} rows from {path.name}.")

    def _find_col_index(self, cols: list[str], hints: list[str]) -> Optional[int]:
        lowered = [c.lower() for c in cols]
        for hint in hints:
            for idx, col in enumerate(lowered):
                if hint in col:
                    return idx
        return None

    def _build_records_from_selection(self) -> None:
        if self.df is None:
            return

        seat_col = self.seat_col_combo.currentText()
        marks_col = self.marks_col_combo.currentText()
        if not seat_col or not marks_col:
            return

        subset = self.df[[seat_col, marks_col]].copy()
        subset = subset.dropna(subset=[seat_col])

        self.records = []
        for _, row in subset.iterrows():
            seat = self._clean_str(row[seat_col])
            marks = self._clean_str(row[marks_col])
            if seat:
                self.records.append(MarkRecord(seat_number=seat, marks=marks, status="PENDING"))

        self._refresh_table()

    def _refresh_table(self) -> None:
        self.table_widget.setRowCount(len(self.records))
        for i, rec in enumerate(self.records):
            self.table_widget.setItem(i, 0, QTableWidgetItem(rec.seat_number))
            self.table_widget.setItem(i, 1, QTableWidgetItem(rec.marks))
            self.table_widget.setItem(i, 2, QTableWidgetItem(rec.status))

        self.progress.setMaximum(max(1, len(self.records)))
        self.progress.setValue(0)
        self.progress_label.setText(f"Progress: 0/{len(self.records)}")

    def _on_open_browser(self) -> None:
        try:
            if self.playwright is None:
                self.playwright = sync_playwright().start()

            if self.browser is None:
                login_url = "https://mu.samarth.ac.in/index.php/site/login"
                launch_args = ["--no-first-run", "--no-default-browser-check"]

                # Prefer system-installed Google Chrome to avoid dependency on Playwright browser download.
                try:
                    self.browser = self.playwright.chromium.launch(
                        channel="chrome",
                        headless=False,
                        args=launch_args,
                    )
                    self._log("Opened installed Google Chrome (channel=chrome).")
                except Exception:
                    self.browser = self.playwright.chromium.launch(headless=False, args=launch_args)
                    self._log("Installed Chrome not found. Opened Playwright Chromium.")

                self.context = self.browser.new_context(no_viewport=True, ignore_https_errors=True)
                self.page = self.context.new_page()

                nav_ok = False
                last_error: Optional[Exception] = None
                for _ in range(3):
                    try:
                        self.page.goto(login_url, wait_until="domcontentloaded", timeout=20000)
                        nav_ok = True
                        break
                    except Exception as exc:
                        last_error = exc
                        time.sleep(0.4)

                if not nav_ok:
                    try:
                        # Fallback when direct goto is flaky due transient browser/network state.
                        self.page.evaluate("url => window.location.href = url", login_url)
                        self.page.wait_for_load_state("domcontentloaded", timeout=15000)
                        nav_ok = True
                    except Exception as exc:
                        last_error = exc

                self.page.expose_binding("esamarthSelectionDone", self._on_selection_done)

                if nav_ok:
                    self._log("Browser opened and login page loaded.")
                else:
                    self._log(f"Browser opened but auto-navigation failed: {last_error}")
                    self._log("Manually open: https://mu.samarth.ac.in/index.php/site/login")
            else:
                self._log("Browser is already open.")
        except Exception as exc:
            self._error(f"Could not open browser: {exc}")

    def _on_selection_done(self, _source, payload: dict) -> None:
        self._selection_queue.put(payload)

    def _on_select_table(self) -> None:
        result = self._pick_from_page("table")
        if not result:
            return

        self.table_selector = result.get("tableSelector", "")
        if self.table_selector:
            self._log(f"Selected table: {self.table_selector}")
        else:
            self._error("Could not capture table selector.")

    def _on_select_seat_column(self) -> None:
        result = self._pick_from_page("column")
        if not result:
            return

        column_index = result.get("columnIndex")
        table_selector = result.get("tableSelector", "")

        if table_selector:
            self.table_selector = table_selector
        if isinstance(column_index, int):
            self.seat_column_index = column_index
            self._log(f"Selected seat-number column index: {column_index}")
        else:
            self._error("Could not capture seat column index.")

    def _pick_from_page(self, mode: str) -> Optional[dict]:
        if self.page is None:
            self._error("Open browser first.")
            return None

        self._clear_queue(self._selection_queue)

        script = """
        (mode) => {
            const old = window.__esamarthPicker;
            if (old && old.cleanup) {
                old.cleanup();
            }

            function cssPath(el) {
                if (!el) return "";
                if (el.id) return `#${el.id}`;
                const parts = [];
                let node = el;
                while (node && node.nodeType === 1 && node !== document.body) {
                    let selector = node.tagName.toLowerCase();
                    if (node.id) {
                        selector += `#${node.id}`;
                        parts.unshift(selector);
                        break;
                    }
                    let sibling = node;
                    let idx = 1;
                    while ((sibling = sibling.previousElementSibling) != null) {
                        if (sibling.tagName === node.tagName) idx += 1;
                    }
                    selector += `:nth-of-type(${idx})`;
                    parts.unshift(selector);
                    node = node.parentElement;
                }
                return parts.join(" > ");
            }

            function getColumnIndex(cell) {
                if (!cell || !cell.parentElement) return null;
                const cells = Array.from(cell.parentElement.children).filter((x) => x.matches("th,td"));
                return cells.indexOf(cell) + 1;
            }

            function highlight(el) {
                if (!el) return;
                el.__prevOutline = el.style.outline;
                el.style.outline = "2px solid #e11d48";
            }

            function unhighlight(el) {
                if (!el) return;
                el.style.outline = el.__prevOutline || "";
            }

            let hovered = null;
            const onMove = (ev) => {
                const target = mode === "table"
                    ? ev.target.closest("table")
                    : ev.target.closest("th,td");
                if (hovered !== target) {
                    unhighlight(hovered);
                    hovered = target;
                    highlight(hovered);
                }
            };

            const onClick = (ev) => {
                ev.preventDefault();
                ev.stopPropagation();

                const cell = ev.target.closest("th,td");
                const table = mode === "table"
                    ? ev.target.closest("table")
                    : cell ? cell.closest("table") : null;

                const payload = {
                    mode,
                    tableSelector: cssPath(table),
                    columnIndex: mode === "column" ? getColumnIndex(cell) : null
                };

                cleanup();
                window.esamarthSelectionDone(payload);
            };

            function cleanup() {
                document.removeEventListener("mousemove", onMove, true);
                document.removeEventListener("click", onClick, true);
                unhighlight(hovered);
                window.__esamarthPicker = null;
            }

            window.__esamarthPicker = { cleanup };
            document.addEventListener("mousemove", onMove, true);
            document.addEventListener("click", onClick, true);
        }
        """

        armed = False
        last_error: Optional[Exception] = None
        for _ in range(5):
            try:
                # Results pages can refresh after navigation/login; wait briefly for a stable context.
                self.page.wait_for_load_state("domcontentloaded", timeout=4000)
            except Exception:
                pass

            try:
                self.page.evaluate(script, mode)
                armed = True
                break
            except Exception as exc:
                last_error = exc
                msg = str(exc).lower()
                if (
                    "execution context was destroyed" in msg
                    or "cannot find context" in msg
                    or "navigation" in msg
                ):
                    QApplication.processEvents()
                    time.sleep(0.25)
                    continue
                self._error(f"Could not arm picker: {exc}")
                return None

        if not armed:
            hint = "Page was still navigating. Wait until the page is fully loaded, then click Select Table again."
            self._error(f"Could not arm picker: {last_error}\n\n{hint}")
            return None

        self._log(f"Picker armed ({mode}). Click on the portal page.")

        deadline = time.time() + 90
        while time.time() < deadline:
            QApplication.processEvents()
            try:
                return self._selection_queue.get_nowait()
            except queue.Empty:
                time.sleep(0.05)

        self._error("Timed out waiting for selection. Try again.")
        return None

    def _on_start_entry(self) -> None:
        if self.df is None:
            self._error("Load source data first.")
            return

        self._build_records_from_selection()
        if not self.records:
            self._error("No rows to process.")
            return

        if self.page is None:
            self._error("Open browser first.")
            return

        if not self.table_selector:
            self._error("Select the portal table first.")
            return

        if self.seat_column_index is None:
            self._error("Select the seat-number column first.")
            return

        marks_col_idx = self.marks_col_index_spin.value()
        delay_ms = self.delay_spin.value()

        total = len(self.records)
        self.progress.setMaximum(total)

        done = 0
        for idx, record in enumerate(self.records):
            QApplication.processEvents()

            try:
                success = self._process_single_record(record, marks_col_idx)
                record.status = "DONE" if success else "NOT FOUND"
            except Exception as exc:
                record.status = f"ERROR: {exc}"

            self.table_widget.setItem(idx, 2, QTableWidgetItem(record.status))
            done += 1
            self.progress.setValue(done)
            self.progress_label.setText(f"Progress: {done}/{total}")
            time.sleep(delay_ms / 1000)

        self._log("Data entry run completed.")

    def _process_single_record(self, record: MarkRecord, marks_col_idx: int) -> bool:
        assert self.page is not None

        rows = self.page.locator(f"{self.table_selector} tr")
        row_count = rows.count()

        seat_key = self._norm(record.seat_number)
        for row_idx in range(row_count):
            row = rows.nth(row_idx)
            seat_cell = row.locator(f"td:nth-child({self.seat_column_index})").first
            if seat_cell.count() == 0:
                continue

            try:
                cell_text = self._norm(seat_cell.inner_text(timeout=250))
            except TimeoutError:
                continue

            if cell_text != seat_key:
                continue

            marks_input = row.locator(f"td:nth-child({marks_col_idx}) input[type='text']").first
            if marks_input.count() == 0:
                return False

            marks_input.click(timeout=1000)
            marks_input.fill(record.marks)

            present_checkbox = row.locator(
                f"td:nth-child({marks_col_idx}) input.component-status[value='PRESENT']"
            ).first
            if present_checkbox.count() > 0:
                present_checkbox.check(force=True)

            save_button = row.locator("button.savefunction").first
            if save_button.count() > 0:
                save_button.click(timeout=1000)
            else:
                return False

            self.page.wait_for_timeout(220)
            return True

        return False

    def _clear_queue(self, q: queue.Queue) -> None:
        while not q.empty():
            try:
                q.get_nowait()
            except queue.Empty:
                return

    def _clean_str(self, value) -> str:
        if pd.isna(value):
            return ""
        text = str(value).strip()
        return text

    def _norm(self, text: str) -> str:
        return re.sub(r"\s+", " ", str(text).strip())

    def _error(self, message: str) -> None:
        QMessageBox.critical(self, "Error", message)
        self._log(f"ERROR: {message}")

    def closeEvent(self, event) -> None:  # noqa: N802
        try:
            if self.context is not None:
                self.context.close()
            if self.browser is not None:
                self.browser.close()
            if self.playwright is not None:
                self.playwright.stop()
        except Exception:
            pass
        super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
