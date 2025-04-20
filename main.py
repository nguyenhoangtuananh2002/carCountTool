import sys
import os
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QScrollArea, QGroupBox, QGridLayout,
    QMessageBox, QInputDialog, QTableWidget, QTableWidgetItem, QDialog
)
from PySide6.QtCore import Qt


class VehicleCounterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Vehicle Counter Tool")
        self.setGeometry(100, 100, 1200, 700)

        self.directions = []  # Start with an empty list for directions
        self.vehicle_types = [
            "Xe hơi", "Xe bus nhỏ", "Xe bus lớn",
            "Xe chở nhỏ", "Xe chở trung", "Xe chở lớn"
        ]

        self.counts = {}  # {(direction, vehicle_type): count}
        self.labels = {}  # {(direction, vehicle_type): QLabel}
        self.direction_editors = {}  # {(direction): QLineEdit}
        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        main_layout = QVBoxLayout(main_widget)

        # Top controls layout
        top_layout = QHBoxLayout()
        main_layout.addLayout(top_layout)

        top_layout.addWidget(QLabel("Time In (HH:MM):"))
        self.time_in_edit = QLineEdit()
        self.time_in_edit.setFixedWidth(100)
        top_layout.addWidget(self.time_in_edit)

        top_layout.addWidget(QLabel("Time Out (HH:MM):"))
        self.time_out_edit = QLineEdit()
        self.time_out_edit.setFixedWidth(100)
        top_layout.addWidget(self.time_out_edit)

        export_button = QPushButton("Export to Excel")
        export_button.clicked.connect(self.export_excel)
        top_layout.addWidget(export_button)

        summary_button = QPushButton("View Summary")
        summary_button.clicked.connect(self.show_summary)
        top_layout.addWidget(summary_button)

        reset_button = QPushButton("Reset All")
        reset_button.clicked.connect(self.reset_all_counts)
        top_layout.addWidget(reset_button)

        clear_file_button = QPushButton("Clear File")
        clear_file_button.clicked.connect(self.clear_excel_file)
        top_layout.addWidget(clear_file_button)

        add_direction_button = QPushButton("Add Direction")
        add_direction_button.clicked.connect(self.add_direction)
        top_layout.addWidget(add_direction_button)

        # Scroll area for vehicle counters
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_layout.addWidget(scroll_area)

        scroll_widget = QWidget()
        scroll_area.setWidget(scroll_widget)

        self.grid_layout = QGridLayout(scroll_widget)
        self.update_direction_ui()

    def update_direction_ui(self):
        # Clear existing widgets in the grid layout
        for i in reversed(range(self.grid_layout.count())):
            widget = self.grid_layout.itemAt(i).widget()
            if widget is not None:
                widget.deleteLater()

        # Add new direction group boxes
        for i, direction in enumerate(self.directions):
            group_box = QGroupBox(f"Direction {direction}")
            group_layout = QGridLayout()
            group_box.setLayout(group_layout)

            self.grid_layout.addWidget(group_box, i // 2, i % 2)

            # Editable direction name (double-click to edit)
            direction_edit = QLineEdit(direction)
            self.direction_editors[direction] = direction_edit
            direction_edit.setFixedWidth(100)
            direction_edit.setStyleSheet("font-size: 16px; padding: 5px;")
            direction_edit.setAlignment(Qt.AlignCenter)

            # Double-click event to allow renaming the direction
            direction_edit.setContextMenuPolicy(Qt.CustomContextMenu)
            direction_edit.customContextMenuRequested.connect(lambda pos, de=direction_edit: self.rename_direction(de))

            group_layout.addWidget(direction_edit, 0, 0, 1, 4)

            for j, vehicle in enumerate(self.vehicle_types):
                key = (direction, vehicle)
                self.counts[key] = 0

                vehicle_label = QLabel(vehicle)
                count_label = QLabel("0")
                count_label.setAlignment(Qt.AlignCenter)
                count_label.setStyleSheet("font-size: 18px; color: blue;")
                count_label.mousePressEvent = lambda event, k=key, l=count_label: self.edit_count(k, l)

                plus_button = QPushButton("+")
                minus_button = QPushButton("-")

                plus_button.clicked.connect(lambda _, k=key, l=count_label: self.update_count(k, l, 1))
                minus_button.clicked.connect(lambda _, k=key, l=count_label: self.update_count(k, l, -1))

                group_layout.addWidget(vehicle_label, j + 1, 0)
                group_layout.addWidget(count_label, j + 1, 1)
                group_layout.addWidget(plus_button, j + 1, 2)
                group_layout.addWidget(minus_button, j + 1, 3)

                self.labels[key] = count_label

    def update_count(self, key, label, change):
        self.counts[key] = max(0, self.counts[key] + change)
        label.setText(str(self.counts[key]))

    def reset_all_counts(self):
        for key in self.counts:
            self.counts[key] = 0
            self.labels[key].setText("0")
        QMessageBox.information(self, "Reset Done", "All counts have been reset to 0.")

    def edit_count(self, key, label):
        value, ok = QInputDialog.getInt(self, "Edit Count", f"Enter count for {key[0]} - {key[1]}:", self.counts[key], 0)
        if ok:
            self.counts[key] = value
            label.setText(str(value))

    def export_excel(self):
        time_in = self.time_in_edit.text()
        time_out = self.time_out_edit.text()

        if not (time_in and time_out):
            QMessageBox.warning(self, "Missing Input", "Please enter Time In and Time Out.")
            return

        time_period = f"{time_in} - {time_out}"

        # Collect data for the Excel file
        data = []
        for direction in self.directions:
            row = [time_period, direction]
            for vehicle in self.vehicle_types:
                key = (direction, vehicle)
                row.append(self.counts.get(key, 0))
            data.append(row)

        # Convert data to a DataFrame
        df = pd.DataFrame(data, columns=["Time Period", "Direction"] + self.vehicle_types)

        # Define the output file
        output_file = "vehicle_counts.xlsx"

        # Write to Excel file and add AutoFilter
        if os.path.exists(output_file):
            with pd.ExcelWriter(output_file, mode="a", engine="openpyxl", if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                workbook = writer.book
                worksheet = writer.sheets["Sheet1"]
        else:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Sheet1", index=False)
                workbook = writer.book
                worksheet = writer.sheets["Sheet1"]

        # Apply AutoFilter to the entire dataset
        first_row = worksheet.min_row
        last_row = worksheet.max_row
        first_column = worksheet.min_column
        last_column = worksheet.max_column

        worksheet.auto_filter.ref = f"{chr(65 + first_column - 1)}{first_row}:{chr(65 + last_column - 1)}{last_row}"

        # Save the workbook
        workbook.save(output_file)

        QMessageBox.information(self, "Exported", "Data exported successfully!")

    def clear_excel_file(self):
        output_file = "vehicle_counts.xlsx"
        if os.path.exists(output_file):
            reply = QMessageBox.question(
                self, "Confirm Clear", "Are you sure you want to clear all data in the Excel file?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                columns = ["Time Period", "Direction"] + self.vehicle_types
                df = pd.DataFrame(columns=columns)
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Sheet1", index=False)
                QMessageBox.information(self, "Cleared", "Excel file has been cleared.")
        else:
            QMessageBox.warning(self, "File Not Found", "The Excel file does not exist yet.")

    def show_summary(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Current Vehicle Counts")
        layout = QVBoxLayout(dialog)

        table = QTableWidget()
        table.setColumnCount(len(self.vehicle_types) + 2)
        table.setHorizontalHeaderLabels(["Direction"] + self.vehicle_types + ["Total"])

        table.setRowCount(len(self.directions))
        for i, direction in enumerate(self.directions):
            table.setItem(i, 0, QTableWidgetItem(direction))
            total = 0
            for j, vehicle in enumerate(self.vehicle_types):
                count = self.counts.get((direction, vehicle), 0)
                total += count
                table.setItem(i, j + 1, QTableWidgetItem(str(count)))
            table.setItem(i, len(self.vehicle_types) + 1, QTableWidgetItem(str(total)))

        layout.addWidget(table)

        close_button = QPushButton("Close")
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)

        dialog.exec()

    def add_direction(self):
        direction, ok = QInputDialog.getText(self, "Add Direction", "Enter new direction name:")
        if ok and direction:
            if direction not in self.directions:
                self.directions.append(direction)
                self.update_direction_ui()
            else:
                QMessageBox.warning(self, "Direction Exists", "This direction already exists.")

    def rename_direction(self, direction_edit):
        new_name, ok = QInputDialog.getText(self, "Rename Direction", "Enter new direction name:", text=direction_edit.text())
        if ok and new_name:
            direction_edit.setText(new_name)
            direction = direction_edit.text()
            self.directions = [direction if d == direction_edit.text() else d for d in self.directions]
            self.update_direction_ui()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VehicleCounterApp()
    window.show()
    sys.exit(app.exec())
