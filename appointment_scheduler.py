from matplotlib.lines import Line2D
import sys
from PyQt5.QtWidgets import QApplication,QCheckBox,QWidgetAction,QDateEdit,QMessageBox,QSpinBox,QTableWidget,QLineEdit, QTableWidgetItem, QMainWindow, QHBoxLayout, QVBoxLayout, QCalendarWidget, QWidget, QPushButton, \
    QLabel, QInputDialog,  QVBoxLayout,QFileDialog, QFormLayout, QScrollArea,QLabel,QSizePolicy,QHeaderView,QGroupBox ,QComboBox, QTextEdit, QFrame,  QDialog, QComboBox, QTabWidget, QMenu, QAction, QVBoxLayout,QAbstractItemView
import datetime as dt
from PyQt5.QtCore import Qt, QDate, QTimer, pyqtSignal,QSize,QUrl,QThread,QDateTime,QFile,QEvent
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
from PyQt5.QtGui import QIcon,QFont,QDesktopServices
from datetime import datetime
import csv, json
import pandas as pd
import math
import openpyxl 
import xlsxwriter
from openpyxl import load_workbook
import requests
from PyQt5.QtGui import QColor,QTextCharFormat, QColor
class TimeSlotDialog(QDialog):
        def __init__(self, locations, engineers, parent=None):
            super().__init__(parent)

            self.setWindowTitle("Book Slot")
            self.setGeometry(200, 200, 400, 200)
            self.parent = parent
            self.location_dropdown = QComboBox(self) 
            self.location_dropdown.addItems(locations)
        
            self.engineer_dropdown = QComboBox(self)
            self.engineer_dropdown.addItems(engineers)

            self.time_options = list(range(1, 9))
            self.time_dropdown = QComboBox(self)
            self.time_dropdown.addItems([f"{hours} Hour(s)" for hours in self.time_options])
            self.time_dropdown.setCurrentIndex(7)
            icon_width = 20
            icon_height = 20

            self.new_button = QPushButton(self)
            self.new_button.setIcon(self.createIcon("file-plus-icon.png", QSize(icon_width, icon_height)))
            self.new_button.setFixedSize(icon_width + 10, icon_height + 10)  
            self.new_button.clicked.connect(self.add_engineer)

            self.new_button1 = QPushButton(self)
            self.new_button1.setIcon(self.createIcon("file-plus-icon.png", QSize(icon_width, icon_height)))
            self.new_button1.setFixedSize(icon_width + 10, icon_height + 10)  
            self.new_button1.clicked.connect(self.add_location)

            self.main_layout = QVBoxLayout(self)
        
            self.engineer_layout = QHBoxLayout()
            self.engineer_layout.addWidget(self.engineer_dropdown)
            self.engineer_layout.addWidget(self.new_button)
            
            self.location_layout = QHBoxLayout()
            self.location_layout.addWidget(self.location_dropdown)
            self.location_layout.addWidget(self.new_button1)

            
            self.main_layout.addLayout(self.engineer_layout)
            self.main_layout.addLayout(self.location_layout)
            self.main_layout.addWidget(self.time_dropdown)

            self.ok_button = QPushButton("OK", self)
            self.ok_button.clicked.connect(self.accept)
            self.main_layout.addWidget(self.ok_button)

            self.cancel_button = QPushButton("Cancel", self)
            self.cancel_button.clicked.connect(self.reject)
            self.main_layout.addWidget(self.cancel_button)

            self.engineers = {} 
            self.locations = {}
            
            self.data = self.parent.data

            self.engineer_dropdown.addItem("Select Engineer")
            self.location_dropdown.addItem("Select Location")
            self.engineer_dropdown.setCurrentIndex(self.engineer_dropdown.findText("Select Engineer"))
            self.location_dropdown.setCurrentIndex(self.location_dropdown.findText("Select Location"))


        def save_json(self, data):
            # save json
            # self.data[what_to_save] = data
            with open ("booked_slot.json","w+") as json_file:
                json.dump(self.data,json_file,indent=4)
        def load_json(self):
                with open("booked_slot.json") as json_file:
                    data = json.load(json_file)
                return data       
        def createIcon(self, path, size):
            icon = QIcon(path)
            pixmap = icon.pixmap(size)
            return QIcon(pixmap)
        def add_engineer(self):
            engineer, ok = QInputDialog.getText(self, "Add Engineer", "Enter the engineer's name:")
            if ok and engineer:
                self.engineer_dropdown.addItem(engineer)
            
                if 'engineers' not in self.data:
                    self.data['engineers'] = {}
                self.data['engineers'][engineer] = []
                api_url = "http://192.168.17.72:5000/api/add_engineer"
                response = requests.post(api_url, json={"engineer": engineer})
                self.save_json(self.data)
                
                if response.status_code == 200:
                    QMessageBox.information(self, "Success", "Engineer added successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Engineer added successfully!")

        def add_location(self):
            location, ok = QInputDialog.getText(self, "Add Location", "Enter the location name:")
            if ok and location:
                self.location_dropdown.addItem(location) 

                if 'locations' not in self.data:
                    self.data['locations'] = {}
                self.data['locations'][location] = []
                api_url = "http://192.168.17.72:5000/api/add_location"
                response = requests.post(api_url, json={"location": location})
                self.save_json(self.data)
                
                if response.status_code == 200:
                    QMessageBox.information(self, "Success", "Location added successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Location added successfully!")
        def exec_(self):
            self.result = super().exec_()
            return self.result
        def get_selected_data(self):
            return {
                'engineer': self.engineer_dropdown.currentText(),
                'location': self.location_dropdown.currentText(), 
                'time': self.time_options[self.time_dropdown.currentIndex()],
            }

class DashboardWidget(QWidget):
    def __init__(self, locations, engineers, parent=None):
        super().__init__(parent)

       

        self.bookings = []
        with open("booked_slot.json") as json_file:
            data = json.load(json_file)
            engineers_data = data.get("engineers", {})
            locations_data = data.get("locations", {})
            self.dates = sorted(set(entry["date"] for entry in data.get("bookings", [])))

        
        self.engineer_dropdown = QComboBox(self)
        self.engineer_dropdown.addItem("All Engineer")
        self.engineer_dropdown.addItems(engineers_data.keys())
        self.engineer_dropdown.currentIndexChanged.connect(self.update_charts)
        self.engineer_dropdown.setStyleSheet("background-color: lightblue;")

        self.location_dropdown = QComboBox(self)
        self.location_dropdown.addItem("All Location")
        self.location_dropdown.addItems(locations_data.keys())
        self.location_dropdown.currentIndexChanged.connect(self.update_charts)
        self.location_dropdown.setStyleSheet("background-color: lightblue;")


        
        self.update_charts_button = QPushButton("Update Charts", self)
        self.update_charts_button.clicked.connect(self.update_charts)
        self.update_charts_button.setStyleSheet("background-color: lightblue;")
        self.update_charts_button.setFixedHeight(20)


       
      
        self.engineer_schedule_chart = ScrollableMatplotlibWidget(self, width=800, height=800)
        self.engineer_schedule_chart.setGeometry(800, 800, 800, 600) 
        self.engineer_schedule_chart.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)

      
        container_widget = QWidget(self)
        container_layout = QVBoxLayout(container_widget)
        container_layout.addWidget(self.engineer_schedule_chart)
        scroll_area.setWidget(container_widget)

       
        self.date = QDateEdit(self)
        self.date.setFixedHeight(20)
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat("yyyy/MM/dd")
        self.date.setDate(QDate.currentDate())
        self.date.dateChanged.connect(self.update_charts)
        self.date.setStyleSheet("background-color: lightblue;")

     
        self.top_layout = QHBoxLayout()
        self.top_layout.addWidget(self.engineer_dropdown)
        self.top_layout.addWidget(self.location_dropdown)
        self.top_layout.addWidget(self.date)
        self.top_layout.addWidget(self.update_charts_button)


        
        self.layout = QVBoxLayout(self)
        self.layout.addWidget(scroll_area)
        self.layout.addLayout(self.top_layout)
        

    def update_charts(self):
        with open("booked_slot.json") as json_file:
            data = json.load(json_file)

        bookings = data.get("bookings", [])

        selected_engineer = self.engineer_dropdown.currentText()
        selected_location = self.location_dropdown.currentText()
        selected_date = self.date.date().toString("yyyy/MM/dd")
        
        # Filter bookings based on selected engineer, location, and date
        engineer_schedule_bookings = [
            entry
            for entry in bookings
            if (
                (entry["engineer"] == selected_engineer or selected_engineer == "All Engineer") and
                (entry["location"] == selected_location or selected_location == "All Location") and
                (entry["date"] == selected_date)
            )
        ]

        # Extract unique engineers and locations from filtered bookings
        unique_engineers = sorted(set(entry["engineer"] for entry in engineer_schedule_bookings))
        unique_locations = sorted(set(entry["location"] for entry in engineer_schedule_bookings))

        engineer_data = {engineer: {"locations": [], "durations": []} for engineer in unique_engineers}

        # Populate engineer data with filtered bookings
        for entry in engineer_schedule_bookings:
            engineer_data[entry['engineer']]["locations"].append(entry["location"])
            engineer_data[entry['engineer']]["durations"].append(entry["duration"])

        # Plot engineer schedule based on filtered data
        self.engineer_schedule_chart.plot_engineer_schedule(
            unique_engineers, engineer_data, "Engineer Schedule"
)

class ScrollableMatplotlibWidget(QScrollArea):
    def __init__(self, parent=None, width=15, height=10):
        super(ScrollableMatplotlibWidget, self).__init__(parent)

        self.fig, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.fig)
        self.setWidgetResizable(True)

        container_widget = QWidget(self)
        container_layout = QVBoxLayout(container_widget)

        self.figure = Figure(figsize=(width, height))
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)
        container_layout.addWidget(self.canvas)

        self.setWidget(container_widget)

    def clear_plot(self):
        if self.ax is not None:
            self.ax.clear()

    def plot_engineer_schedule(self, engineers, engineer_data, title):
        self.clear_plot()
        # Define a list of distinct colors
        distinct_colors = [
                'tab:blue', 'tab:orange', 'tab:green', 'tab:red', 'tab:purple', 
                'tab:brown', 'tab:pink', 'tab:gray', 'tab:olive', 'tab:cyan', 
                'deepskyblue', 'salmon', 'gold', 'darkviolet', 'limegreen', 
                'darkorange', 'dodgerblue', 'mediumorchid', 'darkslategrey', 
                'darkturquoise', 'rosybrown', 'mediumpurple', 'olivedrab', 
                'steelblue', 'peru', 'royalblue', 'indianred', 'darkkhaki', 
                'mediumseagreen', 'orangered', 'slateblue', 'seagreen', 
                'sienna', 'crimson', 'darkolivegreen', 'teal', 'slategray', 
                'firebrick', 'cadetblue', 'chocolate', 'forestgreen', 
                'mediumaquamarine', 'darkgoldenrod', 'dimgray', 'saddlebrown', 
                'darkcyan', 'darkmagenta', 'mediumslateblue', 'darkorchid']
       
        
        bar_width = 0.2  

        
        unique_locations = set()
        location_color_mapping = {}

        for i, engineer in enumerate(engineers):
            durations = engineer_data[engineer]["durations"]
            locations = engineer_data[engineer]["locations"]

            
            for unique_location in set(locations) - unique_locations:
                unique_locations.add(unique_location)

              
                color = distinct_colors[len(unique_locations) % len(distinct_colors)]

                
                location_color_mapping[unique_location] = color

            for j, (duration, location) in enumerate(zip(durations, locations)):
                
                x_position = i + j * (bar_width + 0.0013)
                self.ax.bar(x_position, duration, width=bar_width, color=location_color_mapping[location], label=f'{location}')

       
        self.ax.set_xticks(np.arange(len(engineers)) + (len(engineers) - 2) * 0.025 * (bar_width + 0.124))
        self.ax.set_xticklabels(engineers, rotation=0, ha="center")

       
        legend_handles = [Line2D([0], [0], marker='o', color='w', markerfacecolor=color, markersize=15, label=label) for label, color in location_color_mapping.items()]
        self.ax.legend(handles=legend_handles, loc='upper right', title='Location')

       
        self.ax.set_title(title)
        self.ax.set_xlabel('Engineer')
        self.ax.set_ylabel('Duration')

       
        self.canvas.draw()

class ReportsWidget(QWidget):
    def __init__(self, engineers, locations, parent=None):
        super().__init__(parent)
        self.engineers = engineers
        self.locations = locations
        self.data = []  
        self.layout = QVBoxLayout(self)
        # Filter button
        self.filter_groupbox = QGroupBox("Filter Options", self)    
        filter_layout = QHBoxLayout(self.filter_groupbox)
        
        self.engineer_dropdown = QComboBox(self)
        self.engineer_dropdown.addItems(["All"] + list(self.engineers))
        filter_layout.addWidget(QLabel("Engineer:"))
        filter_layout.addWidget(self.engineer_dropdown)
        self.engineer_dropdown.setStyleSheet("background-color: lightblue;")

        self.location_dropdown = QComboBox(self)
        self.location_dropdown.setEditable(True)  # Enable editing
        self.location_dropdown.setInsertPolicy(QComboBox.NoInsert)  # Prevent insertion of new items
        self.location_dropdown.addItems(["All"] + self.locations)
        filter_layout.addWidget(QLabel("Location:"))
        filter_layout.addWidget(self.location_dropdown)
        self.location_dropdown.setStyleSheet("background-color: lightblue;")


        self.from_date_edit = QDateEdit(self)
        self.from_date_edit.setCalendarPopup(True)
        self.from_date_edit.setDisplayFormat("yyyy/MM/dd")
        self.from_date_edit.setDate(QDate.currentDate()) 
        filter_layout.addWidget(QLabel("From Date:"))
        filter_layout.addWidget(self.from_date_edit)
        self.from_date_edit.setStyleSheet("background-color: lightblue;")


        self.to_date_edit = QDateEdit(self)
        self.to_date_edit.setCalendarPopup(True)
        self.to_date_edit.setDisplayFormat("yyyy/MM/dd")
        self.to_date_edit.setDate(QDate.currentDate())  
        filter_layout.addWidget(QLabel("To Date:"))
        filter_layout.addWidget(self.to_date_edit)
        self.to_date_edit.setStyleSheet("background-color: lightblue;")


        self.filter_button = QPushButton("Filter", self)
        self.filter_button.clicked.connect(self.apply_filters)
        filter_layout.addWidget(self.filter_button)
        self.layout.addWidget(self.filter_groupbox)
        self.filter_button.setStyleSheet("background-color: lightblue;")

        
       
        # Styling for Filter Group Box
        self.filter_groupbox.setStyleSheet("QGroupBox { font-size: 16px; \
                                        border: 2px solid black; \
                                        border-radius: 10px; \
                                        padding-top: 20px; \
                                        background-color: #f0f0f0; } \
                                        QGroupBox::title { subcontrol-origin: margin; \
                                        subcontrol-position: top center; \
                                        padding: 0 3px; \
                                        color: black; \
                                        font-weight: bold; }")

        # Table
        self.table = QTableWidget(self)
        self.table.setGeometry(800, 800, 800, 600) 
        self.layout.addWidget(self.table)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Download button
        self.download_button = QPushButton("Download", self)
        self.download_button.clicked.connect(self.export_to_csv)
        self.download_button.setStyleSheet("background-color: lightblue;")


        self.layout.addWidget(self.download_button)

        # Set the table properties
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

         # Load and display booking data
        self.load_booking_data()
        # Display initial data in the table
        self.display_data(self.data)
        
    def load_booking_data(self):
        try:
            with open('booked_slot.json', 'r') as json_file:
                data = json.load(json_file)
                bookings = data.get('bookings', [])
                self.data = bookings
        except FileNotFoundError:
            print(f"Error: File not found ")
        except json.JSONDecodeError:
            print(f"Error: Unable to decode JSON in file")   
    
    def show_filter_dialog(self):
        filter_options_widget = QWidget(self)
        filter_layout = QFormLayout(filter_options_widget)

        engineer_dropdown = QComboBox(self)
        engineer_dropdown.addItems(self.engineers)
        filter_layout.addRow("Engineer:", engineer_dropdown)

        location_dropdown = QComboBox(self)
        location_dropdown.addItems(self.locations)
        filter_layout.addRow("Location:", location_dropdown)

        from_date_dropdown = QDateEdit(self)
        from_date_dropdown.setCalendarPopup(True)
        # from_date_dropdown.setDate(QDate(2023, 12, 1))  # Set default value to December 1, 2023
        filter_layout.addRow("From Date:", from_date_dropdown)

        to_date_dropdown = QDateEdit(self)
        to_date_dropdown.setCalendarPopup(True)
        filter_layout.addRow("To Date:", to_date_dropdown)

        apply_button = QPushButton("Apply", self)
        apply_button.clicked.connect(lambda: self.apply_filters({
            'engineer': engineer_dropdown.currentText(),
            'location': location_dropdown.currentText(),
            'from_date': from_date_dropdown.date().toString('yyyy/MM/dd'),
            'to_date': to_date_dropdown.date().toString('yyyy/MM/dd') if not to_date_dropdown.date().isNull() else None,
        }))
        filter_layout.addRow(apply_button)

        self.layout.addWidget(filter_options_widget)

    def display_data(self, data):
            self.table.clear()
            if not data or not data[0]:
                return
            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(data[0]) + 1)  # Add one extra column for "Remark"
            header_labels = list(data[0].keys()) + ["Remark"] 
            if 'Created Date' not in header_labels:
                header_labels.append('Created Date')
            self.table.setHorizontalHeaderLabels(header_labels)
            for j, key in enumerate(header_labels):
                self.table.setHorizontalHeaderItem(j, QTableWidgetItem(key))
            for i, row in enumerate(data):
                for j, key in enumerate(header_labels):
                    if key == "Remark":  
                        item = QTableWidgetItem(row.get(key, ''))
                        item.setTextAlignment(Qt.AlignCenter)
                        item.setFlags(item.flags() | Qt.ItemIsEditable) 
                    else:
                        item = QTableWidgetItem(str(row.get(key, '')))
                        if key == 'engineer' or key == 'location' or key == 'date' or key == 'duration' or key == 'Created Date':
                            item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(i, j, item)
            for j in range(self.table.columnCount()):
                self.table.setSortingEnabled(True)
                # Apply style sheet to set background color for specific columns and header
            self.table.setStyleSheet(
        "QHeaderView::section {"
        "   background-color: yellow;"
        "   color: black;"  # Set text color to black
        "}"
        "QTableWidget::item {"
        "   background-color: #f5f5f5;"
        "}"
        "QTableWidget {"
        "   alternate-background-color: #f5f5f5;"
        "   gridline-color: black;"  # Set grid line color to black
        "}"
    )
    def export_to_csv(self):
            options = QFileDialog.Options()
            file_name, _ = QFileDialog.getSaveFileName(self, "Save Report", "", "CSV Files (*.csv);;Excel Files (*.xlsx);;All Files (*)", options=options)

            if file_name:
                if file_name.endswith('.csv'):
                    with open(file_name, 'w', newline='') as csvfile:
                        csv_writer = csv.writer(csvfile)
                        
                        # Write header
                        header_labels = [self.table.horizontalHeaderItem(j).text() for j in range(self.table.columnCount())]
                        csv_writer.writerow(header_labels)

                        # Write data
                        for i in range(self.table.rowCount()):
                            row_data = [self.table.item(i, j).text() for j in range(self.table.columnCount())]
                            csv_writer.writerow(row_data)
                elif file_name.endswith('.xlsx'):
                    workbook = xlsxwriter.Workbook(file_name)
                    worksheet = workbook.add_worksheet()

                    # Write header
                    header_labels = [self.table.horizontalHeaderItem(j).text() for j in range(self.table.columnCount())]
                    for j, label in enumerate(header_labels):
                        worksheet.write(0, j, label)

                    # Write data
                    for i in range(self.table.rowCount()):
                        for j in range(self.table.columnCount()):
                            worksheet.write(i + 1, j, self.table.item(i, j).text())

                    workbook.close()
                    
    def filter_table(self):
        
        pass
    def apply_filters(self):
        
        from_date_dropdown = dt.datetime(2023, 12, 1).strftime('%Y/%m/%d')
        selected_to_date = self.to_date_edit.date().toString('yyyy/MM/dd')
        if not selected_to_date:
            selected_to_date = None
        
        filters = {
            'engineer': self.engineer_dropdown.currentText(),
            'location': self.location_dropdown.currentText(),
            'from_date': from_date_dropdown,
            'to_date': selected_to_date,
            'duration': None 
        }

        filtered_data = self.filter_data(filters)
        self.display_data(filtered_data)

    def filter_data(self, filters):
        if not isinstance(filters, dict):
            return []

        filtered_data = self.data.copy()

        if filters.get('engineer') is not None and filters['engineer'] != "All":
            filtered_data = [entry for entry in filtered_data if entry.get('engineer') == filters['engineer']]

        if filters.get('location') and filters['location'] != "All":
            filtered_data = [entry for entry in filtered_data if entry.get('location') == filters['location']]

        if filters.get('from_date') and filters.get('to_date'):
            # Convert filter date strings to datetime objects
            filter_from_date = dt.datetime.strptime(filters['from_date'], '%Y/%m/%d')
            filter_to_date = dt.datetime.strptime(filters['to_date'], '%Y/%m/%d')

            filtered_data = [
                entry for entry in filtered_data
                if 'date' in entry
                and filter_from_date <= dt.datetime.strptime(entry.get('date', ''), '%Y/%m/%d') <= filter_to_date
            ]

        if filters.get('duration') and filters['duration'] != "All":
            filtered_data = [entry for entry in filtered_data if entry.get('duration') == int(filters['duration'].split()[0])]

        return filtered_data

        
    def get_available_dates(self):
        # Extract and return unique dates from the data
        return list(set(entry['date'] for entry in self.data))
    

    def add_booked_data(self, booked_data):
        # Add the 'Created Date' to the booked data
        # booked_data['Created Date'] = dt.datetime.now().strftime('%Y/%m/%d ')
        self.data.append(booked_data)
        self.display_data(self.data)

class OpenUrlThread(QThread):
    finished = pyqtSignal()
    def __init__(self, url):
        super().__init__()
        self.url = url
    def run(self):
        QDesktopServices.openUrl(QUrl(self.url))
        self.finished.emit()

class CustomTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        
    def mouseDoubleClickEvent(self, event):
        index = self.indexAt(event.pos())
        if index.isValid():
            self.edit(index)  # Make the cell editable on double-click
        super().mouseDoubleClickEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            index = self.indexAt(event.pos())
            if index.isValid():
                column_name = self.horizontalHeaderItem(index.column()).text()
                if column_name != "COMPLAINT NO.":
                    menu = QMenu(self)
                    book_slot_action = menu.addAction("Book Slot")
                    action = menu.exec_(self.viewport().mapToGlobal(event.pos()))
                    if action == book_slot_action:
                        QMessageBox.information(self, "Book Slot", "Do you want to book a slot?")
                    else:
                        self.edit(index)  
        elif event.button() == Qt.RightButton:  
            index = self.indexAt(event.pos())
            if index.isValid():
                column_name = self.horizontalHeaderItem(index.column()).text()
                if column_name == "COMPLAINT NO.":
                    menu = QMenu(self)
                    open_google_action = menu.addAction("Open URL")
                    open_google_action.triggered.connect(self.open_google)  
                    menu.exec_(self.viewport().mapToGlobal(event.pos()))
        super().mousePressEvent(event)

    def set_hyperlink_item(self, row, column, text, url):
        item = QTableWidgetItem(text)
        item.setData(Qt.TextColorRole, Qt.blue)  
        item.setData(Qt.TextDecorationRole, QUrl(url)) 
        item.setTextAlignment(Qt.AlignCenter) 
        self.setItem(row, column, item)

    def open_google(self):
        QDesktopServices.openUrl(QUrl("https://daksh.teamerge.in/admin/ticket_tool/ticket_detail"))
    
class ComplaintTab(QWidget):
    def __init__(self):
        super().__init__()
        self.checkbox_states = {}
        self.file_name=None

        # Initialize table
        self.table = CustomTableWidget(self)
        self.table.setShowGrid(True) 
        # Set the initial horizontal header labels
        self.table.setColumnCount(len(self.header_labels))
        self.table.setHorizontalHeaderLabels(self.header_labels)
        # Set number of rows to display blank rows below the headers
        num_blank_rows = 70  # You can adjust this value as needed
        self.table.setRowCount(num_blank_rows)

        # Apply stylesheet to QHeaderView::section
        self.table.setStyleSheet(
            "QHeaderView::section {"
            "   background-color: yellow;"
            "}"
        )

        # Main layout
        self.layout = QVBoxLayout(self)

        # Group box for filtering elements
        filter_groupbox = QGroupBox("File Menu")
        filter_groupbox_layout = QVBoxLayout(filter_groupbox)

        # Search Filter Edit
        search_filter_edit_layout = QHBoxLayout()
        self.search_filter_edit = QLineEdit()
        self.search_filter_edit.setPlaceholderText("Search...")
        self.search_filter_edit.textChanged.connect(self.filter_table)
        self.search_filter_edit.setStyleSheet("background-color: #ADD8E6;")
        search_filter_edit_layout.addWidget(self.search_filter_edit,stretch=1)

        # Save Button
        self.save_button = QPushButton("Save Excel")
        self.save_button.clicked.connect(self.auto_save_to_excel)
        self.save_button.setStyleSheet("background-color: #ADD8E6;")
        search_filter_edit_layout.addWidget(self.save_button,stretch=1)


        # Download Button
        self.download_button = QPushButton("Download Excel")
        self.download_button.clicked.connect(self.export_to_excel)
        self.download_button.setStyleSheet("background-color: #ADD8E6;")
        search_filter_edit_layout.addWidget(self.download_button,stretch=1)

        # Upload Button
        self.upload_button = QPushButton("Upload Excel")
        self.upload_button.clicked.connect(self.upload_excel)
        self.upload_button.setStyleSheet("background-color: #ADD8E6;")
        search_filter_edit_layout.addWidget(self.upload_button, stretch=1)

        # Setting Button
        self.setting_button = QPushButton()
        self.setting_button.setIcon(QIcon("setting.png"))  # Set icon
        self.setting_button.setIconSize(QSize(16, 16))  # Set icon size
        self.setting_button.setStyleSheet("background-color: #ADD8E6;")
        self.setting_button.clicked.connect(self.open_settings)
        search_filter_edit_layout.addWidget(self.setting_button)

        filter_groupbox_layout.addLayout(search_filter_edit_layout)

        # Add filter group box to main layout
        self.layout.addWidget(filter_groupbox)

        # Group box for action buttons
        button_groupbox = QGroupBox("Actions")
        button_groupbox_layout = QVBoxLayout(button_groupbox)

        # Insert Row Button
        insert_button_layout = QHBoxLayout()
        self.insert_row_button = QPushButton("Insert Row")
        self.insert_row_button.clicked.connect(self.insert_row)
        self.insert_row_button.setStyleSheet("background-color: #ADD8E6;")
        insert_button_layout.addWidget(self.insert_row_button)

        # # Insert Column Button
        # self.insert_column_button = QPushButton("Insert Column")
        # self.insert_column_button.clicked.connect(self.insert_column)
        # self.insert_column_button.setStyleSheet("background-color: #ADD8E6;")
        # insert_button_layout.addWidget(self.insert_column_button)

        # Delete Row Button
        self.delete_row_button = QPushButton("Delete Row")
        self.delete_row_button.clicked.connect(self.delete_row)
        self.delete_row_button.setStyleSheet("background-color: #ADD8E6;")
        insert_button_layout.addWidget(self.delete_row_button)

        # Delete Column Button
        self.delete_column_button = QPushButton("Delete Column")
        self.delete_column_button.clicked.connect(self.delete_column)
        self.delete_column_button.setStyleSheet("background-color: #ADD8E6;")
        insert_button_layout.addWidget(self.delete_column_button)

        button_groupbox_layout.addLayout(insert_button_layout)
        self.layout.addWidget(button_groupbox)

        self.layout.addWidget(self.table)

        # Styling for Filter Group Box
        filter_groupbox.setStyleSheet("QGroupBox { font-size: 16px; \
                                 border: 2px solid black; \
                                 border-radius: 10px; \
                                 padding-top: 20px; \
                                 background-color: #f0f0f0; \
                                 } \
                                 QGroupBox::title { subcontrol-origin: margin; \
                                subcontrol-position: top center; \
                                 padding: 0 3px; \
                                 color: black; \
                                 font-weight: bold; \
                                 }")

        # Styling for Button Group Box
        button_groupbox.setStyleSheet("QGroupBox { font-size: 16px; \
                                  border: 2px solid black; \
                                  border-radius: 10px; \
                                  padding-top: 20px; \
                                  background-color: #f0f0f0; \
                                  } \
                                  QGroupBox::title { subcontrol-origin: margin; \
                                  subcontrol-position: top center; \
                                  padding: 0 3px; \
                                  color: black; \
                                  font-weight: bold; \
                                  }")
     
    def set_hyperlink_item(self, row, column, text, url):
        item = QTableWidgetItem(text)
        item.setData(Qt.TextColorRole, Qt.blue)  # Set text color to blue
        item.setData(Qt.DecorationRole, QUrl(url))  # Set decoration role as QUrl
        item.setTextAlignment(Qt.AlignCenter)  # Align text to center
        self.setItem(row, column, item)

    def reset_complaint_tab(self):
        # Show all columns in the complaint tab
        for column in range(self.table.columnCount()):
            self.table.setColumnHidden(column, False)

        # Clear any applied text filter
        self.search_filter_edit.clear()
        self.filter_table()  # Apply the filter to show all rows 

    def open_settings(self):
        # Parse settings.json file to get columns related to complaints
        with open('settings.json', 'r') as file:
            settings_data = json.load(file)
            complaint_columns = settings_data.get("Complaint", {}).get("columns", [])

        # Create a menu to display columns with checkboxes
        menu = QMenu(self.setting_button)
        menu.setStyleSheet("""
            QMenu {
                padding: 10px;
            }
            QMenu::item {
                padding: 5px;
                border: 1px solid #ccc;
                background-color: #f0f0f0;
            }
            QMenu::item:selected {
                background-color: #e0e0e0;
            }
            QMenu::item:hover {
                padding: 5px;
            background-color: #f0f0f0;
            border: 1px solid #ccc;
            border-radius: 4px;
            }
        """)

        for column in complaint_columns:
            checkbox_widget = QWidget()
            layout = QHBoxLayout(checkbox_widget)
            checkbox = QCheckBox()
            checkbox.setText(column)

            # Check the checkbox if it was previously selected
            if self.checkbox_states.get(column, False):
                checkbox.setChecked(True)

            layout.addWidget(checkbox)

            # Create a custom action with the checkbox widget
            action = QWidgetAction(menu)
            action.setDefaultWidget(checkbox_widget)
            menu.addAction(action)

        # Add OK action
        ok_action = menu.addAction("OK")
       
        ok_action.triggered.connect(lambda: self.apply_filter(menu))


        # Add Reset action
        reset_action = menu.addAction("Reset")
        reset_action.triggered.connect(lambda: self.reset_filter(menu))

        # Show the menu below the setting button
        menu.exec_(self.setting_button.mapToGlobal(self.setting_button.rect().bottomLeft()))

    def reset_filter(self, menu):
        # Reset all checkboxes to default state (unchecked)
        for action in menu.actions():
            if isinstance(action, QWidgetAction):
                widget = action.defaultWidget()
                if widget and isinstance(widget, QWidget):
                    checkbox = widget.findChild(QCheckBox)
                    checkbox.setChecked(False)
        self.checkbox_states = {}
        self.reset_complaint_tab()

    def apply_filter(self, menu):
        # Get the selected column names from the checkboxes
        selected_columns = []
        for action in menu.actions():
            if isinstance(action, QWidgetAction):
                widget = action.defaultWidget()
                if widget and isinstance(widget, QWidget):
                    checkbox = widget.findChild(QCheckBox)
                    if checkbox.isChecked():
                        selected_columns.append(checkbox.text())
                        # Store the state of the checkbox
                        self.checkbox_states[checkbox.text()] = True
                    else:
                        # Unchecked checkboxes should not be stored
                        self.checkbox_states.pop(checkbox.text(), None)

        # Show all rows by default
        for row in range(self.table.rowCount()):
            self.table.setRowHidden(row, False)

        # Hide columns not selected in the checkboxes
        for column in range(self.table.columnCount()):
            header_item = self.table.horizontalHeaderItem(column)
            if header_item.text() not in selected_columns:
                self.table.setColumnHidden(column, True)
            else:
                self.table.setColumnHidden(column, False)

        # Apply text filter if needed
        text = self.search_filter_edit.text()
        if text:
            self.filter_table(text)

    def open_url(self):
        index = self.currentIndex()
        if index.isValid():
            item = self.item(index.row(), index.column())
            if item is not None:
                url = item.data(Qt.DecorationRole).toString()
                if url:
                    QDesktopServices.openUrl(QUrl(url))

    def upload_excel(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Upload Excel File", "", "Excel Files (.xlsx *.xls);;All Files ()", options=options)
        if file_name:
            self.table.clear()  # Clear existing data in the table
            self.load_data(file_name)
            self.file_name = file_name  #

           
    def header_labels(settings_file):
        try:
            with open(settings_file, "r") as f:
                settings_data = json.load(f)
                complaint_columns = settings_data.get("Complaint", {}).get("columns", [])
                return complaint_columns
        except FileNotFoundError:
            print("Settings file not found.")
            return []

    # Usage
    settings_file = "settings.json"
    header_labels = header_labels(settings_file)
    print(header_labels)

    def auto_save_to_excel(self):
        try:
            if self.file_name:
                file_name = self.file_name
                workbook = load_workbook(file_name)
                sheet = workbook.active
                # Clear existing data in the sheet
                sheet.delete_rows(2, sheet.max_row)
                # Update the entire table data to the Excel sheet
                for i in range(self.table.rowCount()):
                    for j in range(self.table.columnCount()):
                        item = self.table.item(i, j)
                        value = item.text() if item is not None else ''
                        sheet.cell(row=i + 2, column=j + 1, value=value)
                # Save the changes back to the file
                workbook.save(file_name)
                print(f"Auto-save successful to {file_name}")
        except FileNotFoundError as e:
            print(f"Error: File not found: {e}")
        except PermissionError as e:
            print(f"Error: Permission issue: {e}")
        except openpyxl.utils.exceptions.InvalidFileException as e:
            print(f"Error: Invalid Excel file: {e}")
        except Exception as e:
            print(f"Unexpected error in auto_save_to_excel: {e}")

    def filter_table(self, text=None):
        if text is None:
            text = self.search_filter_edit.text()
        for row in range(self.table.rowCount()):
            row_hidden = True
            for column in range(self.table.columnCount()):
                item = self.table.item(row, column)
                if item is not None and text.lower() in item.text().lower():
                    row_hidden = False
                    break
            self.table.setRowHidden(row, row_hidden)
            
    def insert_row(self):
        current_rows = self.table.rowCount()
        self.table.insertRow(current_rows)
        self.auto_save_to_excel()

    # def insert_column(self):
    #     current_columns = self.table.columnCount()
    #     header_label, ok = QInputDialog.getText(self, "Input Dialog", "Enter column header:")
    #     if ok:
    #         self.table.setColumnCount(current_columns + 1)
    #         self.table.setHorizontalHeaderItem(current_columns, QTableWidgetItem(header_label))
            
    #         # Update header_labels list only if the header_label is not already in the list
    #         if header_label not in self.header_labels:
    #             self.header_labels.append(header_label)
    #             # Save updated column names to file
    #             self.save_column_names(self.header_labels)

    #             # Set the horizontal header labels again
    #             self.table.setHorizontalHeaderLabels(self.header_labels)
    #     self.auto_save_to_excel()

    def save_column_names(self, header_labels):
        settings_file = "settings.json"
        try:
            with open(settings_file, "r") as f:
                settings_data = json.load(f)
                settings_data["Complaint"]["columns"] = header_labels

            with open(settings_file, "w") as f:
                json.dump(settings_data, f, indent=4)
            
            print("Column names updated successfully.")
        except FileNotFoundError:
            print("Settings file not found.")
        except Exception as e:
            print(f"Error while saving column names: {e}")   

    def delete_row(self):
        selected_rows = set(index.row() for index in self.table.selectionModel().selectedRows())
        for row in sorted(selected_rows, reverse=True):
            self.table.removeRow(row)
        self.auto_save_to_excel()

    def delete_column(self):
        try:
            # Get the selected column indexes
            selected_columns = [index.column() for index in self.table.selectionModel().selectedColumns()]

            if self.file_name:
                # Load the workbook
                workbook = openpyxl.load_workbook(self.file_name)
                sheet = workbook.active

                # Remove column data from the file for each selected column
                for col in sorted(selected_columns, reverse=True):
                    sheet.delete_cols(col + 1)  # col + 1 because column indexes in Excel start from 1

                # Save the modified workbook
                workbook.save(self.file_name)
                print(f"Column data removed successfully from {self.file_name}")

            # Update header labels excluding the deleted columns
            new_header_labels = [label for col, label in enumerate(self.header_labels) if col not in selected_columns]
            self.header_labels = new_header_labels
            self.table.setColumnCount(len(self.header_labels))  # Update the table column count
            self.table.setHorizontalHeaderLabels(self.header_labels)

            # Save updated column names to file
            self.save_column_names(self.header_labels)
            
            # Save changes to Excel file
            self.auto_save_to_excel()

        except Exception as e:
            print(f"Error removing column data from file: {e}")


      

    def load_data(self, file_name):
        try:
            workbook = openpyxl.load_workbook(file_name, read_only=True)
            sheet = workbook.active

            # Clear existing data in the table
            self.table.clearContents()
            self.table.setRowCount(0)

            # Load new data
            loaded_data = []
            header_labels = []

            for row_data in sheet.iter_rows(values_only=True):
                if not header_labels:
                    header_labels = [str(value) for value in row_data]
                    continue  # Skip the first row

                loaded_data.append(row_data)

            # Populate the table with the loaded data
            for row_data in loaded_data:
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)
                for col_position, value in enumerate(row_data):
                    # Check if the value is None or NaN (Not a Number)
                    if value is None or (isinstance(value, float) and math.isnan(value)):
                        value = ''  # Replace None and NaN values with empty string
                    item = QTableWidgetItem()
                    if header_labels[col_position] == "COMPLAINT NO.":
                        # Create a font with blue color and underlined
                        font = QFont()
                        font.setUnderline(True)
                        font.setPointSize(10)
                        item.setFont(font)
                        # Set the text color to blue
                        item.setForeground(QColor(Qt.blue))
                        item.setData(Qt.DisplayRole, value)
                        item.setData(Qt.UserRole, value)  # Set the data role to create a hyperlink
                    else:
                        item.setData(Qt.DisplayRole, value)
                    self.table.setItem(row_position, col_position, item)

            # Set the number of columns in the table based on the data
            if loaded_data:
                self.table.setColumnCount(len(loaded_data[0]))

            # Set the new header labels
            self.table.setHorizontalHeaderLabels(header_labels)

            print("Excel data loaded successfully!")
        except FileNotFoundError:
            print(f"Error: File not found: {file_name}")
        except Exception as e:
            print(f"Error reading Excel file: {e}")
                    
    def export_to_excel(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Complaints to Excel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            header_labels = [self.table.horizontalHeaderItem(j).text() for j in range(self.table.columnCount())]
            sheet.append(header_labels)
            for i in range(self.table.rowCount()):
                row_data = [self.table.item(i, j).text() if self.table.item(i, j) is not None else '' for j in range(self.table.columnCount())]
                sheet.append(row_data)
            workbook.save(file_name)

class AppointmentSchedulerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Appointment Scheduler")
        self.setGeometry(100, 100, 1200, 800)
        self.central_widget = QTabWidget(self)
        self.setCentralWidget(self.central_widget)

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        refresh_button = QPushButton("Refresh Data", self)
        refresh_button.clicked.connect(self.refresh_data)
        layout.addWidget(refresh_button)

        # File menu
        self.file_menu = self.menuBar().addMenu("File")
        # Add Location action
        self.add_location_action = QAction("Add Location", self)
        self.add_location_action.triggered.connect(self.add_location)
        self.file_menu.addAction(self.add_location_action)

        # Add Engineer action
        self.add_engineer_action = QAction("Add Engineer", self)
        self.add_engineer_action.triggered.connect(self.add_engineer)
        self.file_menu.addAction(self.add_engineer_action)

        # Delete Location action
        self.delete_location_action = QAction("Delete Location", self)
        self.delete_location_action.triggered.connect(self.delete_location)
        self.file_menu.addAction(self.delete_location_action)

        # Delete Engineer action
        self.delete_engineer_action = QAction("Delete Engineer", self)
        self.delete_engineer_action.triggered.connect(self.delete_engineer)
        self.file_menu.addAction(self.delete_engineer_action)

        # Refresh action
        self.refresh_action = QAction(QIcon("refresh button.png"), "Refresh", self)
        self.refresh_action.triggered.connect(self.refresh_data)
        self.file_menu.addAction(self.refresh_action)

        self.logout_action = QAction(QIcon("logout button.png"), "Logout", self)
        self.logout_action.triggered.connect(self.show_logout_dialog)
        self.file_menu.addAction(self.logout_action)      

        # Calendar tab
        self.calendar_tab = QWidget(self)
        self.central_widget.addTab(self.calendar_tab, "Calendar")

        # Complaint tab
        self.complaint_tab = ComplaintTab()
        self.central_widget.addTab(self.complaint_tab, "Complaint")

        # Create a vertical layout for the calendar tab
        self.calendar_tab_layout = QVBoxLayout(self.calendar_tab)
        self.calendar_and_slot_layout = QHBoxLayout()

        # Calendar widget
        self.calendar = QCalendarWidget(self)
        self.calendar.clicked.connect(self.show_slot_info)
        self.calendar.setContextMenuPolicy(Qt.CustomContextMenu)
        self.calendar.customContextMenuRequested.connect(self.show_context_menu)
        self.calendar_and_slot_layout.addWidget(self.calendar)

        # Book slot button
        self.book_button = QPushButton("Book Slot", self)
        self.book_button.clicked.connect(self.show_book_slot_dialog)
        self.book_button.setShortcut("Return")

        # Frame to display assigned engineers
        self.engineers_frame = QFrame(self)
        self.engineers_frame.setFrameShape(QFrame.StyledPanel)
        self.engineers_frame.setFrameShadow(QFrame.Raised)
        self.engineers_frame.setFixedWidth(450)

        # GroupBox inside the Frame
        self.engineers_group_box = QGroupBox("Booked Slots")
        self.engineers_group_box.setStyleSheet("font-weight: bold; font-size: 11px;")

        self.inner_group_box = QGroupBox("Inner GroupBox")
        self.inner_group_box_layout = QVBoxLayout(self.inner_group_box)

        # Assigned Engineers
        self.assigned_label = QLabel("Assigned Engineers:", self.inner_group_box)
        self.assigned_text_edit = QTextEdit(self.inner_group_box)

        # Unassigned Engineers
        self.unassigned_label = QLabel("Unassigned Engineers:", self.inner_group_box)
        self.unassigned_text_edit = QTextEdit(self.inner_group_box)
        

        # Add the GroupBox to the Frame
        self.engineers_frame_layout = QVBoxLayout(self.engineers_frame)
        self.engineers_frame_layout.addWidget(self.engineers_group_box)

        self.inner_group_box_layout.addWidget(self.assigned_label)
        self.inner_group_box_layout.addWidget(self.assigned_text_edit)
        self.inner_group_box_layout.addWidget(self.unassigned_label)
        self.inner_group_box_layout.addWidget(self.unassigned_text_edit)
        self.engineers_group_box.setLayout(self.inner_group_box_layout)
        self.calendar_tab_layout.addLayout(self.calendar_and_slot_layout)
        self.calendar_and_slot_layout.addWidget(self.engineers_frame)
        self.calendar_tab_layout.addWidget(self.book_button)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_data)
        refresh_interval = 10000
        self.timer.start(refresh_interval)
        
        # Data
        self.data = {"bookings":[], "engineers":[], "locations":[]}
        self.data = self.load_json()
        if len(self.data["locations"]) != 0:
            self.locations = {i:[] for i in self.data["locations"]}
        else:
            self.locations = {"Office":[],"Leave":[],"Candor Tikri":[], "DLF Ultima":[], "DLF Primus":[], "Candor Dhundahera":[], "Regal Garden":[]}
            self.save_json("locations", self.locations)
        
        if len(self.data["engineers"]) != 0:
            self.engineers = {i:[] for i in self.data["engineers"]}
        else:
            self.engineers =  {"Deepak":[], "Raman":[], "Shishupal":[], "Vikas":[], "Sachin":[]}
            self.save_json("engineers", self.engineers)

        self.booked_slots = {}
        self.bookings = self.data["bookings"]
        self.dashboard_tab = DashboardWidget(self.locations.keys(), self.engineers.keys(), self)
        self.central_widget.addTab(self.dashboard_tab, "Dashboard")
        self.reports_widget = ReportsWidget(list(self.engineers), list(self.locations), self)
        self.central_widget.addTab(self.reports_widget, "Reports")
        self.central_widget.tabBarClicked.connect(self.handle_tab_click)
        self.central_widget.setCurrentWidget(self.calendar_tab)

        if self.load_existing_credentials():
            self.show()
        else:
            self.login_window = LoginWindow(self)
            self.login_window.login_signal.connect(self.show)
            self.login_window.show()

    def fetch_latest_data(self):
        try:
            response = requests.get("http://192.168.17.72:5000/sync")
            if response.status_code == 200:
                return response.json()
            else:
                print("Failed to fetch latest data from server.")
                return None
        except Exception as e:
            print(f"An error occurred: {e}")
            return None        

    def refresh_data(self):
        latest_data = self.fetch_latest_data()
        self.data["bookings"] = self.data.get("data", [])
        self.show_slot_info()
        self.update_calendar_colors()
        if latest_data:
            # Update your UI or perform any other actions based on the latest data
            print("Data refreshed successfully.")
        else:
            print("Failed to refresh data. Please try again.")        

    def handle_tab_click(self, index):
        current_tab_text = self.central_widget.tabText(index)
        if current_tab_text == "Logout":
            self.show_logout_dialog()

    def show_logout_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Logout")
        dialog.setGeometry(500, 300, 300, 100)
        layout = QVBoxLayout(dialog)
        confirmation_label = QLabel("Do you want to logout?", dialog)
        layout.addWidget(confirmation_label)
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK", dialog)
        ok_button.clicked.connect(self.logout)
        button_layout.addWidget(ok_button)
        cancel_button = QPushButton("Cancel", dialog)
        cancel_button.clicked.connect(dialog.reject)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        dialog.exec_()

    def logout(self):
        try:
            with open("license.json", "r") as license_file:
                existing_data = json.load(license_file)
                for user in existing_data.get("user", []):
                    user["status"] = "inactive"

            with open("license.json", "w") as license_file:
                json.dump(existing_data, license_file, indent=4)

            self.close()

        except FileNotFoundError:
            QMessageBox.warning(self, "Logout Error", "License file not found. Logout failed.")
    
    def load_existing_credentials(self):
        try:
            with open("license.json", "r") as license_file:
                existing_data = json.load(license_file)
                for i in existing_data["user"]:
                    if i["status"] == "active":
                        return True
                return False
        except FileNotFoundError:
            return False

    def update_calendar_colors(self):   
        selected_date = self.calendar.selectedDate().toString('yyyy/MM/dd')
        assigned_dates = [booking["date"] for booking in self.bookings]

        for date, _ in self.calendar_date_formats.items():
            self.calendar.setDateTextFormat(QDate.fromString(date, 'yyyy/MM/dd'), QTextCharFormat())
        # Assign colors to assigned and unassigned dates
        for date in assigned_dates:
            color = QColor(255, 0, 0)
            # Check if all engineers are fully assigned for the selected date
            engineers_for_date = [booking["engineer"] for booking in self.bookings if booking["date"] == date]
            unique_engineers_for_date = set(engineers_for_date)
            if len(unique_engineers_for_date) == len(self.engineers):
                # color = self.data["color options"][]
                color = QColor(255, 0, 0)
            elif len(unique_engineers_for_date) > 0:
                color = QColor(255, 165, 0)
            else:
                color = QColor(255, 255, 255)

            self.calendar_date_formats[date] = color
        if selected_date not in assigned_dates:
            color = QColor(255, 255, 255)
            self.calendar_date_formats[selected_date] = color
        for date, color in self.calendar_date_formats.items():
            date_format = QDate.fromString(date, 'yyyy/MM/dd')
            text_format = QTextCharFormat()
            text_format.setBackground(color)
            self.calendar.setDateTextFormat(date_format, text_format)
    def show_context_menu(self, pos):
        selected_date = self.calendar.selectedDate().toString('yyyy/MM/dd')
        filtered_bookings = [booking for booking in self.bookings if booking["date"] == selected_date]
        if not filtered_bookings:
        # If no bookings, show an error message and return
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle('No Bookings')
            error_box.setText('No bookings found for the selected date.')
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            return
        context_menu = QMenu(self)
        show_dialog_action = context_menu.addAction("Edit")
        # Connect the "Edit" action to the show_dialog_function with the filtered bookings and selected date as arguments
        show_dialog_action.triggered.connect(lambda: self.show_dialog_function(filtered_bookings, selected_date))
        context_menu.exec_(self.calendar.mapToGlobal(pos))

    def show_dialog_function(self, filtered_bookings, selected_date):
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Slot")
        dialog.setGeometry(300, 300, 675, 290)
        layout = QVBoxLayout(dialog)
        # Create or update the table directly
        self.populate_table(dialog, filtered_bookings, selected_date)
        # OK button
        ok_button = QPushButton("OK", dialog)
        ok_button.clicked.connect(dialog.accept)
        layout.addWidget(ok_button)
        # Cancel button
        cancel_button = QPushButton("Cancel", dialog)
        cancel_button.clicked.connect(dialog.reject)
        layout.addWidget(cancel_button)
        result = dialog.exec_()
        if result == QDialog.Accepted:
            print("OK button clicked")
            self.save_json("bookings", self.bookings)
        else:
            print("Cancel button clicked")

    def populate_table(self, dialog, filtered_bookings, selected_date):
        booking_table_widget = dialog.findChild(QTableWidget, "booking_table_widget")
        if not booking_table_widget:
            booking_table_widget = QTableWidget(dialog)
            booking_table_widget.setObjectName("booking_table_widget")
            layout = dialog.layout()
            layout.addWidget(booking_table_widget)
        booking_table_widget.clearContents()
        booking_table_widget.setRowCount(len(filtered_bookings))
        booking_table_widget.setColumnCount(6)
        booking_table_widget.setHorizontalHeaderLabels(["Engineer", "Location", "Date", "Duration", "Edit", "Delete"])
        for row, booking in enumerate(filtered_bookings):
            booking_table_widget.setItem(row, 0, QTableWidgetItem(booking["engineer"]))
            booking_table_widget.setItem(row, 1, QTableWidgetItem(booking["location"]))
            booking_table_widget.setItem(row, 2, QTableWidgetItem(booking["date"]))

            duration_str = f"{booking['duration']} {'Hour' if booking['duration'] == 1 else 'Hours'}"
            booking_table_widget.setItem(row, 3, QTableWidgetItem(duration_str))

            edit_button = QPushButton("Edit", booking_table_widget)
            edit_button.clicked.connect(lambda _, b=booking: self.update_booking(b, dialog))
            booking_table_widget.setCellWidget(row, 4, edit_button)

            delete_button = QPushButton("Delete", booking_table_widget)
            delete_button.clicked.connect(lambda _, b=booking: self.delete_booking(b, dialog))
            booking_table_widget.setCellWidget(row, 5, delete_button)

    def delete_booking(self, booking, parent_dialog):
            confirmation = QMessageBox.question(
                self,
                "Delete Booking",
                f"Do you want to delete the booking for {booking['engineer']} on {booking['date']}?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if confirmation == QMessageBox.Yes:
                delete_api_url = "http://192.168.17.72:5000/delete_booking"
                headers = {"Content-Type": "application/json"}

                delete_data = {
                    "engineer": booking["engineer"],
                    "location": booking["location"],
                    "date": booking["date"],
                    "duration": booking["duration"]
                }

                try:
                    response = requests.delete(delete_api_url, data=json.dumps(delete_data), headers=headers)
                    response.raise_for_status()
                    delete_response = response.json()

                    if "message" in delete_response:
                        print(delete_response["message"])
                        self.bookings.remove(booking)
                        self.save_json("bookings", self.bookings)
                        self.show_slot_info()
                        parent_dialog.accept()
                    else:
                        print("Failed to delete booking. API response:", delete_response)
                        QMessageBox.warning(self, "Delete Failed", "Failed to delete booking. Please try again.")
                except requests.exceptions.RequestException as e:
                    print(f"Error making delete_booking API request: {e}")
                    QMessageBox.critical(self, "Delete Error", f"Error making delete_booking API request: {e}")

    def update_booking(self, booking, parent_dialog):
        dialog = QDialog(self)
        dialog.setWindowTitle("Update Booking")
        dialog.setGeometry(300, 300, 400, 200)

        layout = QVBoxLayout(dialog)

        engineer_dropdown = QComboBox(dialog)
        engineer_dropdown.addItems(self.engineers.keys())
        engineer_dropdown.setCurrentText(booking["engineer"])

        location_dropdown = QComboBox(dialog)
        location_dropdown.addItems(self.locations.keys())
        location_dropdown.setCurrentText(booking["location"])

        date_edit = QDateEdit(dialog)
        date_edit.setDate(QDate.fromString(booking["date"], "yyyy/MM/dd"))

        duration_spinbox = QSpinBox(dialog)
        duration_spinbox.setMinimum(1)
        duration_spinbox.setMaximum(8)
        duration_spinbox.setValue(booking["duration"])

        ok_button = QPushButton("OK", dialog)
        ok_button.clicked.connect(lambda: self.update_booking_callback(booking, engineer_dropdown, location_dropdown, date_edit, duration_spinbox, dialog, parent_dialog))

        cancel_button = QPushButton("Cancel", dialog)
        cancel_button.clicked.connect(dialog.reject)

        layout.addWidget(QLabel("Engineer:"))
        layout.addWidget(engineer_dropdown)

        layout.addWidget(QLabel("Location:"))
        layout.addWidget(location_dropdown)

        layout.addWidget(QLabel("Date:"))
        layout.addWidget(date_edit)

        layout.addWidget(QLabel("Duration (hours):"))
        layout.addWidget(duration_spinbox)

        layout.addWidget(ok_button)
        layout.addWidget(cancel_button)

        dialog.exec_()

    def update_booking_callback(self, booking, engineer_dropdown, location_dropdown, date_edit, duration_spinbox, dialog, parent_dialog):
        new_booking_data = {
            "engineer": engineer_dropdown.currentText(),
            "location": location_dropdown.currentText(),
            "date": date_edit.date().toString("yyyy/MM/dd"),
            "duration": duration_spinbox.value()
        }

        update_api_url = "http://192.168.17.72:5000/update_booking"
        headers = {"Content-Type": "application/json"}

        payload = {
            "engineer": booking["engineer"],
            "location": booking["location"],
            "date": booking["date"],
            "duration": booking["duration"],
            "update": new_booking_data
        }

        try:
            response = requests.put(update_api_url, data=json.dumps(payload), headers=headers)
            response.raise_for_status()
            update_response = response.json()

            if "message" in update_response:
                print(update_response["message"])
                booking.update(new_booking_data)
                self.save_json("bookings", self.bookings)
                self.show_slot_info()
                dialog.accept()
            else:
                print("Failed to update booking. API response:", update_response)
                QMessageBox.warning(self, "Update Failed", "Failed to update booking. Please try again.")
        except requests.exceptions.RequestException as e:
            print(f"Error making update_booking API request: {e}")
            QMessageBox.critical(self, "Update Error", f"Error making update_booking API request: {e}")

    def show_slot_info(self,parent=None):
        # if len(self.bookings) == 0:
        #     return

        # Clear existing content in the inner_group_box_layout
        for i in reversed(range(self.inner_group_box_layout.count())):
            widgetToRemove = self.inner_group_box_layout.itemAt(i).widget()
            self.inner_group_box_layout.removeWidget(widgetToRemove)
            widgetToRemove.setParent(None)

        selected_date_str = self.calendar.selectedDate().toString('yyyy/MM/dd')

        assigned_engineers = []
        unassigned_engineers = set(self.engineers.keys())

        booking_info = None

        for booking in self.bookings:
            if selected_date_str == booking["date"]:
                booking_info = {
                    "engineer": booking["engineer"],
                    "site": booking["location"],
                    "duration": f"{booking['duration']} {'Hour' if booking['duration'] == 1 else 'Hours'}",
                    "date": booking["date"]
                }

                
                combination_exists = any(
                    entry and
                    entry.get('engineer') == booking_info['engineer'] and
                    entry.get('site') == booking_info['site'] and
                    entry.get('duration') == booking_info['duration'] and
                    entry.get('date') == booking_info['date']
                    for entry in assigned_engineers
                )

                if not combination_exists:
                    assigned_engineers.append(booking_info)
                    unassigned_engineers.discard(booking["engineer"])

        assigned_table = QTableWidget(self)
        assigned_table.setColumnCount(3)
        assigned_table.setHorizontalHeaderLabels(["Engineer", "Site", "Duration"])

        assigned_table.setRowCount(len(assigned_engineers))
        for row, assigned_engineer in enumerate(assigned_engineers):
            assigned_table.setItem(row, 0, QTableWidgetItem(assigned_engineer['engineer']))
            assigned_table.setItem(row, 1, QTableWidgetItem(assigned_engineer['site']))
            assigned_table.setItem(row, 2, QTableWidgetItem(assigned_engineer['duration']))

        self.inner_group_box_layout.addWidget(assigned_table)


        if unassigned_engineers:
            unassigned_table = QTableWidget(self)
            unassigned_table.setColumnCount(1)
            unassigned_table.setHorizontalHeaderLabels(["Unassigned Engineers"])

            # Populate QTableWidget with unassigned engineer information
            unassigned_table.setRowCount(len(unassigned_engineers))
            for row, unassigned_engineer in enumerate(unassigned_engineers):
                unassigned_table.setItem(row, 0, QTableWidgetItem(unassigned_engineer))

            self.inner_group_box_layout.addWidget(unassigned_table)

            # Make the last column of unassigned_table take up the remaining space
            unassigned_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

        #Update calendar colors
        self.update_calendar_colors()

    def add_location(self):
        location, ok = QInputDialog.getText(self, "Add Location", "Enter the location name:")
        if ok and location:
            self.locations[location] = []
            self.save_json("locations", self.locations)
           
            api_url = "http://192.168.17.72:5000/api/add_location"
            response = requests.post(api_url, json={"location": location})
            
            if response.status_code == 200:
                QMessageBox.information(self, "Success", "Location added successfully!")
            else:
                QMessageBox.warning(self, "Success", "Location added successfully!")

    def delete_location(self):
        location, ok = QInputDialog.getItem(self, "Delete Location", "Select location to delete:", self.locations.keys(), 0, False)
        if ok and location:
            del self.locations[location]
            self.save_json("locations", self.locations)
            
            api_url = "http://192.168.17.72:5000/api/delete_location"
            response = requests.delete(api_url, json={"location": location})
            
            if response.status_code == 200:
                QMessageBox.information(self, "Success", "Location deleted successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to delete location. Please try again.")

    def delete_engineer(self):
        engineer, ok = QInputDialog.getItem(self, "Delete Engineer", "Select engineer to delete:", self.engineers.keys(), 0, False)
        if ok and engineer:
            del self.engineers[engineer]
            self.save_json("engineers", self.engineers)
            self.booked_slots.pop(engineer, None)
            self.show_slot_info()
           
            api_url = "http://192.168.17.72:5000/api/delete_engineer"
            response = requests.delete(api_url, json={"engineer": engineer})
            
            if response.status_code == 200:
                QMessageBox.information(self, "Success", "Engineer deleted successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to delete engineer. Please try again.")

    def add_engineer(self):
        engineer, ok = QInputDialog.getText(self, "Add Engineer", "Enter the engineer's name:")
        if ok and engineer:

            api_url = "http://192.168.17.72:5000/api/add_engineer"
            response = requests.post(api_url, json={"engineer": engineer})
            self.engineers[engineer] = []
            self.save_json("engineers", self.engineers)
          
            
            
            if response.status_code == 200:
                QMessageBox.information(self, "Success", "Engineer added successfully!")
            else:
                QMessageBox.warning(self, "Success", "Engineer added successfully!")

    def select_location_from_dropdown(self, index):
        location = self.location_dropdown.itemText(index)
        self.select_location(location)

    def select_engineer_from_dropdown(self, index):
        engineer = self.engineer_dropdown.itemText(index)
        self.select_engineer(engineer)

    def select_location(self, location):
        self.current_location = location
        # self.location_label.setText(f"Location: {location}\nBooked Slots:\n{', '.join(self.locations[location])}")

    def select_engineer(self, engineer):
        self.current_engineer = engineer

    def show_book_slot_dialog(self):
        dialog = TimeSlotDialog(self.locations.keys(), self.engineers.keys(), self)
        if dialog.exec_():
            selected_data = dialog.get_selected_data()
            self.book_slot(selected_data)



    #Sagar
    def book_slot(self, selected_data):
        date = self.calendar.selectedDate()
        duration = selected_data['time']
        slot = {
            'engineer': selected_data['engineer'],
            'location': selected_data['location'],
            'date': date.toString('yyyy/MM/dd'),
            'duration': duration
         
        }
        if any(
                booking["engineer"] == selected_data["engineer"]
                and booking["location"] == selected_data["location"]
                and booking["date"] == date.toString("yyyy/MM/dd")
                and booking["duration"] == duration
                for booking in self.bookings
        ):
            print("Duplicate entry. Booking not added.")
            QMessageBox.warning(self, "Duplicate Entry", "Booking not added. Duplicate entry.")
            return

        self.reports_widget.add_booked_data(slot)

        if selected_data['engineer'] not in self.booked_slots:
            self.booked_slots[selected_data['engineer']] = []

        if slot not in self.booked_slots[selected_data['engineer']]:
            self.booked_slots[selected_data['engineer']].append(slot)

            data = {
                "engineer": selected_data['engineer'],
                "location": selected_data['location'],
                "duration": duration,
                "date": date.toString('yyyy/MM/dd')
            }

            booking_api_url = "http://192.168.17.72:5000/api/booking"
            headers = {"Content-Type": "application/json"}

            try:
                response = requests.post(booking_api_url, data=json.dumps(data), headers=headers)
                response.raise_for_status()
                booking_response = response.json()

                if "booking_id" in booking_response:
                    booking_id = booking_response["booking_id"]
                    print("Booking successful! Booking ID:", booking_id)
                    data["booking_id"] = booking_id
                    QMessageBox.information(self, "Booking Successful", f"Booking successful! Booking ID: {booking_id}")
                else:
                    print("Booking failed. API response:", booking_response)
                    QMessageBox.warning(self, "Booking Failed", "Booking failed. Please try again.")
            except requests.exceptions.RequestException as e:
                print(f"Error making booking API request: {e}")
                QMessageBox.critical(self, "Booking Error", f"Error making booking API request: {e}")
            self.bookings.append(data)
            self.update_location_label(selected_data['location'])
            self.save_json("bookings", self.bookings)
        else:
            print("Duplicate entry. Booking not added.")
            QMessageBox.warning(self, "Duplicate Entry", "Booking not added. Duplicate entry.")
        #Ankit
        date = self.calendar.selectedDate()
        duration = selected_data['time']
        slot = {
                'engineer': selected_data['engineer'],
                'location': selected_data['location'],
                'date': date.toString('yyyy/MM/dd'),
                'duration': duration,
                'Created Date': dt.datetime.now().strftime('%Y/%m/%d')  
        }
        self.reports_widget.add_booked_data(slot)     
    
        
    def update_location_label(self, location):
        if location in self.locations:
            self.booked_slots = {}
            booked_slots = "\n".join(self.locations[location])
        # self.location_label.setText(f"Location: {location}\nBooked Slots:\n{booked_slots}")
        else:
        # Handle the case when the location is not found
            booked_slots = "No bookings for this location"

    def update_dashboard(self):
        self.dashboard_tab.update_charts(self.bookings)
    
    def save_json(self, what_to_save, data):
        # save json
        self.data[what_to_save] = data

        with open ("booked_slot.json","w+") as json_file:
            json.dump(self.data,json_file,indent=4)
    
    
    def initialize_calendar_colors(self):
        self.calendar_date_formats = {}

        for booking in self.bookings:
            date = booking["date"]
            if date not in self.calendar_date_formats:
                self.calendar_date_formats[date] = {"assigned": 0, "unassigned": 0}

            if booking["engineer"] in self.booked_slots:
                # Fully assigned slots in red
                self.calendar_date_formats[date]["assigned"] += booking["duration"]
            else:
                # Unassigned slots in green
                self.calendar_date_formats[date]["unassigned"] += booking["duration"]

        selected_date = self.calendar.selectedDate().toString('yyyy/MM/dd')
        if selected_date not in self.calendar_date_formats:
            self.calendar_date_formats[selected_date] = {"assigned": 0, "unassigned": 0}

        self.update_calendar_colors()

        
    def fetch_data_from_api(self, from_date=None, to_date=None):
        api_url = "http://192.168.17.72:5000/api/fetchdata"
        payload = {"from": from_date, "to": to_date}
        headers = {"Content-Type": "application/json"}

        try:
            response = requests.post(api_url, json=payload, headers=headers)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data from API: {e}")
            return {"status": "Failure", "error_code": 500, "message": "Error fetching data from API", "data": {}}

        
    def load_json(self):
        try:
            data = self.fetch_data_from_api()["data"]
            bookings = data.get("bookings", [])
            
            if "locations" in data and "engineers" in data:
                locations = data["locations"]
                engineers = data["engineers"]
            else:
                print("No 'locations' and 'engineers' keys found in API response. Loading from local file.")
                with open("booked_slot.json") as json_file:
                    local_data = json.load(json_file)

                locations = local_data.get("locations", [])
                engineers = local_data.get("engineers", [])

            print({"bookings": bookings, "engineers": engineers, "locations": locations})
            return {"bookings": bookings, "engineers": engineers, "locations": locations}

        except Exception as e:
            print(f"An exception occurred while loading data: {e}")
            return {"bookings": [], "engineers": [], "locations": []}
        
    def showEvent(self, event):
        super().showEvent(event)
        self.initialize_calendar_colors()
    
    # def refresh_fetchdata(self):
    #     api_data = self.fetch_data_from_api()
    #     self.data["bookings"] = self.data.get("data", [])
    #     self.show_slot_info()
    #     self.update_calendar_colors()

    #     if api_data["status"] == "Success":
    #         self.log_message("Data has been refreshed successfully.")
    #     else:
    #         self.log_message("Failed to refresh data. Please try again.")

    # def log_message(self, message):
    #    print(message) 
       
class LoginWindow(QDialog):
    login_signal = pyqtSignal()
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Login")
        self.setGeometry(500, 200, 300, 200)

        layout = QVBoxLayout(self)

        self.email_label = QLabel("Email:")
        self.email_input = QLineEdit(self)

        self.password_label = QLabel("Password:")
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)

        self.login_button = QPushButton("Login", self)
        self.login_button.clicked.connect(self.login)

        layout.addWidget(self.email_label)
        layout.addWidget(self.email_input)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)
        layout.addWidget(self.login_button)

        self.setLayout(layout)


    def login(self):
        email = self.email_input.text()
        password = self.password_input.text()
        if email == "user" and password == "123":
            self.login_signal.emit()  # Emit the login signal
            self.close()
        else:
            error_message = "Invalid credentials. Please try again."
            QMessageBox.warning(self, "Login Error", error_message)

    def login(self):
        email = self.email_input.text()
        password = self.password_input.text()

        # Make the login API request
        login_url = "http://192.168.17.72:5000/api/login"
        payload = {"email": email, "password": password}
        headers = {"Content-Type": "application/json"}

        response = requests.post(login_url, data=json.dumps(payload), headers=headers)

        print("Response Status Code:", response.status_code)
        print("Response Content:", response.content)

        if response.status_code == 200 and response.json().get("message") == "Login successful":
            self.login_signal.emit()
            self.accept()

            # Save credentials to license.json after successful login
            self.create_license_file(email)
        else:
            error_message = "Invalid credentials. Please try again."
            QMessageBox.warning(self, "Login Error", error_message)

    def load_existing_credentials(self):
        try:
            with open("license.json", "r") as license_file:
                existing_data = json.load(license_file)
                for i in existing_data["user"]:
                    if i["status"] == "active":
                        return True
                return False
        except FileNotFoundError:
            return False

    def create_license_file(self, email, is_active=True):
        current_datetime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        user_name = email.split('@')[0][:6]

        license_entry = {
            "user_name": user_name,
            "email": email,
            "datetime": current_datetime,
            "status": "active" if is_active else "inactive"
        }

        try:
            with open("license.json", "r") as license_file:
                existing_data = json.load(license_file)

            user_entry = next((user for user in existing_data.get("user", []) if user["email"] == email), None)

            if user_entry:
                user_entry.update(license_entry)
            else:
                existing_data.setdefault("user", []).append(license_entry)
            with open("license.json", "w") as license_file:
                json.dump(existing_data, license_file, indent=4)

        except FileNotFoundError:
            with open("license.json", "w") as license_file:
                json.dump({"user": [license_entry]}, license_file, indent=4)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_app = AppointmentSchedulerApp()
    # main_app.show()
    sys.exit(app.exec_())

