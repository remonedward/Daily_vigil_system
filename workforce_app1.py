import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QComboBox, QLineEdit, QDateEdit, QPushButton, QTableWidget, 
                             QTableWidgetItem, QLabel, QGroupBox, QMessageBox)
from PyQt5.QtCore import QDate, Qt
from sqlalchemy import create_engine, Column, Integer, String, Date
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import datetime, date
import pandas as pd


Base = declarative_base()

class WorkforceAllocation(Base):
    __tablename__ = 'workforce_allocation'
    id = Column(Integer, primary_key=True)
    department = Column(String)
    cairo_count = Column(Integer)
    tenth_count = Column(Integer)
    date = Column(Date)


engine = create_engine('sqlite:///workforce.db', echo=False)
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)
session = Session()


def load_departments():
    departments = session.query(WorkforceAllocation.department).distinct().all()
    return sorted([dept[0] for dept in departments])

class WorkforceApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('نظام إدارة السهر اليومي >>BY**REMO_OX**<<')
        self.setGeometry(100, 100, 800, 600)
        
        
        self.edit_id = None
        
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)
        
        
        main_widget.setLayoutDirection(Qt.RightToLeft)
        
        
        entry_group = QGroupBox('إدخال البيانات')
        entry_layout = QVBoxLayout()
        
        self.dept_combo = QComboBox()
        self.dept_combo.addItems([''] + load_departments())
        self.dept_combo.setEditable(True)  
        entry_layout.addWidget(QLabel('القسم'))
        entry_layout.addWidget(self.dept_combo)
        
        self.cairo_input = QLineEdit()
        self.cairo_input.setPlaceholderText('0')
        entry_layout.addWidget(QLabel('عدد العاملين - القاهرة'))
        entry_layout.addWidget(self.cairo_input)
        
        self.tenth_input = QLineEdit()
        self.tenth_input.setPlaceholderText('0')
        entry_layout.addWidget(QLabel('عدد العاملين - العاشر'))
        entry_layout.addWidget(self.tenth_input)
        
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(QDate.currentDate())
        entry_layout.addWidget(QLabel('التاريخ'))
        entry_layout.addWidget(self.date_input)
        
        self.submit_btn = QPushButton('حفظ')
        self.submit_btn.clicked.connect(self.save_data)
        entry_layout.addWidget(self.submit_btn)
        
        entry_group.setLayout(entry_layout)
        main_layout.addWidget(entry_group)
        
        
        report_group = QGroupBox('إصدار التقارير')
        report_layout = QVBoxLayout()
        
        self.report_dept_combo = QComboBox()
        self.report_dept_combo.addItems(['الكل'] + load_departments())
        report_layout.addWidget(QLabel('القسم'))
        report_layout.addWidget(self.report_dept_combo)
        
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(QDate.currentDate().addDays(-30))
        report_layout.addWidget(QLabel('تاريخ البداية'))
        report_layout.addWidget(self.start_date)
        
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(QDate.currentDate())
        report_layout.addWidget(QLabel('تاريخ النهاية'))
        report_layout.addWidget(self.end_date)
        
        
        button_layout = QHBoxLayout()
        self.generate_btn = QPushButton('إصدار التقرير')
        self.generate_btn.clicked.connect(self.generate_report)
        button_layout.addWidget(self.generate_btn)
        
        self.edit_btn = QPushButton('تعديل')
        self.edit_btn.clicked.connect(self.load_for_edit)
        button_layout.addWidget(self.edit_btn)
        
        self.export_btn = QPushButton('تصدير إلى إكسيل')
        self.export_btn.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_btn)
        
        self.exit_btn = QPushButton('خروج')
        self.exit_btn.clicked.connect(QApplication.quit)
        button_layout.addWidget(self.exit_btn)
        
        report_layout.addLayout(button_layout)
        
        self.report_table = QTableWidget()
        self.report_table.setColumnCount(5)
        self.report_table.setHorizontalHeaderLabels(['المعرف', 'القسم', 'القاهرة', 'العاشر', 'التاريخ'])
        report_layout.addWidget(self.report_table)
        
        
        self.total_cairo_label = QLabel('إجمالي العاملين - القاهرة: 0')
        report_layout.addWidget(self.total_cairo_label)
        self.total_tenth_label = QLabel('إجمالي العاملين - العاشر: 0')
        report_layout.addWidget(self.total_tenth_label)
        self.total_daily_label = QLabel('الإجمالي اليومي: 0')
        report_layout.addWidget(self.total_daily_label)
        
        report_group.setLayout(report_layout)
        main_layout.addWidget(report_group)
        
    def save_data(self):
        department = self.dept_combo.currentText().strip()
        try:
            cairo_count = int(self.cairo_input.text()) if self.cairo_input.text() else 0
            tenth_count = int(self.tenth_input.text()) if self.tenth_input.text() else 0
        except ValueError:
            self.cairo_input.setText('0')
            self.tenth_input.setText('0')
            QMessageBox.warning(self, 'تحذير', 'يرجى إدخال أرقام صحيحة لعدد العاملين!')
            return
        
        if not department or cairo_count < 0 or tenth_count < 0:
            QMessageBox.warning(self, 'تحذير', 'يرجى إدخال قسم صالح وأعداد غير سالبة!')
            return
        
        date = self.date_input.date().toPyDate()
        
        if self.edit_id:  
            record = session.query(WorkforceAllocation).filter_by(id=self.edit_id).first()
            if record:
                record.department = department
                record.cairo_count = cairo_count
                record.tenth_count = tenth_count
                record.date = date
                session.commit()
                QMessageBox.information(self, 'نجاح', 'تم تحديث البيانات بنجاح!')
        else:  
            session.add(WorkforceAllocation(
                department=department,
                cairo_count=cairo_count,
                tenth_count=tenth_count,
                date=date
            ))
            session.commit()
            QMessageBox.information(self, 'نجاح', 'تم حفظ البيانات بنجاح!')
        
        
        if department not in [self.dept_combo.itemText(i) for i in range(self.dept_combo.count())]:
            self.dept_combo.addItem(department)
            self.report_dept_combo.addItem(department)
        
        
        self.reset_form()
        
    def reset_form(self):
        self.cairo_input.clear()
        self.tenth_input.clear()
        self.dept_combo.setCurrentIndex(0)
        self.date_input.setDate(QDate.currentDate())
        self.edit_id = None
        self.submit_btn.setText('حفظ')
        
    def load_for_edit(self):
        selected_rows = self.report_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, 'تحذير', 'يرجى تحديد صف للتعديل!')
            return
        
        row = selected_rows[0].row()
        record_id = int(self.report_table.item(row, 0).text())
        department = self.report_table.item(row, 1).text()
        cairo_count = self.report_table.item(row, 2).text()
        tenth_count = self.report_table.item(row, 3).text()
        date_str = self.report_table.item(row, 4).text()
        
        
        self.dept_combo.setCurrentText(department)
        self.cairo_input.setText(cairo_count)
        self.tenth_input.setText(tenth_count)
        date = datetime.strptime(date_str, '%d/%m/%Y')
        self.date_input.setDate(QDate(date.year, date.month, date.day))
        
        
        self.edit_id = record_id
        self.submit_btn.setText('حفظ التعديلات')
        
    def generate_report(self):
        department = self.report_dept_combo.currentText()
        start_date = self.start_date.date().toPyDate()
        end_date = self.end_date.date().toPyDate()
        
        query = session.query(WorkforceAllocation)
        if department != 'الكل':
            query = query.filter_by(department=department)
        query = query.filter(WorkforceAllocation.date.between(start_date, end_date))
        
        results = query.all()
        self.report_table.setRowCount(len(results))
        
        
        total_cairo = 0
        total_tenth = 0
        for row, alloc in enumerate(results):
            self.report_table.setItem(row, 0, QTableWidgetItem(str(alloc.id)))
            self.report_table.setItem(row, 1, QTableWidgetItem(alloc.department))
            self.report_table.setItem(row, 2, QTableWidgetItem(str(alloc.cairo_count)))
            self.report_table.setItem(row, 3, QTableWidgetItem(str(alloc.tenth_count)))
            self.report_table.setItem(row, 4, QTableWidgetItem(alloc.date.strftime('%d/%m/%Y')))
            total_cairo += alloc.cairo_count
            total_tenth += alloc.tenth_count
        
        self.report_table.resizeColumnsToContents()
        
        
        self.total_cairo_label.setText(f'إجمالي العاملين - القاهرة: {total_cairo}')
        self.total_tenth_label.setText(f'إجمالي العاملين - العاشر: {total_tenth}')
        self.total_daily_label.setText(f'الإجمالي اليومي: {total_cairo + total_tenth}')
        
    def export_to_excel(self):
        if self.report_table.rowCount() == 0:
            QMessageBox.warning(self, 'تحذير', 'لا توجد بيانات لتصديرها!')
            return
        
        
        data = []
        for row in range(self.report_table.rowCount()):
            row_data = {
                'المعرف': self.report_table.item(row, 0).text(),
                'القسم': self.report_table.item(row, 1).text(),
                'القاهرة': int(self.report_table.item(row, 2).text()),
                'العاشر': int(self.report_table.item(row, 3).text()),
                'التاريخ': self.report_table.item(row, 4).text()
            }
            data.append(row_data)
        
        
        df = pd.DataFrame(data, columns=['المعرف', 'القسم', 'القاهرة', 'العاشر', 'التاريخ'])
        
        
        total_cairo = sum(int(self.report_table.item(row, 2).text()) for row in range(self.report_table.rowCount()))
        total_tenth = sum(int(self.report_table.item(row, 3).text()) for row in range(self.report_table.rowCount()))
        total_row = pd.DataFrame([{
            'المعرف': '', 'القسم': 'الإجمالي', 'القاهرة': total_cairo, 
            'العاشر': total_tenth, 'التاريخ': f'الإجمالي اليومي: {total_cairo + total_tenth}'
        }])
        df = pd.concat([df, total_row], ignore_index=True)
        
        
        today = date.today().strftime('%Y-%m-%d')
        filename = f'report_{today}.xlsx'
        
        
        df.to_excel(filename, index=False, engine='openpyxl')
        
        
        QMessageBox.information(self, 'نجاح', f'تم تصدير التقرير إلى {filename}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WorkforceApp()
    window.show()
    sys.exit(app.exec_())