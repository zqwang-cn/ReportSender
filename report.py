#!/bin/python
import os
import sys
import json
import datetime
import calendar
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from docxtpl import DocxTemplate, R
import openpyxl
from jinja2 import Template
from PyQt5 import QtCore
from PyQt5.QtWidgets import QDialog, QLabel, QLineEdit, QTextEdit, QGridLayout, QApplication, QPushButton, QMessageBox


daily_template = 'daily_template.docx'
weekly_template = 'weekly_template.xlsx'
content_filename = 'content.json'
archive_folder = 'archives/'
daily_filename_format = 'daily-{name}-{date}.docx'
weekly_filename_format = 'weekly-{name}-{date}.xlsx'
daily_subject_format = 'daily-{name}-{date}'
weekly_subject_format = 'weekly-{name}-{date}'

default_content = {
    'settings': {},
    'daily': {},
    'weekly': {},
}
content_types = {
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}

# UI configs
settings_widgets_config = [
    ('server_name', 'SMTP server name:'),
    ('server_port', 'SMTP server port:'),
    ('sender', 'Email account:'),
    ('password', 'Email password'),
    ('to', 'To:'),
    ('cc', 'Cc:'),
    ('name', 'Name:'),
]
weekly_widgets_config = [
    ('conclusion', 'Conclusions'),
    ('progress', 'Progresses'),
    ('plan', 'Plans'),
]
daily_label_texts = [
    'Mon.',
    'Tue.',
    'Wed.',
    'Thur.',
    'Fri.',
]
daily_buttons_config = [
    ('settings', 'Settings'),
    ('edit_weekly', 'Edit weekly'),
    ('clear', 'Clear'),
    ('send_daily', 'Send daily'),
]
daily_widgets_config = [
    ('conclusion', 'Conclusions'),
    ('plan', 'Plans'),
]


def send_email(server_name, server_port, sender, password, to, cc, subject, filenames):
    """Send email with attachments.

    Args:
        server_name (str): SMTP server name
        server_port (str): SMTP server port
        sender (str): email account
        password (str): email password
        to (str): receiver emails (separated by ',')
        cc (str): cc receiver emails (separated by ',')
        subject (str): email subject
        filenames (list): filenames of attachments
    """
    # Create message
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = to
    message['Cc'] = cc
    message['Subject'] = Header(subject, 'utf-8')
    # Attach files
    for filename in filenames:
        with open(filename, 'rb') as f:
            data = f.read()
        att = MIMEText(data, 'base64', 'utf-8')
        att["Content-Type"] = content_types[os.path.splitext(filename)[1]]
        att.add_header("Content-Disposition", "attachment", filename=os.path.basename(filename))
        message.attach(att)
    # Send message
    with smtplib.SMTP(server_name, server_port) as server:
        server.starttls()
        server.login(sender, password)
        server.sendmail(sender, to.split(',') + cc.split(','), message.as_string())


def get_widgets_content(widgets):
    """Get content of edit widgets

    Args:
        widgets (dict): edit widgets

    Returns:
        dict: content
    """
    return {name: [edit.toPlainText() for edit in edits] for name, edits in widgets.items()}


def set_widgets_content(widgets, content):
    """Set content of edit widgets

    Args:
        widgets (dict): edit widgets
        content (dict): content
    """
    for name in content.keys():
        for edit, text in zip(widgets[name], content[name]):
            edit.setPlainText(text)


def render_docx(template, output, context):
    """Render docx template with context and save to file.

    Args:
        template (str): template filename
        output (str): output filename
        context (dict): render context
    """
    template = DocxTemplate(template)
    template.render(context, autoescape=True)
    template.save(output)


def render_daily(template, output, content):
    """Render and save daily report

    Args:
        template (str): template filename
        output (str): output filename
        content (dict): content of daily report
    """
    context = {}
    for name, texts in content.items():
        context[name] = [R(text) for text in texts]
    render_docx(template, output, context)


def render_xlsx(template, output, context):
    """Render xlsx template with context and save to file.

    Args:
        template (str): template filename
        output (str): output filename
        context (dict): render context
    """
    wb = openpyxl.load_workbook(template)
    sheet = wb.active
    for row in sheet.rows:
        for cell in row:
            value = cell.value
            if value is not None and type(value) == str and '{{' in value:
                cell.value = Template(cell.value).render(context)
    wb.save(output)


def render_weekly(template, output, content, settings):
    """Render and save weekly report

    Args:
        template (str): template filename
        output (str): output filename
        content (dict): content of weekly report
        settings (dict): settings
    """
    today = datetime.date.today()
    oneday = datetime.timedelta(days=1)
    next_friday = today + oneday
    while next_friday.weekday() != calendar.FRIDAY:
        next_friday += oneday

    context = {
        'today': today.strftime("%Y-%m-%d"),
        'next_friday': next_friday.strftime("%Y-%m-%d"),
        'name': settings['name'],
        'date': today.strftime("%Y/%m/%d"),
    }
    context.update(content)
    render_xlsx(template, output, context)


class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # Set title
        self.setWindowTitle('Settings')
        # Set layout
        grid = QGridLayout()
        self.setLayout(grid)

        # Create label and edit for each setting
        self.edits = {}
        for i, (name, text) in enumerate(settings_widgets_config):
            label = QLabel(text)
            grid.addWidget(label, i, 0)
            self.edits[name] = QLineEdit()
            self.edits[name].setMinimumWidth(300)
            grid.addWidget(self.edits[name], i, 1)
        # Show mask in password edit
        self.edits['password'].setEchoMode(QLineEdit.Password)

    def set_content(self, content):
        """Set content of edit widgets

        Args:
            content (dict): settings content
        """
        for name, text in content.items():
            self.edits[name].setText(text)

    def get_content(self):
        """Get content of edit widgets

        Returns:
            dict: settings content
        """
        return {name: edit.text() for name, edit in self.edits.items()}


class WeeklyDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # Set title
        self.setWindowTitle('Weekly')
        # Set layout
        grid = QGridLayout()
        self.setLayout(grid)

        # Create label and edits for each column
        self.widgets = {}
        for i, (name, text) in enumerate(weekly_widgets_config):
            label = QLabel(text)
            label.setAlignment(QtCore.Qt.AlignCenter)
            grid.addWidget(label, 0, i)
            self.widgets[name] = []
            for j in range(1, 6):
                edit = QTextEdit()
                edit.setMinimumSize(50, 50)
                grid.addWidget(edit, j, i)
                self.widgets[name].append(edit)

        # Create buttons
        self.clear_button = QPushButton('Clear')
        self.clear_button.clicked.connect(self.clear_button_clicked)
        grid.addWidget(self.clear_button, 6, 0)
        self.send_weekly_button = QPushButton('Send weekly')
        self.send_weekly_button.clicked.connect(self.send_weekly_button_clicked)
        grid.addWidget(self.send_weekly_button, 6, 2)

    def clear_button_clicked(self):
        """Clear all edits
        """
        for edits in self.widgets.values():
            for edit in edits:
                edit.setPlainText('')

    def send_weekly_button_clicked(self):
        """Send weekly report
        """
        # Get weekly report content
        self.content['weekly'] = get_widgets_content(self.widgets)

        name = self.content['settings']['name']
        date = datetime.date.today().strftime("%Y%m%d")
        # Render daily report and save
        daily_filename = daily_filename_format.format(name=name, date=date)
        daily_filename = os.path.join(archive_folder, daily_filename)
        render_daily(daily_template, daily_filename, self.content['daily'])
        # Render weekly report and save
        weekly_filename = weekly_filename_format.format(name=name, date=date)
        weekly_filename = os.path.join(archive_folder, weekly_filename)
        render_weekly(weekly_template, weekly_filename, self.content['weekly'], self.content['settings'])

        # Create email subject
        subject = weekly_subject_format.format(name=name, date=date)
        # Send email
        settings = self.content['settings']
        try:
            send_email(settings['server_name'],
                       settings['server_port'],
                       settings['sender'],
                       settings['password'],
                       settings['to'],
                       settings['cc'],
                       subject,
                       [daily_filename, weekly_filename])
            QMessageBox.about(self, "Message", "Succeeded")
        except:
            QMessageBox.about(self, "Message", "Failed")


class DailyDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.init_ui()

        # Load content from file
        self.load_content()
        # Open setting dialog for the first time
        if not self.content['settings']:
            self.settings_button_clicked()

    def init_ui(self):
        # Set title
        self.setWindowTitle('Daily')
        # Set layout
        grid = QGridLayout()
        self.setLayout(grid)

        # Create settings and weekly dialog
        self.settings_dialog = SettingsDialog()
        self.weekly_dialog = WeeklyDialog()

        # Create weekday labels
        for i, text in enumerate(daily_label_texts):
            label = QLabel(text)
            label.setAlignment(QtCore.Qt.AlignRight)
            grid.addWidget(label, i+1, 0)

        # Create label and edits for each column
        self.widgets = {}
        for i, (name, text) in enumerate(daily_widgets_config):
            label = QLabel(text)
            label.setAlignment(QtCore.Qt.AlignCenter)
            grid.addWidget(label, 0, i*2+1, 1, 2)
            self.widgets[name] = []
            for j in range(1, 6):
                edit = QTextEdit()
                edit.setMinimumSize(50, 50)
                grid.addWidget(edit, j, i*2+1, 1, 2)
                self.widgets[name].append(edit)

        # Create buttons
        for i, (name, text) in enumerate(daily_buttons_config):
            button = QPushButton(text)
            button.clicked.connect(getattr(self, '%s_button_clicked' % name))
            grid.addWidget(button, 6, i+1)

    def load_content(self):
        """Load content file
        """
        # if content file exists, load and display
        if os.path.exists(content_filename):
            with open(content_filename) as f:
                self.content = json.load(f)
            set_widgets_content(self.widgets, self.content['daily'])
        # else use default content
        else:
            self.content = default_content

    def save_content(self):
        """Save content file
        """
        self.content['daily'] = get_widgets_content(self.widgets)
        with open(content_filename, 'w') as f:
            json.dump(self.content, f, indent=4)

    def settings_button_clicked(self):
        """Show settings dialog
        """
        self.settings_dialog.set_content(self.content['settings'])
        self.settings_dialog.exec_()
        self.content['settings'] = self.settings_dialog.get_content()

    def edit_weekly_button_clicked(self):
        """Show weekly dialog
        """
        # Save daily report content
        self.content['daily'] = get_widgets_content(self.widgets)
        # Pass content to weekly dialog
        self.weekly_dialog.content = self.content

        set_widgets_content(self.weekly_dialog.widgets, self.content['weekly'])
        self.weekly_dialog.exec_()
        self.content['weekly'] = get_widgets_content(self.weekly_dialog.widgets)

    def clear_button_clicked(self):
        """Clear all edits
        """
        for edits in self.widgets.values():
            for edit in edits:
                edit.setPlainText('')

    def send_daily_button_clicked(self):
        """Send daily report
        """
        # Get daily report content
        self.content['daily'] = get_widgets_content(self.widgets)

        name = self.content['settings']['name']
        date = datetime.date.today().strftime("%Y%m%d")
        # Render daily report and save
        daily_filename = daily_filename_format.format(name=name, date=date)
        daily_filename = os.path.join(archive_folder, daily_filename)
        render_daily(daily_template, daily_filename, self.content['daily'])

        # Create email subject
        subject = daily_subject_format.format(name=name, date=date)
        # Send email
        settings = self.content['settings']
        try:
            send_email(settings['server_name'],
                       settings['server_port'],
                       settings['sender'],
                       settings['password'],
                       settings['to'],
                       settings['cc'],
                       subject,
                       [daily_filename])
            QMessageBox.about(self, "Message", "Succeeded")
        except:
            QMessageBox.about(self, "Message", "Failed")

    def closeEvent(self, event):
        """Save content file when dialog closed

        Args:
            event (QCloseEvent): close event
        """
        self.save_content()
        super().closeEvent(event)

    def reject(self):
        """Save content file when dialog rejected
        """
        self.save_content()
        super().reject()


if __name__ == '__main__':
    # Create archive folder if not exists
    if not os.path.exists(archive_folder):
        os.mkdir(archive_folder)

    # Start app and show daily dialog
    app = QApplication(sys.argv)
    daily_dialog = DailyDialog()
    daily_dialog.show()
    sys.exit(app.exec_())
