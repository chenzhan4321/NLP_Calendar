import sys
import os
from typing import Optional, List
from datetime import datetime, timedelta
import tempfile
import json

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QTextEdit, QPushButton, QMessageBox, QSplitter
)
from PyQt5.QtCore import Qt
from openai import OpenAI
from icalendar import Calendar, Event
from pydantic import BaseModel

import subprocess
import time

class CalendarEvent(BaseModel):
    name: str
    start_date: str
    end_date: Optional[str] = None
    start_time: Optional[str] = None
    end_time: Optional[str] = None
    timezone: Optional[str] = None
    participants: Optional[List[str]] = None
    location: Optional[str] = None
    content: Optional[str] = None

    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat(),
            timedelta: lambda v: str(v)
        }

class CalendarApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # API Key input
        api_layout = QHBoxLayout()
        api_layout.addWidget(QLabel('API Key:'))
        self.api_input = QLineEdit()
        api_layout.addWidget(self.api_input)
        layout.addLayout(api_layout)

        # Create a splitter for input and output
        splitter = QSplitter(Qt.Vertical)

        # Event description input
        input_widget = QWidget()
        input_layout = QVBoxLayout(input_widget)
        input_layout.addWidget(QLabel('Event Description(s):'))
        self.event_input = QTextEdit()
        input_layout.addWidget(self.event_input)
        splitter.addWidget(input_widget)

        # Structured data output
        output_widget = QWidget()
        output_layout = QVBoxLayout(output_widget)
        output_layout.addWidget(QLabel('Structured Data:'))
        self.structured_output = QTextEdit()
        self.structured_output.setReadOnly(True)
        output_layout.addWidget(self.structured_output)
        splitter.addWidget(output_widget)

        layout.addWidget(splitter)

        # Generate button
        self.generate_button = QPushButton('Generate ICS')
        self.generate_button.clicked.connect(self.generate_ics)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)
        self.setWindowTitle('Calendar Event Generator')
        self.setGeometry(300, 300, 600, 500)

    def generate_ics(self):
        api_key = self.api_input.text()
        event_descriptions = self.event_input.toPlainText().split('\n')

        if not api_key or not event_descriptions:
            QMessageBox.warning(self, 'Input Error', 'Please enter both API key and at least one event description.')
            return

        client = OpenAI(api_key=api_key)
        all_parsed_events = []

        for event_description in event_descriptions:
            if event_description.strip():  # Skip empty lines
                try:
                    parsed_events = self.parse_event(client, event_description)
                    all_parsed_events.extend(parsed_events)
                    for parsed_event in parsed_events:
                        ics_filename = self.create_ics_file(parsed_event)
                        self.import_ics_to_calendar(ics_filename)
                    print(f'ICS file(s) created and opened for: {event_description}')
                except ValueError as ve:
                    QMessageBox.critical(self, 'Parsing Error', f'Error parsing event: {str(ve)}')
                except IOError as ie:
                    QMessageBox.critical(self, 'File Error', f'Error creating or importing ICS file: {str(ie)}')
                except Exception as e:
                    QMessageBox.critical(self, 'Unexpected Error', f'An unexpected error occurred: {str(e)}')
                    print(f"Unexpected error details: {type(e).__name__}, {str(e)}")

        # Display structured data in the output text edit
        self.structured_output.setPlainText(json.dumps([event.dict() for event in all_parsed_events], indent=2, ensure_ascii=False))

        self.event_input.clear()  # Clear the input field after processing all events

    def parse_event(self, client: OpenAI, event_description: str) -> List[CalendarEvent]:
        current_date = datetime.now().strftime("%Y-%m-%d")
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": f"""Extract the event information from the given description and return it as a JSON array of event objects. Use today's date ({current_date}) as a reference for relative dates like 'next Tuesday'. Note that the speaker often sleeps late and when they say 'tomorrow' before the morning, e.g. at 6 a.m., it means the same calendar day. If start_time, end_time, timezone, or location is not specified, leave them as null. For Chinese input, ensure you correctly interpret dates and times. Return the start_date and end_date in YYYY-MM-DD format and start_time and end_time in HH:MM format if available. If only start_time is provided, assume the event lasts for 1 hour. Extract the location if provided. For multiple dates mentioned (e.g., "31/10, 8/11"), create separate event objects for each date. Preserve as much original information as possible in the 'content' field. The JSON should be an array of objects with the following structure: [{{"name": string, "start_date": string, "end_date": string|null, "start_time": string|null, "end_time": string|null, "timezone": string|null, "participants": array|null, "location": string|null, "content": string}}]"""},
                {"role": "user", "content": event_description},
            ],
            response_format={"type": "json_object"}
        )
        events_data = completion.choices[0].message.content
        try:
            parsed_data = json.loads(events_data)
            if isinstance(parsed_data, list):
                events = parsed_data
            elif isinstance(parsed_data, dict):
                events = parsed_data.get('events', [parsed_data])
            else:
                raise ValueError(f"Unexpected data format: {parsed_data}")

            calendar_events = []
            for event in events:
                # Ensure required fields are present
                if 'name' not in event or not event['name']:
                    event['name'] = "Untitled Event"
                if 'start_date' not in event or not event['start_date']:
                    event['start_date'] = current_date

                try:
                    calendar_event = CalendarEvent.parse_obj(event)
                    calendar_events.append(calendar_event)
                except Exception as e:
                    print(f"Error parsing event: {e}")
                    print(f"Event data: {event}")
                    # Skip this event and continue with the next one

            return calendar_events
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON response: {events_data}")

    def create_ics_file(self, event: CalendarEvent) -> str:
        cal = Calendar()
        cal.add('prodid', '-//Calendar Event Generator//mxm.dk//')
        cal.add('version', '2.0')

        ics_event = Event()
        ics_event.add('summary', event.name)

        start_date = datetime.strptime(event.start_date, "%Y-%m-%d").date()
        if event.end_date:
            end_date = datetime.strptime(event.end_date, "%Y-%m-%d").date()
        else:
            end_date = start_date

        if event.start_time:
            start_datetime = datetime.strptime(f"{event.start_date} {event.start_time}", "%Y-%m-%d %H:%M")
            ics_event.add('dtstart', start_datetime)

            if event.end_time:
                end_datetime = datetime.strptime(f"{event.end_date or event.start_date} {event.end_time}", "%Y-%m-%d %H:%M")
            else:
                end_datetime = start_datetime + timedelta(hours=1)
            ics_event.add('dtend', end_datetime)
        else:
            ics_event.add('dtstart', start_date)
            ics_event.add('dtend', end_date + timedelta(days=1))  # Add one day to make it inclusive

        if event.location:
            ics_event.add('location', event.location)

        if event.participants:
            for participant in event.participants:
                ics_event.add('attendee', f'MAILTO:{participant}')

        cal.add_component(ics_event)

        # Use a temporary directory to save the ICS file
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.ics', delete=False) as temp_file:
            temp_file.write(cal.to_ical())
            filename = temp_file.name

        return filename

    def import_ics_to_calendar(self, ics_filename: str):
        try:
            if sys.platform == "darwin":  # macOS
                subprocess.run(["open", ics_filename], check=True)
                # Wait a short time to allow the Calendar app to process the file
                time.sleep(2)
                # Try to delete the temporary file
                try:
                    os.remove(ics_filename)
                except OSError:
                    print(f"Unable to delete temporary file: {ics_filename}")
            elif sys.platform == "win32":  # Windows
                os.startfile(ics_filename)
            else:  # Linux and other systems
                subprocess.run(["xdg-open", ics_filename], check=True)
        except Exception as e:
            raise IOError(f"Failed to open the ICS file: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CalendarApp()
    ex.show()
    sys.exit(app.exec_())