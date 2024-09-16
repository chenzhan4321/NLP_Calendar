from pydantic import BaseModel
from openai import OpenAI
import os
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Optional, List
from icalendar import Calendar, Event
import subprocess

# Set your API key
api_key = ""

client = OpenAI(api_key=api_key)

class CalendarEvent(BaseModel):
    name: str
    date: str
    start_time: Optional[str] = None
    end_time: Optional[str] = None
    timezone: Optional[str] = None
    participants: Optional[List[str]] = None
    location: Optional[str] = None  # 添加 location 字段

def get_current_date():
    return datetime.now().strftime("%Y-%m-%d")

def parse_event(event_description):
    current_date = get_current_date()
    completion = client.beta.chat.completions.parse(
        model="gpt-4o-2024-08-06",
        messages=[
            {"role": "system", "content": f"Extract the event information from the given description. Use today's date ({current_date}) as a reference for relative dates like 'next Tuesday'. Note that the speaker often sleep late and when he says 'tomorrow' before the morning, e.g. at 6 a.m., it means the same calendar day. If start_time, end_time, timezone, or location is not specified, leave them as null. For Chinese input, ensure you correctly interpret dates and times. Return the date in YYYY-MM-DD format and start_time and end_time in HH:MM format if available. If only start_time is provided, assume the event lasts for 1 hour. Extract the location if provided."},
            {"role": "user", "content": event_description},
        ],
        response_format=CalendarEvent,
    )
    return completion.choices[0].message.parsed

def create_ics_file(event):
    cal = Calendar()
    cal.add('prodid', '-//My Calendar Event//example.com//')
    cal.add('version', '2.0')

    ical_event = Event()
    ical_event.add('summary', event.name)

    event_date = datetime.strptime(event.date, "%Y-%m-%d").date()

    if event.start_time:
        start_time = datetime.strptime(event.start_time, "%H:%M").time()
        start_datetime = datetime.combine(event_date, start_time)
        
        if event.end_time:
            end_time = datetime.strptime(event.end_time, "%H:%M").time()
            end_datetime = datetime.combine(event_date, end_time)
        else:
            end_datetime = start_datetime + timedelta(hours=1)
        
        if event.timezone:
            start_datetime = start_datetime.replace(tzinfo=ZoneInfo(event.timezone))
            end_datetime = end_datetime.replace(tzinfo=ZoneInfo(event.timezone))
        
        ical_event.add('dtstart', start_datetime)
        ical_event.add('dtend', end_datetime)
    else:
        # All-day event
        ical_event.add('dtstart', event_date)
        ical_event.add('dtend', event_date + timedelta(days=1))

    if event.participants:
        ical_event['attendee'] = event.participants

    # 添加地点信息
    if event.location:
        ical_event.add('location', event.location)

    cal.add_component(ical_event)

    # 使用 os.path.join 来确保文件路径正确
    filename = os.path.join(os.path.dirname(__file__), f"{event.name.replace(' ', '_')}_{event.date}.ics")
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    
    return filename

def import_ics_to_calendar(ics_filename):
    try:
        full_path = os.path.abspath(ics_filename)
        print(f"尝试打开 ICS 文件: {full_path}")
        result = subprocess.run(["open", full_path], capture_output=True, text=True, check=True)
        print(f"打开命令输出: {result.stdout}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"打开 ICS 文件时出错: {e}")
        print(f"错误输出: {e.stderr}")
        return False

def main():
    while True:
        event_description = input("请输入事件描述 (或输入 'quit' 退出): ")
        if event_description.lower() == 'quit':
            break

        parsed_event = parse_event(event_description)
        print("解析的事件:", parsed_event)

        ics_filename = create_ics_file(parsed_event)
        print(f"ICS文件 '{ics_filename}' 已创建。")

        if import_ics_to_calendar(ics_filename):
            print(f"已尝试打开ICS文件。请检查Calendar应用是否弹出导入对话框。")
        else:
            print("打开ICS文件时出现错误。您可以手动导入生成的ICS文件。")

if __name__ == "__main__":
    main()