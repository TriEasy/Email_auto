# Save this file as: outlookHelp.py
import datetime
import pytz
from exchangelib import EWSDateTime, EWSTimeZone

def get_riyadh_datetime():
    """Returns the current datetime in Riyadh (UTC+3) using EWSTimeZone."""
    riyadh_tz = EWSTimeZone.timezone('Asia/Riyadh')
    return EWSDateTime.now(tz=riyadh_tz)

def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if the due date is within the next 2 days.
    Properly handles EWSDateTime and EWSTimeZone objects.
    """
    try:
        # Convert both to EWSDateTime if needed
        if isinstance(due_date_obj, EWSDateTime):
            due_date = due_date_obj
        elif hasattr(due_date_obj, 'astimezone'):
            # Convert regular datetime to EWSDateTime
            riyadh_tz = EWSTimeZone.timezone('Asia/Riyadh')
            due_date = EWSDateTime.from_datetime(due_date_obj.astimezone(pytz.timezone('Asia/Riyadh')))
        else:
            due_date = due_date_obj

        # Get current time in EWSDateTime format
        if isinstance(now_in_riyadh, EWSDateTime):
            now = now_in_riyadh
        else:
            riyadh_tz = EWSTimeZone.timezone('Asia/Riyadh')
            now = EWSDateTime.now(tz=riyadh_tz)

        # Calculate two days from now
        two_days_from_now = now + datetime.timedelta(days=2)

        # Compare dates
        is_due = now <= due_date <= two_days_from_now
        
        print(f"  🕐 Now: {now}")
        print(f"  🎯 Due: {due_date}")
        print(f"  ⏰ Two days from now: {two_days_from_now}")
        
        return is_due
        
    except Exception as e:
        print(f"  ⚠️ Error comparing dates: {e}")
        # Fallback to naive datetime comparison
        try:
            # Strip timezone info and compare as naive datetimes
            if hasattr(due_date_obj, 'replace'):
                due_date_naive = due_date_obj.replace(tzinfo=None)
            else:
                due_date_naive = due_date_obj

            if hasattr(now_in_riyadh, 'replace'):
                now_naive = now_in_riyadh.replace(tzinfo=None)
            else:
                now_naive = datetime.datetime.now()

            two_days_from_now = now_naive + datetime.timedelta(days=2)

            return now_naive <= due_date_naive <= two_days_from_now
        except Exception as e2:
            print(f"  ❌ Error in fallback comparison: {e2}")
            return False

def format_due_date_for_email(due_date_obj, riyadh_tz=None):
    """Formats the due date into a clean string for the email body."""
    try:
        # Handle EWSDateTime objects
        if isinstance(due_date_obj, EWSDateTime):
            # Convert to Riyadh timezone
            riyadh_tz = EWSTimeZone.timezone('Asia/Riyadh')
            due_date_riyadh = due_date_obj.astimezone(riyadh_tz)
            return due_date_riyadh.strftime('%Y-%m-%d %H:%M')
        
        # Handle regular datetime objects
        if hasattr(due_date_obj, 'astimezone'):
            if riyadh_tz is None:
                riyadh_tz = pytz.timezone('Asia/Riyadh')
            due_date_riyadh = due_date_obj.astimezone(riyadh_tz)
            return due_date_riyadh.strftime('%Y-%m-%d %H:%M')
        
        # Fallback - just format as is
        return due_date_obj.strftime('%Y-%m-%d %H:%M')
    except Exception as e:
        print(f"  ⚠️ Error formatting date: {e}")
        return str(due_date_obj)

def get_reminder_subject(original_subject):
    """Returns the formatted subject for the reminder email."""
    return f"🔔 تذكير بالمتابعة: {original_subject}"

def get_reminder_body(original_subject, due_date_str):
    """Returns the formatted body for the reminder email."""
    return (
        f"السلام عليكم ورحمة الله وبركاته،\n\n"
        f"نود تذكيركم بأن الرسالة التالية بلغت موعدها المحدد للمتابعة:\n\n"
        f"📩 العنوان: {original_subject}\n"
        f"📅 الموعد: {due_date_str}\n\n"
        f"يرجى اتخاذ اللازم.\n\n"
        f"قسم المتابعة - هيئة الغذاء والدواء"
    )

def add_sent_category(existing_categories, sent_category):
    """
    Safely adds the sent category to existing categories.
    """
    if not existing_categories:
        return [sent_category]

    if isinstance(existing_categories, str):
        category_list = [cat.strip() for cat in existing_categories.split(',')]
    else:
        category_list = list(existing_categories)

    if sent_category not in category_list:
        category_list.append(sent_category)

    return category_list