# Save this file as: OutlookHelpers.py
import datetime

def get_riyadh_datetime():
    """Returns the current datetime in Riyadh (UTC+3)."""
    riyadh_tz = datetime.timezone(datetime.timedelta(hours=3))
    return datetime.datetime.now(riyadh_tz)

def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if the Outlook due date is within the next 2 days.
    """
    try:
        # Convert the Outlook due date object to Riyadh timezone
        due_date_riyadh = due_date_obj.astimezone(now_in_riyadh.tzinfo)
        
        # Define the 2-day window
        two_days_from_now = now_in_riyadh + datetime.timedelta(days=2)
        
        # Check if due date is between now and 2 days from now
        return now_in_riyadh <= due_date_riyadh <= two_days_from_now
    except Exception as e:
        print(f"Error comparing dates: {e}")
        return False

def format_due_date_for_email(due_date_obj, riyadh_tz):
    """Formats the due date into a clean string for the email body."""
    due_date_riyadh = due_date_obj.astimezone(riyadh_tz)
    return due_date_riyadh.strftime('%Y-%m-%d %H:%M')

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

def add_sent_category(existing_categories):
    """Safely adds the 'AutoReminderSent' category."""
    if not existing_categories:
        return "AutoReminderSent"
    if "AutoReminderSent" in existing_categories:
        return existing_categories
    
    return (existing_categories + ", AutoReminderSent").strip(", ")