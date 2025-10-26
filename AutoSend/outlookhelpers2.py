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
        due_date_riyadh = due_date_obj.astimezone(now_in_riyadh.tzinfo)
        two_days_from_now = now_in_riyadh + datetime.timedelta(days=2)
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

def add_sent_category(existing_categories, sent_category):
    """
    Safely adds the sent category to existing categories.
    
    Args:
        existing_categories: List, tuple, or string of existing categories
        sent_category: The category string to add
    
    Returns:
        List of categories
    """
    if not existing_categories:
        return [sent_category]
    
    # Convert to list if it's a string or tuple
    if isinstance(existing_categories, str):
        category_list = [cat.strip() for cat in existing_categories.split(',')]
    else:
        category_list = list(existing_categories)
    
    # Add the sent category if not already present
    if sent_category not in category_list:
        category_list.append(sent_category)
    
    return category_list