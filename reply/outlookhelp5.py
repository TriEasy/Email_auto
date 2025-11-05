# Save this file as: outlookHelp.py
import datetime
import pytz
from exchangelib import EWSDateTime, EWSTimeZone

def get_riyadh_datetime():
    """Returns the current datetime in Riyadh (UTC+3)."""
    # Use pytz to get current time in Riyadh
    riyadh_tz = pytz.timezone('Asia/Riyadh')
    return datetime.datetime.now(riyadh_tz)

def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if the due date is within the next 2 days.
    Properly handles EWSDateTime objects from Exchange.
    """
    try:
        # Work with the objects as-is, converting to naive datetimes for comparison
        # This avoids timezone compatibility issues between pytz and EWSTimeZone
        
        # Convert due_date to naive datetime
        if hasattr(due_date_obj, 'replace'):
            due_date_naive = due_date_obj.replace(tzinfo=None)
        else:
            due_date_naive = due_date_obj
        
        # Convert now to naive datetime
        if hasattr(now_in_riyadh, 'replace'):
            now_naive = now_in_riyadh.replace(tzinfo=None)
        else:
            now_naive = datetime.datetime.now()
        
        # Calculate two days from now
        two_days_from_now = now_naive + datetime.timedelta(days=2)
        
        # Compare dates
        is_due = now_naive <= due_date_naive <= two_days_from_now
        
        print(f"  🕐 Now (naive): {now_naive}")
        print(f"  🎯 Due (naive): {due_date_naive}")
        print(f"  ⏰ Two days from now: {two_days_from_now}")
        print(f"  ✓ Is due: {is_due}")
        
        return is_due
        
    except Exception as e:
        print(f"  ⚠️ Error comparing dates: {e}")
        import traceback
        traceback.print_exc()
        return False

def format_due_date_for_email(due_date_obj, riyadh_tz=None):
    """Formats the due date into a clean string for the email body."""
    try:
        # 1. Ensure we have the Riyadh timezone
        if riyadh_tz is None:
            riyadh_tz = pytz.timezone('Asia/Riyadh')

        # 2. Handle EWSDateTime/datetime objects
        if hasattr(due_date_obj, 'astimezone'):
            
            # 3. Check if it's timezone-aware
            if hasattr(due_date_obj, 'tzinfo') and due_date_obj.tzinfo:
                # It's aware, so we can safely convert it
                due_date_riyadh = due_date_obj.astimezone(riyadh_tz)
                return due_date_riyadh.strftime('%Y-%m-%d %H:%M')
            else:
                # It's naive (no timezone). Assume it's already local time.
                return due_date_obj.strftime('%Y-%m-%d %H:%M')

        # Fallback for other types
        if hasattr(due_date_obj, 'strftime'):
            return due_date_obj.strftime('%Y-%m-%d %H:%M')
        
        return str(due_date_obj)
        
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