# Save this file as: OutlookHelpers.py
import datetime
import pytz

def get_riyadh_datetime():
    """Returns the current datetime in Riyadh (UTC+3) using pytz."""
    riyadh_tz = pytz.timezone('Asia/Riyadh')
    return datetime.datetime.now(riyadh_tz)

def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if the due date is within the next 2 days.
    Handles both EWSTimeZone and standard Python timezone objects.
    """
    try:
        # Convert both to UTC for comparison to avoid timezone issues
        # Get the UTC offset from now_in_riyadh
        if hasattr(due_date_obj, 'astimezone'):
            # Convert due date to UTC
            due_date_utc = due_date_obj.astimezone(pytz.UTC)
        else:
            # If it's already naive or can't convert, assume UTC
            due_date_utc = due_date_obj
        
        # Convert current time to UTC
        now_utc = now_in_riyadh.astimezone(pytz.UTC)
        
        # Define the 2-day window
        two_days_from_now = now_utc + datetime.timedelta(days=2)
        
        # Check if due date is between now and 2 days from now
        result = now_utc <= due_date_utc <= two_days_from_now
        
        return result
    except Exception as e:
        print(f"Error comparing dates: {e}")
        # Try a simpler comparison without timezone conversion
        try:
            # Remove timezone info and compare
            if hasattr(due_date_obj, 'replace'):
                due_date_naive = due_date_obj.replace(tzinfo=None)
            else:
                due_date_naive = due_date_obj
            
            now_naive = now_in_riyadh.replace(tzinfo=None)
            two_days_from_now = now_naive + datetime.timedelta(days=2)
            
            return now_naive <= due_date_naive <= two_days_from_now
        except Exception as e2:
            print(f"Error in fallback comparison: {e2}")
            return False

def format_due_date_for_email(due_date_obj, riyadh_tz):
    """Formats the due date into a clean string for the email body."""
    try:
        # Try to convert to Riyadh timezone
        if hasattr(due_date_obj, 'astimezone'):
            due_date_riyadh = due_date_obj.astimezone(riyadh_tz)
        else:
            due_date_riyadh = due_date_obj
        return due_date_riyadh.strftime('%Y-%m-%d %H:%M')
    except:
        # Fallback: just format as-is
        return due_date_obj.strftime('%Y-%m-%d %H:%M')

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