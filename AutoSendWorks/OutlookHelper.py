# Save this file as: outlookHelp.py
import datetime
import pytz
from exchangelib import EWSDateTime, EWSTimeZone

def get_riyadh_datetime():
    """Returns the current datetime in Riyadh (UTC+3)."""
    tz = EWSTimeZone.timezone("Asia/Riyadh")
    return EWSDateTime.now(tz)

def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if the due date is within the next 2 days, using Riyadh timezone.
    """
    try:
        tz_riyadh = EWSTimeZone.timezone("Asia/Riyadh")

        # Convert to Riyadh time if possible
        if isinstance(due_date_obj, EWSDateTime):
            due_date_local = due_date_obj.astimezone(tz_riyadh)
        else:
            due_date_local = EWSDateTime.from_datetime(
                due_date_obj.astimezone(datetime.timezone.utc)
            ).astimezone(tz_riyadh)

        now_local = now_in_riyadh

        two_days_from_now = now_local + datetime.timedelta(days=2)
        is_due = now_local <= due_date_local <= two_days_from_now

        print(f"  🕐 Now (Riyadh): {now_local}")
        print(f"  🎯 Due (Riyadh): {due_date_local}")
        print(f"  ⏰ Two days from now: {two_days_from_now}")
        print(f"  ✓ Is due: {is_due}")

        return is_due

    except Exception as e:
        print(f"  ⚠️ Error comparing dates: {e}")
        import traceback
        traceback.print_exc()
        return False


def format_due_date_for_email(due_date_obj):
    """Formats the due date into Riyadh time for the reminder email."""
    try:
        tz_riyadh = EWSTimeZone.timezone("Asia/Riyadh")

        if isinstance(due_date_obj, EWSDateTime):
            due_date_riyadh = due_date_obj.astimezone(tz_riyadh)
        else:
            # Handle naive datetime or pytz
            if due_date_obj.tzinfo is None:
                due_date_obj = pytz.UTC.localize(due_date_obj)
            due_date_riyadh = EWSDateTime.from_datetime(due_date_obj).astimezone(tz_riyadh)

        return due_date_riyadh.strftime('%Y-%m-%d %H:%M')

    except Exception as e:
        print(f"  ⚠️ Error formatting date: {e}")
        import traceback
        traceback.print_exc()
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
        f"📅 الموعد (بتوقيت الرياض): {due_date_str}\n\n"
        f"يرجى اتخاذ اللازم.\n\n"
        f"قسم المتابعة - هيئة الغذاء والدواء"
    )


def add_sent_category(existing_categories, sent_category):
    """Safely adds the sent category to existing categories."""
    if not existing_categories:
        return [sent_category]

    if isinstance(existing_categories, str):
        category_list = [cat.strip() for cat in existing_categories.split(',')]
    else:
        category_list = list(existing_categories)

    if sent_category not in category_list:
        category_list.append(sent_category)

    return category_list