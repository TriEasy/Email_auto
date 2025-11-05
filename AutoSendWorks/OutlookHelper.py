# outlookHelp.py
import datetime
import pytz
from exchangelib import EWSDateTime, EWSTimeZone

def _get_ews_riyadh_tz():
    """
    Return an EWSTimeZone instance for Asia/Riyadh in a way that works
    across exchangelib versions.
    """
    # Try common constructor first (works when EWSTimeZone accepts a string key)
    try:
        return EWSTimeZone('Asia/Riyadh')
    except Exception:
        pass

    # Try the helper 'from_timezone' if available and pytz can provide zone
    try:
        if hasattr(EWSTimeZone, 'from_timezone'):
            return EWSTimeZone.from_timezone(pytz.timezone('Asia/Riyadh'))
    except Exception:
        pass

    # Last fallback: attempt to use timezone() if present (older docs show this)
    try:
        return EWSTimeZone.timezone('Asia/Riyadh')
    except Exception:
        pass

    # As a safe fallback, return UTC EWSTimeZone (not ideal, but prevents crashes)
    try:
        return EWSTimeZone('UTC')
    except Exception:
        # If even that fails, raise so user sees the real problem
        raise RuntimeError("Could not construct an EWSTimeZone for Asia/Riyadh")


def get_riyadh_datetime():
    """
    Returns current datetime in Riyadh as an EWSDateTime (tz-aware with EWSTimeZone).
    """
    tz = _get_ews_riyadh_tz()
    # EWSDateTime.now accepts tz keyword in most versions
    try:
        return EWSDateTime.now(tz)
    except Exception:
        # fallback: create from Python datetime localized with pytz and convert
        dt = datetime.datetime.now(pytz.timezone('Asia/Riyadh'))
        return EWSDateTime.from_datetime(dt)


def is_due_soon(due_date_obj, now_in_riyadh):
    """
    Checks if due_date_obj is within the next 2 days (Riyadh time).
    due_date_obj may be an EWSDateTime or a regular datetime.
    now_in_riyadh is expected to be an EWSDateTime (as returned by get_riyadh_datetime).
    """
    try:
        tz_riyadh = _get_ews_riyadh_tz()

        # Normalize now_in_riyadh to EWSDateTime if necessary
        if not isinstance(now_in_riyadh, EWSDateTime):
            # try to convert Python datetime to EWSDateTime (keeps tzinfo)
            if isinstance(now_in_riyadh, datetime.datetime):
                now_in_riyadh = EWSDateTime.from_datetime(now_in_riyadh)
            else:
                now_in_riyadh = EWSDateTime.now(tz_riyadh)

        # Convert due_date_obj to EWSDateTime in Riyadh tz
        if isinstance(due_date_obj, EWSDateTime):
            due_riyadh = due_date_obj.astimezone(tz_riyadh)
        elif isinstance(due_date_obj, datetime.datetime):
            # Make sure it's timezone-aware
            if due_date_obj.tzinfo is None:
                due_date_obj = pytz.UTC.localize(due_date_obj)
            due_ews = EWSDateTime.from_datetime(due_date_obj)
            due_riyadh = due_ews.astimezone(tz_riyadh)
        else:
            # As a last resort, try string conversion
            due_ews = EWSDateTime.from_datetime(datetime.datetime.fromisoformat(str(due_date_obj)))
            due_riyadh = due_ews.astimezone(tz_riyadh)

        # Ensure now is in Riyadh tz as EWSDateTime
        now_riyadh = now_in_riyadh.astimezone(tz_riyadh)

        two_days_from_now = now_riyadh + datetime.timedelta(days=2)

        is_due = now_riyadh <= due_riyadh <= two_days_from_now

        print(f"  🕐 Now (Riyadh): {now_riyadh.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"  🎯 Due (Riyadh): {due_riyadh.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"  ⏰ Two days from now (Riyadh): {two_days_from_now.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"  ✓ Is due: {is_due}")

        return is_due

    except Exception as e:
        print(f"  ⚠️ Error comparing dates: {e}")
        import traceback
        traceback.print_exc()
        return False


def format_due_date_for_email(due_date_obj):
    """
    Format the due date into a readable Riyadh-time string for email body.
    Example output: '2025-11-05 14:00'
    """
    try:
        tz_riyadh = _get_ews_riyadh_tz()

        if isinstance(due_date_obj, EWSDateTime):
            due_riyadh = due_date_obj.astimezone(tz_riyadh)
        elif isinstance(due_date_obj, datetime.datetime):
            if due_date_obj.tzinfo is None:
                due_date_obj = pytz.UTC.localize(due_date_obj)
            due_ews = EWSDateTime.from_datetime(due_date_obj)
            due_riyadh = due_ews.astimezone(tz_riyadh)
        else:
            # Fallback: try parsing string then convert
            parsed = datetime.datetime.fromisoformat(str(due_date_obj))
            parsed = pytz.UTC.localize(parsed) if parsed.tzinfo is None else parsed
            due_riyadh = EWSDateTime.from_datetime(parsed).astimezone(tz_riyadh)

        return due_riyadh.strftime('%Y-%m-%d %H:%M')

    except Exception as e:
        print(f"  ⚠️ Error formatting date: {e}")
        import traceback
        traceback.print_exc()
        # fallback to string
        return str(due_date_obj)


def get_reminder_subject(original_subject):
    return f"🔔 تذكير بالمتابعة: {original_subject}"


def get_reminder_body(original_subject, due_date_str):
    return (
        f"السلام عليكم ورحمة الله وبركاته،\n\n"
        f"نود تذكيركم بأن الرسالة التالية بلغت موعدها المحدد للمتابعة:\n\n"
        f"📩 العنوان: {original_subject}\n"
        f"📅 الموعد (بتوقيت الرياض): {due_date_str}\n\n"
        f"يرجى اتخاذ اللازم.\n\n"
        f"قسم المتابعة - هيئة الغذاء والدواء"
    )


def add_sent_category(existing_categories, sent_category):
    if not existing_categories:
        return [sent_category]
    if isinstance(existing_categories, str):
        category_list = [cat.strip() for cat in existing_categories.split(',')]
    else:
        category_list = list(existing_categories)
    if sent_category not in category_list:
        category_list.append(sent_category)
    return category_list