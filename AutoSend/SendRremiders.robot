*** Settings ***
Documentation     This robot checks a specific Outlook folder for flagged emails
...               due soon and sends a reminder.
Library           RPA.Outlook.Application
Library           OutlookHelpers.py    # Import our custom Python helper file
Library           Collections          # For list counting

*** Variables ***
${FOLDER_NAME}         Flag
${TARGET_FOLDER_PATH}  Inbox/${FOLDER_NAME}
${RECIPIENT_EMAIL}     FollowUp_TAM@sfda.gov.sa

*** Tasks ***
Check Flagged Emails and Send Reminders
    Log To Console    Starting Outlook reminder process...
    
    # 1. Get current time from our Python helper
    ${now_riyadh} =    Get Riyadh Datetime
    
    # 2. Get all emails from the target folder
    # 'account_name=None' uses the default Outlook account
    ${messages} =    Get Emails
    ...    account_name=None
    ...    folder_name=${TARGET_FOLDER_PATH}
    
    ${total_count} =    Get Length    ${messages}
    Log To Console    Found ${total_count} messages in '${TARGET_FOLDER_PATH}'.
    
    ${reminder_count} =    Set Variable    ${0}

    # 3. Loop through each message
    FOR    ${msg}    IN    @{messages}
        # We use the 'raw_object' to access detailed properties
        # just like in the original win32com script.
        ${raw_msg} =    Evaluate    $msg.raw_object

        # 4. Check all conditions
        ${should_process} =    Run Keyword And Return Status
        ...    Email Should Be Processed    ${raw_msg}

        # 5. If all checks pass, process the reminder
        IF    ${should_process}
            ${was_sent} =    Run Keyword And Return Status
            ...    Process And Send Reminder    ${raw_msg}    ${now_riyadh}
            
            IF    ${was_sent}
                ${reminder_count} =    Set Variable    ${reminder_count + 1}
            END
        END
    END

    Log To Console    \n📨 تم إرسال ${reminder_count} تذكير من مجلد '${FOLDER_NAME}'.

*** Keywords ***
Email Should Be Processed
    [Arguments]    ${raw_msg}
    
    # Check 1: Is it a mail item (Class 43)?
    Should Be Equal As Integers    ${raw_msg.Class}    43
    
    # Check 2: Is it marked as a task?
    Should Be True    ${raw_msg.IsMarkedAsTask}
    
    # Check 3: Does it have a due date?
    Should Not Be Equal    ${raw_msg.TaskDueDate}    ${None}
    
    # Check 4: Has a reminder already been sent?
    Should Not Contain    ${raw_msg.Categories}    AutoReminderSent

Process And Send Reminder
    [Arguments]    ${raw_msg}    ${now_riyadh}
    
    # Get the due date object
    ${due_date} =    Set Variable    ${raw_msg.TaskDueDate}

    # Check 1: Is it due within the 2-day window?
    ${is_due} =    Is Due Soon    ${due_date}    ${now_riyadh}
    
    # If it's not due soon, stop this keyword.
    Run Keyword Unless    ${is_due}    RETURN

    # --- If we are here, the email is due soon ---
    Log    Email "${raw_msg.Subject}" is due soon. Sending reminder.
    
    # 2. Get formatted text from our Python helpers
    ${subject} =       Get Reminder Subject     ${raw_msg.Subject}
    ${due_date_str} =  Format Due Date For Email    ${due_date}    ${now_riyadh.tzinfo}
    ${body} =          Get Reminder Body        ${raw_msg.Subject}    ${due_date_str}

    # 3. Send the new reminder email
    Send Email
    ...    recipients=${RECIPIENT_EMAIL}
    ...    subject=${subject}
    ...    body=${body}

    # 4. Mark the original email as "Sent"
    ${new_categories} =    Add Sent Category    ${raw_msg.Categories}
    Evaluate    $raw_msg.Categories = $new_categories
    Evaluate    $raw_msg.Save()
    
    Log To Console    ✅ تم إرسال تذكير لـ: ${raw_msg.Subject}