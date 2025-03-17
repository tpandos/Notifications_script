import win32com.client

def send_email(recipient, subject, body, attachment_path=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body

    if attachment_path:
        mail.Attachments.Add(attachment_path)
    
    mail.Send()

send_email("sebastian.blondetvillafana08@lagcc.cuny.edu", "Test Email", "This is a test email from Python")