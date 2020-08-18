import win32com.client

conn = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 3  Deleted Items
# 4  Outbox
# 5  Sent Items
# 6  Inbox
# 9  Calendar
# 10 Contacts
# 11 Journal
# 12 Notes
# 13 Tasks
# 14 Drafts
inbox = conn.GetDefaultFolder(6) 

# reading mailbox items
messages = inbox.Items

# reading message
message = messages.GetLast()

# message content
body_content = message.body

print(body_content)