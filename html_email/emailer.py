import win32com.client as win32
from django.contrib.staticfiles import finders

def sendEmail(subject, to, body):
    olMailItem = 0x0
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')
    new_mail = outlook.CreateItem(olMailItem)


    attachment = new_mail.Attachments.Add(finders.find('images\\attestation.gif'), win32.constants.olEmbeddeditem, 0,
                                         "Attestation")
    imageCid = "attestation.gif@123"
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid)

    new_mail.Subject = subject
    new_mail.To = to
    new_mail.HTMLBody = body.replace('attestation_gif_src',imageCid)
    new_mail.Send()