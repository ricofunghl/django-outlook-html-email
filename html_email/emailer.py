import win32com.client as win32
from django.contrib.staticfiles import finders
from django.template.loader import get_template

from html_email.models import Fruit


def send_email(body):
    olMailItem = 0x0
    outlook = win32.gencache.EnsureDispatch('Outlook.Application')
    new_mail = outlook.CreateItem(olMailItem)


    attachment = new_mail.Attachments.Add(finders.find('images\\fruit.jpg'), win32.constants.olEmbeddeditem, 0,
                                         "Attestation")
    imageCid = "fruit@image"
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imageCid)

    new_mail.Subject = 'HTML Email Subject'
    new_mail.To = 'abc@gmail.om'
    new_mail.HTMLBody = body.replace('fruit_img_src',imageCid)
    new_mail.Send()

def construct_email_body():
    fruits = Fruit.objects.all()
    email_temp = get_template('email_html.html')
    email_ctx = dict({'fruits': fruits,})

    body = email_temp.render(email_ctx)

    send_email(body)
