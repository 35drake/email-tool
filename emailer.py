# A simple program that can open a complete Outlook email, ready-to-send.
# The major issue is that send_email()'s "attachment" parameter must be the absolute path for some reason. How do I get it to work with the relative path? I've seen people do: os.path.join(os.getcwd(), 'file.png')
# I use this website to create the html body:	 https://onlinehtmleditor.dev/

import sys
import win32com.client as client

def send_email(recipient,cc,subject,html_body_file,attachment):
	outlook = client.Dispatch("outlook.application")
	message = outlook.CreateItem(0)
	message.To = recipient
	message.CC = cc
	message.Subject = subject
	with open(html_body_file) as myfile: # open the body's text file and read it
		html_body = myfile.read()
	message.HTMLBody = html_body
	message.Display(False) # false for "modal Outlook window"
	message.Attachments.Add(attachment)	

send_email("dsorkhab@mgb.org ; 35drake@gmail.com", "phys152boi@gmail.com","Subject 41","body.txt","C:\\Users\\DS001\\OneDrive - Mass General Brigham\\ONEDRIVE DOCUMENTS\\CODE\\emailer\\test-attachment.txt")
