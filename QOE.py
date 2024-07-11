# This program will send Quartlery Offsite Memo requests to people as emails. First you paste data from the Construction Survey Excel sheet into the "Inputs" text files.
# URL of the Excel sheet here: https://partnershealthcare.sharepoint.com/:x:/r/sites/mghEnviroHealthSafety/FireLifeSafety/EOCdashboard/_layouts/15/Doc.aspx?sourcedoc=%7B5FB85D79-BA95-4F73-A0EF-3C25E09B9157%7D&file=2024%20Off%20Site%20Construction%20Survey.xlsx&action=default&mobileredirect=true&DefaultItemOpen=1

# NOTE: Each new quarter you run this program, make sure to actually update the offsites memo PDF and also to update the quarter and year inside of send_email (it's only one line of code in it)

import sys
import win32com.client as client

# This function sends the email. Watch out for the very long strings in it that Notepad is forced to wrap to new lines even when Word Wrap is turned off.
# Note that send_email()'s "attachment" parameter must be the absolute path for some reason. 
# I use this website to create the html body:	 https://onlinehtmleditor.dev/
def send_email(site, title_name, email_address):
	outlook = client.Dispatch("outlook.application")
	message = outlook.CreateItem(0)
	message.To = email_address
	# message.CC = ???
	message.Subject = "Construction/Renovation Projects Survey for 2024 Q2 at " + site
	
	# Create the email body, ad-libbed with send_email()'s arguments

	html_body_A = r"""<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">Hello """

	html_body_B = title_name

	html_body_C = r""",</span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">I hope your week is going well. I'm reaching out on behalf of MGH's Safety Office.&nbsp;</span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">We like to keep track of all construction and renovation projects at buildings with MGH locations. This is so we can inform first responders if an incident occurs, and so we can better assist with any safety matters.</span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">I have you on file as the contact for </span><span style="color:black;font-family:Calibri, sans-serif;font-size:11pt;">"""

	html_body_D = site

	html_body_E = r"""</span><span style="font-family:Calibri, sans-serif;font-size:11pt;">. Please fill out the table on this form with any construction activity that's pending or in progress at the location. If there is none, please reply to this email to notify me.</span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"><span style="color:black;"></span></span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">If you have any questions or are no longer the appropriate person to contact for this, let me know.</span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">Thank you,</span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;">~Drake</span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;"><span style="font-family:Calibri, sans-serif;font-size:11pt;"></span></p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;">&nbsp;</p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">Drake L. Sorkhab</span><span style="font-family:Calibri, sans-serif;font-size:11pt;"><strong><span style="color:#009CA6;font-size:10.0pt;"></span></strong></span></p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">Safety Technician</span></p>\n<p style="margin:0in;"><span style="background-color:white;color:#333333;font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">EH&amp;S Dept (Safety)</span></p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">(he/him/his)</span><br><span style="background-color:white;color:#333333;font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">Massachusetts General Hospital</span></p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">50 Staniford Street, Suite 410</span></p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">Boston, MA 02114-2514</span></p>\n<p style="margin:0in;"><span style="font-family:&quot;Times New Roman&quot;,serif;font-size:10.0pt;">T 617-72<strong>4-4572</strong></span></p>"""
	
	html_body = html_body_A + html_body_B + html_body_C + html_body_D + html_body_E
	html_body = html_body.replace(r"""\n""","\n") #I messed up the newline processing so here's the fix
	message.HTMLBody = html_body
	message.Display(False) # false for "modal Outlook window"
	message.Attachments.Add("C:\\Users\\DS001\\OneDrive - Mass General Brigham\\ONEDRIVE DOCUMENTS\\CODE\\QUARTERLY OFFSITES EMAILER\\Offsites_Construction_Memo.doc")	






# The main function opens the input files and makes 3 lists out of them, and iterates through the lists with send_email() while pausing after each one so the user can press send in Outlook. It also skips over blank or partially-blank lines which obviously don't correspond to an actual item on the Excel sheet.

with open(r"""C:\Users\DS001\OneDrive - Mass General Brigham\ONEDRIVE DOCUMENTS\CODE\QUARTERLY OFFSITES EMAILER\Inputs\Sites.txt""" , "r") as f:
	Sites = f.read().split("\n")
with open(r"""C:\Users\DS001\OneDrive - Mass General Brigham\ONEDRIVE DOCUMENTS\CODE\QUARTERLY OFFSITES EMAILER\Inputs\Email Addresses.txt""" , "r") as f:
	Email_Addresses = f.read().split("\n")
with open(r"""C:\Users\DS001\OneDrive - Mass General Brigham\ONEDRIVE DOCUMENTS\CODE\QUARTERLY OFFSITES EMAILER\Inputs\Titles and Names.txt""" , "r") as f:
	Titles_and_Names = f.read().split("\n")

# Throw an error if the lists lengths don't line up (they definitely should, since they're each columns pasted from an Excel sheet)
if not ( len(Sites) == len(Email_Addresses) and len(Sites) == len(Titles_and_Names) ):
	print("\nError: lists not same length (" + str(len(Sites)) + "," + str(len(Email_Addresses)) + "," + str(len(Titles_and_Names)) + ").\n")
	exit()

for number in range(len(Sites)): 
	if Sites[number] != "" and Email_Addresses[number] != "" and Titles_and_Names[number] != "" : #Make sure none of the lines are blank so that we know we're on a data row of the Excel file and not a blank or "category title" row
		if Sites[number] != "???" and Email_Addresses[number] != "???" and Titles_and_Names[number] != "???" : #For convenience, ignore any lines that have ??? as an entry in one of the fields.
			send_email(Sites[number] , Titles_and_Names[number] , Email_Addresses[number] )
			unused_var = input()
			print("Just did email for " , Sites[number])
		


send_email("MGH Danvers", "Chris", "cburt@mgb.org")