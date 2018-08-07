import openpyxl
from datetime import datetime,timedelta
import calendar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def convert(s):
    return datetime.strptime(s, '%Y.%m.%d. %H:%M:%S')

wb = openpyxl.load_workbook('export.xlsx')
sheet = wb.active

schedule = []

for row in range(2, sheet.max_row + 1):
    # Each row in the spreadsheet has data for one census tract.
    start = sheet['A' +str(row)].value
    end  = sheet['B' + str(row)].value
    desc = sheet['C' + str(row)].value
    loc = sheet['D' + str(row)].value
    # print( calendar.day_name[convert(start).weekday()])
    schedule.append([start,end,desc,loc])

schedule = sorted(schedule, key = lambda x: convert(x[0]))

classes_tomorrow = []

for value in schedule:

    if datetime.now().date() == (convert(value[0]) - timedelta(1)).date():
        #print(calendar.day_name[convert(value[0]).weekday()])
        classes_tomorrow.append(value)
print(classes_tomorrow)

start = ', '.join(x[0] for x in classes_tomorrow)
end = ', '.join(x[1] for x in classes_tomorrow)
desc = ', '.join(x[2] for x in classes_tomorrow)
loc = ', '.join(x[3] for x in classes_tomorrow)

# message to be sent
message = """<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-2"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style><![endif]--><style><!--
/* Font Definitions */
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;
	mso-fareast-language:EN-US;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:#0563C1;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:#954F72;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal;
	font-family:"Calibri",sans-serif;
	color:windowtext;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-family:"Calibri",sans-serif;
	mso-fareast-language:EN-US;}
@page WordSection1
	{size:612.0pt 792.0pt;
	margin:70.85pt 70.85pt 70.85pt 70.85pt;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]--></head><body lang=HU link="#0563C1" vlink="#954F72"><div class=WordSection1><p class=MsoNormal>Hi Hoang,<o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Here are all classes you will have <b>tomorrow (</b><i>In ascending order</i><b>):<o:p></o:p></b></p>
<p class=MsoNormal><b><o:p>&nbsp;</o:p></b></p><p class=MsoNormal><b>               Subjects: <o:p></o:p></b>"""+desc+"""</p>
<p class=MsoNormal><b>               Start at: <o:p></o:p></b>"""+start+"""</p>
<p class=MsoNormal><b>               End: <o:p></o:p></b>"""+end+"""</p>
<p class=MsoNormal><b>               Locations: </b>"""+loc+"""<o:p></o:p></p>
<p class=MsoNormal>               <o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal style='text-autospace:none'><b><span lang=EN-US style='font-size:10.0pt;color:#548DD4;mso-fareast-language:HU'>Phi Huy Hoang</span></b><b><span lang=EN-US style='font-size:10.0pt;color:navy;mso-fareast-language:HU'><br></span></b><span lang=EN-US style='font-size:8.0pt;color:#1F497D;mso-fareast-language:HU'>Intern<o:p></o:p></span></p><p class=MsoNormal style='text-autospace:none'><span lang=EN-US style='font-size:8.0pt;color:#1F497D;mso-fareast-language:HU'>Machine Learning |&nbsp;SAP Leonardo<o:p></o:p></span></p></div></body></html>"""


# Define SMTP email server details
smtp_server = 'smtp.gmail.com'
smtp_user = '0'
smtp_pass = '0'
#
# # Construct email
msg = MIMEMultipart('alternative')
msg['To'] = "luumanhlapdi@gmail.com"
msg['From'] = "bohemian.crush.123@gmail.com"
msg['Subject'] = 'Reminder: Class Schedule'

message_encrypt = MIMEText(message, 'html')


msg.attach(message_encrypt)
#
s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
try:
# Send the message via an SMTP server
    s.login(smtp_user, smtp_pass)
    s.sendmail(smtp_user, "luumanhlapdi@gmail.com", message_encrypt.as_string())
except Exception as e:
    print(e)

finally:
    s.quit()

