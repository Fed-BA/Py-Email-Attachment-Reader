import win32com.client
from datetime import datetime, timedelta
import pandas as pd
from zipfile import ZipFile
from pretty_html_table import build_table
import time
pd.set_option('display.max_columns', None)

sendEmailAlert = False
sendEmailWarning = False
today = datetime.now() - timedelta(hours=18)
today1 = today.strftime('%m/%d/%Y %H:%M %p')
fileDate = today.strftime('%m%d%Y')
outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
outlookSend = win32com.client.Dispatch('outlook.application')
inbox = outlook.GetDefaultFolder(6)
mail = outlookSend.CreateItem(0)
mail.Importance = 2
messages = inbox.Items
messages = messages.Restrict("[Subject] = 'FAB Daily Make Sheet Accuracy'")
messages = messages.Restrict("[ReceivedTime] > '" + today1 + "'")
outputDir = "C:/Users/bhill1/Documents/Attachments"

# loop through inbox looking for email sent in the last 12 hours that contains the MSA report attachment
# and save it out to local directory
for message in messages:
    attachments = message.Attachments
    attachment = attachments.Item(1)
    attachment_name = str(attachment).lower()
    attachment.SaveASFile(outputDir + '/' + attachment_name)

# attachment is in a zip file so we use zipfile to open it and then pandas to read in the excel file
# using pandas we clean up the raw excel file dropping blank rows and formating null and blank values.
# we only need to read in the values of the columns as we dont care about the realizations
with ZipFile(outputDir + '/fab daily make sheet accuracy cycledatename.xlsx.zip') as zip:
    zip.printdir()
    fields = [4, 8, 9, 14, 18, 19, 24, 28, 29, 34]
    shift_MSA = pd.read_excel(zip.read('FAB Daily Make Sheet Accuracy ' + fileDate + '.xlsx'),
                              sheet_name='FAB', usecols=fields, skiprows=2)
shift_MSA = shift_MSA.iloc[2:, :]
shift_MSA = shift_MSA.dropna(how='all')
shift_MSA = shift_MSA.replace('-', 0)
shift_MSA.columns = ['Greeley A shift', 'Greeley B shift', 'Greeley Combined',
                     'GI A shift', 'GI B shift', 'GI Combined',
                     'Cactus A shift', 'Cactus B shift', 'Cactus Combined',
                     'Hyrum A shift']

# Count the total number of values that are under 50% and under 20%.
# To many values under 50% indicate there may be an issue with data
# to many values under 20% indicate an immediate issue with data.
# Check for overstatement in quintiq or missing par data.
lessThan50Perc = shift_MSA[shift_MSA < .5].count()
lessThan20Perc = shift_MSA[shift_MSA < .2].count()

for index, value in lessThan20Perc.items():
    if value > 5:
        sendEmailAlert = True

for index, value in lessThan50Perc.items():
    if value > 5 and not sendEmailAlert:
        sendEmailWarning = True

# build HTML tables for email message
lessThan20Perc = pd.DataFrame({'Plant Shift': lessThan20Perc.index, 'Under %20 totals': lessThan20Perc.values})
lessThan50Perc = pd.DataFrame({'Plant Shift': lessThan50Perc.index, 'Under %50 totals': lessThan50Perc.values})
under50TableCount = build_table(lessThan50Perc, 'blue_light')
under20TableCount = build_table(lessThan20Perc, 'red_light')

# send the email alert based on priority.
if sendEmailAlert:
    mail.To = 'benjamin.hill@jbssa.com'
    mail.Subject = 'ALERT: MSA Data inconsistencies'
    mail.HTMLBody = under20TableCount + under50TableCount
    mail.Send()

if sendEmailWarning:
    mail.To = 'benjamin.hill@jbssa.com'
    mail.Subject = 'Warning: possible MSA Data inconsistencies'
    mail.HTMLBody = under50TableCount + under20TableCount
    mail.Send()
