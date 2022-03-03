import pandas as pd
import smtplib

# ENTER EMAIL ID, PASSWORD AND SUBJECT HERE
your_email = "EMAIL ID"
your_password = "PASSWORD"
subject = "SUBJECT"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

attendance_sheet = pd.read_excel('sample.xlsx')

names = attendance_sheet['Name'].to_list()
emails = attendance_sheet['email'].to_list()
id_numbers = attendance_sheet['ID Number'].to_list()

attendance_sheet = attendance_sheet.drop(['email'], axis=1)
attendance_sheet = attendance_sheet.drop(['Name'], axis=1)
attendance_sheet = attendance_sheet.drop(['ID Number'], axis=1)

attendance_sheet.columns = [col.strftime('%d/%m/%y') for col in attendance_sheet.columns]

attendance_sheet.insert(0, "Name", names)
attendance_sheet.insert(0, "ID Number", id_numbers)

attendance_sheet['Classes Attended'] = attendance_sheet.iloc[:,2:].sum(axis=1)
attendance_sheet['Total Classes Held'] = attendance_sheet.iloc[:,3:].count(axis=1)

for i in range(len(emails)):
    name = names[i]
    email = emails[i]
    row = attendance_sheet.iloc[i].to_string()
    message = 'Subject: {}\n\n{}'.format(subject, row)
    server.sendmail(your_email, [email], message)

server.close()