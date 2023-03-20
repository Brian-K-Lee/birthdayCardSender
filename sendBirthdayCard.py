##############################################
#############Birthday Card Sender#############
##############Author: Brian Lee###############
##########Last Updated On 3/19/2023###########
##############################################

# This Python program reads in a list of names and their corresponding birthdays from an Excel file
# and checks to see if any of the birthdays match the current date. 
# If there is a match, the program sends an email to the person wishing them a happy birthday
# and attaching a birthday card image I designed. 
# The email is sent using the smtplib library
# and the program connects to an SMTP server using the sender's email and password.

### Load Libraries
import datetime
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

### Read in a file with 100 random names, birthday and emails
# Used 100 random names from random-name-generator.info
names = pd.read_excel('C:/Users/KT/Documents/Python/ForFun/birthdayCard/hundredRandomNames.xlsx')
# Keep all columns except for the last 
names = names.iloc[:,:-1]
# Remove the number label in the "Name" column
names['Name'] = names['Name'].str.replace(r'^\d+\.\s*', '', regex=True)

### Get today's date as "mm/dd"
today = str(datetime.date.today().strftime('%m/%d'))
# Find whose brithday is today 
matchDate = names['Birthday'].isin([today])
matchPerson = names.loc[matchDate]
# Print matching Name and Birthday
for i in range(len(matchPerson)):
    print("Today is ", matchPerson['Birthday'].values[i],", and it is ", matchPerson['Name'].values[i],"'s Birthday!!", sep="")

### Email information
sender_email = "ADD YOUR EMAIL@outlook.com"
password = "ADD YOUR PASSWORD"
subject = "HAPPY BIRTHDAY to YOU!"

### Image to be sent
with open("C:/Users/KT/Documents/Python/ForFun/birthdayCard/brithdayCard.jpg", "rb") as f:
    img_data = f.read()

### Create a Loop to Send Email To Each Person Whose Birthday Is Today
for i in range(len(matchPerson)):
    # Get recipient's email(s)
    receiver_email = matchPerson['Email'].values[i]
    # Create the email message
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    # Email body
    text = MIMEText("Wishing you a very happy birthday and a splendid year ahead.")
    msg.attach(text)
    # Image attachment
    image = MIMEImage(img_data, name="image.jpg")
    msg.attach(image)
    # Send email
    with smtplib.SMTP('smtp.office365.com', 587) as smtp:
        smtp.starttls()
        smtp.login(sender_email, password)
        smtp.send_message(msg)
        print("Email sent!")