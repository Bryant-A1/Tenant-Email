"""Sends emails.

Reads in data from a text file and an xlsx file, then logs
into my email and sends a bunch of emails.

Note: currently, this is using a TEST file so I don't send a bunch of emails to
actual tenants.

The Excel Workbook is called 'test_ML.xlsx'
"""
import openpyxl


class Tenant(object):
    """Tenant class used for clean iterative implementation."""

    def __init__(self, Name, Property, LeaseEndDate, Email):
        """Init Tenant class."""
        self.Name = Name
        self.Property = Property
        self.LeaseEndDate = LeaseEndDate
        self.Email = Email

    def GetVals(self):
        """Return values."""
        return self.Name, self.Property, self.LeaseEndDate, self.Email

    def GetEmail(self):
        """Return email."""
        return self.Email


def GetInfo():
    for i in range(1, sheet.max_row):
        tName = sheet.cell(row=i, column=1)
        tProp = sheet.cell(row=i, column=2)
        tLease = sheet.cell(row=i, column=3)
        tEmail = sheet.cell(row=i, column=4)
        NewTenant = Tenant(tName, tProp, tLease, tEmail)
        Tenants.append(NewTenant)


# Let's read some data from the text file now.
def ReadEmailMsg():
    email_message = open('ConfirmMoveOut.txt')
    print("Let's print the message just to be sure")
    print(email_message.read())
    email_message.seek(0)

    str_email_message = str(email_message.read())
    email_message.close()
    return str_email_message


def FormatEmailMsg():
    str_email_message = ReadEmailMsg()

    e_messages = []
    for t in Tenants:
        e_messages.append(str_email_message.format(t.GetVals()))

    return e_messages


'''
Next Step: Connect to the email server
'''


def ConnectToEmailServer():
    import smtplib
    import getpass

    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    # Now let's establish a connection to the server:
    smtpObj.ehlo()
    # Now let's enable encryption
    smtpObj.starttls()
    # Should return something with a value of 220
    my_Email = input("What is your email address?")
    my_Password = getpass.getpass()

    return smtpObj, my_Email, my_Password


def SendEmail():
    smtpObj, my_Email, my_Password = ConnectToEmailServer()

    smtpObj.login(my_Email, my_Password)
    # Okay, now that we're logged in, it's time to send the email.

    e_messages = FormatEmailMsg()
    i = 0
    for t in Tenants:
        smtpObj.sendmail(my_Email, t.GetEmail(),
                         'Subject: Please Confirm Move Out or Extension\n' +
                         e_messages[i])
        i += 1

    smtpObj.quit()


if __name__ == "__main__":
    print("Opening Workbook...")
    wb = openpyxl.load_workbook('test_ML.xlsx')
    sheet = wb.get_sheet_by_name('Sheet1')

    Tenants = []
    GetInfo()

    SendEmail()
