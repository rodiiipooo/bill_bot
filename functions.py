billing_dist = []
overdue_dist = []
focus_dist = []
budgetSpend_dist =[]


def email(option):
    pass
def tasks(option):
    if option == "d-Overdue Invoices":
        pass
    elif option == "d-Focus File":
        pass
    elif option == "d-Posted/Unposted":
        pass
    elif option == "d-All Daily":
        from datetime import datetime
        import win32com.client as win32
        import matplotlib.pyplot as plt
        import pandas as pd
        ### time
        from datetime import datetime, timedelta
        def yesterday(frmt='%Y-%m-%d', string=True):
            yesterday = datetime.now() - timedelta(1)
            if string:
                return yesterday.strftime(frmt)
            return yesterday

        ### calendar
        start_date = datetime(2023, 3, 1)
        end_date = datetime(2023, 4, 1)

        ### read files for reports
        csp_file = pd.read_excel()
        unposted_invoices = pd.read_excel()
        project_infExtract = pd.read_csv()
        billing_register = pd.read_excel()
        gbs_extract = pd.read_csv("", delimiter=True)

        # prepare reports
        base_file = csp_file.groupby(unposted_invoices, by="") \
            .groupby(project_infExtract, by="") \
            .groupby(billing_register, by="") \
            .groupby(gbs_extract, by="")

        ### prepare visuals needed
        plt.plot()

        ### distribution lists

        ### email template for billing
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = billing
        mail.Subject = 'Posted Unposted ' + str(yesterday)
        mail.Body = 'Good morning, attached to this email are the posted/unposted invoices for ' + str(yesterday)
        mail.HTMLBody = '<h2>HTML Message body</h2>'  # this field is optional
        # To attach a file to the email (optional):
        attachment = "Path to the attachment"
        mail.Attachments.Add(attachment)
        # send email
        mail.Send()
    else:
        pass

def all_tasks(list):
    for i in list:
        print(i)