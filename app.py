import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

#Get directory with the correct files
while True:
    directory = single_input("Please enter the directory where the Product Data, Master Invoice, and Order forms are stored.")
    directory = directory.strip('"')
    directory = directory.strip("'")
    directory = directory + "\\"
    try:
        filenames = os.listdir(directory)
        print(filenames)
    except:
        show_message("Error","That is not a valid directory. Please try again.")
        continue
    if "ProductData.xlsx" in filenames and "MasterInvoice.pdf" in filenames and "Order1.xlsx" in filenames and "Order2.xlsx" in filenames and "Order3.xlsx" in filenames and "Order4.xlsx" in filenames:
        break
    show_message("Error","That directory does not have the Product Data, Master Invoice, and Order forms.")

#create product description and product price dictionaries
product_descriptions = {}
product_prices = {}
product_data_descriptions = pd.read_excel(directory+"ProductData.xlsx",sheet_name="Descriptions")
for i, row in product_data_descriptions.iterrows():
    product_descriptions[row["Unit number"]] = row.Description
product_data_prices = pd.read_excel(directory+"ProductData.xlsx",sheet_name="Prices")
for i, row in product_data_prices.iterrows():
    product_prices[row["Unit number"]] = row["Price per unit"]

#Get data from Order Form
def get_order_data(invoice_num, filename, directory):
    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    file_workbook = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = file_workbook["OrderForm"]
    date = selected_sheet["C9"].value
    date = str(date.month) + "/" + str(date.day) + "/" + str(date.year)
    order_data["Date"] = date
    name = selected_sheet['C14'].value
    order_data['CustomerName'] = name
    email = selected_sheet['C13'].value
    order_data['CustomerAddress1'] = selected_sheet["C16"].value
    order_data['CustomerAddress2'] = f"{selected_sheet["C17"].value}, {selected_sheet["C18"].value} {selected_sheet["C19"].value}"
    order_data['CustomerCompany'] = selected_sheet["C15"].value
    i = 0
    while True:
        row = i + 24
        unit_num_cell = "B"+str(row)
        unit_num = selected_sheet[unit_num_cell].value
        if unit_num is None:
            break
        description = product_descriptions[unit_num]
        order_data["Description"+str(i)] = description
        price = product_prices[unit_num]
        order_data["Price"+str(i)] = price
        quantity_cell = "C"+str(row)
        quantity = selected_sheet[quantity_cell].value
        order_data["Quantity"+str(i)] = quantity
        amount = order_data["Price"+str(i)] * order_data["Quantity"+str(i)]
        order_data["Amount"+str(i)] = amount
        i += 1
    return order_data, name, email, date

#Create Invoice from order data
def create_invoice(invoice_num, order_data, directory):
    total = 0
    for n in range(5):
        amount_n = float(order_data.get("Amount"+str(n), 0))
        total += amount_n
        if amount_n != 0:
            order_data["Amount"+str(n)] = f"${order_data["Amount"+str(n)]:,.2f}"
    order_data["Total"] = f"${total:,.2f}"
    create_pdf(directory+"MasterInvoice.pdf", directory+"Invoice " + str(invoice_num) + ".pdf", order_data)

#Send email
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory, gmail_user, gmail_password):
    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
E-mail: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = f"{directory}Invoice {str(invoice_num)}.pdf"
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

#Enter email
while True:
    gmail_user = single_input("Please enter your gmail")
    if gmail_user.endswith("@gmail.com"):
        break
    show_message("Error","That was not a google gmail address.")

#Enter email password file
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

#Create each invoice
invoice_num = 8515
for filename in filenames:
    if "Order" in filename and ".xlsx" in filename:
        order_data, name, email, date = get_order_data(invoice_num,filename, directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory,gmail_user,gmail_password)
        invoice_num +=1