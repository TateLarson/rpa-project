# Insert your python code here# INSERT ANSWER FOR QUESTION 1 HERE
pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input('Enter the file path here',mask = False, width= 400, height= 500)
    if directory.endswith('\\'):
        pass
    else: 
        directory = directory +'\\'
    try:
        list_files = os.listdir(directory)
        if 'ProductData.xlsx' and 'MasterInvoice.pdf'and 'Order1.xlsx' and 'Order2.xlsx' and 'Order3.xlsx' and 'Order4.xlsx' in list_files:
            break
        else:
            show_message('Error','Directory entered does not contain the required fields. Re-enter directory')
    except:
        show_message('Error','Not a directory in your PC, enter a valid directory')
# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
Descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')
for index,row in Descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
Prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')
for index,row in Prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename,directory):
    template = openpyxl.load_workbook(directory+filename)
    selected_sheet = template['OrderForm']

    order_data = {}
    order_data['InvoiceNum']= invoice_num
    order_data['Date'] = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerAddress2'] = str(selected_sheet['C17'].value) +', ' + str(selected_sheet['C18'].value) +', ' + str(selected_sheet['C19'].value)
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    

    total = 0
    cell_num = 24
    product_num =0
    while True:
        unit_num = selected_sheet[f'B{cell_num}'].value
        if unit_num is not None:
            order_data[f'Price{product_num}'] = product_prices[unit_num]
            order_data[f'Quantity{product_num}'] = selected_sheet[f'C{cell_num}'].value
            order_data[f'Amount{product_num}'] = order_data[f'Price{product_num}'] * order_data[f'Quantity{product_num}']
            order_data[f'Description{product_num}'] = product_descriptions[unit_num]
            total = order_data[f'Amount{product_num}']

            order_data[f'Price{product_num}'] = f"${order_data[f'Price{product_num}']:,.2f}"
            order_data[f'Amount{product_num}'] = f"${order_data[f'Amount{product_num}']:,.2f}"

            cell_num += 1
            product_num += 1
        else:
            break
    order_data['Total'] = f'${total:.2f}'
    name = order_data['CustomerName']
    email = selected_sheet['C13'].value
    date = order_data['Date']
    # invoice_num = Customer
    return order_data, name, email, date

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num,order_data,directory):
   create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice{invoice_num}.pdf', order_data)
order_data, name, email, date = get_order_data(8515,'Order1.xlsx',directory)
# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "tate.everett.larson@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')

    gmail_password = open(gmail_password_path).read()
    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    # name = name

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice{invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()
# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data,name,email,date = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num +1
    else:
        pass
    

