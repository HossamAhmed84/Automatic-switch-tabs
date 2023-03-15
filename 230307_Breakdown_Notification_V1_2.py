import cx_Oracle
import smtplib
from email.mime.text import MIMEText
import datetime
import time
import win32com.client
import calendar
import subprocess

#Actual Time variable
start_time = time.time()
currentSecond= datetime.datetime.now().second
currentMinute = datetime.datetime.now().minute
currentHour = datetime.datetime.now().hour

currentDay = datetime.datetime.now().day
currentMonth = datetime.datetime.now().month
currentYear = datetime.datetime.now().year
#https://docs.python.org/3/library/calendar.html
currentMonthName=calendar.month_name[currentMonth] #for the full name 
currentFullDate=datetime.datetime.today().strftime('%Y-%m-%d')

# database connections
databases = [
    {'ip': '10.15.174.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Juhayna',
     'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA-869351017': 'L5', 'L5':'CFA312-35',  'L5_CFA312-35':'V31.0',
     'CFA_853251046': 'L6', 'L6':'CFA209-32',  'L6_CFA209':'V04.0',
     'Cfa869351080': 'L7', 'L7':'CFA312-35', 'L7_CFA312-35':'V132.0',
     'CFA873051115': 'L8', 'L8':'CFA124-36', 'L8_CFA124-36':'V19.0',
     'CFA873051087': 'L9', 'L9':'CFA124-36', 'L9_CFA124-36':'V19.0',
     'CFA873051118': 'L12', 'L12':'CFA124-36', 'L12_CFA124-36':'V19.0'},
    {'ip': '10.15.174.27', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Juhayna',
    'CFA112-32':'12000' ,'CFA124-36':'24000' ,
     'CFA-870151143': 'L1', 'L1':'CFA112-32',  'L1_CFA112-32':'V23.0',
     'CFA_873051003': 'L2', 'L2':'CFA124-36',  'L2_CFA124-36':'V05.0',
     'CFA_873051004': 'L3', 'L3':'CFA124-36', 'L3_CFA124-36':'V05.0',
     'CFA873051061': 'L4', 'L4':'CFA124-36', 'L4_CFA124-36':'V19.0'},
    {'ip': '10.15.162.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Beyti',
     'CFA124-36':'24000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA873051098': 'L8', 'L8':'CFA124-36',  'L8_CFA124-36':'V19.0',
     'CFA873051088': 'L9', 'L9':'CFA124-36',  'L9_CFA124-36':'V19.0',
     'CFA873051102': 'L10', 'L10':'CFA124-36', 'L10_CFA124-36':'V19.0',
     'CFA871251020': 'L17', 'L17':'CFA124-36', 'L17_CFA124-36':'V19.0',},
    {'ip': '10.15.162.27', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Beyti',
    'CFA209-32':'9000' ,'CFA1224-36':'24000' ,'CFA124-36':'24000' ,
     'CFA853251057': 'L12', 'L12':'CFA209-32',  'L12_CFA209-32':'V05.0',
     'CFA853251058': 'L13', 'L13':'CFA209-32',  'L13_CFA209':'V05.0',
     'CFA853251059': 'L14', 'L14':'CFA209-32', 'L14_CFA209-32':'V05.0',
     'CFA871251015': 'L15', 'L15':'CFA1224-36', 'L15_CFA1224-36':'V19.0',
     'CFA871251016': 'L16', 'L16':'CFA1224-36', 'L16_CFA1224-36':'V19.0'},
    {'ip': '10.34.22.146', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Lactalis-ZA',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA812-35':'12000' ,
     'CFA869351092': 'COMBI_1', 'COMBI_1':'CFA312-35',  'COMBI_1_CFA312-35':'V31.0',
     'CFA-869351036': 'COMBI_2', 'COMBI_2':'CFA312-35',  'COMBI_2_CFA312-35':'V31.0',
     'CFA869351114': 'COMBI_3', 'COMBI_3':'CFA312-35', 'COMBI_3_CFA312-35':'V132.0',
     'CFA869851058': 'COMBI_4', 'COMBI_4':'CFA812-35', 'COMBI_4_CFA812-35':'V134.0',
     'CFA869851063': 'COMBI_5', 'COMBI_5':'CFA812-35', 'COMBI_5_CFA812-35':'V134.0',},
    {'ip': '10.14.70.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Aljaid',
    'CFA712-32':'12000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA870751218': 'L1', 'L1':'CFA712-32',  'L1_CFA712-32':'V134.0',
     'CFA870751210': 'L2', 'L2':'CFA712-32',  'L2_CFA712-32':'V134.0',
     'CFA871251041': 'L3', 'L3':'CFA124-36', 'L3_CFA124-36':'V20.0',
     'CFA871251046': 'L4', 'L4':'CFA124-36', 'L4_CFA124-36':'V20.0',
     'CFA869851042': 'L5', 'L5':'CFA812-35', 'L5_CFA812-35':'V134.0'},
    {'ip': '10.15.155.16', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'TchinLait-Alger',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA869351064': 'L1', 'L1':'CFA312-35', 'L1_CFA312-35':'V132.0',
     'CFA869351066': 'L2', 'L2':'CFA312-35', 'L2_CFA312-35':'V132.0',
     'CFA873051082': 'L3', 'L3':'CFA124-36', 'L3_CFA124-36':'V20.0',
     'CFA873051095': 'L4', 'L4':'CFA124-36', 'L4_CFA124-36':'V20.0',},
    {'ip': '10.15.184.16', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'TchinLait-Bejaia',
    'CFA310-32':'10000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA_860351064': 'L1', 'L1':'CFA310-32',  'L1_CFA310-32':'V31.0',
     'CFA-869351035': 'L2', 'L2':'CFA312-35',  'L2_CFA312-35':'V04.0',},
    {'ip': '10.15.146.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'ORCL', 'customerName': 'TchinLait-Setif',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA869351063': 'L1', 'L1':'CFA312-35',  'L1_CFA312-35':'V31.0',
     'CFA873051071': 'L2', 'L2':'CFA124-36', 'L2_CFA124-36':'V19.0',
     'CFA873051099': 'L3', 'L3':'CFA124-36', 'L3_CFA124-36':'V19.0',},
    {'ip': '10.14.50.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Baladna',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'cfa853251039': 'L1', 'L1':'CFA209-32',  'L1_CFA209-32':'V4.0',
     'cfa873051146': 'L2', 'L2':'CFA124-36', 'L2_CFA124-36':'V19.0',
     'CFA869351116': 'L3', 'L3':'CFA124-36', 'L3_CFA124-36':'V19.0'},
    {'ip': '10.15.226.17', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Almarai',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA869351039': 'U22', 'U22':'CFA312-35',  'U22_CFA312-35':'V132.0',
     'CFA_869351047': 'U28', 'U28':'CFA312-35',  'U28_CFA312-35':'V132.0',
     'CFA_869351029': 'U30', 'U30':'CFA312-35',  'U30_CFA312-35':'V132.0'},
    {'ip': '10.15.226.22', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Almarai',
    'CFA209-32':'9000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA_873051006': 'U24', 'U24':'CFA124-36',  'U24_CFA124-36':'V20.0',
     'CFA873051013': 'U25', 'U25':'CFA124-36',  'U25_CFA124-36':'V20.0',
     'CFA873051057': 'U26', 'U26':'CFA124-36',  'U26_CFA124-36':'V20.0',
     'CFA873051058': 'U27', 'U27':'CFA124-36',  'U27_CFA124-36':'V20.0',
     'CFA_873051002': 'U29', 'U29':'CFA124-36',  'U29_CFA124-36':'V20.0',},
    {'ip': '10.15.226.20', 'user': 'ecs3rtdb', 'password': 'ecs', 'service': 'SIGDB1', 'customerName': 'Almarai',
    'CFA612-35':'6000' ,'CFA312-35':'12000' ,'CFA124-36':'24000' ,
     'CFA869651011': 'U21', 'U21':'CFA612-35', 'U21_CFA612-35':'V23.0',
     'CFA869651004': 'U23', 'U23':'CFA612-35', 'U23_CFA612-35':'V23.0',},
    
]

# table name
status_table_name = 'LOG_CURRENTOPSTATES'
machine_name_table_name = 'STAT_MACHINE'
destination_table_name = 'all_status'

# columns to select from each table
#status_columns = ['MACHINE_REF','COMPATIBILITY_REF', 'COMMUNICATIONSTATE', 'DT', 'VAR_ERROR_REF', ]
status_columns = ['PRLNAME','COMPATIBILITY_REF', 'COMMUNICATIONSTATE', 'DT', 'VAR_ERROR_REF']
status_columns_list= ','.join(status_columns)
machine_name_columns = ['ID', 'PRLNAME']

# destination table name
destination_table_name = 'all_status'

# email configuration
email_from = 'reliabilitycenter.mea@sig.biz'
email_to = 'hossam.ahmed@sig.biz; ana.sedarati@sig.biz;mostafa.shalaby@sig.biz; waleed.radwan@sig.biz; abdelouahed.benbelgacem@sig.biz'
#email_to = 'hossam.ahmed@sig.biz; hossamahmed999@gmail.com'
email_cc= 'hossam.ahmed@sig.biz'
email_subject = 'Breakdown Automated Notification for '
smtp_server = 'smtp_server_address'
smtp_port = 'smtp_server_port'
smtp_username = 'username'
smtp_password = 'Password'
#customerName='Juhayna'

# Set the log file path
log_file_path = ""
log_file_name = '_log.txt'
# Function to write logs to a daily log file
def write_logs(log_file_path, message):
    now = datetime.datetime.now()
    log_file_path = log_file_path+now.strftime('%Y-%m-%d') + log_file_name
    # Write a log entry
    with open(log_file_path, 'a') as log_file:
        log_file.write(now.strftime('%Y-%m-%d %H:%M:%S') + ': ' + message + '\n')

        
#Function to ping IP
def ping(ip):
    # Ping the IP address and capture the output
    ping_output = subprocess.Popen(['ping', '-n', '1', '-w', '100', ip], stdout=subprocess.PIPE).communicate()[0]

    # Check if the ping was successful
    if 'Reply from' in ping_output.decode('utf-8'):
        print(f'{ip} is connected')
        message= "VPN is connected successfully"
        write_logs(log_file_path, message)
    else:
        print(f'{ip} is not connected')
        message= "VPN didn't connect successfully"
        write_logs(log_file_path, message)

# function to execute a query on a database connection
def execute_query(connection, query):
    cursor = connection.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    cursor.close()
    return rows

# function to execute a query on a database connection and insert the results into a destination table
def insert_into_destination_table(connection, query, destination_table):
    cursor = connection.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    for row in rows:
        cursor.execute('INSERT INTO {} VALUES {}'.format(destination_table, row))
    cursor.close()
    connection.commit()
    
# function to send an email with the given message
def send_email(message):
    msg = MIMEText(message)
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = email_subject
    smtp_server = smtplib.SMTP(smtp_server, smtp_port)
    smtp_server.starttls()
    smtp_server.login(smtp_username, smtp_password)
    smtp_server.sendmail(email_from, email_to, msg.as_string())
    smtp_server.quit()
    
# function to send an email with the given message    
def send_emails(customerName, lineName,machineName,filler_Type,filler_Speed,filler_SW, duration, timeUpdate, LPO, Operator_Name, Operator_LASTENTRY, Product_Name, Product_LASTENTRY):
    ol = win32com.client.Dispatch('Outlook.Application')
    # size of the new email
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    fullSubject=' '.join ([email_subject, customerName, 'Line_', lineName])
    newmail.Subject = fullSubject
    newmail.To = email_to
    newmail.CC = email_cc
    newmail.Body ="Test should not appear"
    #<p>The Lost Production Opportunity so far are {LPO} packs.</p>
    #<p> The current rnning Product Name <strong> {Product_Name} </strong> update on HMI as of last entry on <strong>{Product_LASTENTRY} </strong>. </p>
    #<p> The current Operator Name <strong> {Operator_Name} </strong> update on HMI as of last entry on <strong> {Operator_LASTENTRY}</strong>.</p>
    #<a href=“tel:555-666-7777”>call us now</a>           
    html_body = f"""\
        <html>
            <body>
                <p> Dear Addressed, </p>
                
                <p> This message to inform you that your Line_ <strong>{lineName} </strong> {filler_Type} filler number <strong>{machineName} </strong> has been in breakdown for more than <strong> {duration} </strong> minutes.</p>
                <p>This is RC automated montoring system running over all <strong> {customerName}'s</strong> Lines and this notification about status update as of <strong>{timeUpdate}</strong>.</p>
                <p>The Lost Production Opportunity so far are <strong> {LPO} Packs </strong>. </p>
                <p> The current running Product Name as per HMI : <strong> {Product_Name} </strong>. </p>
                <p> The current User login on HMI Name is : <strong> {Operator_Name} </strong>. his contact is : <a href=“tel:971-588-224778”>971-588-224778</a></p>
                <p>Best regards,</p>
                <p> RC Robot</p>
             </body>
             
          </html>
          """
    newmail.HTMLBody=html_body
    #attach = 'C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
    #newmail.Attachments.Add(attach)
    # To display the mail before sending it
    # newmail.Display() 
    #########################
    #check if there is a mail body
    try:
        newmail.Subject = newmail.Subject
    except:
        newmail.Subject = 'No subject'

    #########################
    #check if there is a mail body
    try:
        newmail.HTMLBody = html_body
    except:
        newmail.HTMLBody = 'No email body defined'

    #########################
    #add first attachement if available
    try:
        newmail.Attachments.Add(attachment1)
    except:
        pass

    #########################
    #add second attachement if available
    try:
        newmail.Attachments.Add(attachment2)
    except:
        pass
    
    #########################
    #Send email
    newmail.Send()
    message= "The notification email has been send successfuly to "+ customerName+ " for line "+lineName
    write_logs(log_file_path, message)
        
while True:
    message= "Cycle of the program started "
    write_logs(log_file_path, message)
    ###########################
    #Connect to VPN
    subprocess.call("ConnectVPN.bat")
    # Ping VPN
    ping('93.122.84.23')
    for db in databases:
        # loop through each database and connect to it
        all_status = []
        # filter out rows where the datetime value is older than 120 minutes
        filtered_status = []
        #print('Database1')
        try:
            connection = cx_Oracle.connect(db['user'], db['password'], db['ip'] + '/' + db['service'])
            # get the current timestamp from the remote machine
            timestamp_rows = execute_query(connection, 'SELECT SYSTIMESTAMP FROM DUAL')
            remote_timestamp = timestamp_rows[0][0]
            
            #to be used as filter the machine was not in operation since long time (10 hr) to make sure status table is updated
            ten_hours_ago = remote_timestamp - datetime.timedelta(minutes=600)
            #Now we need to filter for the running machine which status remain more than 2 hr
            two_hours_ago = remote_timestamp - datetime.timedelta(minutes=120)
            two_minutes_ago = remote_timestamp - datetime.timedelta(minutes=4)
            two_minutes_ago_test = '.'.join(str(two_minutes_ago).split('.')[:1])
            print(two_minutes_ago)
            #SQL statement with filter non update machine for less than 10 hr (to make sure the machine is connected and breakdown more than 2hr
            #sql = 'SELECT PRLNAME,COMPATIBILITY_REF, COMMUNICATIONSTATE, DT, VAR_ERROR_REF FROM LOG_CURRENTOPSTATES JOIN STAT_MACHINE ON LOG_CURRENTOPSTATES.MACHINE_REF = STAT_MACHINE.ID WHERE LOG_CURRENTOPSTATES.DT > :ten_hours_ago_var and LOG_CURRENTOPSTATES.DT < :two_hours_ago_var'
            #SQL statement with filter non update machine for only more than 2 hr
            #sql = 'SELECT PRLNAME,COMPATIBILITY_REF, COMMUNICATIONSTATE, DT, VAR_ERROR_REF FROM LOG_CURRENTOPSTATES JOIN STAT_MACHINE ON LOG_CURRENTOPSTATES.MACHINE_REF = STAT_MACHINE.ID WHERE LOG_CURRENTOPSTATES.COMPATIBILITY_REF in (251,400,401,410,411,420,421,430,431,440,441) AND LOG_CURRENTOPSTATES.DT < :two_hours_ago_var and LOG_CURRENTOPSTATES.COMMUNICATIONSTATE=1'
            #SQL statement with filter non update machine for only more than 2 hr with User login 
            #sql = 'SELECT PRLNAME,COMPATIBILITY_REF, COMMUNICATIONSTATE, DT, VAR_ERROR_REF, LOG_USER.STRNAME as Operator_Name, LOG_USER.LASTENTRY as Operator_LastEntry_DT  FROM LOG_CURRENTOPSTATES JOIN STAT_MACHINE ON LOG_CURRENTOPSTATES.MACHINE_REF = STAT_MACHINE.ID JOIN LOG_USER ON LOG_CURRENTOPSTATES.MACHINE_REF = LOG_USER.STAT_MACHINE_REF WHERE LOG_CURRENTOPSTATES.COMPATIBILITY_REF in (251,400,401,410,411,420,421,430,431,440,441) AND LOG_CURRENTOPSTATES.DT < :two_hours_ago_var and LOG_CURRENTOPSTATES.COMMUNICATIONSTATE=1'
            #SQL statement with filter non update machine for only more than 2 hr with User login and Product Name 
            sql = 'SELECT PRLNAME,COMPATIBILITY_REF, COMMUNICATIONSTATE, DT, VAR_ERROR_REF, LOG_USER.STRNAME as Operator_Name, LOG_USER.LASTENTRY as Operator_LastEntry_DT, LOG_PRODUCT.STRNAME as Product_Name, LOG_PRODUCT.LASTENTRY as Product_LastEntry_DT  FROM LOG_CURRENTOPSTATES JOIN STAT_MACHINE ON LOG_CURRENTOPSTATES.MACHINE_REF = STAT_MACHINE.ID JOIN LOG_USER ON LOG_CURRENTOPSTATES.MACHINE_REF = LOG_USER.STAT_MACHINE_REF JOIN LOG_PRODUCT ON LOG_CURRENTOPSTATES.MACHINE_REF = LOG_PRODUCT.STAT_MACHINE_REF WHERE LOG_CURRENTOPSTATES.COMPATIBILITY_REF in (251,400,401,410,411,420,421,430,431,440,441) AND LOG_CURRENTOPSTATES.DT < :two_hours_ago_var and LOG_CURRENTOPSTATES.COMMUNICATIONSTATE=1'
            #status_rows = execute_query(connection, 'SELECT {} FROM {} JOIN {} ON {}.MACHINE_REF = {}.ID WHERE {}.DT > SYSDATE - 2/24'.format(', '.join(status_columns),status_table_name, machine_name_table_name, status_table_name, machine_name_table_name,status_table_name))
            #status_rows = execute_query(connection, 'SELECT {} FROM {} JOIN {} ON {}.MACHINE_REF = {}.ID WHERE {}.DT > :1'.format(', '.join(status_columns),status_table_name, machine_name_table_name, status_table_name, machine_name_table_name,status_table_name,[remote_timestamp - 2/24/60]))
            #status_rows = execute_query(connection, 'SELECT {} FROM {} JOIN {} ON {}.MACHINE_REF = {}.ID WHERE {}.DT > :1'.format(', '.join(status_columns),status_table_name, machine_name_table_name, status_table_name, machine_name_table_name,status_table_name,[two_minutes_ago]))
            # execute the query with the bind variable value
            cursor = connection.cursor()
            #To excute the SQl with tow filter 10 hr and 2 hr
            #cursor.execute(sql, ten_hours_ago_var=ten_hours_ago, two_hours_ago_var=two_hours_ago)
            #To excute the SQl with only 2 hr filter
            cursor.execute(sql,  two_hours_ago_var=two_hours_ago)
            # fetch the results
            status_rows= cursor.fetchall()
            print('Connection done for customer '+ db['customerName'])
            message= "Connection done for customer : "+ db['customerName'] + " With IP :" + db['ip']
            write_logs(log_file_path, message)
            for row in status_rows:
                
                all_status.append(row)
                #print(row)
                #print((datetime.now() - row[3]).total_seconds()/ 60)
                #if (two_minutes_ago - row[3]).total_seconds() / 60 >= 120 and 'cfa' in row[0].lower():
                if 'cfa' in row[0].lower() and 'track' not in row[0].lower():
                    filtered_status.append(row)
                
            connection.close()
        except Exception as e:
            print('Error connecting to database {}: {}'.format(db['ip'], str(e)))
            message= 'Error connecting to database {}: {}'.format(db['ip'], str(e)) +" for customer : " + db['customerName'] 
            write_logs(log_file_path, message)

        # send an email with the status information
        #message = 'Status update as of {}: \n{}'.format(datetime.datetime.now().strftime('%m/%d/%Y %I:%M %p'), filtered_status)
        # create a message with the filtered status rows
        if len(filtered_status) ==0:
            message= "There is no Breakdown exceed 120 min for all lines in " + db['customerName'] + " With IP :" + db['ip']
            write_logs(log_file_path, message)
        message = ''
        for row in filtered_status:
            #########################
            #check customer name
            try:
                customerName=db['customerName']
            except:
                customerName = 'No customer name defined'
            #########################
            #check machine name
            try:
                machineName= row[0]
            except:
                machineName= 'No machine name defined'
            #########################
            #check line name
            try:
                lineName=db[machineName]
            except:
                lineName= 'No line name defined'
            #########################
            #check filler Type
            try:
                filler_Type=db[lineName]
            except:
                filler_Type= 'No filler Type defined'
            #########################
            #check filler Speed
            try:
                filler_Speed=db[filler_Type]
            except:
                filler_Speed= 'No filler speed defined'
            #########################
            #check filler Speed
            try:
                filler_Speed=db[filler_Type]
            except:
                filler_Speed= 'No filler speed defined'
            #########################
            #check filler Software
            try:
                filler_SW=db['_'.join([lineName,filler_Type])]
            except:
                filler_SW= 'No filler speed defined'
            #########################
            #check Operator name
            try:
                Operator_Name=row[5]
            except:
                Operator_Name= 'No operator defined'
            #########################
            #check LOG_USER.LASTENTRY
            try:
                Operator_LASTENTRY=row[6]  - datetime.timedelta(minutes=4)
                Operator_LASTENTRY='.'.join(str(Operator_LASTENTRY).split('.')[:1])
            except:
                Operator_LASTENTRY= 'No operator LAST ENTRY defined'
            #########################
            #check Product name
            try:
                Product_Name=row[7]
            except:
                Product_Name= 'No Product defined'
            #########################
            #check LOG_Product.LASTENTRY
            try:
                Product_LASTENTRY=row[8]  - datetime.timedelta(minutes=4)
                Product_LASTENTRY= '.'.join(str(Product_LASTENTRY).split('.')[:1])
            except:
                Product_LASTENTRY= 'No Product LAST ENTRY defined'
            #########################
            #check downtime duration
            try:
                duration = int((two_minutes_ago - row[3]).total_seconds() / 60)
            except:
                duration = 'Duration not defined'
            #########################
            #check Lost Production Opportunity
            try:
                LPO = int((duration/60) * int(filler_Speed))
            except:
                LPO = 'LPO not defined'
            #print(row[0])
            #For now we will send one per line(Filler)
            message = ' {}\t{}\t{}\t{}\t{}\n'.format(two_minutes_ago_test, row[0], row[1], row[2], row[3])
            send_emails(customerName, lineName,machineName,filler_Type,filler_Speed,filler_SW, duration, two_minutes_ago_test, LPO, Operator_Name, Operator_LASTENTRY, Product_Name, Product_LASTENTRY)
        #if you want to Append all under same message do below command but remove the send outside the for loop
        #message += ' {}\t{}\t{}\t{}\n'.format(two_minutes_ago_test, row[0], row[1], row[2], row[3])
        #send_emails(db['customerName'], machineName, message, two_minutes_ago_test)
    message= "Process has been completed and waiting for the next cycle"
    write_logs(log_file_path, message)

    #################################
    #VPN Disconnect
    # set the path to the VPN client program
    vpn_path = r'C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe'
    # disconnect from the VPN
    cmd = f'"{vpn_path}" disconnect'
    subprocess.run(cmd, shell=True, check=True)

    # wait for 2 hours before repeating the process
    time.sleep(2 * 60 * 60)
