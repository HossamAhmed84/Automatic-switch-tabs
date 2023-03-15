from twilio.rest import Client
# Your Twilio Account SID and Auth Token
#account_sid = 'your_account_sid'
#auth_token = 'your_auth_token'
account_sid = 'AC5ae7efa99c27b94a2bbd9202f9c862b6'
auth_token = '080addc42c3b25146735118d882dcaf9'

# Your Twilio WhatsApp sandbox number
whatsapp_number = 'whatsapp:+14155238886'

# Your personal phone number for receiving SMS and WhatsApp messages
# The list of recipients
#to_numbers = ['+15555555555', '+16666666666']
#to_number = 'your_phone_number'
to_number = '+971588224778'

#lineName = 'L9'
lineName = 'U27'
filler_Type= 'CFA124-36'
#customerName= 'Juhayna'
customerName= 'Almarai'
#machineName = 'CFA873051118'
machineName = 'CFA873051058'
#Product_Name = 'Cocktail Classic 235ml'
Product_Name = '200 ML Double Choco Milk '
Operator_Name = 'SIG Global Service' 
#duration = 132
duration = 227
#LPO= 52800
LPO= 90800
timeUpdate= '2023-03-10 07:20:19'
message= '<h1>Hello, world!</h1><p>This is some HTML content.</p>'
message2 = 'This message to inform you that line '+ lineName + ' '+ filler_Type + 'in '+ customerName + ' has been in breakdown for more than ' + str(duration) + ' minutes. The Lost Production Opportunity so far are ' + str(LPO) + ' Packs'
message3 = 'This message to inform you that line '+ lineName + ' '+ filler_Type + 'in '+ customerName + ' has track/s switched-off for more than ' + str(duration) + ' minutes. The Lost Production Opportunity so far are ' + str(LPO) + ' Packs'

html_body = f"""\
        <html>
            <body>
                <p> Dear Addressed, </p>
                
                <p> This message to inform you that <strong>{customerName} </strong> Line_ <strong>{lineName} </strong> {filler_Type} filler number <strong>{machineName} </strong> has been in breakdown for more than <strong> {duration} </strong> minutes.</p>
                <p>This is RC automated montoring system running over all <strong> {customerName}'s</strong> Lines and this notification about status update as of <strong>{timeUpdate}</strong>.</p>
                <p>The Lost Production Opportunity so far are <strong> {LPO} Packs </strong>. </p>
                <p> The current running Product Name as per HMI : <strong> {Product_Name} </strong>. </p>
                <p> The current User login on HMI Name is : <strong> {Operator_Name} </strong>. his contact is : <a href=“tel:971-588-224778”>971-588-224778</a></p>
                <p>Best regards,</p>
                <p> RC Robot</p>
             </body>
             
          </html>
          """
# Create a Twilio client
client = Client(account_sid, auth_token)

# Send an SMS message
try:
    sms_message = client.messages.create(
        #body='Hello from Twilio!',
        body=message3,
        from_='+15673501525', # Your Twilio phone number
        to=to_number,
        #media_url=['data:text/html;charset=utf-8,' + message],
        provide_feedback=True
    )
except Exception as e:
    print(f"Error sending message: {e}")

print(f'SMS message sent to {to_number}: {sms_message.sid}')

# Send the WhatsApp message
try:
    whatsapp_message = client.messages.create(
        to='whatsapp:'+to_number,
        from_=whatsapp_number,
        body=message3,
        #media_url=['data:text/html;charset=utf-8,' + message],
        provide_feedback=True
    )
    print(f"Message SID: {whatsapp_message.sid}")
except Exception as e:
    print(f"Error sending message: {e}")


print(f'WhatsApp message sent to {to_number}: {whatsapp_message.sid}')
