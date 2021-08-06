import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.message import EmailMessage
from email import encoders
import smtplib

# data = pd.read_excel('hackBiz_Gov2021_Response.xlsx', sheet_name='Form Responses 1', usecols=['email','filename','path'])     # ryujin  ganti nama file excell
data = pd.read_excel('coba_email_spam.xlsx', sheet_name='Sheet1', usecols=['email','filename','path'])                          ####### uji coba
df = pd.DataFrame(data)
name = df['email']
data_length = len(name)
for x in list(range(data_length)) :
    ###########################################################################################
    ## EMAIL TEMPLATE                                                                   #######
    ###########################################################################################
    # styling HTML for email message     ryujin  ganti isi message
    message = """      
              <!DOCTYPE html>
              <html>
                <head>
                  <title>HackBiz x HackGov 2021 Webinar Certificate</title>
                </head>
                <header> 
                <img src="'F:/python/WebScrapping/logo.png', 'rb'?w=640" alt="img" />
                </header>
                <body>
                """ + asdasda + """
                </body>
              <footer style="background-color: gray;"> 
              
                  <p style="text-align: center;">Contact us on :  
                                <a href="mailto:haaply.apps@gmail.com">
                                haaply.apps@gmail.com</a></p>
                  </footer>
        </body>
      </html>
    """
    ### Email Content ###
    email_from = 'dimasmaehendra@gmail.com'                 #### ryujin  ganti email sender
    email_to = df.loc[x, 'email']

    # email compost requirement
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to
    msg['Subject'] = 'HackBiz x HackGov 2021 Webinar Certificate'
    msg.attach(MIMEText(message, 'html')) 


    # Attach file to email
    filename = df.loc[x, 'filename']  # fill the name here
    path = df.loc[x, 'path']
    attachment = open(path, 'rb')

    p = MIMEBase('application', 'octet-stream')

    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(p)


    #connect to SMTP server
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()

    # Login to account
    s.login(email_from, "password") # (email, password) sender    ## ryujin  ganti password
    # open the link bellow
    # enable the checklist (make sure click checklist again!!)
    ## https://myaccount.google.com/lesssecureapps?pli=1&rapt=AEjHL4PUwBrlFHpq-rldpsrcg0iJwj3H_qmSFtcTzUxKSSpvUBPjs1KSeJi8O1ABnstwFUG5Dn-WGfxytp4jxaVlGUSNntBkNA

    text = msg.as_string()
    s.sendmail(email_from, email_to,text)

    print('email are sent :)' + (x))