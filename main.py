import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

data = p.read_excel("giz_result.xlsx")
# print(type(data))
email_col = data.get("EMAIL")
name_col = data.get("NAME")
theory_col = data.get("THEORY")
practical1_col = data.get("PRACTICAL1")
practical2_col = data.get("PRACTICAL2")
practical3_col = data.get("PRACTICAL3")
practical4_col = data.get("PRACTICAL4")
practical5_col = data.get("PRACTICAL5")
total_col = data.get("TOTAL")
remarks_col = data.get("REMARKS")
list_of_emails = list(email_col)
list_of_name = list(name_col)
list_of_theory = list(theory_col)
list_of_practical1 = list(practical1_col)
list_of_practical2 = list(practical2_col)
list_of_practical3 = list(practical3_col)
list_of_practical4 = list(practical4_col)
list_of_practical5 = list(practical5_col)
list_of_total = list(total_col)
list_of_remarks = list(remarks_col)
print(list_of_emails)
print(list_of_name)
print(list_of_practical1)
print(list_of_practical2)
print(list_of_practical3)
print(list_of_practical4)
print(list_of_practical5)
print(list_of_total)
print(list_of_remarks)
# print(email_col)
plain1 = "TESTING THIS SYSTEM"
t1 = t2 = t3 = t4 = t5 = t6 = 12
t7 = 81
theory_score = 23
practical1 = 13
practical2 = 14
practical3 = 7
practical4 = 7
practical5 = 14
total_score = 89
name = "CHINEDU ERNEST"
rmks = "EXCELLENT"
msg = '''
THEORY: %s ,
PRACTICAL 1: %s ,
PRACTICAL 2: %s ,
PRACTICAL 3: %s ,
PRACTICAL 4: %s ,
PRACTICAL 5: %s ,
TOTAL:( %s )
''' % (t1, t2, t3, t4, t5, t6, t7)

for que in range(len(list_of_name)):
    print(que)
    try:
        server = sm.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login("edutech4logic@gmail.com", "kdacthrpswbcckhi")
        from_ = "edutech4logic@gmail.com"
        to_ = list_of_emails
        message = MIMEMultipart("alternative")
        message['Subject'] = "JUNIOR ROBOTICS EXAM SCORE"
        message['from'] = "edutech4logic@gmail.com"
        # theory_score = 28
        html = f'''
        <html>
        <head>
        
        </head>
        
        <body>
            <div style="text-align:center; border:6px dotted purple;">
            <img src="cid:banner_logo.jpg">
            <u><h2>JUNIOR ROBOTICS EXAMINATION RESULT</h2></u>	
            <h3> NAME: {list_of_name[que]}.</h2>
        <table align="center" border="2px" style="width:50%;">
          <caption>EXAMINATION SCORES</caption>
          <tr>
            <th>SESSIONS</th>
            <th>SCORES</th>
          </tr>
          <tr>
            <td>THEORY(30%)</td>
            <td>{list_of_theory[que]}</td>
          </tr>
          <tr>
            <td>PRACTICAL 1(14%)</td>
            <td>{list_of_practical1[que]}</td>
          </tr>
           <tr>
            <td>PRACTICAL 2(14%)</td>
            <td>{list_of_practical2[que]}</td>
          </tr>
          <tr>
            <td>PRACTICAL 3(14%)</td>
            <td>{list_of_practical3[que]}</td>
          </tr>
           <tr>
            <td>PRACTICAL 4(14%)</td>
            <td>{list_of_practical4[que]}</td>
          </tr>
          <tr>
            <td>PRACTICAL 5(14%)</td>
            <td>{list_of_practical5[que]}</td>
          </tr>
           <tr>
          
            <td><h2><b>TOTAL SCORE(100%):</b></h2></td>
            <td><h2><b>{list_of_total[que]}</b></h2></td>
            
          </tr>
        </table>
        <p></p>
        <h2>REMARKS: {list_of_remarks[que]}</h2>
        <p></p>
        </div>	
        
        </body>
        
        </html>     
    '''
        # msgAlternative = MIMEMultipart('alternative')
        # message.attach(msgAlternative)
        text = MIMEText(html,"html")
        # plain2 = MIMEText(msg, "plain")
        # text = text.format(12, 23, 23)
        # message.attach(plain2)
        # msgText = MIMEText('<b>Some <i>HTML</i> text</b> and an image.<br><img src="cid:image1"><br>KPI-DATA!', 'html')
        # msgAlternative.attach(msgText)
        # fp = open('banner_logo.jpg', 'rb')  # Read image
        # msgImage = MIMEImage(fp.read())
        # fp.close()

        # Define the image's ID as referenced above
        # msgImage.add_header('Content-ID', '<image1>')
        # message.attach(msgImage)

        # msg = plain1+text
        message.attach(text)
        # mail.sendmail('from', 'to', html.format(score1="20" score2="10", score3="30"))
        # server.sendmail(from_, to_, message.as_string())
        server.sendmail(from_, list_of_emails[que], message.as_string())
        print("mail sent successfully")
        server.quit()
    except Exception as e:
        print(e)
