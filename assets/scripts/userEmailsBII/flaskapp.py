from flask import Flask
from flask import request
import time
import smtplib  
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

app = Flask(__name__)

@app.route('/')
def hello_world():
    return 'Hello World!'

@app.route('/', methods=['POST'])
def parse_request():
    data = request.form 
    print("Received request")
    print(data)    
    # Email parameters 
    recipient = request.form['to']
    subject = request.form['subject']
    body_text = request.form['body']
    body_html = request.form.get('html')
    name = request.form["name"]
    bcc = "admin@innovation-engineering.net"
    sender = "admin@innovation-engineering.net"
    
    # AWS setup
    username_smtp = "AKIAJLJPVFIC5ZOQNW6A"
    password_smtp = "AkK4P5QW6caVjJtGQZaa77CzThVs71hgykvBOBmmwvyg"
    host = "email-smtp.us-west-2.amazonaws.com"
    port = 587
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = email.utils.formataddr((name, sender))
    msg['To'] = recipient
    msg['Bcc'] = bcc
    part1 = MIMEText(body_text, 'plain')
    msg.attach(part1)
    part2 = MIMEText(body_html, 'html')
    msg.attach(part2)

    try:  
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        #stmplib docs recommend calling ehlo() before & after starttls()
        server.ehlo()
        server.login(username_smtp, password_smtp)
        server.send_message(msg)
        server.close()
    # Display an error message if something goes wrong.
    except Exception as e:
        response = "Post request received at " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()) + ". Email failed to send."
        print(response)
        print(e)
    else:
        response = "Post request received at " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()) + ". Email sent success!"
        print (response)
    return response 

@app.route('/instructors/', methods=['POST'])
def parse_instructor_request():
    data = request.form 
    print("Received instructor request")
    print(request.form)
    print(request.values)
    print(request.args) 
    response = "Instructor post request received at " + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
    return response

if __name__ == "__main__":
	app.run(threaded = True)
