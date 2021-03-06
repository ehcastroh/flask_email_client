## -------------------------------------------------------------------- ##
# Import Required Libs
from flask import Flask, render_template, url_for
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

# create instance of class '__name__' since we are using a single module
app = Flask(__name__)

# create sqlite db named 'myFlaskDB' stored in relative path
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///myFlaskDB.db'
# initialize database by passing app object defined above
db = SQLAlchemy(app)


# create class
class emailData(db.Model):
	# create visit id
	id = db.Column(db.Integer, primary_key=True)
	# set max-characters to 200 and require that field be completed
	email = db.Column(db.String(200), nullable = False)
	# nonaccessible column to keep track of submissions
	completed = db.Column(db.Integer, default = 0)
	# track date of email submission
	date_submitted = db.Column(db.DateTime, default=datetime.utcnow)

	# funciton to return notice of data entry
	def __repr__(self):
		# string formating converts the self.id value to a string using repr()
		return '<Task %r>' % self.id


## routing binds urls to functions via decorators 
@app.route('/')
def home():
	"""
	Function responds to browser URL (local Host) and renders template
	@home.html is a template for Jinja2
	"""
	return render_template('home.html')

# if running in stand alone, run the application
if __name__ == "__main__":
	# enable debbuging
	app.run(debug=True)
## -------------------------------------------------------------------- ##

