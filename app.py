from flask import Flask, render_template, request, jsonify, json, redirect, url_for
from flask_sqlalchemy import SQLAlchemy  
from wtforms import SelectField
from flask_migrate import Migrate
import xlrd
import MySQLdb

# Establish a MySQL connection
database = MySQLdb.connect (user= 'root',password='root',host= 'localhost',database='latihansqlalchemy')

app = Flask(__name__)
app.config['SECRET_KEY'] = 'sdfsfsfevsrfsf'
app.config['SQLALCHEMY_DATABASE_URI']='mysql+pymysql://root:root@localhost:3306/latihansqlalchemy'

db = SQLAlchemy(app)
migrate = Migrate(app, db)

class orders(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product = db.Column(db.String(300))
    customer = db.Column(db.String(300))
    rep = db.Column(db.String(300))
    date = db.Column(db.Integer)


@app.route("/", methods=['GET','POST'])
def index():

	return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():

	if request.method == 'POST':
		fileuploaded = request.form['excelfile']
		
		# Open the workbook and define the worksheet
		book = xlrd.open_workbook(fileuploaded)
		sheet = book.sheet_by_name("source")

		# Get the cursor, which is used to traverse the database, line by line
		cursor = database.cursor()

		# Create the INSERT INTO sql query
		query = """INSERT INTO orders (product, customer, rep, date) VALUES (%s, %s, %s, %s)"""

		# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
		for r in range(1, sheet.nrows):
			product		= sheet.cell(r,0).value
			customer	= sheet.cell(r,1).value
			rep			= sheet.cell(r,2).value
			date		= sheet.cell(r,3).value


			# Assign values from each row
			values = (product, customer, rep, date)

			# Execute sql Query
			cursor.execute(query, values)

		# Close the cursor
		cursor.close()

		# Commit the transaction
		database.commit()

		# Close the database connection
		# database.close()

	return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)