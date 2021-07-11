import xlrd
import MySQLdb

# Open the workbook and define the worksheet
book = xlrd.open_workbook("pytest.xlsx")
sheet = book.sheet_by_name("source")

# Establish a MySQL connection
database = MySQLdb.connect (user= 'root',password='root',host= 'localhost',database='latihansqlalchemy')


# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()

# Create the INSERT INTO sql query
query = """INSERT INTO orders (product, customer_type, rep, date) VALUES (%s, %s, %s, %s)"""

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
database.close()

# Print results
print ("")
print ("All Done! Bye, for now.")
print ("")
print ("I just imported rows to MySQL!")