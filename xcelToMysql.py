#!/usr/bin/python
import xlrd
import MySQLdb
import sys


# command line args
sheetName = str(sys.argv[1])
dataBase = str(sys.argv[2])
usr = str(sys.argv[3])
pd = str(sys.argv[4])

# Open the workbook and define the worksheet
book = xlrd.open_workbook(sheetName)
sheet = book.sheet_by_name("source")

# Establish a MySQL connection
database = MySQLdb.connect (host="localhost", user = usr, passwd = pd, db = dataBase)

# Get the cursor, which is used to traverse the database, line by line
cursor = database.cursor()

# Create the INSERT INTO sql query
query = """INSERT INTO orders (date,name,lab,height,width,background,charge,speedtype,comments) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
for r in range(1, sheet.nrows):
      date = sheet.cell(r,).value
      name = sheet.cell(r,1).value
      lab = sheet.cell(r,2).value
      height = sheet.cell(r,3).value
      width = sheet.cell(r,4).value
      background = sheet.cell(r,5).value
      charge = sheet.cell(r,6).value
      speedtype = sheet.cell(r,7).value
      comments = sheet.cell(r,8).value

      # Assign values from each row
      values = (date,name,lab,height,width,background,charge,speedtype,comments)

      # Execute sql Query
      cursor.execute(query, values)

# Close the cursor
cursor.close()

# Commit the transaction
database.commit()

# Close the database connection
database.close()

# Print results
print ""
print "All Done! Bye, for now."
print ""
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print "I just imported " %2B columns %2B " columns and " %2B rows %2B " rows to MySQL!"