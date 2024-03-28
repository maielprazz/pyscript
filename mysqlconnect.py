import mysql.connector

# Replace these with your MySQL database connection details
host = "localhost"
user = "ismail"
password = "Wellings@24"
database = "db_ip"

# Establish a connection to the MySQL database
conn = mysql.connector.connect(
    host=host,
    user=user,
    password=password,
    database=database
)

# Create a cursor to interact with the database
cursor = conn.cursor()

try:
    # Example SELECT query
    query = """
    SELECT 	REPORT_NAME, Subject, GROUP_CONCAT(email_to order by email_to SEPARATOR ', ') AS email_to,
		    GROUP_CONCAT(email_cc order by email_cc SEPARATOR ', ') AS email_cc,
		    GROUP_CONCAT(email_bc order by email_bc SEPARATOR ', ') AS email_bc
    FROM REF_AUTOMAIL
    where REPORT_NAME = 'Daily Stock Report' 
    and status = 1
    group by REPORT_NAME, Subject
"""  
    # Execute the query
    cursor.execute(query)
    
    # Fetch all rows
    rows = cursor.fetchall()
    # Print the result
    for row in rows:
        print(row[1])

except mysql.connector.Error as err:
    print("Error: {}".format(err))

finally:
    # Close the cursor and connection
    cursor.close()
    conn.close()
