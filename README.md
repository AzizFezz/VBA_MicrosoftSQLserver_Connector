# Excel SQL Connector Application

The Excel SQL Connector Application is a VBA-based application that allows you to connect an Excel file to a Microsoft SQL Server and manipulate the data using a user-friendly form.

## Table Creation

Before using the application, make sure to create the required table in your Microsoft SQL Server database. Execute the following SQL statement to create the table:


    Create Table TBL_Customer(
        CustomerId int identity(1,1),
        CustomerName nvarchar(50),
        CustomerAddress nvarchar(100),
        MobileNumber int unique,
        EmailId nvarchar(50) Unique,
        UpdateTimeStamp datetime default GETDATE()
      )

  
## Configuration
To connect the application to your Microsoft SQL Server database, you need to update the connection string in the "myform" module code. Locate the following line in the code:

    Const Connection_String = "Provider=SQLOLEDB; Data Source=[Server_Name]; Initial Catalog=[Database_Name]; Trusted_Connection=yes;"
  
    
Replace [Server_Name] with the name of your SQL Server and [Database_Name] with the name of your database.

## Usage

1. Open the Excel file containing the application.

2. Ensure that the table is created in your SQL Server database as described in the "Table Creation" section.

3. Enable macros if prompted.

4. Navigate to the form interface provided by the application.

5. Use the form to interact with the data stored in the SQL Server table. You can perform operations such as adding new customers, updating existing customer details, and retrieving customer information.

6.  The application will establish a connection to the SQL Server database using the provided connection string and perform the necessary SQL queries to manipulate the data.

