# ExcelVTEntegrasyonProjesi

This project provides data integration between Excel files and an SQL Server database using a Windows Forms application. The project performs data reading from an Excel file to the database and data writing from the database to an Excel file.

## Features

- **Write Data from Database to Excel:** Reads data from an SQL Server database and writes it to a new Excel file.
- **Read Data from Excel to Database:** Reads data from an existing Excel file and adds it to the SQL Server database.

## Technologies Used

- **C#**
- **Windows Forms**
- **Microsoft.Office.Interop.Excel**
- **SQL Server**

## Requirements

- .NET Framework
- Microsoft Excel
- SQL Server

## Installation

1. Clone this project to your local machine:

2. Open the project with Visual Studio.

3. Install the necessary dependencies.

4. Update the database connection string (`SqlConnection baglanti`) in the `Form1.cs` file with your own database details.

## Usage

### Write Data from Database to Excel

1. Run the application.
2. Click the "Read from DB" button.
3. Data from the database will be written to a new Excel file.

### Read Data from Excel to Database

1. Run the application.
2. Click the "Read from Excel" button.
3. Data from the `C:\Users\Mustafa\Desktop\Personel.xlsx` file will be transferred to the database.
