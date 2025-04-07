# Excel VBA SQL & Email Automation

This repository contains an Excel VBA solution that performs the following tasks:

- **Connects to the current workbook** using an ADODB connection.
- **Executes a SQL query** to retrieve data from multiple Excel tables.
- **Transfers the query results** into a specified worksheet.
- **Sends an email** via Outlook with the SQL query details and workbook attachment.

## Overview

The VBA code in this project automates the process of extracting data from Excel tables by:
- Opening an ADODB connection to the active workbook.
- Dynamically constructing a SQL query by replacing placeholder table names with real table ranges from a specific worksheet.
- Outputting the retrieved data into a designated worksheet ("Planilha2") with headers and formatting.
- Creating and sending an email using Outlook, attaching the workbook, and including the executed SQL query in the email body.

## Features

- **Dynamic Table Handling:** Uses the Excel ListObjects to dynamically build table names for SQL execution.
- **Data Retrieval & Display:** Executes SQL queries to retrieve data and displays the results on a dedicated worksheet.
- **Email Integration:** Automates sending an email with the SQL query details and the workbook as an attachment.
- **Modular Design:** Separate subroutines and a function for connection setup, data retrieval, and email sending.

## Requirements

- Microsoft Excel with macro support.
- References set in VBA for:
  - **Microsoft ActiveX Data Objects Library** (for ADODB connections).
  - **Microsoft Outlook Object Library** (for email automation) – alternatively, late binding is used via `CreateObject("Outlook.Application")`.
- The Excel workbook must include:
  - A worksheet named **Planilha1** containing the tables (`Table_1`, `Table_2`, `Table_3`, `Table_4`, `Table_5`).
  - A worksheet named **Planilha2** where the query results will be pasted.
- Microsoft Access Database Engine (ACE) installed to use `Microsoft.ACE.OLEDB.12.0`.

## Setup & Installation

1. **Clone or Download the Repository:**  
   Clone the repository to your local machine or download it as a ZIP file.

2. **Open the Workbook:**  
   Open the provided Excel workbook (`prova_vba_e_sql_2020_Allec1.xltm`) that contains the VBA code.

3. **Configure References:**  
   - Open the VBA editor (ALT + F11).
   - Ensure the following references are enabled:
     - **Microsoft ActiveX Data Objects x.x Library**
     - **Microsoft Outlook xx.x Object Library** (if not using late binding)

4. **Adjust Parameters if Needed:**  
   - Verify the worksheet names (`Planilha1` and `Planilha2`) and table names (`Table_1` to `Table_5`) match those in your workbook.
   - Update the recipient email address in the `enviar_email` subroutine.

## Usage

1. **Data Extraction:**
   - Run the `Tarefa` subroutine. This will:
     - Open a connection to the workbook.
     - Execute the SQL query, dynamically replace table names with actual ranges.
     - Output the results in "Planilha2" with headers and auto-fitted columns.

2. **Email Sending:**
   - Run the `enviar_email` subroutine to:
     - Create an Outlook email.
     - Populate it with the SQL query and attachment (the workbook).
     - Send the email to the specified recipient.

## Code Structure

- **OpenConnection:**  
  Initializes and opens an ADODB connection to the active workbook.

- **Tarefa:**  
  Main subroutine that opens the connection, executes a SQL query, and pastes the resulting data into "Planilha2".

- **GetData:**  
  Function that:
  - Retrieves dynamic table names from "Planilha1".
  - Replaces placeholder table names in the SQL query with actual table range names.
  - Executes the SQL query and returns a recordset.

- **enviar_email:**  
  Subroutine that creates and sends an email via Outlook, attaching the workbook and including the SQL query in the body.

## Customization

- **SQL Query Modifications:**  
  You can modify the SQL query in the `Tarefa` subroutine to suit your data extraction needs.  
- **Email Customization:**  
  Change the recipient, subject, and body text in the `enviar_email` subroutine to match your requirements.

## Troubleshooting

- **Connection Errors:**  
  Ensure the ACE OLEDB provider is installed and that the connection string correctly points to the active workbook.
- **Table Name Mismatches:**  
  Verify that the ListObjects in "Planilha1" are named exactly as expected (`Table_1` to `Table_5`).
- **Outlook Issues:**  
  If the email is not sending, ensure Outlook is installed and configured correctly. You might need to adjust security settings to allow programmatic access.


## Author

Allec Terrezo  
*Email: xxxxx@email.com*

---

