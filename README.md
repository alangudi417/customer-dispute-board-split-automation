# Customer Dispute Board-Split Automation

## 📊 Project Overview
This project automates the distribution of the customer dispute information, by transforming a centralized dispute-databse into region-specific and customer-specific workbooks using Microsoft Excel VBA. 

The automation reads a master dispute file, enriches the data with dispute notes from a secondary workbook, separates records by region, and generates individual Excel files for each customer account. The resulting files are automatically organized into regional folders, eliminating a repetitive manual process and significantly improving reporting efficiency.

This solution was designed to reduce processing time, improve consistency, and minimize human error in dispute management workflows.

## 📂 Data Sources
The automation processes two Excel workbooks:

    UDM_Dispute.xlsm
    - This the primary dispute database containing customer and invoice information. The most relevant columns are:
        - AOR "Customer Region"
        - Days 
        - Case ID "Dispute Number"
        - External Reference "Purchase Order"
        - Sales Order
        - Case Title "Invoice Number"
        - Inv Date
        - Customer Number
        - Customer Name
        - Cause Desc
        - Disputed Amount
        - Processor Name "Dispute Responsible"

    UDM_Notes.xlsx
    - this secondary database contains additional dispute notes and related info, just the dispute notes are relevant to be matched against the master dataset, using a unique identifier "Dispute Number (Case ID)"

## ⚙️ Automation Workflow
This project follows a fully automated Excel Workflow. 
- First Macro "UDM Comments":
    1. Load Source Files. 
    - Opens the master dispute workbook
    - Opens the notes workbook
    - Validates both files before processing

    2. Worksheet creation inside master workbook.
    - Creates a worksheet named "Main"
    - Reorganize the original dataset removing "no needed columns"
    - Change the headers of some columns
    - organize them in relevant order:
        - AOR "Customer Region"
        - Days
        - Dispute
        - PO
        - Sales Order
        - Invoice
        - Inv Date
        - Customer Number
        - Customer Name
        - Cause Desc
        - Disputed Amount
        - Owner "area in charge"
        - Processor "id"
        - Processor Name
        - Manager Name
        - Dispute Status
        - UDM Notes

    3. Merge Dispute Notes
        - Matches disptue records using a unique identifier "Case ID"
        - Appends notes to the corresponding dispute
        - Preserves the integrity of the original dataset
        - Fill column "Dispute Status" with a formula, highlighting if the dispute is in working status "In Process", or not worked yet "New"

- Second Macro "Split by Customer":
    4. Split Records by Region "AOR"
        - The automation identifies every region present in the dataset and creates dedicated worksheets for each one. 
        - Example: 
            - North America
            - Latin
            - Africa
        - Each worksheet contains only the disputes and customers assigned to that specific region

    5. Generate Customer Workbooks
        - Within each regional worksheet, the macro:
            - Identifies every unique customer account
            - Creates a new workbook
            - Copies the corresponding dispute records
            - Preserves headers and formatting
            - Applies a standardized file naming convention
        - Example:
            - North America/<br>
            |<br>
            ├── customer_1.xlsx <br>
            ├── customer_2.xlsx
            
            - Africa/ <br>
            |<br>
            ├── customer_3.xlsx <br>
            ├── customer_4.xlsx <br>

            - Latam/ <br>
            |<br>
            ├── customer_5.xlsx <br>
    
    6. Organize Output
        - The generated workbooks are automatically saved into folders corresponding to their assigned region, making them immediately available for distribution to regional dispute teams. 

## 📈 Key Insights
- Automatic opening of source workbooks
- Merging dispute notes from multiple files
- Record matching using unique identifiers
- Dynamic worksheet creation by region
- Generation of one workbook per customer account
- Standardized file naming
- Automatic folder organization
- End-to-end workflow automation
- Minimal user interaction required

## 🛠️ Technologies Used
- Microsoft Excel VBA
- Excel Object Model
- Workbook & Worksheet Automation
- AutoFilter
- Dynamic Range Management
- VBA Dictionaries
- File System Operations
- Excel File Generation

## 🔄 Workflow
Open UDM_Dispute.xlsm

        |
        ▼

Open UDM_Notes.xlsx

        |
        ▼

Worksheet creation inside Master Workbook

        |
        ▼

Match and Merge Notes

        |
        ▼

Split by Region

        |
        ▼

Create Regional Worksheets

        |
        ▼

Identify Customer Accounts in each Worksheet

        |
        ▼

Generate individual Files

        |
        ▼

Save to Regional Folders

NOTE: the sample files included in this repository contain anonymized data to protect confidential business information.

## 📈 Project Highlights
- Automated a previously manual reporting process, dropping the time up to 94%
- Eliminated repetitive Excel tasks
- Improved consistency across generated reports
- Reduced the risk of manual human processing errors
- Organized output into a standardized folder structure
- Scalable to thousands of dispute records

## 💼 Business Impact
Before automation, generating customer dispute reports required manually filtering data, copying information between workbooks, creating dozens of files, and organizing them into regional folders. This VBA solution transformed the workflow into a single automated process that:

- Reduced report generation time (up to 94% of the time)
- Standardized dispute reporting
- Improved data accuracy
- Increased operational efficiency
- Allowed to focus on higher-value tasks instead of repetitive administrative work

## ▶️ How to Run the Project
1. git clone https://github.com/alangudi417/customer-dispute-board-split-automation.git
2. Open UDM Generator.xlsm (Excel workbook where the VBA projects live)
3. Update the file paths for:
    - UDM_Dispute.xlsm
    - UDM_Notes.xlsx
4. Run the first macro. This automation will:
    - Load Source Files.
    - Create Worksheet inside master workbook
    - Merge Dispute Notes
5. Run the second macro. This automation will:
    - Split Records by Region "AOR"
    - Generate Customer Workbooks
    - Organize Output