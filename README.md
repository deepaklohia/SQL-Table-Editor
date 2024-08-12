# SQL Server Table Editor Tool for Windows #

## Youtube Guide : How to use SQL Server Table Editor / Manager Tool ##
[![SQL Table Data Editor](https://github.com/user-attachments/assets/8ed45e61-b59c-4c4d-87ec-7fb334cf3e9e)](https://youtu.be/aazDK1a4Nuk)

![image](https://github.com/user-attachments/assets/72c31527-49d7-4891-8c72-9bad3b4de4c3)

## Tool Features ##
- Export Table to Excel
- Duplicate a table
- Delete a table
- Insert Records in bulk
- Update records in bulk
- Delete records in bulk
- Export to excel and Upload data with changes in excel.

## How to Setup ##
### Software Requirements ## 
- MS SQL Server (Management Studio)
- Visual Studio (Optional for Setup Purpose / Amendments)

### Settings Requirements ###
#### 1. Install SQL Server and make changes below: ####
   - Open SQL Server Configuration Manager.
   - SQL Server Network Configuration > Protocols for MSSQLSERVER
     - Shared Memory ENABLED
     - Named Pipes ENABLED
     - TCP/IP ENABLED
       
   ![export error fix](https://github.com/user-attachments/assets/e4d5002c-811a-4c5f-b76b-1943c3035afa)

   - Restart SQL Server

   ![sql service start](https://github.com/user-attachments/assets/183fe6c9-a646-43ff-a38b-3862f27264a6)

#### 3. Install "Microsoft Access Database Engine 2016" ####
   - Search for "Microsoft Access Database Engine 2016 Redistributable"
   - Choose Download exe.(not X64)
   - Install using CMD as /quiet (shown below)

  ![MS Access DB Engine Guide](https://github.com/user-attachments/assets/0b4498f5-c410-469d-aa7d-d283069d83ac)

#### 3.Make changes in Visual Studio Settings for ADODB. ####
  - Open Project in VS Studio
  - Choose "ADODB" under references and hit properties.
  - Go to "Embed Interop Types" in properties and Choose "False"
    
  ![ADODB Setting](https://github.com/user-attachments/assets/c5cf2074-77d0-40d0-aab8-42d6619e88d2)

## Important Information ##
Works on .NET version 4.0
https://download.visualstudio.microsoft.com/download/pr/e2bef28a-75a0-4e9f-895b-c5d2fb864b1e/acd252d4ddaeda8db1a4829d8b2ec74be62ec40f1b806334ecd13aa216c999fb/vs_Enterprise.exe

Link for .NET Version:
https://learn.microsoft.com/en-us/visualstudio/releases/2019/history#installing-an-earlier-release
use version : 16.9.11

### $${\color{yellow}Powered \space by: Deepak  \space Lohia \space Automations}$$ ###

