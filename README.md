# Easy SQL Editor Tool for Windows #

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
