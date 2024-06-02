# Data-Transfer-VBA-Macro
This VBA macro automates the process of transferring data between two Excel workbooks. The script prompts the user to select a production file and specify a target column in the reporting workbook. It then searches for matching values between a specified column in the production file and a specified column in the reporting file. When a match is found, the macro transfers data from the production file to the reporting file, handling cases where the source cell might be empty by using the value from the next row.

**Features:**
* Prompts user to select the production file.
* Allows user to specify the target column for data transfer.
* Searches for matching values between the specified columns in both workbooks.
* Transfers data from the production file to the reporting file, with special handling for empty cells.
* Displays a message upon successful completion of the transfer.

**Use Case:**
This macro was created to solve a workplace problem where data was being manually copied between two files. With this automation, the process is streamlined, saving time and reducing errors.
