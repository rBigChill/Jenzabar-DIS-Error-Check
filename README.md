DIS Import Log Monitoring Script

This PowerShell script automates the monitoring, error handling, and correction of issues related to the DIS Import Log process by interacting with SQL Server databases and system services.

Features
- Connects to SQL Server to retrieve DIS Import Log records.
- Detects and handles errors such as missing users or sections.
- Stops and starts critical services as part of the troubleshooting process.
- Deletes specific user and section records from the database.
- Reposts batch jobs in case of failures.
- Sends email notifications for important events or errors.

Prerequisites
- PowerShell (latest version recommended)
- SQL Server PowerShell module (`SqlServer`)
- SMTP server for email notifications
- Permissions to access and modify SQL Server databases
- Administrative access to start/stop services on the target machine

Installation
1. Ensure the `SqlServer` module is installed:
    ```powershell
    if (Get-Module -ListAvailable SqlServer) {
        Import-Module -Name SqlServer
    } else {
        Install-Module -Name SqlServer -Scope CurrentUser -Force
        Import-Module -Name SqlServer
    }
    ```
2. Configure the following script variables:
    ```powershell
    $Computer = "YourComputerName"
    $ErrorFileLocation = "PathToErrorLogs"
    
    $SqlServer = "YourSQLServer"
    $SqlDBName = "YourDatabase"
    $SchemaName = "YourSchema"
    $TableName = "YourTable"
    
    $smtpServer = "your.smtp.server"
    $fromEmail = "your-email@example.com"
    $toEmail = "recipient@example.com"
    $subject = "DIS Import Log Monitoring Alert"
    ```

Usage
1. Save the script as `DISImportMonitor.ps1`.
2. Run the script using PowerShell:
    ```powershell
    .\DISImportMonitor.ps1
    ```
3. The script will:
    - Retrieve the latest DIS Import Log entry from the database.
    - Analyze its status and handle errors if necessary.
    - Stop and restart services if required.
    - Parse log files to identify missing users or sections and delete them from the database.
    - Repost batch jobs to attempt issue resolution.
    - Send email alerts for detected issues or resolution failures.

Error Handling
- If an error occurs, the script will send an email with the error message.
- If the target computer is unreachable, a notification will be sent.
- If an issue persists after multiple retries, an alert is triggered with the failed log file location.

Customization
- Modify the `Send-Email` function to integrate with a different email provider.
- Adjust the `Parse-File` function to detect and handle additional error conditions.
- Change database queries as needed to fit your organization’s data structure.

License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

