# winservicemonitoring
Windows Service Monitoring and Notification

This script is to monitor a specific Windows event and trigger email notifications upon detection. It is intended to be used in conjunction with the Windows Task Scheduler, which will be configured to execute the script in response to the monitored event.

The script, by default, monitors the following event:

- Event ID: 1026
- Event Source: .NET Runtime
- Keyword: Xpo.Svc.Agent

How to Use

1. Setup: Copy `ServiceMonitor.exe` and `config-servicemonitor.txt` to a designated directory. Ensure that this directory remains unchanged.
   
2. Configuration: Edit `config-servicemonitor.txt`. You'll need a SMTP2GO account for email notifications.

3. SMTP2GO Setup: Create an SMTP2GO API key and template. Links for reference:
   - https://support.smtp2go.com/hc/en-gb/articles/20733554340249-API-Keys
   - https://support.smtp2go.com/hc/en-gb/articles/4402929434777-API-Templates

4. In a Command Prompt window (run as administrator), navigate to the selected folder and run the following command:
   .\ServiceMonitor.exe -setup 1
   
5. Verification: Confirm that a Windows task named `xpoSvcAgent_monitor` has been created.

How It Works

This script utilizes native Windows event management. The Windows task created above is triggered when a specified event (ID: 1026, Source: .NET Runtime, Keyword: Xpo.Svc.Agent) occurs. Upon triggering, the task sends an email notification using the SMTP2GO API.

System Impacts

- Database Access: This script does not access databases and has no impact on them.
- File Operations: It reads/writes three text files (configuration, log, last matched event ID) and does not involve any other file system I/O.
- Communication: The system utilizes standard HTTP POST requests to send emails through the SMTP2GO API.

Security Note

This script is intended for use within a trusted, internal, and non-shared environment. Credentials are not encrypted.

Copyright: Minghui Yu (myu@southarm.ca), 2024, South Arm Technology Services Ltd.
License: GPL-3
