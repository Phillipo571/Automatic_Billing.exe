# BillingMaster Project Overview
Hello! This project is a billing automation program developed in Python with the assistance of ChatGPT.

The tool significantly reduced the time required to prepare customer billing reports — from an average of 1 hour to just 5 minutes. It is integrated with Microsoft Outlook, enabling the program to automatically attach generated billing files and dynamically populate the email body using custom templates for each customer.

Please note that the uploaded code currently contains placeholder customer names and redacted sensitive content. As such, it is not immediately usable in a production environment. I plan to update the repository with detailed comments and improved structure when time permits.

The core of the application is written in Python, and it generates billing reports in Excel format.
Some parts of the code include Korean text, as the reports were originally created for Korean clients. You may adjust the language as needed to fit the context of your region or customers.

In addition, billing reports vary depending on each customer's contract type — such as CSP, EA, or markup agreements — so the logic has been customized individually per client.

When you finish edit your code, with your logo file(Extension ico) run this code for make .exe<br>
pyinstaller --onefile --noconsole --icon (icon file name) --add-data ("source_path;destination_path") --name ("FileName") (BuildCodefileName)

Thank you!