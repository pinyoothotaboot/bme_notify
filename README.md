How to use
1.	Install environment support 
1.1	Install Python from https://www.python.org/downloads/ 
-	Choose Python3 version 3.7 up to stable
-	Check version : Open terminal  console and Enter “python --version”
1.2	Install pip from guide https://www.liquidweb.com/kb/install-pip-windows/
-	Check version : Open terminal console and Enter “pip --version”
1.3	Install Libs from pip command 
-	Open terminal console and Enter “pip install -r requirements.txt”

2.	Open filename “config.json” for edit data
-	CSI_URL = URL Link from CSI form
-	LINE_TOKEN = 
o	Department name = example “Intensive Care Unit” if not Choose name of BME Site , Ex. SSH
o	Line Token = enter line notify token , we can create from https://medium.com/@nattaponsirikamonnet/%E0%B8%A1%E0%B8%B2%E0%B8%A5%E0%B8%AD%E0%B8%87-line-notify-%E0%B8%81%E0%B8%B1%E0%B8%99%E0%B9%80%E0%B8%96%E0%B8%AD%E0%B8%B0-%E0%B8%9E%E0%B8%B7%E0%B9%89%E0%B8%99%E0%B8%90%E0%B8%B2%E0%B8%99-65a7fc83d97f
-	USERNAME = Enter user to login N Smart
-	PASSWORD = Enter password to login N Smart
-	SITE_NAME = Enter name of Site BME ex. SSH

3.	Start process
-	Open terminal console and enter “python bme_notify.py”
