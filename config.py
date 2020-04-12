
import datetime
x = datetime.datetime.now()

YEAR = x.year
# Database Excel file
LOC = 'database.xls' # Do not Edit
LOGIN_URL = "http://nsmart.nhealth-asia.com/MTDPDB01/index.php" # Do not Edit
URL = "http://nsmart.nhealth-asia.com/MTDPDB01/reports/master_kpi.php?sp_tyear={}&sp_branchid=10024&sp_dept=1002401".format(YEAR) # Do not Edit
URL_ONLINE_REPAIR = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/BJOBA_01online.php" # Do not Edit
LINE_URL = "https://notify-api.line.me/api/notify" # Do not Edit
