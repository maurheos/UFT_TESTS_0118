dim p, ds, ic, uid, pwd
p="OraOLEDB.Oracle"
ds="localhost:1521/xe" 'SBMEDLAPBANAR7:1521/XE
ic="DBEASYCALCULATION"
uid="SYSTEM"
pwd="Sophos.2017"
	if not createDBHandler(p, ds, ic, uid, pwd) then
		Reporter.ReportEvent micFail, "DB Init", "Failed to create a DB Handler instance."
		exittest
	end if

'Open the connection
oDBHandler.openDBConnection()

'Here goes code to query the DB (see recipe Using SQL queries programmatically)
'...
'...

'Close the connection
oDBHandler.closeDBConnection()

'Dispose the DB Handler
set oDBHandler=nothing
