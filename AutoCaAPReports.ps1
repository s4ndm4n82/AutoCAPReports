

#"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport ItemsPerJob %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL
#"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport TimeSpent %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL
#"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport AuditLogFieldChanges %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL /fields:"Voucher Type,Payment Method,Currency,Voucher Date"

$cmd_command = '"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport AuditLogFieldChanges'
$dates =  $from_date +" "+ $to_date
$pm = '"C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL /fields:"Voucher Type,Payment Method,Currency,Voucher Date"'

$cmd = $cmd_command + " " + $dates + " " + $pm