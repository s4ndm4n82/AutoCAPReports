<# ******************************************************
   *                                                    *
   *            Individual Reports Script               *
   *                                                    *
   *  Created by: Sandaruwan Samaraweera                *
   *  Version: 1.0.1                                    *
   *                                                    *
   ****************************************************** #>

# Report names
# ~~~~~~~~~~~~
# ItemsPerJob
# ItemsPerJobDetailed
# ImportedExported
# GroupByFields
# GroupByFieldsJobColumns
# ProcessingTime
# AuditLogFieldChanges
# AuditLogCount
# TimeSpent

# Below option can be passed into ManaagementConsole when used through the command prompt.
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#  /CreateReport: (Report name to generate.)
#  /j:            (Pass just one job name.)
#  /jobs:         (Pass more than one job name.)
#  /folder:       (Pass a set of jobs with in a folder in CAP job list.)
#  /fields:       (Specify what fields to use.)
#  /type:         (Type of the report.)
#  /dbserver:     (Database server name.)
#  /dbname:       (Database table name.)
#  /dbusername:   (Database user name.)
#  /dbpassword:   (Database user password.)
#  /format:       (Out put file type. If not specified, default will be PDF. Other supported formats are excel and word.)

# Examples for usage:
#  "C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport ItemsPerJob %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL
#  "C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport TimeSpent %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL
#  "C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport AuditLogFieldChanges %FromDate% %ToDate% "C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL /fields:"Voucher Type,Payment Method,Currency,Voucher Date"