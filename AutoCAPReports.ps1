<# ******************************************************
   *                                                    *
   *                 AutoCAPReport                      *
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

<#
$cmd_command = '"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport AuditLogFieldChanges'
$dates =  $from_date +" "+ $to_date
$pm = '"C:\Reports" dmz-sql-01.dmz.local\pronav,59116 ATC_Archive sa Asi@16 /type:EXCEL /fields:"Voucher Type,Payment Method,Currency,Voucher Date"'

$cmd = $cmd_command + " " + $dates + " " + $pm
#>

# Variable decleration.
# Dates
$StartDate  = (Get-Date).AddDays(-1).ToString("dd.MM.yyyy")
$EndDate    = (Get-Date).ToString("dd.MM.yyyy")
$DateForAuditLogFc   = $StartDate + " " + $EndDate
$DateForTimeSpent    = $StartDate + " " + $StartDate

# Email addresses
$ToEmailAddressesList   = @("Sandaruwan Samaraweera<sandaruwan.s@pro-account.lk>","Heman Krishantha<heman.krishantha@pro-account.lk>","Priyantha Kumarasiri<priyantha.kumarasiri@pro-account.lk>","Sachini Madara<sachini.madara@pro-account.lk>","Nipun Karunarathna<nipun.karunarathna@pro-account.lk>")

# Executable path
$McCommand  = '"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport'

# Parametter values
$DataBaseAddress  = 'dmz-sql-01.dmz.local\pronav,59116'
$DataBaseTable    = 'ATC_Archive'
$DataBaseUser     = 'sa'
$DataBasePassword = 'Asi@16'
$ReportType       = '/type:EXCEL'
$UsedFields       = '/fields:"Voucher Type,Payment Method,Currency,Voucher Date"'

# Report names
$AuditLogFieldChanges   = 'AuditLogFieldChanges'
$TimeSpent              = 'TimeSpent'

# Report output folder
$OutPutFolder = "$($PWD)\Reports"

# Comands to run
# For AuditLogFieldChanges
$CMD_AuditLogFc = $McCommand + " " + $AuditLogFieldChanges + " " + $DateForAuditLogFc + " " + $OutPutFolder + " " + $DataBaseAddress + " " + $DataBaseTable + " " + $DataBaseUser + " " + $DataBasePassword + " " + $ReportType + " " + $UsedFields

# For TimeSpent
$CMD_TimeSpent  = $McCommand + " " + $DateForTimeSpent + " " + $TimeSpent + " " + $OutPutFolder + " " + $DataBaseAddress + " " + $DataBaseTable + " " + $DataBaseUser + " " + $DataBasePassword + " " + $ReportType

# TODO: use the file name check to see if theres an file which got created to day. If yes just send the email no need to run the command.
#To get the file name with out the extention.
#$AuditLogFcFileNames  =   @(Get-ChildItem -Path $OutPutFolder | Where-Object {($_.CreationTime -gt (Get-Date).Date -and $_.Name -match "AuditLogFieldChanges")} | ForEach-Object -Process {[System.IO.Path]::GetFileNameWithoutExtension($_)})

#To get the file name with the extention.
$AuditLogFcFileNames  =   @(Get-ChildItem -Path $OutPutFolder | Where-Object {($_.CreationTime -gt (Get-Date).Date -and $_.Name -match "AuditLogFieldChanges")} | ForEach-Object -Process {[System.IO.Path]::GetFileName($_)})

$AuditLogFcFileNameCount = $AuditLogFcFileNames.count

if ($AuditLogFcFileNameCount -gt 1){
   $AuditLogFcFileName = $AuditLogFcFileNames[0]
}else {
   $AuditLogFcFileName = $AuditLogFcFileNames
}

if ($AuditLogFcFileNameCount -gt 1) {
   # Send email with AuditLogFieldchanges report attached.
   foreach ($ToEmailAddress in $ToEmailAddressesList) {
      Send-MailMessage -From "no-reply@navitro.com" -To "$($ToEmailAddress)" -Subject "$($AuditLogFieldChanges) for $($StartDate)" -Body "Generated $($AuditLogFieldChanges) report for $($StartDate)." -Attachment "$($PWD)\Reports\$($AuditLogFcFileName)" -SmtpServer mx01.minedata.no -Port 25
      if ($?){
         Write-Host "Email sent ..."
      }
   } else {
      # Run the ManagementConsole using commandprompt uses the above variables which makes the command.
      cmd /c $CMD_AuditLogFc
   }
}


#cmd /c $CMD_TimeSpent

