<# ******************************************************
   *                                                    *
   *                 AutoCAPReport                      *
   *                                                    *
   *  Created by: Sandaruwan Samaraweera                *
   *  Version: 1.0.2                                    *
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
#Directory locations
$WorkingDirectory = "E:\Tools\AutoCAPReports" # Directory location of the script. Please set this path if not script wont work corrrectly.

# Active computer name.
$ComputerName = [Environment]::MachineName

# Dates
$StartDate  = (Get-Date).AddDays(-1).ToString("dd.MM.yyyy")
$EndDate    = (Get-Date).ToString("dd.MM.yyyy")
$DateForAuditLogFc   = $StartDate + " " + $EndDate
$DateForTimeSpent    = $StartDate + " " + $StartDate

# Email addresses
$ToEmailAddressesList   = @("Heman Krishantha<heman.krishantha@pro-account.lk>","Priyantha Kumarasiri<priyantha.kumarasiri@pro-account.lk>","Sachini Madara<sachini.madara@pro-account.lk>","Nipun Karunarathna<nipun.karunarathna@pro-account.lk>")
$ErrorMailgList = @("Sandaruwan Samaraweera<sandaruwan.s@pro-account.lk>","Andreas Moen Arnet<andreas.moen.arnet@digitalcapture.no>","Fredrik Skjellum<fredrik.skjellum@digitalcapture.no>")

# Executable path
$McCommand  = '"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport'

# Parametter values
$DataBaseAddress  = '<Your DB Address>'
$DataBaseTable    = '<Your table name>'
$DataBaseUser     = '<Your database user name>'
$DataBasePassword = '<Your data base password>'
$ReportType       = '/type:EXCEL'
$UsedFields       = '/fields:"Voucher Type,Payment Method,Currentcy,Voucher Date,Due Date,Voucher Number,Bank Account,KID,Mva Amount,Total Invoice Amount Mva,Total Invoice Amount Excl Mva,Rounding Amount,Prepaid Amount,Supplier Organization No,Supplier Name,Supplier Address,Supplier City,Supplier Postal Code,Supplier Country,Customer Organization No,Customer Name,Customer Address,Customer City,Customer Country,Invoice Reference,Contract Document,Accounting Cost,Order Reference,Supplier Reference,Customer Reference,Delivery Date,Delivery Address,Delivery City,Delivery Postal Code,Delivery Country,Product No,Description,Mva,Unit,Unit Price,Quantity,Discount,Unit Price Incl Discount,Net Amount,Amount Incl Mva"'

# Report names
$AuditLogFieldChanges   = 'AuditLogFieldChanges'
$TimeSpent              = 'TimeSpent'

# Report output folder
$OutPutFolder = "$($WorkingDirectory)\Reports"

# Log file
$ToDaysDate = Get-Date -Format "dd-MM-yyyy"
$LogFileName = "ACR_$($ToDaysDate).log"
$LogFileSavePath = "$($WorkingDirectory)\Logs\$($LogFileName)"

# Checks the log file
if (!(Test-Path $LogFileSavePath)){
   New-Item -ItemType File -Path $LogFileSavePath | Out-Null
}

#Functions
function Write-Log($Message){
   $TimeStamp = (Get-Date).ToString("HH:mm:ss")
   "$($ToDaysDate)::$($TimeStamp)::$($Message)" >> $LogFileSavePath
}

function Get-AuditLogFcFileName{
   #To get the file name with the extention.
   $AuditLogFcFileNames    =  @(Get-ChildItem -Path $OutPutFolder | Where-Object {($_.CreationTime -gt (Get-Date).Date -and $_.Name -match "AuditLogFieldChanges")} | ForEach-Object -Process {[System.IO.Path]::GetFileName($_)})

   if ($AuditLogFcFileNames){
      # Count the filenames in the array.
      $AuditLogFcFileNameCount = $AuditLogFcFileNames.count
      
      # Assignes the first file name if there's more than one similar file names in the array.
      if ($AuditLogFcFileNameCount -gt 1){
         $AuditLogFcFileName = $AuditLogFcFileNames[0]
      } else {
         $AuditLogFcFileName = $AuditLogFcFileNames
      }
      # Returns all the values. To be used by the file gen. and email sending codes.
      return $AuditLogFcFileName, $AuditLogFcFileNames, $AuditLogFcFileNameCount
   }   
}

function Get-TimeSpentFileName{
   $TimeSpentFileNames     =  @(Get-ChildItem -Path $OutPutFolder | Where-Object {($_.CreationTime -gt (Get-Date).Date -and $_.Name -match "TimeSpent")} | ForEach-Object -Process {[System.IO.Path]::GetFileName($_)})

   if ($TimeSpentFileNames){
      # Count the filenames in the array.
      $TimeSpentFileNameCount = $TimeSpentFileNames.count

      # Assignes the first file name if there's more than one similar file names in the array.
      if ($TimeSpentFileNameCount -gt 1){
         $TimeSpentFileName = $TimeSpentFileNames[0]
      } else {
         $TimeSpentFileName = $TimeSpentFileNames
      }
      # Returns all the values. To be used by the file gen. and email sending codes.
      return $TimeSpentFileName, $TimeSpentFileNames, $TimeSpentFileNameCount
   }
}

# Comands to run
# For AuditLogFieldChanges
$CMD_AuditLogFc = $McCommand + " " + $AuditLogFieldChanges + " " + $DateForAuditLogFc + " " + $OutPutFolder + " " + $DataBaseAddress + " " + $DataBaseTable + " " + $DataBaseUser + " " + $DataBasePassword + " " + $ReportType + " " + $UsedFields

# For TimeSpent
$CMD_TimeSpent  = $McCommand + " " + $TimeSpent + " " + $DateForTimeSpent + " " + $OutPutFolder + " " + $DataBaseAddress + " " + $DataBaseTable + " " + $DataBaseUser + " " + $DataBasePassword + " " + $ReportType

#To get the file name with out the extention.
#$AuditLogFcFileNames  =   @(Get-ChildItem -Path $OutPutFolder | Where-Object {($_.CreationTime -gt (Get-Date).Date -and $_.Name -match "AuditLogFieldChanges")} | ForEach-Object -Process {[System.IO.Path]::GetFileNameWithoutExtension($_)})

# Get the generated file details for AuditLogFieldChanges report.
$AuditLogFcFileName, $AuditLogFcFileNames, $AuditLogFcFileNameCount = Get-AuditLogFcFileName

if ($AuditLogFcFileNameCount -gt 0) {

   Write-Log "File $($AuditLogFcFileName) already exists ... Starting to send emails."

   # Send email with AuditLogFieldchanges report attached.
   foreach ($ToEmailAddress in $ToEmailAddressesList) {
      Send-MailMessage -From "no-reply@navitro.com" -To "$($ToEmailAddress)" -Subject "$($AuditLogFieldChanges) for $($StartDate) on $($ComputerName)" -Body "Generated $($AuditLogFieldChanges) report on $($ComputerName) for $($StartDate)." -Attachment "$($WorkingDirectory)\Reports\$($AuditLogFcFileName)" -SmtpServer mx01.minedata.no -Port 25
      if ($?){
         Write-Log "Email sent to $($ToEmailAddress)..."
      }
   }
} else {
   
   Write-Log "File not found ... running command to generate the file."

   # Run the ManagementConsole using commandprompt uses the above variables which makes the command.
   cmd /c $CMD_AuditLogFc

   if ($?){
      #Get the generated file name.
      $AuditLogFcFileName, $AuditLogFcFileNames, $AuditLogFcFileNameCount = Get-AuditLogFcFileName

      if (Test-Path $WorkingDirectory\Reports\$AuditLogFcFileName){
         Write-Log "File Generated ... $($AuditLogFcFileName)."
         
         # Send email with AuditLogFieldchanges report attached.
         foreach ($ToEmailAddress in $ToEmailAddressesList) {
            Send-MailMessage -From "no-reply@navitro.com" -To "$($ToEmailAddress)" -Subject "$($AuditLogFieldChanges) for $($StartDate) on $($ComputerName)" -Body "Generated $($AuditLogFieldChanges) report on $($ComputerName) for $($StartDate)." -Attachment "$($WorkingDirectory)\Reports\$($AuditLogFcFileName)" -SmtpServer mx01.minedata.no -Port 25
               if ($?){
                  Write-Log "Email sent to $($ToEmailAddress) ..."
               }
         }
      } else {
         Write-Log "Report file was not found."
      }
   }
}

# Get the genereated file details for TimeSpent report.
$TimeSpentFileName, $TimeSpentFileNames, $TimeSpentFileNameCount = Get-TimeSpentFileName

if ($TimeSpentFileNameCount -gt 0) {

   Write-Log "File $($TimeSpentFileName) already exists ... Starting to send emails."

   # Send email with AuditLogFieldchanges report attached.
   foreach ($ToEmailAddress in $ToEmailAddressesList) {
      Send-MailMessage -From "no-reply@navitro.com" -To "$($ToEmailAddress)" -Subject "$($TimeSpent) for $($StartDate) on $($ComputerName)" -Body "Generated $($TimeSpent) report on $($ComputerName) for $($StartDate)." -Attachment "$($WorkingDirectory)\Reports\$($TimeSpentFileName)" -SmtpServer mx01.minedata.no -Port 25
      if ($?){
         Write-Log "Email sent to $($ToEmailAddress) ..."
      }
   }
} else {

   Write-Log "File not found ... running command to generate the file."

   # Run the ManagementConsole using commandprompt uses the above variables which makes the command.   
   cmd /c $CMD_TimeSpent

   if ($?){
      #Get the generated file name.
      $TimeSpentFileName, $TimeSpentFileNames, $TimeSpentFileNameCount = Get-TimeSpentFileName

      if (Test-Path $WorkingDirectory\Reports\$TimeSpentFileName){
         Write-Log "File Generated ... $($TimeSpentFileName)"
         
         # Send email with AuditLogFieldchanges report attached.
         foreach ($ToEmailAddress in $ToEmailAddressesList) {
            Send-MailMessage -From "no-reply@navitro.com" -To "$($ToEmailAddress)" -Subject "$($TimeSpent) for $($StartDate) on $($ComputerName)" -Body "Generated $($TimeSpent) report on $($ComputerName) for $($StartDate)." -Attachment "$($WorkingDirectory)\Reports\$($TimeSpentFileName)" -SmtpServer mx01.minedata.no -Port 25
            if ($?){
               Write-Log "Email sent to $($ToEmailAddress) ..."
            }
         }
      } else {
         Write-Log "Report file was not found."
      }
   }
}

if ($?){
   Write-Log "Operation completed successfully ..."

   foreach ($ErrorMail in $ErrorMailgList){
      Send-MailMessage -From "no-reply@navitro.com" -To "$($ErrorMail)" -Subject "CAP Auto Report Script" -Body "Report genration successfully ended on the $($ComputerName) server at $($(Get-Date).ToString("HH:mm:ss")) on $($(Get-Date).ToString("dd.MM.yyyy"))." -SmtpServer mx01.minedata.no -Port 25
   }

   # Terminate the script. If generation is successfull.
   [Environment]::Exit(0)
} else {
   Write-Log "Somthing went wrong. Did not complete properly ..."

   foreach ($ErrorMail in $ErrorMailgList){
      Send-MailMessage -From "no-reply@navitro.com" -To "$($ErrorMail)" -Subject "CAP Auto Report Script" -Body "Report genration not successfull on the server $($ComputerName). Something went wrong." -SmtpServer mx01.minedata.no -Port 25
   }

   # Terminate the script. If generation is unsuccessfull.
   [Environment]::Exit(1)
}