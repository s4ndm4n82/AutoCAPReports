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

# Variable Assignment #

# Menu text.
# Top banner.
$NewLine = [Environment]::Newline

$Banner += "            ******************************************************
            *                                                    *
            *              Manual Reports Generator              *
            *                                                    *
            *  Created by: Sandaruwan Samaraweera                *
            *  Version: 1.0.1                                    *
            *                                                    *
            ******************************************************$($NewLine)
---------------------------------------------------------------------------------$($NewLine)"

$TableSelectOptions +="       Select the Server.$($NewLine)
        1. Navitro - 02 (Old server.)
        2. DCProd - 01 (New server.)
        x. Exit.$($NewLine)"

# Working directory.
#$WorkingDirectory = 'C:\Tools\Reports'
$WorkingDirectory = 'F:\Users\SiNUX\MEGA\GitRepoShare\Repos\AutoCAPReports'

# Dates.
$StartDate = ""
$EndDate = ""

# Command.
$McCommand = '"C:\Program Files (x86)\Navitro\CAP\ManagementConsole.exe" /CreateReport'

# Database.
$DataBaseAddress    = ''
$DataBaseTableV1    = ''
$DataBaseTableV2    = ''
$DataBaseUser     = ''
$DataBasePassword = ''

# Report output type.
$ReportOutPdf   = '/type:PDF'
$ReportOutExl   = '/type:EXCEL'
$ReportOutWrd   = '/type:WORD'

# Used field.
$UsedField = '/fields:"Voucher Type"'

# Report names.
$ItemsPerJob                = 'ItemPerJob'
$ItemsPerJobDetailed        = 'ItemsPerJobDetailed'
$ImportedExported           = 'ImportedExported'
$GroupByFields              = 'GroupByFields'
$GroupByFieldsJobColumns    = 'GroupByFieldsJobColumns'
$ProcessingTime             = 'GroupByFieldsJobColumns'
$AuditLogFieldChanges       = 'AuditLogFieldChanges'
$AuditLogCount              = 'AuditLogCount'
$TimeSpent                  = 'TimeSpent'

# Report output folders.
$ItemsPerJobFolder                = 'ItemPerJob'
$ItemsPerJobDetailedFolder        = 'ItemsPerJobDetailed'
$ImportedExportedFolder           = 'ImportedExported'
$GroupByFieldsFolder              = 'GroupByFields'
$GroupByFieldsJobColumnsFolder    = 'GroupByFieldsJobColumns'
$ProcessingTimeFolder             = 'GroupByFieldsJobColumns'
$AuditLogFieldChangesFolder       = 'AuditLogFieldChanges'
$AuditLogCountFolder              = 'AuditLogCount'
$TimeSpentFolder                  = 'TimeSpent'

# Folder name array.
$OutputFolderNameArray  = @($ItemsPerJobFolder, $ItemsPerJobDetailedFolder, $ImportedExportedFolder, $GroupByFieldsFolder
                            , $GroupByFieldsJobColumnsFolder, $ProcessingTimeFolder, $AuditLogFieldChangesFolder
                            , $AuditLogCountFolder, $TimeSpentFolder)

# Output folder path.
$ReportMainFolderPath   = "$($WorkingDirectory)\Reports"

function CheckFolders() {
    # Output folder check.
    if (!(Test-Path $ReportMainFolderPath)){
        New-Item -ItemType Directory -Path $ReportMainFolderPath | Out-Null
    }

    foreach ($OutputFolderName in $OutputFolderNameArray){
        $OutputFolderPathFull = "$($ReportMainFolderPath)\$($OutputFolderName)"

        if (!(Test-Path $OutputFolderPathFull)){
            New-Item -ItemType Directory -Path $OutputFolderPathFull | Out-Null
        }
    }
}

function MenuDatabase() {
    Write-Output "$($Banner)"
    Write-Output "$($TableSelectOptions)"
}

function MakeReportNavitro02() {
    Write-Output "Navitro 02 selected."
}

function MakeReportDCProd01() {
    Clear-Host
    Write-Output "DCProd 01 selected."
}

do {
    MenuDatabase
    $UserServerSelect = Read-Host "       You selection: "

    switch ($UserServerSelect) {
        '1' {MakeReportNavitro02}
        '2' {MakeReportDCProd01}
    }
} until (($UserServerSelect -eq 'x') -or ($UserServerSelect -eq '1') -or ($UserServerSelect -eq '2'))