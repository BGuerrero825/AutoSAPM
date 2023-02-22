# 1st Lt Brian Guerrero
# PullUpdates version 1.6

# PowerShell online Documentation: https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Core/?view=powershell-7.2

# TODO 
# Red/Green dates and program name in printed SAS if they have been assessed for the month
# ----- Replace program string when a new one is parsed
# Add "Unassigned" category in SAS print for programs parsed in excel sheet but not in SAS.txt

# SCRIPT TITLE
Write-Host -ForegroundColor DarkGray " ______   __  __   __       __         
/_____/\ /_/\/_/\ /_/\     /_/\        
\:::_ \ \\:\ \:\ \\:\ \    \:\ \       
 \:(_) \ \\:\ \:\ \\:\ \    \:\ \      
  \: ___\/ \:\ \:\ \\:\ \____\:\ \____ 
   \ \ \    \:\_\:\ \\:\/___/\\:\/___/\
    \_\/     \_____\/ \_____\/ \_____\/"
Write-Host -ForegroundColor Gray " ___ __ __    ________  ______  _________ 
/__//_//_/\  /_______/\/_____/\/________/\
\::\| \| \ \ \__.::._\/\:::__\/\__.::.__\/
 \:.      \ \   \::\ \  \:\ \  __ \::\ \  
  \:.\-/\  \ \  _\::\ \__\:\ \/_/\ \::\ \ 
   \. \  \  \ \/__\::\__/\\:\_\ \ \ \::\ \
    \__\/ \__\/\________\/ \_____\/  \__\/                                          "
Write-Host -ForegroundColor White " __  __   ______   ______   ________   _________  ______   ______     
/_/\/_/\ /_____/\ /_____/\ /_______/\ /________/\/_____/\ /_____/\    
\:\ \:\ \\:::_ \ \\:::_ \ \\::: _  \ \\__.::.__\/\::::_\/_\::::_\/_   
 \:\ \:\ \\:(_) \ \\:\ \ \ \\::(_)  \ \  \::\ \   \:\/___/\\:\/___/\  
  \:\ \:\ \\: ___\/ \:\ \ \ \\:: __  \ \  \::\ \   \::___\/_\_::._\:\ 
   \:\_\:\ \\ \ \    \:\/.:|| \:.\ \  \ \  \::\ \   \:\____/\ /____\:\
    \_____\/ \_\/     \____/   \__\/\__\/   \__\/    \_____\/ \_____\/"

# MICT URLs
$dashboard_url = "https://mict.cce.af.mil/ViewAssessmentDashboard.aspx"
$poc_url = "https://mict.cce.af.mil/MICTReports/ReportChecklistPOC.aspx"

# Globals
$programs = @()
$excel = New-Object -com "Excel.Application"
$date = Get-Date -Format yyyyMMdd

# Filepaths
$downloads_path = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$downloads_dash_path = $downloads_path + "\(FOUO)91 COS-UnitWorkcenterDashboard-" + $date + ".xlsx"
$downloads_poc_path = $downloads_path + "\(ForOfficialUseOnly)91 COSChecklistPOCReport.xlsx"
$src_path = "$PSScriptRoot"
$dash_path = $src_path + "\(FOUO)91 COS-UnitWorkcenterDashboard-" + $date + ".xlsx"
$poc_path = $src_path + "\(ForOfficialUseOnly)91 COSChecklistPOCReport.xlsx"
$sas_path = "$PSScriptRoot" + "\SAS.txt"


# Helper Functions

function Get-Program{
# Get a program from the programs collection based on name
    param($name, $index)
    foreach ($p in $programs){
        if($p[0] -like "*$name*"){
            return $p
        }
    }
    return $null
}

function Get-ProgramIndex{
# Get a program index from the programs collection using a program name
    param($name)
    $i = 0
    foreach ($p in $programs){
        if($p[0] -eq $name){
            return $i
        }
        $i += 1
    }
    return $null
}

function Copy-DownloadedFiles{
    # moves files from downloads directory to script src directory
    Remove-Item .\*UnitWorkcenterDashboard*
    Remove-Item .\*ChecklistPOCReport*
    Copy-Item $downloads_dash_path -Destination $dash_path
    Copy-Item $downloads_poc_path -Destination $poc_path
    Write-Host -ForegroundColor DarkGray "Copied files over to script directory."
}

function Print-ProgramDetails{
# given a program collection, print details to specified format
    param($program)
    if ($program -eq $null){
        return
    }
    Write-Host -ForegroundColor Cyan $program[0]
    Write-Host $program[1] "|" $program[2] "|" $program[3] "|" $program[4] "`n"
}

function Check-Files{
    Write-Host -ForegroundColor DarkGray "`nLocating https://mict.cce.af.mil/ Excel files on local machine."
    $found = $true
    # test if documents exist in script dir
    Write-Host $dash_path ", " $poc_path
    if (-not ((Test-Path -Path $dash_path) -and (Test-Path -Path $poc_path))){
        # if they do not, copy them from the downloads folder if they exist
        if ((Test-Path -Path $downloads_dash_path) -and (Test-Path -Path $downloads_poc_path)){ 
            Copy-DownloadedFiles
        # otherwise, proceed to download from the web
        } else {
            $found = $false
            Download-Files
        }
    }
    # if files found, ask user if they want to grab new ones anyway
    if ($found -eq $true){
    Write-Host -ForegroundColor Green "Files found in script directory are from" (Get-Item $dash_path).LastWriteTime "and" (Get-Item $poc_path).LastWriteTime "."
        $option =  Read-Host "Do you want to grab new files? (Y/N)"
        if ($option -eq "y" -or $option -eq "Y" -or $option -eq "yes"){
            Download-Files
        }
    }
}

function Download-Files{
    # remove old MICT files in downloads
    $to_remove = Join-Path $downloads_path '\*UnitWorkcenterDashboard*'
    Remove-Item $to_remove
    $to_remove = Join-Path $downloads_path '\*ChecklistPOCReport*'
    Remove-Item $to_remove
    # open two tabs to relevant MICT pages to download files and prompt user action
    Start-Process "chrome.exe" $dashboard_url
    Start-Process "chrome.exe" $poc_url
    Write-Host -ForegroundColor DarkYellow  "`r`nYou either do not have the most recent MICT data downloaded or are missing documents, opening your browser to MICT now. Press the ""Export to Excel"" buttons on each tab that opens, then press ENTER when both documents have been saved to Downloads (default location)..."
    Read-Host "Press ENTER"
    # while path still can't be found
    while (-not ((Test-Path -Path $downloads_dash_path) -and (Test-Path -Path $downloads_poc_path))){
        Write-Host -ForegroundColor DarkYellow  "`r`nStill could not find file paths to Downloads\(FOUO)91 COS-UnitWorkcenterDashboard-"$date ".xlsx or Downloads\(ForOfficialUseOnly)91 COSChecklistPOCReport.xlsx"
        Read-Host "Press ENTER"
    }
    # when found, copy to script dir then proceed
    Copy-DownloadedFiles
    Write-Host -ForegroundColor Green "`r`nFound files. Proceeding..."
}

# Print all programs related to current and next SAS month, optionally print all months
function Write-SAS{
    param($full)
    $current_month = Get-Date
    $next_month = (Get-Date).AddMonths(1)
    $current_month_name = (Get-Culture).DateTimeFormat.GetMonthName($current_month.Month)
    $next_month_name = (Get-Culture).DateTimeFormat.GetMonthName($next_month.Month)

    $current = 0
    # iterate through SAS text file
    foreach($line in (Get-Content $sas_path)){
        # if month listed (denoted with a leading '-')
        if($line.StartsWith("-")){
            $current = 0
            $month_name = $line.Substring(1, $line.Length-1)
            # if the month is the current or next month, print specially highlighted
            if($line.Substring(1) -eq $current_month_name -or $line.Substring(1) -eq $next_month_name){
                $current = 1
                Write-Host "`n"
                Write-Host -BackgroundColor DarkCyan "----------------------" $month_name "----------------------"
            # else print other months if full selected
            } elseif ($full) {
                Write-Host "`n"
                Write-Host -BackgroundColor DarkGray "----------------------" $month_name "----------------------"
            }
        # if the line isn't blank or commented, and is needed for current SAS scope, check for a program name and print
        } elseif($line -ne '' -and $line -ne ' ' -and (-not $line.StartsWith('#')) -and ($current -eq 1 -or $full)) {
            $program = Get-Program -name $line
            if ($program -eq $null){
                Write-Host -ForegroundColor Red "Couldn't find match for program '$line'. Try editing keywords in SAS.txt. `n"
            } else {
            Print-ProgramDetails -program $program
            }
        } else {
            #do nothing if line is blank or commented
        }
    }
}


# =================== Begin Script =====================

# ensure files exist in Downloads, otherwise open Chrome
Check-Files

# open the excel files from Downloads, then create cells objects for each file
Write-Host -ForegroundColor DarkGray -NoNewline "`r`nOpening files... "
$dash = $excel.Workbooks.open($dash_path)
$pocs = $excel.Workbooks.open($poc_path)
$dash_cells = $dash.ActiveSheet.Cells
$pocs_cells = $pocs.ActiveSheet.Cells
Write-Host -ForegroundColor DarkGray "Files opened."


# parse programs in POC sheet and insert objects to the programs array
$i = 2
Write-Host -ForegroundColor DarkGray -NoNewLine "Collecting data... "
while(($pocs_cells.item($i, 5).text) -ne ""){
    $program_name = $pocs_cells.item($i, 5).text
    $poc1 = $pocs_cells.item($i, 7).text
    $poc2 = $pocs_cells.item($i, 9).text
    $assessed = $null
    $validated = $null
    $programs += ,@($program_name, $poc1, $poc2, $assessed, $validated)
    $programs[$i-2][1] = if($poc1 -ne ''){$poc1} else{'VACANT'}
    $programs[$i-2][2] = if($poc2 -ne ''){$poc2} else{'VACANT'}
    $i += 1
}
# parse programs in Dashboard sheet and insert new data to existing objects in programs array
$j = 3
while(($dash_cells.item($j, 2).text) -ne ""){
    # skip this row if there is no assessment (this is a header row)
    if(($dash_cells.item($j, 3).text) -eq ""){
        $j += 1
        continue
    }
    $k = Get-ProgramIndex -name ($dash_cells.item($j, 2).text)
    $assessed = $dash_cells.item($j, 10).text
    $validated = $dash_cells.item($j, 12).text
    $programs[$k][3] = if($assessed -ne ''){$assessed} else{'NO DATE'}
    $programs[$k][4] = if($validated -ne ''){$validated} else{'NO DATE'}
    $j += 1
}

Write-Host "Data is ready."

# close excel session when finished
$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[GC]::Collect()

# print out programs
Write-SAS
Read-Host "Press ENTER to print full SAS"
Write-SAS -full $true

# TODO: create SAS month in excel table for easy copy paste to powerpoint 
Read-Host "Press ENTER to exit."

# =================== End Script =====================