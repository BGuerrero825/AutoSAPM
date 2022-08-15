# 1st Lt Brian Guerrero
# PullUpdates version 1.0

# PowerShell online Documentation: https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Core/?view=powershell-7.2

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
# Get a program in the programs collection based on either name or index
    param($name, $index)
    foreach ($p in $programs){
        if($p[0] -eq $name){
            return $p
        }
    }
    return $null
}

function Get-ProgramIndex{
# Get a program index in the programs collection based on either name
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
    Copy-Item $downloads_dash_path -Destination $dash_path
    Copy-Item $downloads_poc_path -Destination $poc_path
    Write-Host -ForegroundColor DarkGray "Copied files over to script directory."
}

function Write-ProgramDetails{
# given a program collection, print details to specified format
    param($program)
    Write-Host -BackgroundColor DarkCyan $program[0]
    Write-Host $program[1] "|" $program[2] "|" $program[3] "|" $program[4] "`n"
}

function Check-Files2{
    Write-Host -ForegroundColor DarkGray "`nLocating https://mict.cce.af.mil/ Excel files on local machine."
    # test if documents exist in script dir
    if (-not ((Test-Path -Path $dash_path) -and (Test-Path -Path $poc_path))){
        # if they do not, copy them from the downloads folder if they exist
        if ((Test-Path -Path $downloads_dash_path) -and (Test-Path -Path $downloads_poc_path)){ 
            Copy-DownloadedFiles
        # otherwise, proceed to download from the web
        } else {
            Download-Files
        }
    } else {
        # if files found, ask user if they want to grab new ones anyway
        $option = Read-Host "Files found in script directory are from" (Get-Item $dash_path).LastWriteTime "and" (Get-Item $poc_path).LastWriteTime ". Do you want to grab new files? (Y/N)"
        if ($option -eq "y" -or $option -eq "Y" -or $option -eq "yes"){
            Download-Files
        }
    }
}

function Download-Files{
    # remove old MICT files in downloads
    Remove-Item .\*UnitWorkcenterDashboard*
    Remove-Item .\*ChecklistPOCReport*
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

function Check-Files{
    Write-Host -ForegroundColor DarkGray "`nLocating necessary Excel files from https://mict.cce.af.mil/"
    # test if documents exist in ShareDrive
    if (-not ((Test-Path -Path $dash_path) -and (Test-Path -Path $poc_path))){
        # if not, open two tabs to relevant MICT pages to download files and prompt user action
        Write-Host -ForegroundColor DarkYellow  "`r`nYou either do not have the most recent MICT data downloaded or are missing documents, opening your browser to MICT now. Press the ""Export to Excel"" buttons on each tab that opens, then press ENTER when both documents have been saved to Downloads (default location)..."
        Read-Host "Press ENTER"
        # while path still can't be found
        while (-not ((Test-Path -Path $dash_path) -and (Test-Path -Path $poc_path))){
            Write-Host -ForegroundColor DarkYellow  "`r`nStill could not find file paths to Downloads\(FOUO)91 COS-UnitWorkcenterDashboard-"$date ".xlsx or Downloads\(ForOfficialUseOnly)91 COSChecklistPOCReport.xlsx"
            Read-Host "Press ENTER"
        }
# if found, proceed
    } else { Write-Host -ForegroundColor Green "`r`nFound files in Downloads."}
}

function Write-SAS{
    # Print all programs related to current and next SAS month
    $month = Get-Date
    $next_month = (Get-Date).AddMonths(1)
    $month_name = (Get-Culture).DateTimeFormat.GetMonthName($month.Month)
    $next_month_name = (Get-Culture).DateTimeFormat.GetMonthName($next_month.Month)

    $current = 0
    # iterate through SAS text file
    foreach($line in (Get-Content $sas_path)){
        # if in a month that string matched month or next month, print all programs until end found
        if($current -eq 1){
            if($line -like "Xend*"){
                $current = 0
            }else{
            $program = Get-Program -name $line
            Write-ProgramDetails -program $program
            }  
        } 
        # else if not in a month, look for a months signifier, and print name if found
        elseif($line -like '+*'){
            if ($line.Substring(1) -eq $month_name){
                $current = 1
                Write-Host "----------------------" $month_name "----------------------" "`n"
            } elseif ($line.Substring(1) -eq $next_month_name){
                $current = 1
                Write-Host "----------------------" $next_month_name "----------------------" "`n"
            }
        }
    }
}

function Write-FullSAS{
    $current = 0
    # iterate throus SAS text file
    foreach($line in(Get-Content $sas_path)){
        # if in any month print programs until end found
        if($current -eq 1){
            if($line -like "Xend*"){
                $current = 0
            }else{
                $program = Get-Program -name $line
                Write-ProgramDetails -program $program
            }
        }
        # else if not in a month, look for any month signifier, and print name if found
        elseif($line -like '+*'){
            $month_name = $line.Substring(1, $line.Length-1)
            $current = 1
            Write-Host "----------------------" $month_name "----------------------" "`n"
        }
    }
}


# =================== Begin Script =====================

# ensure files exist in Downloads, otherwise open Chrome
Check-Files2

# open the excel files from Downloads, then create cells objects for each
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
    $assessed = "None Found"
    $validated = "None Found"
    $programs += ,@($program_name, $poc1, $poc2, $assessed, $validated)
    $i += 1
}
# parse programs in Dashboard sheet and insert new data to existing objects in programs array
$j = 3
while(($dash_cells.item($j, 2).text) -ne ""){
    if(($dash_cells.item($j, 3).text) -eq ""){
        
    }
    else{
        $k = Get-ProgramIndex -name ($dash_cells.item($j, 2).text)
        $assessed = $dash_cells.item($j, 10).text
        $validated = $dash_cells.item($j, 11).text
        $programs[$k][3] = $assessed
        $programs[$k][4] = $validated 
    }
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
Write-FullSAS

# TODO: create SAS month in excel table for easy copy paste to powerpoint 
Read-Host "Press ENTER to exit."

# =================== End Script =====================