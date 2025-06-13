Add-Type -AssemblyName System.IO.Compression
#region ObjectType
class IPAddress {
    <#遠端位置#>    [string] $IP
    <#路徑#>        [string] $PATH
    #C'tor
    IPAddress([string] $IP, [string] $PATH) {
        $this.IP = $IP
        $this.PATH = $PATH
    }
}
#檔案物件
class FileInfo {
    <#檔案狀態(新增/刪除)#>      [string]$FileStatus
    <#檔案名稱#>                [string]$FileName
    <#機器種類(Ap/Web)#>        [string]$MachineType
    <#From(擷取路徑)#>          [string]$FromPath
    <#To(上版路徑)#>            [string]$ToPath
    <#檔案說明#>                [string]$FileMemo
    #C'tor
    FileInfo([string]$FileStatus, [string]$FileName, [string]$MachineType, [string]$FromPath, [string]$ToPath, [string]$FileMemo) {
        if ([string]::IsNullOrEmpty($FileStatus) -or [string]::IsNullOrEmpty($FileName) -or [string]::IsNullOrEmpty($MachineType) `
                -or [string]::IsNullOrEmpty($FromPath) -or [string]::IsNullOrEmpty($ToPath) -or [string]::IsNullOrEmpty($FileMemo)) { throw "Error" }
        $this.FileStatus = $FileStatus
        $this.FileName = $FileName
        $this.MachineType = $MachineType
        $this.FromPath = $FromPath
        $this.ToPath = $ToPath
        $this.FileMemo = $FileMemo
    }
    FileInfo([string]$FileStatus, [string]$FileName, [string]$FromPath) {
        $this.FileStatus = $FileStatus
        $this.FileName = $FileName
        $this.FromPath = $FromPath
    }
}
#endregion ObjectType
#壓縮檔案
function CompressArchive {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$Path,    # Input folder paths (array of strings)
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath      # Output zip file path
    )

    # Current directory
    $currentPath = (Get-Location).Path

    # Check if zipFilePath is relative and convert to absolute if needed
    if (-not [System.IO.Path]::IsPathRooted($DestinationPath)) {
        $DestinationPath = Join-Path -Path $currentPath -ChildPath $DestinationPath
    }

    # Check if the zip file exists and remove it
    if (Test-Path $DestinationPath) {
        Remove-Item $DestinationPath -Force
    }

    # Open the zip file for writing
    $DestinationPathStream = [System.IO.File]::Open($DestinationPath, [System.IO.FileMode]::Create)
    $zipArchive = New-Object System.IO.Compression.ZipArchive($DestinationPathStream, [System.IO.Compression.ZipArchiveMode]::Update)

    try {
        foreach ($path in $Path) {
            # Resolve each folder path to absolute path
            $resolvedPath = Resolve-Path -Path $path -ErrorAction Stop
            $rootFolderName = Split-Path -Path $resolvedPath -Leaf  # Get the root folder name
            $parentPath = Split-Path -Path $resolvedPath -Parent    # Get the parent directory

            # Get all files in the folder (including subdirectories)
            $filesToCompress = Get-ChildItem -Path $resolvedPath -Recurse -File

            foreach ($file in $filesToCompress) {
                # Get the relative path including the root folder
                $relativePath = Join-Path -Path $rootFolderName -ChildPath ($file.FullName.Substring($resolvedPath.Path.Length + 1))

                # If the file is directly in the root folder, just use the root folder and filename
                if ($file.DirectoryName -eq $resolvedPath.Path) {
                    $relativePath = Join-Path -Path $rootFolderName -ChildPath $file.Name
                }

                # Create an entry in the zip archive with the relative path
                $zipEntry = $zipArchive.CreateEntry($relativePath)

                # Open the entry and copy the file's content to the zip entry
                $entryStream = $zipEntry.Open()
                $fileStream = [System.IO.File]::OpenRead($file.FullName)
                $fileStream.CopyTo($entryStream)
                $fileStream.Close()
                $entryStream.Close()
            }
        }
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        # Close the zip archive and stream
        $zipArchive.Dispose()
        $DestinationPathStream.Close()
    }
}
#紀錄Log
function _logInfo {
    Param([string]$message, [string]$color = "Green")
    $Time = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "$Time :$message" -ForegroundColor $color
    $ToolStripStatusLabel1.Text = $message
    $StatusStrip1.Update()
}
#執行指令(本地/遠端)
function _exeCommand {
    Param([string]$ip, [scriptblock] $fun, [Object[]]$param )
    $result = $null
    if ($ip -ne "localhost") {
        $result = Invoke-Command -ComputerName $ip -ScriptBlock $fun -ArgumentList $param
    }
    else {
        $result = Invoke-Command -ScriptBlock $fun -ArgumentList $param
    }
    return $result

}
#取得設定頁籤參數
function GetSettingValue {
    $RootPath = $textBox1.Text                                      
    $HoldPath = $textBox2.Text                                      
    $DateFormat = $ComboBox2.SelectedItem.ToString()
    $SourceFolder = $textBox3.Text                              
    $ObjFolder = $textBox4.Text                                 
    $DeployDate = ($datetimepciker1.Value).ToString($DateFormat)
    $FileListSetting = $CheckBox1.Checked
    $CsvFile = "FileInfo_$($ComboBox1.SelectedItem.ToString()).csv"
    _logInfo "GetSettingValue"
    _logInfo "換版文件根目錄:$RootPath" -color "Blue"
    _logInfo "凍版程式路徑:$HoldPath" -color "Blue"
    _logInfo "source路徑:$SourceFolder" -color "Blue"
    _logInfo "Obj路徑:$ObjFolder" -color "Blue"
    _logInfo "換版日期:$DeployDate" -color "Blue"
    _logInfo "FileList區分AP/web:$FileListSetting" -color "Blue"
    _logInfo "檔案設定清單:$CsvFile" -color "Blue"
    return $RootPath, $HoldPath, $DateFormat, $SourceFolder, $ObjFolder`
        , $DeployDate, $FileListSetting, $CsvFile
}
#解析目標路徑(本地/遠端)
function ParsePath {
    param([string]$target)
    $result = $null
    $isRemote = $target -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
    if ($isRemote) {
        $tmp = @($target -split "\\\\")[1] -split "\\"
        # $tmp #Debug
        $result = New-Object IPAddress $tmp[0] , $(($tmp[1..($tmp.Count - 1)] -join "\").Replace("$", ":"))
    }
    else {
        $result = New-Object IPAddress "localhost" , $target
    }
    return $result
}
#寫檔案(.txt)到遠端目標
function WriteContent {
    param([string] $ip, [string]$path, [string]$fileName, [string]$content)
    _logInfo "WriteContent $fileName"
    function temp {
        param([string]$path, [string]$fileName, [string]$content)
        #Set Default Global variable to log file,Prevent garbled
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'
        $originPath = Get-Location
        Set-Location $path
        #Delete file, if exists
        if (Test-Path -Path $fileName) { Remove-Item -Path $fileName }
        New-Item -itemType "file" -Value $content -Path $fileName
        Set-Location $originPath
    }
    #Caputure Ouput which is no need to return
    $_ = _exeCommand -ip $ip -fun $function:temp -param @($path, $fileName, $content) 
}
#顯示檔案清單到畫面供使用者確認
function DisplayFileList {
    Param([string]$CsvFile, [string]$RootIP, [string]$RootPath, [string]$HoldIP, [string]$HoldPath
        , [string]$SourceFolder, [string]$ObjFolder, [string]$DeployDate)
    _logInfo "顯示檔案清單到畫面供使用者確認"
    function GetFile {
        Param([string]$path, [int] $len)
        if (-not (Test-Path -Path $path)) {
            throw "$path is not exists"
        }
        $result = @(Get-ChildItem -Recurse -file -Path "$path" -exclude *.zip) | foreach-Object { $_.FullName.substring($len) }
        return $result
    }
    #Summary Changed File
    $ChangedFiles = New-Object System.Collections.Generic.List[System.Object]
    # Get Freezed File
    $HoldSource = _exeCommand -ip $HoldIP -fun $function:GetFile -param @("$HoldPath\$SourceFolder", $HoldPath.Length)
    $HoldObj = _exeCommand -ip $HoldIP -fun $function:GetFile -param @("$HoldPath\$ObjFolder", $HoldPath.Length)
    # Get Changed File
    $NewSource = _exeCommand -ip $RootIP -fun $function:GetFile -param @("$RootPath\$DeployDate\$SourceFolder", "$RootPath\$DeployDate".Length)
    $NewObj = _exeCommand -ip $RootIP -fun $function:GetFile -param @("$RootPath\$DeployDate\$ObjFolder", "$RootPath\$DeployDate".Length)
    # Compoare Source
    $NewSource | foreach-Object `
    { 
        $tmp = New-Object FileInfo (& { if ($HoldSource.Contains($_)) { "modify" } else { "new" } }), @($_ -split "\\")[-1], $_
        $ChangedFiles.Add($tmp)
    }
    # Compare Obj
    $NewObj | foreach-Object `
    { 
        $tmp = New-Object FileInfo (& { if ($HoldObj.Contains($_)) { "modify" } else { "new" } }), @($_ -split "\\")[-1], $_
        $ChangedFiles.Add($tmp)
    }
    # #Check FileInfo Setting，第一次執行將將所有本次的檔案先放進去之後用檢查的方式(新增)<用預設格式建立，方便之後讀取>
    if ( -not (Test-Path -Path $CsvFile)) {
        New-Item -itemType "file" -Path $CsvFile
        $ChangedFiles | Export-Csv $CsvFile
    }
    # # #Read FileInfo Setting
    $FileInfo = Import-Csv -Path $CsvFile
    # #Add value when not found in csv (Will Add FileStatus temporarily)
    $ChangedFiles | Where-Object { !$FileInfo.FromPath.Contains($_.FromPath) } | Export-Csv -Append $CsvFile    
    # # #Refresh FileInfo
    $FileInfo = Import-Csv -Path $CsvFile
    #Display Data to gridView
    $DataGridView1.Rows.Clear()
    foreach ($elem in $ChangedFiles) {
        $target = $FileInfo[[array]::IndexOf($FileInfo.FromPath, $elem.FromPath)]
        $DataGridView1.Rows.Add($elem.FileStatus, $elem.FileName, $target.MachineType
            , $elem.FromPath, $target.ToPath, $target.FileMemo)
    }
    $DataGridView1.Refresh()
    # #Display MessageBox
    [MessageBox]::Show("在此工具程式內編輯或關閉此程式到外面另開CSV編輯", "提示訊息")
} 
#產生檔案清單
function GenerateFileList {
    param([bool] $FileListSettings, [string]$CsvFile, [string]$RootIP, [string]$RootPath, [string]$DeployDate)
    _logInfo "產生檔案清單"
    $ChangedFiles = New-Object System.Collections.Generic.List[System.Object]
    $IsAnyValueIsNullOrEmpty = $false
    #Capture value from gridView，Check value is not null or empty
    $DataGridView1.Update()
    #檢查有無空值，如有則不作後續
    foreach ($row in $DataGridView1.Rows) {
        try {
            $ChangedFiles += New-Object FileInfo $row.Cells["FileStatus"].Value, $row.Cells["FileName"].Value, `
                $row.Cells["MachineType"].Value, $row.Cells["FromPath"].Value, `
                $row.Cells["ToPath"].Value, $row.Cells["FileMemo"].Value
        }
        catch {
            $IsAnyValueIsNullOrEmpty = $true
            break
        }
    }
    if ($IsAnyValueIsNullOrEmpty) {
        [MessageBox]::Show("尚有未填入的資訊", "提示訊息")
        return $null
    }
    #先儲存資訊到Csv，儲存前清空檔案狀態
    $FileInfos = Import-Csv -Path $CsvFile
    foreach ($elem in $ChangedFiles) {
        $tar = $FileInfos[[array]::IndexOf($FileInfos.FromPath, $elem.FromPath)]
        $tar.FileStatus = ""    #清空檔案狀態
        $tar.MachineType = $elem.MachineType
        $tar.ToPath = $elem.ToPath
        $tar.FileMemo = $elem.FileMemo
    }
    $FileInfos | Export-Csv -Path $CsvFile
    #產生檔案清單(全部及機器種類區分)
    #ModifyFile
    $ModifyFile = "Modified Files -------`r`n"
    $NewFile = "New Files -------`r`n"
    $DeployFile = @{}
    foreach ($elem in $ChangedFiles){
        #Modify or New 
        if($elem.FileStatus -eq "modify"){
            $ModifyFile += "$($elem.FromPath)`r`n" 
        }
        else{
            $NewFile += "$($elem.FromPath)`r`n" 
        }
        #Deploy
        if($elem.MachineType -ne "NULL"){
            if(-not $DeployFile.Contains($elem.MachineType)){
                $DeployFile.Add($elem.MachineType,"Deploy Files -------$($elem.MachineType)`r`n")
                $DeployFile[$elem.MachineType] += "$($elem.FromPath),$($elem.ToPath)$($elem.FileName)`r`n" 
            }
            else{
                $DeployFile[$elem.MachineType] += "$($elem.FromPath),$($elem.ToPath)$($elem.FileName)`r`n" 
            }
        }
    }
    $content = "$ModifyFile`r`n$NewFile`r`n"
    foreach($elem in $DeployFile.Getenumerator()){
        $content += "$($elem.Value)`r`n"
        if ($FileListSettings -eq $true) {
            WriteContent -ip $RootIP -path "$RootPath\$DeployDate" -fileName "FileList_$($elem.Name).txt" -content $elem.Value
        }
    }
    WriteContent -ip $RootIP -path "$RootPath\$DeployDate" -fileName "FileList.txt" -content $content.Substring(0,$content.Length - 2)
    return $ChangedFiles 
}
#備份凍版程式資料夾(供退版還原使用)
function BackUpHoldSource {
    param([string]$IP, [string]$Path, [string]$DeployDate, [string]$SourceFolder, [string]$ObjFolder)
    _logInfo "備份凍版程式資料夾(在凍版程式路徑底下)"
    function temp {
        param([string]$path, [string]$deploydate, [string]$sourcefolder, [string]$objfolder)
        #HoldSource folder to backup.zip
        if (-not (Test-Path -Path "$path\backup_$deploydate.zip")) { 
            CompressArchive -Path "$path\$sourcefolder", "$path\$objfolder" -DestinationPath "$path\backup_$deploydate.zip"
        }
    }
    _exeCommand -ip $ip -fun $function:temp -param @($Path, $DeployDate, $SourceFolder, $ObjFolder)    
}
#複製檔案到(本次上線資料夾)
function CopyFile {
    param([System.Collections.Generic.List[System.Object]]$ChangedFiles, [string]$RootIP, [string]$RootPath, [string]$HoldIP, [string]$HoldPath
        , [string]$DeployDate)
    _logInfo "複製檔案到(本次上線資料夾)"
    $script = ""
    #Adjust Path
    $RootPath = (& { If ($RootIP -eq "localhost") { $RootPath } Else { "\\$RootIP\$RootPath".Replace(":", "$") } })
    $HoldPath = (& { If ($HoldIP -eq "localhost") { $HoldPath } Else { "\\$HoldIP\$HoldPath".Replace(":", "$") } })
    #Source files
    foreach ($elem in $ChangedFiles) {
        $fromPath = $elem.FromPath.Substring(1)
        $toPath = $elem.FromPath.Substring(1, $elem.FromPath.lastIndexOf("\"))
        #將檔案複製到本次上線資料夾
        #Hold
        if ($elem.FileStatus -ne "new") {
            $script += "xcopy /y $HoldPath\$fromPath $RootPath\$DeployDate\HoldSource\$toPath`r`n"
        }
        #New
        $script += "xcopy /y $RootPath\$DeployDate\$fromPath $RootPath\$DeployDate\NewSource\$toPath`r`n"
        #將NewSource複製到凍版程式路徑
        # $script += "xcopy /y  $RootPath\$global:DeployDate\$fromPath $HoldPath\$toPath`r`n"
    }
    #Publish files，seperate by MachineType
    foreach ($type in $ChangedFiles | Select-Object -ExpandProperty MachineType -Unique | Where-Object { $_.MachineType -ne "NULL" }) {
        # Write-Host $type -BackgroundColor Green #Debug
        foreach ($elem in $ChangedFiles | Where-Object { $_.FromPath.Contains($ObjFolder) -and $_.MachineType -eq $type }) {
            $fromPath = $elem.FromPath.Substring(1)
            $toPath = $elem.FromPath.Substring(1, $elem.FromPath.lastIndexOf("\"))
            #Hold
            if ($elem.FileStatus -ne "new") {
                $script += "xcopy /y $HoldPath\$fromPath $RootPath\$DeployDate\Deploy_Hold_$type\$toPath`r`n"
            }
            #New
            $script += "xcopy /y $RootPath\$DeployDate\$fromPath $RootPath\$DeployDate\Deploy_$type\$toPath`r`n"
        }
    }
    # Write-Host $script -BackgroundColor Red #Debug
    _exeCommand -ip "localhost" -fun {
        param([string]$script)
        Invoke-Expression $script
    } -param @($script)
}
#產生檔案的時間戳記
function GenerateTimeStamp {
    param([System.Collections.Generic.List[System.Object]]$ChangedFiles, [string] $RootIP, [string] $RootPath
    , [string]$SourceFolder, [string]$ObjFolder)
    _logInfo "產生檔案的時間戳記"
    $folders = @($SourceFolder, $ObjFolder)
    $folders += $ChangedFiles | Select-Object -ExpandProperty MachineType -Unique | Where-Object { $_ -ne "NULL" } | foreach-Object { "Deploy_$_" }
    _exeCommand -ip $RootIP -fun {
        param([string]$path, [array]$folders)
        $originPath = Get-Location
        Set-Location $path
        #Call cmd command
        # Write-Host $folders -BackgroundColor Green #Debug
        foreach ($elem in $folders) {
            # Write-Host "/c dir $elem /a-d/s > TimeStamp_$elem.txt" -BackgroundColor Green #Debug
            # Start-Process cmd.exe -ArgumentList "/c dir $elem /a-d/s > TimeStamp_$elem.txt" #TODO這段有問題
            & "C:\WINDOWS\system32\cmd.exe" "/c dir $elem /a-d/s > TimeStamp_$elem.txt"
        }
        Set-Location $originPath
    } -param @("$RootPath\$DeployDate", $folders)
}
#產生壓縮檔和MD5 hash值
function GenerateZipAndMD5 {
    param([string] $RootIP, [string] $RootPath, [string] $DeployDate, [string]$SourceFolder)
    _logInfo "產生壓縮檔和MD5 hash值"
    _exeCommand -ip $RootIP -fun {
        param([string]$path, [string]$folder)
        #Set Default Global variable to log file,Prevent garbled
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'
        $originPath = Get-Location
        Set-Location $path
        #Source folder to zip
        if (Test-Path -Path ".\source.zip") { Remove-Item -Path ".\source.zip" }
        CompressArchive -Path ".\$folder" -DestinationPath ".\source.zip"
        $(Get-FileHash -Path ".\source.zip" -Algorithm "MD5").Hash > Vendor_Src_MD5.txt
        Set-Location $originPath
    } -param @("$RootPath\$DeployDate", $SourceFolder)
}
#產生文件(程式名稱及檔案名稱、檔案名稱說明)
function GenerateFileMemo {
    param([System.Collections.Generic.List[System.Object]]$ChangedFiles, [string] $RootIP, [string] $RootPath
    , [string] $DeployDate)
    _logInfo "產生文件(程式名稱及檔案名稱、檔案名稱說明)"
    $content_1 = ""; $content_2 = "";
    #ModifyFile
    $index = 1
    $content_1 += "Modified Files -------`r`n"; $content_2 += "Modified Files -------`r`n";
    foreach ($elem in $ChangedFiles | Where-Object { $_.FileStatus -eq "modify" }) {
        $content_1 += "$index.	$($elem.FromPath)`r`n" 
        $content_2 += "$index.	$($elem.FromPath),$($elem.FileMemo)`r`n"
        $index++
    }
    #NewFile
    $index = 1
    $content_1 += "`r`nNew Files -------`r`n"; $content_2 += "`r`nNew Files -------`r`n";
    foreach ($elem in $ChangedFiles | Where-Object { $_.FileStatus -eq "new" }) {
        $content_1 += "$index.	$($elem.FromPath)`r`n" 
        $content_2 += "$index.	$($elem.FromPath),$($elem.FileMemo)`r`n"
        $index++
    }
    #DeployFile
    $index = 1
    foreach ($type in $ChangedFiles | Select-Object -ExpandProperty MachineType -Unique | Where-Object { $_ -ne "NULL" }) {
        $content_1 += "`r`nDeploy Files -------$type`r`n"; $content_2 += "`r`nDeploy Files -------$type`r`n"
        foreach ($elem in $ChangedFiles | Where-Object { $_.FromPath.Contains($ObjFolder) -and $_.MachineType -eq $type }) {
            $content_1 += "$index.	$($elem.FromPath)`r`n" 
            $content_2 += "$index.	$($elem.FromPath),$($elem.FileMemo)`r`n"
            $index++
        }
    }
    WriteContent -ip $RootIP -path "$RootPath\$DeployDate" -fileName "Vendor_fileName.txt" -content $content_1
    WriteContent -ip $RootIP -path "$RootPath\$DeployDate" -fileName "Vendor_memo.txt" -content $content_2
}