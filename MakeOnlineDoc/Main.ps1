#Program:
#       This program will generate the documents needed to go online
#       Advice run this script with powershell 7
#       https://learn.microsoft.com/zh-tw/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4#msi
# History:
# 2024/03/06    Jie	First release
# 2024/05/16    Jie Adjust program can handle loacl folder and remote folder simultaneously
# 2024/05/30    Jie Improve Performance
#import module
. .\GUI.ps1
. .\Module.ps1
#Set Default Global variable to log file,Prevent garbled
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
# 檢查
# 顯示檔案清單到畫面供使用者確認
function Check {
    _logInfo "檢查-Start" "Yellow"
    # $RootPath = $null               <#換版文件根目錄#>                   
    # $HoldPath = $null               <#凍版程式路徑#>           
    # $DateFormat = ""                <#慣用日期格式#>        
    # $SourceFolder = ""              <#Source資料夾名稱#>    
    # $ObjFolder = ""                 <#Obj資料夾名稱#>         
    # $DeployDate = ""                <#換版日期#>
    # $FileListSetting = $null        <#產生FileList時部署程式區分AP/WEB#>
    # $CsvFile = ""                   <#檔案資訊#>
    # 0_Step-Get Setting Value
    $RootPath, $HoldPath, $DateFormat, $SourceFolder, $ObjFolder`
        , $DeployDate, $FileListSetting, $CsvFile = GetSettingValue
    # 1_Step-Check RootPath、HoldPath、 is local or remote
    $RootPath = ParsePath $RootPath
    $HoldPath = ParsePath $HoldPath
    # 2_Step-Display Filelist
    DisplayFileList -CsvFile $CsvFile -RootIP $RootPath.IP -RootPath $RootPath.PATH -HoldIP $HoldPath.IP -HoldPath $HoldPath.PATH `
        -SourceFolder $SourceFolder -ObjFolder $ObjFolder -DeployDate $DeployDate

    _logInfo "檢查-End" "Yellow"
}
# 執行
# 產生filelist
# 備份凍版程式資料夾(供退版還原使用)
# 複製檔案(Hold/New Source，Hold/New Obj)
# 產生TimeStamp
# 產生壓縮檔和MD5 hash值
# 產生文件(程式名稱及檔案名稱、檔案名稱說明)
function Main {
    #Powershell local scope variable is share between function
    _logInfo "執行-Start" "Yellow"
    # 0_Step-Get Setting Value
    $RootPath, $HoldPath, $DateFormat, $SourceFolder, $ObjFolder`
        , $DeployDate, $FileListSetting, $CsvFile = GetSettingValue
    # 1_Step-Check RootPath、HoldPath、 is local or remote
    $RootPath = ParsePath $RootPath
    $HoldPath = ParsePath $HoldPath
    #產生filelist
    $ChangedFiles = GenerateFileList -FileListSettings $FileListSetting  -CsvFile $CsvFile `
        -RootIP $RootPath.IP -RootPath $RootPath.PATH -DeployDate $DeployDate
    if ($null -ne $ChangedFiles -and $ChangedFiles.Count -gt 0) {
        #備份凍版程式資料夾(供退版還原使用)
        BackUpHoldSource -IP $HoldPath.IP -Path $HoldPath.PATH -DeployDate $DeployDate -SourceFolder $SourceFolder -ObjFolder $ObjFolder
        #複製檔案到(本次上線資料夾及凍版程式資料夾)
        CopyFile -ChangedFiles $ChangedFiles -RootIP $RootPath.IP -RootPath $RootPath.PATH -HoldIP $HoldPath.IP `
            -HoldPath $HoldPath.PATH -DeployDate $DeployDate
        #產生TimeStamp
        GenerateTimeStamp -ChangedFiles $ChangedFiles -RootIP $RootPath.IP -RootPath $RootPath.PATH `
            -SourceFolder $SourceFolder -ObjFolder $ObjFolder -DeployDate $DeployDate
        #產生壓縮檔和MD5 hash值
        GenerateZipAndMD5 -RootIP $RootPath.IP -RootPath $RootPath.PATH -DeployDate $DeployDate `
            -SourceFolder $SourceFolder
        #產生文件(程式名稱及檔案名稱、檔案名稱說明)
        GenerateFileMemo -ChangedFiles $ChangedFiles -RootIP $RootPath.IP -RootPath $RootPath.PATH `
            -DeployDate $DeployDate
    }
    _logInfo "執行-End" "Yellow"
} 
$result = $form.ShowDialog()


