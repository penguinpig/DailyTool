using namespace System.IO
using namespace System.Collections.Generic
using namespace System.Drawing
using namespace System.Windows.Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#region Instruction
$instructionText = " `# Readme.md

# 遠端執行設定

- 遠端及本地的信任清單，加入各自對方的IP

## 本地電腦設定

1. Command + x -> A，以系統管理員身分執行Powershell
2. 確認 winrm狀態 `"Get-Service -Name winrm`"
3. 確認目前信任清單 `"winrm get winrm/config/client`"，確認 TrustedHosts 有哪些
4. 將遠端電腦IP加入信任清單(需串上舊資料)  `"winrm set winrm/config/client '@{TrustedHosts=`"serverIP,serverIP`"}'`"
5. 測試連線 `"Enter-PSsession {serverIP}`"，成功後輸入exit離開
6. 如不行到到遠端電腦設定將自己的IP加到信任清單後再試一次

## 遠端電腦設定

1. 先用mstsc.exe遠端登入進去，並打開Powershell
2. 確認 winrm狀態 `"Get-Service -Name winrm`"
3. 確認目前信任清單 `"winrm get winrm/config/client`"，確認 TrustedHosts 有哪些
4. 將自己電腦IP加入信任清單(需串上舊資料) `"winrm set winrm/config/client '@{TrustedHosts=`"serverIP,serverIP`"}'`"
"
#endregion Instruction
#region GUI
#region DateTimePicker
$datetimepciker1 = New-Object System.Windows.Forms.DateTimePicker
$datetimepciker1.Anchor = 'Bottom, Right'
$datetimepciker1.Location = New-Object System.Drawing.Point(896, 428)
$datetimepciker1.Size = New-Object System.Drawing.Size(200, 30)
#endregion DateTimePicker
#region Button
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Anchor = 'Bottom, Right'
$Button1.Location = New-Object System.Drawing.Point(557, 430)
$Button1.Size = New-Object System.Drawing.Size(92, 33)
$Button1.Text = '檢查'
$Button1.Add_Click({ Check })

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Anchor = 'Bottom, Right'
$Button2.Location = New-Object System.Drawing.Point(670, 428)
$Button2.Size = New-Object System.Drawing.Size(92, 33)
$Button2.ForeColor = [Color]::Red
$Button2.Text = '執行'
$Button2.Add_Click({ Main })
#endregion Button
#region CheckBox
$CheckBox1 = New-Object System.Windows.Forms.CheckBox
$CheckBox1.Anchor = 'Bottom, Right'
$CheckBox1.Location = New-Object System.Drawing.Point(211, 374)
$CheckBox1.Size = New-Object System.Drawing.Size(367, 28)
$CheckBox1.Text = '產生FileList時部署程式區分AP/WEB'
#endregion CheckBox
#region TextBox
$instruction = New-Object System.Windows.Forms.TextBox
$instruction.Location = New-Object System.Drawing.Point(0, 3)
$instruction.Size = New-Object System.Drawing.Size(1100, 460)
$instruction.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$instruction.ReadOnly = $true
$instruction.Multiline = $true
$instruction.ScrollBars = [ScrollBars]::Vertical
$instruction.Text = $instructionText

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(211, 138)
$textBox1.Size = New-Object System.Drawing.Size(512, 30)
$textBox1.Font = [Font]::new("Microsoft Sans Serif", 12)
$textBox1.Text = ""

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(211, 174)
$textBox2.Size = New-Object System.Drawing.Size(512, 30)
$textBox2.Font = [Font]::new("Microsoft Sans Serif", 12)
$textBox2.Text = ""

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(211, 250)
$textBox3.Size = New-Object System.Drawing.Size(512, 30)
$textBox3.Font = [Font]::new("Microsoft Sans Serif", 12)
$textBox3.Text = ""

$textBox4 = New-Object System.Windows.Forms.TextBox
$textBox4.Location = New-Object System.Drawing.Point(211, 286)
$textBox4.Size = New-Object System.Drawing.Size(512, 30)
$textBox4.Font = [Font]::new("Microsoft Sans Serif", 12)
$textBox4.Text = ""
#endregion TextBox
#region ComboBox
$ComboBox1 = New-Object System.Windows.Forms.ComboBox
$ComboBox1.Anchor = 'Top, Left'
$ComboBox1.Location = New-Object System.Drawing.Point(211, 29)
$ComboBox1.Size = New-Object System.Drawing.Size(243, 31)
$ComboBox1.Items.AddRange(@("Watm", "LandbankMS", "Landbank"))
$ComboBox1.Add_SelectedIndexChanged({
        $local:RooPathList = @("D:\Drop_Watm", "D:\Drop_LandBankMS", "D:\Landbank_Drop")
        $local:HoldPathList = @("D:\Drop_Watm\Hold_WatmSource", "D:\Drop_LandBankMS\Hold_LandBankMS_Source", "\\10.253.27.126\d$")
        $local:SourceFolderList = @("WatmSource", "LandBankMSSource", "LandBankSource")
        $local:ObjFolderList = @("WatmPublish", "LandBankMSPublish", "LandBankPublish")
        $local:FileListSettings = @($true, $true, $true)
        $textBox1.Text = $RooPathList[$ComboBox1.SelectedIndex]                                 
        $textBox2.Text = $HoldPathList[$ComboBox1.SelectedIndex]                                      
        $textBox3.Text = $SourceFolderList[$ComboBox1.SelectedIndex]                              
        $textBox4.Text = $ObjFolderList[$ComboBox1.SelectedIndex]    
        $CheckBox1.Checked = $FileListSettings[$ComboBox1.SelectedIndex] 
    })
$ComboBox1.SelectedIndex = 0

$ComboBox2 = New-Object System.Windows.Forms.ComboBox
$ComboBox2.Anchor = 'Top, Left'
$ComboBox2.Location = New-Object System.Drawing.Point(211, 213)
$ComboBox2.Size = New-Object System.Drawing.Size(331, 31)
$ComboBox2.Items.AddRange(@("yyyyMMdd"))
$ComboBox2.SelectedIndex = 0
#endregion ComboBox
#region label
$label1 = New-Object System.Windows.Forms.Label
$label1.Anchor = 'Bottom, Right'
$label1.Location = New-Object System.Drawing.Point(784, 434)
$label1.Size = New-Object System.Drawing.Size(94, 24)
$label1.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label1.Text = "換版日期"

$label2 = New-Object System.Windows.Forms.Label
$label2.Anchor = 'Bottom, Right'
$label2.Location = New-Object System.Drawing.Point(98, 32)
$label2.Size = New-Object System.Drawing.Size(94, 24)
$label2.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label2.Text = "專案代號"

$label3 = New-Object System.Windows.Forms.Label
$label3.Anchor = 'Top, Left'
$label3.Location = New-Object System.Drawing.Point(56, 144)
$label3.Size = New-Object System.Drawing.Size(136, 24)
$label3.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label3.Text = "換版文件路徑"

$label4 = New-Object System.Windows.Forms.Label
$label4.Anchor = 'Top, Left'
$label4.Location = New-Object System.Drawing.Point(56, 180)
$label4.Size = New-Object System.Drawing.Size(136, 24)
$label4.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label4.Text = "凍版文件路徑"

$label5 = New-Object System.Windows.Forms.Label
$label5.Anchor = 'Top, Left'
$label5.Location = New-Object System.Drawing.Point(56, 220)
$label5.Size = New-Object System.Drawing.Size(136, 24)
$label5.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label5.Text = "慣用日期格式"

$label6 = New-Object System.Windows.Forms.Label
$label6.Anchor = 'Top, Left'
$label6.Location = New-Object System.Drawing.Point(12, 256)
$label6.Size = New-Object System.Drawing.Size(180, 24)
$label6.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label6.Text = "source資料夾名稱"

$label7 = New-Object System.Windows.Forms.Label
$label7.Anchor = 'Top, Left'
$label7.Location = New-Object System.Drawing.Point(48, 292)
$label7.Size = New-Object System.Drawing.Size(144, 24)
$label7.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label7.Text = "obj資料夾名稱"

$label8 = New-Object System.Windows.Forms.Label
$label8.Anchor = 'Top, Left'
$label8.Location = New-Object System.Drawing.Point(77, 331)
$label8.Size = New-Object System.Drawing.Size(115, 24)
$label8.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$label8.Text = "製版機路徑"
#endregion label
#region DataGridViewColumn
$DataGridView1_col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col1.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col1.HeaderText = "檔案狀態"
$DataGridView1_col1.MinimumWidth = 6;
$DataGridView1_col1.Name = "FileStatus";
$DataGridView1_col1.ReadOnly = $false;

$DataGridView1_col2 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col2.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col2.HeaderText = "異動檔案"
$DataGridView1_col2.MinimumWidth = 6;
$DataGridView1_col2.Name = "FileName";
$DataGridView1_col2.ReadOnly = $false;

$DataGridView1_col3 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col3.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col3.HeaderText = "機器種類(AP/WEB/NULL)"
$DataGridView1_col3.MinimumWidth = 6;
$DataGridView1_col3.Name = "MachineType";
$DataGridView1_col3.ReadOnly = $false;

$DataGridView1_col4 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col4.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col4.HeaderText = "From(擷取路徑)"
$DataGridView1_col4.MinimumWidth = 6;
$DataGridView1_col4.Name = "FromPath";
$DataGridView1_col4.ReadOnly = $false;

$DataGridView1_col5 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col5.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col5.HeaderText = "To(上版路徑)"
$DataGridView1_col5.MinimumWidth = 6;
$DataGridView1_col5.Name = "ToPath";
$DataGridView1_col5.ReadOnly = $false;

$DataGridView1_col6 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$DataGridView1_col6.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill;
$DataGridView1_col6.HeaderText = "檔案說明"
$DataGridView1_col6.MinimumWidth = 6;
$DataGridView1_col6.Name = "FileMemo";
$DataGridView1_col6.ReadOnly = $false;
#endregion DataGridViewColumn
#region DataGridView
$DataGridView1 = New-Object System.Windows.Forms.DataGridView
$DataGridView1.Anchor = 'Bottom, Top, Left, Right'
$DataGridView1.Location = New-Object System.Drawing.Point(0, 0);
$DataGridView1.Name = "dataGridView1";
$DataGridView1.RowHeadersWidth = 51;
$DataGridView1.RowTemplate.Height = 27;
$DataGridView1.Size = New-Object System.Drawing.Size(1096, 417);
$DataGridView1.AllowUserToAddRows = $false
$DataGridView1.AutoSizeColumnsMode = [DataGridViewAutoSizeColumnsMode]::AllCells
$DataGridView1.AutoSizeRowsMode = [DataGridViewAutoSizeRowsMode]::DisplayedCells
$DataGridView1.ColumnHeadersHeightSizeMode = [DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DataGridView1.RowHeadersWidthSizeMode = [DataGridViewRowHeadersWidthSizeMode]::AutoSizeToAllHeaders
$DataGridView1.ScrollBars = [ScrollBars]::Both
$DataGridView1.TabIndex = 0;
$DataGridView1.AllowUserToAddRows = $false
$DataGridView1.Columns.AddRange(
    [System.Windows.Forms.DataGridViewColumn[]]@(
        $DataGridView1_col1,
        $DataGridView1_col2,
        $DataGridView1_col3,
        $DataGridView1_col4,
        $DataGridView1_col5,
        $DataGridView1_col6
    )
);
$DataGridView1.PerformLayout()
#endregion DataGridView
#region ToolStripStatusLabel1
$ToolStripStatusLabel1 = New-Object System.Windows.Forms.ToolStripStatusLabel
$ToolStripStatusLabel1.Size = New-Object System.Drawing.Size(158, 19)
$ToolStripStatusLabel1.Text = "1234"
#endregion ToolStripStatusLabel1
#region StatusStrip
$StatusStrip1 = New-Object System.Windows.Forms.StatusStrip
$StatusStrip1.ImageScalingSize = New-Object System.Drawing.Size(20, 20)
$StatusStrip1.Location = New-Object System.Drawing.Point(3, 438);
$StatusStrip1.Items.AddRange([System.Windows.Forms.ToolStripItem[]]@(
        $ToolStripStatusLabel1
    ))
#endregion StatusStrip
#region TabPage
$TabPage1 = New-Object System.Windows.Forms.TabPage
$TabPage1.Size = New-Object System.Drawing.Size(1122, 474)
$TabPage1.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$TabPage1.Text = '檔案清單'
$TabPage1.Controls.Add($datetimepciker1)
$TabPage1.Controls.Add($label1)
$TabPage1.Controls.Add($Button1)
$TabPage1.Controls.Add($Button2)
$TabPage1.Controls.Add($DataGridView1)
$TabPage1.Controls.Add($StatusStrip1)

$TabPage2 = New-Object System.Windows.Forms.TabPage
$TabPage2.Size = New-Object System.Drawing.Size(1122, 474)
$TabPage2.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$TabPage2.Text = '設定'
$TabPage2.Controls.Add($label2)
$TabPage2.Controls.Add($label3)
$TabPage2.Controls.Add($label4)
$TabPage2.Controls.Add($label5)
$TabPage2.Controls.Add($label6)
$TabPage2.Controls.Add($label7)
$TabPage2.Controls.Add($label8)
$TabPage2.Controls.Add($CheckBox1)
$TabPage2.Controls.Add($textBox1)
$TabPage2.Controls.Add($ComboBox2)
$TabPage2.Controls.Add($ComboBox1)
$TabPage2.Controls.Add($textBox2)
$TabPage2.Controls.Add($textBox3)
$TabPage2.Controls.Add($textBox4)

$TabPage3 = New-Object System.Windows.Forms.TabPage
$TabPage3.Size = New-Object System.Drawing.Size(1122, 474)
$TabPage3.Font = [Font]::new("Microsoft Sans Serif", 12, [FontStyle]::Bold)
$TabPage3.Text = '說明'
$TabPage3.Controls.Add($instruction)

#endregion TabPage
#region TabControl
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Anchor = 'Bottom, Top, Left, Right'
$TabControl.Location = New-Object System.Drawing.Point(0, 1)
$TabControl.Size = New-Object System.Drawing.Size(1130, 510)
$TabControl.Font = [Font]::new("Microsoft Sans Serif", 12)
$TabControl.SelectedIndex = 0
$TabControl.TabIndex = 0
$TabControl.Controls.Add($TabPage1) #檔案清單
$TabControl.Controls.Add($TabPage2) #全域設定
$TabControl.Controls.Add($TabPage3) #說明
#endregion TabControl
#region Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "換版工具程式"
$form.Size = New-Object System.Drawing.Size(1130, 550)
$form.StartPosition = 'CenterScreen'
$form.Controls.Add($TabControl)
#endregion Form
#endregion GUI