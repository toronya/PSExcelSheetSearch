Set-Location $PSScriptRoot

$appVversion = "0.2"
$appReleaseDate = "2024-11-26"

# debug message �\�����邩
$debugOut = $false

# �K�v�ȃ��W���[�����C���|�[�g
Add-Type -AssemblyName System.Windows.Forms
#Import-Module ImportExcel
# �C���X�g�[�������Ƀ��[�J���t�H���_�Ƀ��C�u�����������Ďg���ꍇ
#Import-Module "$PSScriptRoot\ImportExcel-7.8.10\ImportExcel.psd1"
#Add-Type -Path "$PSScriptRoot\ImportExcel-7.8.10\EPPlus.dll"

# �ݒ�t�@�C���̃p�X
$configFilePath = "$PSScriptRoot\config.json"

# �ݒ�t�@�C�������݂���ꍇ�A�ݒ��ǂݍ���
if (Test-Path $configFilePath) {
    $config = Get-Content $configFilePath | ConvertFrom-Json
    $importExcelPath = $config.ImportExcelPath
} else {
    $importExcelPath = ""
}

# �ݒ��ۑ�����֐�
function Save-Config {
    param (
        [string]$importExcelPath
    )
    $config = @{
        ImportExcelPath = $importExcelPath
    }
    $config | ConvertTo-Json | Set-Content -Path $configFilePath
}

# Import-Module �� Add-Type �̎��s
if ($importExcelPath) {
    Import-Module "$importExcelPath\ImportExcel.psd1"
    Add-Type -Path "$importExcelPath\EPPlus.dll"
}

# Visual Styles��L���ɂ���
[System.Windows.Forms.Application]::EnableVisualStyles()

# �t�H�[���̍쐬
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Sheet Search"
$form.Size = New-Object System.Drawing.Size(1000, 600)

# �t�H�[���̕\���ʒu�𒆉��ɐݒ�
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

#--

# ���j���[�o�[�̍쐬
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# �t�@�C�����j���[�̍쐬
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("�t�@�C��")
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("�I��")
$exitMenuItem.Add_Click({ $form.Close() })
[void]$fileMenu.DropDownItems.Add($exitMenuItem)

# �I�v�V�������j���[�̍쐬
$optionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("�I�v�V����")
$useComObjectMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("ComObject ���g����Excel�ɐڑ�����")
[void]$optionsMenu.DropDownItems.Add($useComObjectMenuItem)
# ImportExcel �t�H���_�̃p�X��ݒ肷�郁�j���[�A�C�e���̍쐬
$setImportExcelPathMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("ImportExcel �t�H���_�̃p�X��ݒ�iImportExcel���C���X�g�[�������Ɏg���ꍇ�j")
[void]$optionsMenu.DropDownItems.Add($setImportExcelPathMenuItem)

# �w���v���j���[�̍쐬
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("�w���v")

$usageMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("�g����")
$usageMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show(
            "�G�N�Z���V�[�g�𕶎���������A�v���ł��B`n`n" + 
            " - �������܂܂��s������\�����܂��B`n" + 
            " - �ǂݍ��ރG�N�Z���V�[�g�͂P�s�ڂ��w�b�_�[�s�ɂȂ�悤�ɂ��Ă��������B") })
[void]$helpMenu.DropDownItems.Add($usageMenuItem)

$versionMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("�o�[�W����")
$versionMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("�o�[�W�����F" + $appVversion + " (" + $appReleaseDate + ")") })
[void]$helpMenu.DropDownItems.Add($versionMenuItem)


# ���j���[�o�[�Ƀ��j���[��ǉ�
[void]$menuStrip.Items.Add($fileMenu)
#[void]$menuStrip.Items.Add($optionsMenu)
[void]$menuStrip.Items.Add($helpMenu)

# �t�H�[���Ƀ��j���[�o�[��ǉ�
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

#--

# �I�v�V�������j���[�̃C�x���g
$setImportExcelPathMenuItem.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $importExcelPath = $folderBrowserDialog.SelectedPath
        Save-Config -importExcelPath $importExcelPath
        [System.Windows.Forms.MessageBox]::Show("�ݒ肪�ۑ�����܂����B")
        Import-Module "$importExcelPath\ImportExcel.psd1"
        Add-Type -Path "$importExcelPath\EPPlus.dll"
        $setImportExcelPathMenuItem.Checked = $true
        $useComObjectMenuItem.Checked = $false
    }
})

$useComObjectMenuItem.Add_Click({
    if ($useComObjectMenuItem.Checked -eq $true) {
        # ComObject �g��Ȃ�
        $useComObjectMenuItem.Checked = $false
    }
    else{
        # ComObject �g���ɂ���
        $useComObjectMenuItem.Checked = $true
        $setImportExcelPathMenuItem.Checked = $false
    }
})

function Check-Confg {
    # ComObject �� Importxcel ��
    # �ݒ�t�@�C������p�X���Ǎ��ł��Ă��邩
    if ($importExcelPath -ne ""){
        # use Importxcel
        $setImportExcelPathMenuItem.Checked = $true
        $useComObjectMenuItem.Checked = $false
    }
    else{
        # use ComObject
        $useComObjectMenuItem.Checked = $true
        $setImportExcelPathMenuItem.Checked = $false
    }
}

Check-Confg

#--

# �e�[�u�����C�A�E�g�p�l���̍쐬
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
# �e�[�u�����C�A�E�g�p�l���̍쐬
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
#$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill  # ���j���[�o�[�Ɗ����邽��Fill�͂��Ȃ�
$tableLayoutPanel.Top = $menuStrip.Height  # ���j���[�o�[�̍������������ɔz�u
$tableLayoutPanel.Height = $form.ClientSize.Height - $menuStrip.Height  # �c��̍�����ݒ�
$tableLayoutPanel.Width = $form.ClientSize.Width  # �t�H�[���̕��ɍ��킹��
$tableLayoutPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
    [System.Windows.Forms.AnchorStyles]::Left -bor 
    [System.Windows.Forms.AnchorStyles]::Right -bor 
    [System.Windows.Forms.AnchorStyles]::Bottom

# �e�[�u�����C�A�E�g�p�l�����t�H�[���ɒǉ�
[void]$form.Controls.Add($tableLayoutPanel)


# �e�[�u�����C�A�E�g�p�l���̗�ƍs�̐ݒ�
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 0���
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))  # 1���
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 2���
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 3���

[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 0�s��
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 1�s��
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 2�s��
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))       # 3�s��


# UI �e�[�u�����C�A�E�g
#
#  Controls.Add( XXXXXX,  ��, �s)
#
#  | 0   | 1   | 2   | 3   |
#  |-----|-----|-----|-----|
# 0| 0,0 | 1,0 | 2,0 | 3,0 |
# 1| 0,1 | 1,1 | 2,1 | 3,1 |
# 2| 0,2 | 1,2 | 2,2 | 3,2 |
# 3| 0,3 | 1,3 | 2,3 | 3,3 |
#

#-- 0 �s��

# ���x���̍쐬�i�t�@�C���p�X�\���p�j
$filePathLabel = New-Object System.Windows.Forms.Label
$filePathLabel.Text = "�G�N�Z���F"
$filePathLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$filePathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($filePathLabel, 0, 0)

# �e�L�X�g�{�b�N�X�̍쐬�i�t�@�C���p�X�\���p�j
$filePathTextBox = New-Object System.Windows.Forms.TextBox
$filePathTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$filePathTextBox.ReadOnly = $true
$tableLayoutPanel.Controls.Add($filePathTextBox, 1, 0)
$tableLayoutPanel.SetColumnSpan($filePathTextBox, 2)

# �t�@�C���I���{�^���̍쐬
$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Text = "�I��"
$fileButton.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($fileButton, 3, 0)


#-- 1 �s��

# �v���_�E�����j���[���x���̍쐬�i�V�[�g�I���j
$sheetLabel = New-Object System.Windows.Forms.Label
$sheetLabel.Text = "�V�[�g�F"
$sheetLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$sheetLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($sheetLabel, 0, 1)

# �v���_�E�����j���[�̍쐬�i�V�[�g�I���j
$sheetComboBox = New-Object System.Windows.Forms.ComboBox
$sheetComboBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($sheetComboBox, 1, 1)

#-- 2 �s��

# ���x���̍쐬�i�����p�j
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "�����F"
$searchLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$searchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($searchLabel, 0, 2)

# �e�L�X�g�{�b�N�X�̍쐬�i�����p�j
$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($searchTextBox, 1, 2)
$tableLayoutPanel.SetColumnSpan($searchTextBox, 2)

# �����{�^���̍쐬
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "����"
$searchButton.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($searchButton, 3, 2)

#-- 3 �s��

# �f�[�^�O���b�h�r���[�̍쐬
$dataGridView = New-Object System.Windows.Forms.DataGridView
#
#$dataGridView.VirtualMode = $true

# ������h�~�̂��߁ADoubleBuffered�v���p�e�B��L���ɂ���
$dataGridView.GetType().GetProperty('DoubleBuffered', 
    [System.Reflection.BindingFlags]::NonPublic -bor 
    [System.Reflection.BindingFlags]::Instance).SetValue($dataGridView, $true, $null
    )

# �s�ԍ���\������C�x���g�n���h����ǉ� �i�W���ł͍s�ԍ��͏o�Ȃ����ߕ`�ʂ���j
$dataGridView.add_RowPostPaint({
        param($sndr, $e)
        $brush = New-Object System.Drawing.SolidBrush($sndr.RowHeadersDefaultCellStyle.ForeColor)
        $format = New-Object System.Drawing.StringFormat
        $format.Alignment = [System.Drawing.StringAlignment]::Far
        $format.LineAlignment = [System.Drawing.StringAlignment]::Center
        $e.Graphics.DrawString(($e.RowIndex + 1).ToString(),
            $dataGridView.Font,
            $brush,
            $e.RowBounds.Location.X + 30,
            $e.RowBounds.Location.Y + 10,
            $format)
    })

$dataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($dataGridView, 0, 3)
$tableLayoutPanel.SetColumnSpan($dataGridView, 4)

#--

# �t�@�C���I���_�C�A���O�̍쐬
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls"

# �t�@�C���I���{�^���̃N���b�N�C�x���g
$fileButton.Add_Click({
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $excelPath = $openFileDialog.FileName
            $filePathTextBox.Text = $excelPath
            $sheetComboBox.Items.Clear()

            if ($useComObjectMenuItem.Checked) {
                # ComObject��
                # �G�N�Z���t�@�C���̓ǂݍ���
                $excel = New-Object -ComObject Excel.Application
                # �ǂݎ���p���[�h
                $workbook = $excel.Workbooks.Open($excelPath, [Type]::Missing, $true)

                # �V�[�g�����v���_�E�����j���[�ɒǉ�
                foreach ($worksheet in $workbook.Worksheets ) {
                    $sheetComboBox.Items.Add($worksheet.Name)
                }
                $workbook.Close($false)
                $excel.Quit()

                # �v���Z�X���c��̂Ń����[�X����
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            else{
                #ImportExcel��
                # �G�N�Z���t�@�C���̓ǂݍ���
                $excelPackage = [OfficeOpenXml.ExcelPackage]::new()
                #$fileStream = [System.IO.File]::OpenRead($excelPath)
                $fileStream = [System.IO.File]::Open($excelPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $excelPackage.Load($fileStream)
                $fileStream.Close()

                # �V�[�g�����v���_�E�����j���[�ɒǉ�
                foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
                    $sheetComboBox.Items.Add($worksheet.Name)
                }
            }
        
            # �P�ڂ̃V�[�g���������l�őI����Ԃɂ���
            $sheetComboBox.SelectedIndex = 0
        }
    })


# ���������̒�`
$searchAction = {
    try {
        $excelPath = $filePathTextBox.Text
        $searchString = $searchTextBox.Text
        $selectedSheet = $sheetComboBox.SelectedItem

        $i = 1

        if ($excelPath -and $selectedSheet) {

            # �f�[�^�e�[�u���̍쐬
            If ($debugOut -eq $true) { Write-Host "�f�[�^�e�[�u���̍쐬" }
            $dataTable = New-Object System.Data.DataTable

            If ($debugOut -eq $true) { Write-Host "�G�N�Z���t�@�C���̓ǂݍ���" }

            if ($useComObjectMenuItem.Checked) {
                # COM�I�u�W�F�N�g���g�p��������
                $excel = New-Object -ComObject Excel.Application
                # �ǂݎ���p���[�h
                $workbook = $excel.Workbooks.Open($excelPath, [Type]::Missing, $true)
                $worksheet = $workbook.Sheets.Item($selectedSheet)
                $range = $worksheet.UsedRange
                $rowCount = $range.Rows.Count
                $colCount = $range.Columns.Count
            
                # �w�b�_�[�̒ǉ�
                for ($col = 1; $col -le $colCount; $col++) {
                    $dataTable.Columns.Add($range.Cells.Item(1, $col).Text)
                }
            
                # �f�[�^�̌���
                for ($row = 2; $row -le $rowCount; $row++) {
                    $rowData = @()
                    for ($col = 1; $col -le $colCount; $col++) {
                        $rowData += $range.Cells.Item($row, $col).Text
                    }
                    if ($rowData -join " " -match $searchString) {
                        If ($debugOut -eq $true) { Write-Host "${i}`t${rowData}" }
                        $dataTable.Rows.Add($rowData)
                    }
                    $i++
                }
                $workbook.Close($false)
                $excel.Quit()
                
                # �v���Z�X���c��̂Ń����[�X����
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            else {
                # Import-Excel���W���[�����g�p��������
                Import-Excel -Path $excelPath -WorksheetName $selectedSheet  | ForEach-Object {
                    $data = $_

                    if ($i -eq 1) {
                        # �w�b�_�[�̒ǉ�
                        If ($debugOut -eq $true) { Write-Host "�w�b�_�[�̒ǉ�" }
                        $data.PSObject.Properties.Name | ForEach-Object { $dataTable.Columns.Add($_) }
                    }

                    # �S���ڂɂ������Č���
                    foreach ($row in $data) {
                        $rowData = @()
                        foreach ($column in $data.PSObject.Properties.Name) {
                            $rowData += $row.$column
                        }
                        if ($rowData -join " " -match $searchString) {
                            If ($debugOut -eq $true) { Write-Host "${i}`t${rowData}" }
                            $dataTable.Rows.Add($rowData)
                        }
                    }
                    $i = $i + 1
                }
            }
            # �f�[�^�O���b�h�r���[�Ƀf�[�^���o�C���h
            If ($debugOut -eq $true) { Write-Host "�f�[�^�O���b�h�r���[�Ƀf�[�^���o�C���h" }
            $dataGridView.DataSource = $dataTable
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("�t�@�C���ƃV�[�g��I�����Ă��������B")
        }
    }
    catch {
        $errorMessage = "�G���[���������܂���: $_"
        $errorDetails = $_.Exception.Message
        $stckTrc = $_.Exception.StackTrace
        [System.Windows.Forms.MessageBox]::Show("$errorMessage`n`n�ڍ�: $errorDetails`n`n�X�^�b�N�g���[�X: $stckTrc")
        if ($null -ne $excel){
            $excel.Quit()
        }
    }
}


# �����{�^���̃N���b�N�C�x���g
$searchButton.Add_Click($searchAction)

# �e�L�X�g�{�b�N�X�̃L�[�����C�x���g
$searchTextBox.Add_KeyDown({
        param($sndr, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $searchAction.Invoke()
        }
    })

# �t�H�[���̕\��
[void]$form.ShowDialog()
