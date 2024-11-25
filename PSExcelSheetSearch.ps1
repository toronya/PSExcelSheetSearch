Set-Location $PSScriptRoot

$appVversion = "0.2"
$appReleaseDate = "2024-11-26"

# debug message 表示するか
$debugOut = $false

# 必要なモジュールをインポート
Add-Type -AssemblyName System.Windows.Forms
#Import-Module ImportExcel
# インストールせずにローカルフォルダにライブラリをおいて使う場合
#Import-Module "$PSScriptRoot\ImportExcel-7.8.10\ImportExcel.psd1"
#Add-Type -Path "$PSScriptRoot\ImportExcel-7.8.10\EPPlus.dll"

# 設定ファイルのパス
$configFilePath = "$PSScriptRoot\config.json"

# 設定ファイルが存在する場合、設定を読み込む
if (Test-Path $configFilePath) {
    $config = Get-Content $configFilePath | ConvertFrom-Json
    $importExcelPath = $config.ImportExcelPath
} else {
    $importExcelPath = ""
}

# 設定を保存する関数
function Save-Config {
    param (
        [string]$importExcelPath
    )
    $config = @{
        ImportExcelPath = $importExcelPath
    }
    $config | ConvertTo-Json | Set-Content -Path $configFilePath
}

# Import-Module と Add-Type の実行
if ($importExcelPath) {
    Import-Module "$importExcelPath\ImportExcel.psd1"
    Add-Type -Path "$importExcelPath\EPPlus.dll"
}

# Visual Stylesを有効にする
[System.Windows.Forms.Application]::EnableVisualStyles()

# フォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "Excel Sheet Search"
$form.Size = New-Object System.Drawing.Size(1000, 600)

# フォームの表示位置を中央に設定
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

#--

# メニューバーの作成
$menuStrip = New-Object System.Windows.Forms.MenuStrip

# ファイルメニューの作成
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("ファイル")
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("終了")
$exitMenuItem.Add_Click({ $form.Close() })
[void]$fileMenu.DropDownItems.Add($exitMenuItem)

# オプションメニューの作成
$optionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("オプション")
$useComObjectMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("ComObject を使ってExcelに接続する")
[void]$optionsMenu.DropDownItems.Add($useComObjectMenuItem)
# ImportExcel フォルダのパスを設定するメニューアイテムの作成
$setImportExcelPathMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("ImportExcel フォルダのパスを設定（ImportExcelをインストールせずに使う場合）")
[void]$optionsMenu.DropDownItems.Add($setImportExcelPathMenuItem)

# ヘルプメニューの作成
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("ヘルプ")

$usageMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("使い方")
$usageMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show(
            "エクセルシートを文字検索するアプリです。`n`n" + 
            " - 文字が含まれる行だけを表示します。`n" + 
            " - 読み込むエクセルシートは１行目がヘッダー行になるようにしてください。") })
[void]$helpMenu.DropDownItems.Add($usageMenuItem)

$versionMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("バージョン")
$versionMenuItem.Add_Click({ [System.Windows.Forms.MessageBox]::Show("バージョン：" + $appVversion + " (" + $appReleaseDate + ")") })
[void]$helpMenu.DropDownItems.Add($versionMenuItem)


# メニューバーにメニューを追加
[void]$menuStrip.Items.Add($fileMenu)
#[void]$menuStrip.Items.Add($optionsMenu)
[void]$menuStrip.Items.Add($helpMenu)

# フォームにメニューバーを追加
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

#--

# オプションメニューのイベント
$setImportExcelPathMenuItem.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $importExcelPath = $folderBrowserDialog.SelectedPath
        Save-Config -importExcelPath $importExcelPath
        [System.Windows.Forms.MessageBox]::Show("設定が保存されました。")
        Import-Module "$importExcelPath\ImportExcel.psd1"
        Add-Type -Path "$importExcelPath\EPPlus.dll"
        $setImportExcelPathMenuItem.Checked = $true
        $useComObjectMenuItem.Checked = $false
    }
})

$useComObjectMenuItem.Add_Click({
    if ($useComObjectMenuItem.Checked -eq $true) {
        # ComObject 使わない
        $useComObjectMenuItem.Checked = $false
    }
    else{
        # ComObject 使うにする
        $useComObjectMenuItem.Checked = $true
        $setImportExcelPathMenuItem.Checked = $false
    }
})

function Check-Confg {
    # ComObject か Importxcel か
    # 設定ファイルからパスが読込できているか
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

# テーブルレイアウトパネルの作成
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
# テーブルレイアウトパネルの作成
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
#$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill  # メニューバーと干渉するためFillはしない
$tableLayoutPanel.Top = $menuStrip.Height  # メニューバーの高さ分だけ下に配置
$tableLayoutPanel.Height = $form.ClientSize.Height - $menuStrip.Height  # 残りの高さを設定
$tableLayoutPanel.Width = $form.ClientSize.Width  # フォームの幅に合わせる
$tableLayoutPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
    [System.Windows.Forms.AnchorStyles]::Left -bor 
    [System.Windows.Forms.AnchorStyles]::Right -bor 
    [System.Windows.Forms.AnchorStyles]::Bottom

# テーブルレイアウトパネルをフォームに追加
[void]$form.Controls.Add($tableLayoutPanel)


# テーブルレイアウトパネルの列と行の設定
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 0列目
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))  # 1列目
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 2列目
[void]$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))  # 3列目

[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 0行目
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 1行目
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))           # 2行目
[void]$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))       # 3行目


# UI テーブルレイアウト
#
#  Controls.Add( XXXXXX,  列, 行)
#
#  | 0   | 1   | 2   | 3   |
#  |-----|-----|-----|-----|
# 0| 0,0 | 1,0 | 2,0 | 3,0 |
# 1| 0,1 | 1,1 | 2,1 | 3,1 |
# 2| 0,2 | 1,2 | 2,2 | 3,2 |
# 3| 0,3 | 1,3 | 2,3 | 3,3 |
#

#-- 0 行目

# ラベルの作成（ファイルパス表示用）
$filePathLabel = New-Object System.Windows.Forms.Label
$filePathLabel.Text = "エクセル："
$filePathLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$filePathLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($filePathLabel, 0, 0)

# テキストボックスの作成（ファイルパス表示用）
$filePathTextBox = New-Object System.Windows.Forms.TextBox
$filePathTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$filePathTextBox.ReadOnly = $true
$tableLayoutPanel.Controls.Add($filePathTextBox, 1, 0)
$tableLayoutPanel.SetColumnSpan($filePathTextBox, 2)

# ファイル選択ボタンの作成
$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Text = "選択"
$fileButton.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($fileButton, 3, 0)


#-- 1 行目

# プルダウンメニューラベルの作成（シート選択）
$sheetLabel = New-Object System.Windows.Forms.Label
$sheetLabel.Text = "シート："
$sheetLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$sheetLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($sheetLabel, 0, 1)

# プルダウンメニューの作成（シート選択）
$sheetComboBox = New-Object System.Windows.Forms.ComboBox
$sheetComboBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($sheetComboBox, 1, 1)

#-- 2 行目

# ラベルの作成（検索用）
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "文字："
$searchLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$searchLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tableLayoutPanel.Controls.Add($searchLabel, 0, 2)

# テキストボックスの作成（検索用）
$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($searchTextBox, 1, 2)
$tableLayoutPanel.SetColumnSpan($searchTextBox, 2)

# 検索ボタンの作成
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "検索"
$searchButton.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($searchButton, 3, 2)

#-- 3 行目

# データグリッドビューの作成
$dataGridView = New-Object System.Windows.Forms.DataGridView
#
#$dataGridView.VirtualMode = $true

# ちらつき防止のため、DoubleBufferedプロパティを有効にする
$dataGridView.GetType().GetProperty('DoubleBuffered', 
    [System.Reflection.BindingFlags]::NonPublic -bor 
    [System.Reflection.BindingFlags]::Instance).SetValue($dataGridView, $true, $null
    )

# 行番号を表示するイベントハンドラを追加 （標準では行番号は出ないため描写する）
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

# ファイル選択ダイアログの作成
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xls"

# ファイル選択ボタンのクリックイベント
$fileButton.Add_Click({
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $excelPath = $openFileDialog.FileName
            $filePathTextBox.Text = $excelPath
            $sheetComboBox.Items.Clear()

            if ($useComObjectMenuItem.Checked) {
                # ComObject版
                # エクセルファイルの読み込み
                $excel = New-Object -ComObject Excel.Application
                # 読み取り専用モード
                $workbook = $excel.Workbooks.Open($excelPath, [Type]::Missing, $true)

                # シート名をプルダウンメニューに追加
                foreach ($worksheet in $workbook.Worksheets ) {
                    $sheetComboBox.Items.Add($worksheet.Name)
                }
                $workbook.Close($false)
                $excel.Quit()

                # プロセスが残るのでリリースする
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            else{
                #ImportExcel版
                # エクセルファイルの読み込み
                $excelPackage = [OfficeOpenXml.ExcelPackage]::new()
                #$fileStream = [System.IO.File]::OpenRead($excelPath)
                $fileStream = [System.IO.File]::Open($excelPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $excelPackage.Load($fileStream)
                $fileStream.Close()

                # シート名をプルダウンメニューに追加
                foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
                    $sheetComboBox.Items.Add($worksheet.Name)
                }
            }
        
            # １つ目のシート名を初期値で選択状態にする
            $sheetComboBox.SelectedIndex = 0
        }
    })


# 検索処理の定義
$searchAction = {
    try {
        $excelPath = $filePathTextBox.Text
        $searchString = $searchTextBox.Text
        $selectedSheet = $sheetComboBox.SelectedItem

        $i = 1

        if ($excelPath -and $selectedSheet) {

            # データテーブルの作成
            If ($debugOut -eq $true) { Write-Host "データテーブルの作成" }
            $dataTable = New-Object System.Data.DataTable

            If ($debugOut -eq $true) { Write-Host "エクセルファイルの読み込み" }

            if ($useComObjectMenuItem.Checked) {
                # COMオブジェクトを使用した検索
                $excel = New-Object -ComObject Excel.Application
                # 読み取り専用モード
                $workbook = $excel.Workbooks.Open($excelPath, [Type]::Missing, $true)
                $worksheet = $workbook.Sheets.Item($selectedSheet)
                $range = $worksheet.UsedRange
                $rowCount = $range.Rows.Count
                $colCount = $range.Columns.Count
            
                # ヘッダーの追加
                for ($col = 1; $col -le $colCount; $col++) {
                    $dataTable.Columns.Add($range.Cells.Item(1, $col).Text)
                }
            
                # データの検索
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
                
                # プロセスが残るのでリリースする
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            else {
                # Import-Excelモジュールを使用した検索
                Import-Excel -Path $excelPath -WorksheetName $selectedSheet  | ForEach-Object {
                    $data = $_

                    if ($i -eq 1) {
                        # ヘッダーの追加
                        If ($debugOut -eq $true) { Write-Host "ヘッダーの追加" }
                        $data.PSObject.Properties.Name | ForEach-Object { $dataTable.Columns.Add($_) }
                    }

                    # 全項目にたいして検索
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
            # データグリッドビューにデータをバインド
            If ($debugOut -eq $true) { Write-Host "データグリッドビューにデータをバインド" }
            $dataGridView.DataSource = $dataTable
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("ファイルとシートを選択してください。")
        }
    }
    catch {
        $errorMessage = "エラーが発生しました: $_"
        $errorDetails = $_.Exception.Message
        $stckTrc = $_.Exception.StackTrace
        [System.Windows.Forms.MessageBox]::Show("$errorMessage`n`n詳細: $errorDetails`n`nスタックトレース: $stckTrc")
        if ($null -ne $excel){
            $excel.Quit()
        }
    }
}


# 検索ボタンのクリックイベント
$searchButton.Add_Click($searchAction)

# テキストボックスのキー押下イベント
$searchTextBox.Add_KeyDown({
        param($sndr, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $searchAction.Invoke()
        }
    })

# フォームの表示
[void]$form.ShowDialog()
