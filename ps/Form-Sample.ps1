Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data

# Formとその関連コントロールを定義する関数
function Create-DesertListForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "デザートリスト"
    $form.ClientSize = New-Object System.Drawing.Size(300, 250)
    $form.StartPosition = "CenterScreen"

    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Dock = "Fill"
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.AllowUserToDeleteRows = $false
    $dataGridView.ReadOnly = $true
    $dataGridView.SelectionMode = "FullRowSelect"

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Bottom"
    $panel.Height = 40

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "選択したデザートを表示"
    $button.AutoSize = $true
    $button.Left = 8
    $button.Top = ($panel.Height - $button.Height) / 2

    $button.Add_Click({
        $selectedRow = $dataGridView.CurrentRow
        if ($selectedRow -ne $null) {
            $selectedDessert = $selectedRow.Cells["デザート"].Value
            $price = $selectedRow.Cells["値段"].Value
            if ($selectedDessert) {
                [System.Windows.Forms.MessageBox]::Show("選択されたデザート: $selectedDessert, 値段: $price 円", "情報")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("デザートを選択してください。", "注意")
        }
    })

    $panel.Controls.Add($button)
    $form.Controls.Add($dataGridView)
    $form.Controls.Add($panel)

    # フォームとデータグリッドビューを返す
    return $form, $dataGridView
}

# フォームを表示する関数
function Show-DesertListForm {
    param (
        [System.Data.DataTable]$DataTable
    )
    $form, $dataGridView = Create-DesertListForm
    $dataGridView.DataSource = $DataTable
    $null = $form.ShowDialog()
}

# 使用例
$sampleDataTable = New-Object System.Data.DataTable
$sampleDataTable.Columns.Add("デザート") | Out-Null
$sampleDataTable.Columns.Add("値段", [int]) | Out-Null

$sampleDataTable.Rows.Add("モンブラン", 500) | Out-Null
$sampleDataTable.Rows.Add("イチゴショート", 450) | Out-Null
$sampleDataTable.Rows.Add("バナナパフェ", 600) | Out-Null

# フォームを表示
Show-DesertListForm -DataTable $sampleDataTable
