# ManualBuilder の操作記録を安全に検証するための架空業務画面。
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class MbRecorderFixtureNative {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int command);
}
'@

$form = New-Object Windows.Forms.Form
$form.Text = '経費申請 操作テスト'
$form.ClientSize = New-Object Drawing.Size 620, 360
$form.StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
$form.Font = New-Object Drawing.Font 'Yu Gothic UI', 10
$form.BackColor = [Drawing.Color]::White
$form.Add_Shown({ [void][MbRecorderFixtureNative]::ShowWindow($form.Handle, 5); $form.Activate() })

$title = New-Object Windows.Forms.Label
$title.Text = '経費申請の検索'
$title.Font = New-Object Drawing.Font 'Yu Gothic UI', 16, ([Drawing.FontStyle]::Bold)
$title.Location = New-Object Drawing.Point 28, 24
$title.AutoSize = $true
$form.Controls.Add($title)

$numberLabel = New-Object Windows.Forms.Label
$numberLabel.Text = '申請番号'
$numberLabel.Location = New-Object Drawing.Point 30, 86
$numberLabel.AutoSize = $true
$form.Controls.Add($numberLabel)

$numberBox = New-Object Windows.Forms.TextBox
$numberBox.Name = 'ApplicationNumber'
$numberBox.AccessibleName = '申請番号'
$numberBox.Location = New-Object Drawing.Point 30, 112
$numberBox.Size = New-Object Drawing.Size 250, 30
$form.Controls.Add($numberBox)

$statusLabel = New-Object Windows.Forms.Label
$statusLabel.Text = '状態'
$statusLabel.Location = New-Object Drawing.Point 310, 86
$statusLabel.AutoSize = $true
$form.Controls.Add($statusLabel)

$statusBox = New-Object Windows.Forms.ComboBox
$statusBox.Name = 'ApplicationStatus'
$statusBox.AccessibleName = '状態'
$statusBox.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
$statusBox.Items.AddRange(@('すべて', '申請中', '承認済み'))
$statusBox.SelectedIndex = 0
$statusBox.Location = New-Object Drawing.Point 310, 112
$statusBox.Size = New-Object Drawing.Size 170, 30
$form.Controls.Add($statusBox)

$searchButton = New-Object Windows.Forms.Button
$searchButton.Text = '検索'
$searchButton.Name = 'SearchButton'
$searchButton.AccessibleName = '検索'
$searchButton.Location = New-Object Drawing.Point 500, 108
$searchButton.Size = New-Object Drawing.Size 90, 36
$form.Controls.Add($searchButton)

$resultPanel = New-Object Windows.Forms.Panel
$resultPanel.Location = New-Object Drawing.Point 30, 176
$resultPanel.Size = New-Object Drawing.Size 560, 116
$resultPanel.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
$resultPanel.BackColor = [Drawing.Color]::FromArgb(247, 249, 252)
$form.Controls.Add($resultPanel)

$resultLabel = New-Object Windows.Forms.Label
$resultLabel.Text = '検索条件を入力してください。'
$resultLabel.Location = New-Object Drawing.Point 18, 16
$resultLabel.Size = New-Object Drawing.Size 520, 28
$resultPanel.Controls.Add($resultLabel)

$detailButton = New-Object Windows.Forms.Button
$detailButton.Text = '詳細を表示'
$detailButton.Name = 'DetailButton'
$detailButton.AccessibleName = '詳細を表示'
$detailButton.Location = New-Object Drawing.Point 400, 60
$detailButton.Size = New-Object Drawing.Size 130, 36
$detailButton.Visible = $false
$resultPanel.Controls.Add($detailButton)

$searchButton.Add_Click({
    $searchButton.Enabled = $false
    $resultLabel.Text = '検索しています…'
    $script:fixtureTimer = New-Object Windows.Forms.Timer
    $script:fixtureTimer.Interval = 450
    $script:fixtureTimer.Add_Tick({
        $script:fixtureTimer.Stop(); $script:fixtureTimer.Dispose(); $script:fixtureTimer = $null
        $resultLabel.Text = 'REQ-1042　出張旅費申請　申請中'
        $detailButton.Visible = $true
        $searchButton.Enabled = $true
    })
    $script:fixtureTimer.Start()
})

$detailButton.Add_Click({
    $resultLabel.Text = 'REQ-1042 の詳細を表示しました。合計 12,800 円'
    $detailButton.Text = '表示済み'
    $detailButton.Enabled = $false
})

[void]$form.ShowDialog()
