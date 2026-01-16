Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -------------------------
# Audio-Datei auswählen
# -------------------------
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Audiodateien|*.mp3;*.wav;*.m4a;*.flac"
$openFileDialog.Title  = "Audiodatei auswählen"

if ($openFileDialog.ShowDialog() -ne "OK") { exit }
$audioPath = $openFileDialog.FileName

Write-Host "$audioPath"

# -------------------------
# Modell-Auswahl GUI
# -------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Whisper - Modell auswählen"
$form.Size = New-Object System.Drawing.Size(300,150)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Whisper-Modell:"
$label.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($label)

$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(120,18)
$comboBox.Width = 140
$comboBox.DropDownStyle = "DropDownList"
$comboBox.Items.AddRange(@("tiny","base","small","medium","large","turbo"))
$comboBox.SelectedIndex = 2
$form.Controls.Add($comboBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100,60)
$okButton.Add_Click({ $form.Close() })
$form.Controls.Add($okButton)

$form.ShowDialog()
$model = $comboBox.SelectedItem

# -------------------------
# Ausgabeordner auswählen
# -------------------------
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "Speicherort für Transkript auswählen"
$folderDialog.ShowNewFolderButton = $true

if ($folderDialog.ShowDialog() -ne "OK") { exit }
$outputDir = $folderDialog.SelectedPath


# -------------------------
# Whisper ausführen
# -------------------------
$modelDir = "C:\Users\Admin\Skripte\whisperenv\models"

& "c:\users\admin\skripte\whisperenv\Scripts\activate.ps1"

whisper $audioPath --model_dir "$modelDir" --model $model --output_format txt --output_dir $outputDir

pause

