<#
    .SYNOPSIS
    MassPrint
    Version: 0.07 18.01.2019
    
    © Anton Kosenko mail:Anton.Kosenko@gmail.com
    Licensed under the Apache License, Version 2.0

    .DESCRIPTION
    This script print many files in folder
#>

# requires -version 3

# Declare Variable
    $LogFile = "./info.log"
    $nt=0
    $nf=0  
# Start writing log
    Start-Transcript -path "$LogFile" -append
#   Main function
function Main {
# Description GUI
    Add-Type -assembly System.Windows.Forms
    $Form=New-Object System.Windows.Forms.Form
    $Form.ClientSize  = New-object System.Drawing.Size(600,200)
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
    $Form.MaximizeBox = $false
    $Form.BackColor = "0xDCDCDC"
    $Form.Text="Mass Printing"
# Select folder	
    $LabelFolder = New-Object System.Windows.Forms.Label
    $LabelFolder.Text = "Select folder"
    $LabelFolder.Location  = New-Object System.Drawing.Point(20,25)
    $LabelFolder.AutoSize = $true
    $Form.Controls.Add($LabelFolder)
# Description TextBox "Select"
    $TextBoxFolder = New-Object System.Windows.Forms.TextBox
    $TextBoxFolder.Location  = New-Object System.Drawing.Point(150,25)
    $TextBoxFolder.Size = New-Object System.Drawing.Size(330,20)
    $TextBoxFolder.Text = $Folder
    $Form.Controls.Add($TextBoxFolder)
# Description TextBox "Select"
    $buttonFolder = New-Object System.Windows.Forms.Button
    $buttonFolder.Text = 'Select'
    $buttonFolder.Location = New-Object System.Drawing.Point(500,25)
    $buttonFolder.Add_Click({
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            RootFolder = 'MyComputer'
            ShowNewFolderButton = $false
            }
        [void]$FolderBrowser.ShowDialog()
        $Script:Folder = $FolderBrowser.SelectedPath
        $Script:nt = (Get-ChildItem $Folder | Where-Object {$_.Extension -notmatch "printed"}).Count
        $Form.Refresh()
        $TextBoxFolder.Text = $Folder
        $LabelCount.Text = "In folder $nt files."
        $LabelProgress.Text = "Printed $nf from $Script:nt files"
        })
    $Form.Controls.Add($buttonFolder)
# Description text field "count files"
    $LabelCount = New-Object System.Windows.Forms.Label
    $LabelCount.Font = New-Object System.Drawing.Font("Arial",8,[Drawing.FontStyle]'Bold') 
    $LabelCount.Text = "In foleder $nt files."
    $LabelCount.Location  = New-Object System.Drawing.Point(350,55)
    $LabelCount.AutoSize = $true
    $Form.Controls.Add($LabelCount)
# Description field "Select printer"
    $LabelPrinter = New-Object System.Windows.Forms.Label
    $LabelPrinter.Text = "Select printer"
    $LabelPrinter.Location  = New-Object System.Drawing.Point(20,85)
    $LabelPrinter.AutoSize = $true
    $Form.Controls.Add($LabelPrinter)
# Description list printers
    $ComboBoxPrinter = New-Object System.Windows.Forms.ComboBox
    $ComboBoxPrinter.DataSource = (Get-WmiObject win32_printer).Name
    $ComboBoxPrinter.Location  = New-Object System.Drawing.Point(150,85)
    $ComboBoxPrinter.Size = New-Object System.Drawing.Size(330,20)
    $ComboBoxPrinter.Add_Click({
                            $Form.Refresh()
                            $buttonPrinter.Enabled = $true
                                })
    $Form.Controls.Add($ComboBoxPrinter)
# Description button "Select printer"
    $buttonPrinter = New-Object System.Windows.Forms.Button
    $buttonPrinter.Text = 'Select'
    $buttonPrinter.Location = New-Object System.Drawing.Point(500,85)
    $buttonPrinter.Enabled = $false
    $buttonPrinter.Add_Click({
        $Script:Printer = $ComboBoxPrinter.SelectedItem
        })
        $Form.Controls.Add($buttonPrinter)
# Description ProgressBar
    $ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar.Location  = New-Object System.Drawing.Point(20,130)
    $ProgressBar.Size = New-Object System.Drawing.Size(330,20)
    $ProgressBar.Value = $CntProgress
    $Form.Controls.add($ProgressBar)
# Description text field "count files"
    $LabelProgress = New-Object System.Windows.Forms.Label
    $LabelProgress.Font = New-Object System.Drawing.Font("Arial",8,[Drawing.FontStyle]'Bold') 
    $LabelProgress.Text = "Printed $nf from $Script:nt files"
    $LabelProgress.Location  = New-Object System.Drawing.Point(20,160)
    $LabelProgress.AutoSize = $true
    $Form.Controls.Add($LabelProgress)
# Description button "Print"
    $buttonPrint = New-Object System.Windows.Forms.Button
    $buttonPrint.Text = 'Print'
    $buttonPrint.Location = New-Object System.Drawing.Point(500,160)
    $buttonPrint.Add_Click({
        [array]$Files= (Get-ChildItem $Folder).FullName
        Foreach ($File in $Files)
            {
                [string]$File=$File
                if ($File -match "printed") { continue }
                Start-Process –FilePath $File –Verb printTo $Printer -PassThru | ForEach-Object {Start-Sleep 10;$_} | Stop-Process
                $filenew = $file + "printed"
                Rename-Item -Path $File –NewName $filenew
                $nf=$nf+1
                $CntProgress=$nf/$nt*100
                $Form.Refresh()
                $LabelProgress.Text = "Printed $nf from $Script:nt files"
                $ProgressBar.Value = $CntProgress
                }
         $Form.Refresh()
         $ButtonDone.Enabled = $true
       })
    $Form.Controls.Add($buttonPrint)
# Description button "Done"
    $buttonDone = New-Object System.Windows.Forms.Button
    $buttonDone.Text = 'Done'
    $buttonDone.Location = New-Object System.Drawing.Point(410,160)
    $ButtonDone.Enabled = $false
    $buttonDone.Add_Click({
        $Form.Close()
        })
        $Form.Controls.Add($buttonPrinter)
    $Form.Controls.Add($buttonDone)
# Run dialog
    $Form.ShowDialog()
    }
# Run Function
main
# Done
Write-Host "`t######## End ###################" -foregroundcolor green    
Stop-Transcript
