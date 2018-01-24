########################################################################
#  Date:           24/01/2018
#  Created by :    Andrzej Mega
#  Version:        1.0
#  Source Code :   https://github.com/AMega/SSL-Certificate-Splitter
########################################################################



#----------------------------------------------
#  1) Create the Windows Form
#----------------------------------------------

# Create Form
Add-Type -AssemblyName System.Windows.Forms
$SSLCertificateSplit = New-Object system.Windows.Forms.Form
$SSLCertificateSplit.Text = "SSL Certificate Splitter "
$SSLCertificateSplit.TopMost = $true
$SSLCertificateSplit.Width = 843
$SSLCertificateSplit.Height = 451
$SSLCertificateSplit.FormBorderStyle = 'Fixed3D'
$SSLCertificateSplit.ShowInTaskbar = $True
$SSLCertificateSplit.MaximizeBox = $false
$SSLCertificateSplit.StartPosition = "CenterScreen"
$SSLCertificateSplit.Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

# Status Bar
$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Height = 25
$statusBar.Text = "Ready..."
$statusBar.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.Controls.Add($statusBar)


# Header label
$label_header_text = New-Object system.windows.Forms.Label
$label_header_text.Text = "This tool will extract the SSL Cert, Key (unencrypted), CA and ROOT chain in PEM format."
$label_header_text.AutoSize = $true
$label_header_text.ForeColor = "#656565"
$label_header_text.Width = 25
$label_header_text.Height = 10
$label_header_text.location = new-object system.drawing.point(44,68)
$label_header_text.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($label_header_text)

# Header sub label
$label_header_sub = New-Object system.windows.Forms.Label
$label_header_sub.Text = "SSL Certificate Splitter"
$label_header_sub.AutoSize = $true
$label_header_sub.Width = 25
$label_header_sub.Height = 10
$label_header_sub.location = new-object system.drawing.point(41,33)
$label_header_sub.Font = "Microsoft Sans Serif,20,style=Bold"
$SSLCertificateSplit.controls.Add($label_header_sub)

# About Icon
$image_about = New-Object system.windows.Forms.PictureBox
$image_about.Image = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\HelpPane.exe')
$image_about.SizeMode = 'StretchImage'
$image_about.Width = 28
$image_about.Height = 28
$image_about.location = new-object system.drawing.point(790,8)
$image_about.add_Click({[system.Diagnostics.Process]::start("https://github.com/AMega/SSL-Certificate-Splitter")})
$SSLCertificateSplit.controls.Add($image_about)


# PFX Location label
$label_pfx_location = New-Object system.windows.Forms.Label
$label_pfx_location.Text = "PFX Cert File"
$label_pfx_location.AutoSize = $true
$label_pfx_location.Width = 25
$label_pfx_location.Height = 10
$label_pfx_location.location = new-object system.drawing.point(45,131)
$label_pfx_location.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($label_pfx_location)

# PFX Location text
$textBox_pfx_location = New-Object system.windows.Forms.TextBox
$textBox_pfx_location.Width = 440
$textBox_pfx_location.Height = 30
$textBox_pfx_location.location = new-object system.drawing.point(192,133)
$textBox_pfx_location.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($textBox_pfx_location)

# PFX Location browse button
$btn_pfx_browse = New-Object system.windows.Forms.Button
$btn_pfx_browse.Text = "Browse"
$btn_pfx_browse.Width = 105
$btn_pfx_browse.Height = 28
$btn_pfx_browse.location = new-object system.drawing.point(679,133)
$btn_pfx_browse.Font = "Microsoft Sans Serif,10"
$btn_pfx_browse.add_Click(
    {
    # File select
    $FileSelect=New-Object System.Windows.Forms.OpenFileDialog
    $FileSelect.InitialDirectory = "\\$env:Computername\c$\"
    $FileSelect.title = "Select PFX SSL cert..."   
    $FileSelect.filter = "PFX Files (*.pfx)| *.pfx"  
    $FileSelect.ShowDialog()

    $textBox_pfx_location.Text = $FileSelect.filename 
    $statusBar.Text = "PFX Cert file selected..."
    }
)
$SSLCertificateSplit.controls.Add($btn_pfx_browse)

# PFX Import password label
$label_import_pass = New-Object system.windows.Forms.Label
$label_import_pass.Text = "Import Password"
$label_import_pass.AutoSize = $true
$label_import_pass.Width = 25
$label_import_pass.Height = 10
$label_import_pass.location = new-object system.drawing.point(45,174)
$label_import_pass.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($label_import_pass)

# PFX Import password text
$textBox_import_pass = New-Object system.windows.Forms.MaskedTextBox
# If you want to mask the password - uncomment the 2 lines below
#$textBox_import_pass = New-Object system.windows.Forms.MaskedTextBox
#$textBox_import_pass.PasswordChar = '*'
$textBox_import_pass.Width = 206
$textBox_import_pass.Height = 30
$textBox_import_pass.location = new-object system.drawing.point(193,174)
$textBox_import_pass.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($textBox_import_pass)


# Export folder label
$label_export_folder = New-Object system.windows.Forms.Label
$label_export_folder.Text = "Export Folder"
$label_export_folder.AutoSize = $true
$label_export_folder.Width = 25
$label_export_folder.Height = 10
$label_export_folder.location = new-object system.drawing.point(45,216)
$label_export_folder.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($label_export_folder)

# Export folder text
$textBox_export_folder = New-Object system.windows.Forms.TextBox
$textBox_export_folder.Width = 440
$textBox_export_folder.Height = 30
$textBox_export_folder.location = new-object system.drawing.point(192,217)
$textBox_export_folder.Font = "Microsoft Sans Serif,10"
$SSLCertificateSplit.controls.Add($textBox_export_folder)

# Export folder button
$btn_browse_export_folder = New-Object system.windows.Forms.Button
$btn_browse_export_folder.Text = "Browse"
$btn_browse_export_folder.Width = 105
$btn_browse_export_folder.Height = 28
$btn_browse_export_folder.location = new-object system.drawing.point(678,217)
$btn_browse_export_folder.Font = "Microsoft Sans Serif,10"
$btn_browse_export_folder.add_Click(
    {
    # Folder select
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
	$FolderSelect = New-Object System.Windows.Forms.FolderBrowserDialog 
	$FolderSelect.ShowDialog() | Out-Null

	$textBox_export_folder.Text = $FolderSelect.selectedPath
    $statusBar.Text = "Export folder selected..."
    }
)

$SSLCertificateSplit.controls.Add($btn_browse_export_folder)


# (No)Warranty label
$label_warranty = New-Object system.windows.Forms.Label
$label_warranty.Text = "* This Tool comes with absolutely NO warranty"
$label_warranty.AutoSize = $true
$label_warranty.ForeColor = "#828282"
$label_warranty.Width = 25
$label_warranty.Height = 10
$label_warranty.location = new-object system.drawing.point(15,355)
$label_warranty.Font = "Microsoft Sans Serif,8"
$SSLCertificateSplit.controls.Add($label_warranty)


# OK button
$btn_OK = New-Object system.windows.Forms.Button
$btn_OK.Text = "OK"
$btn_OK.Width = 105
$btn_OK.Height = 40
$btn_OK.location = new-object system.drawing.point(679,315)
$btn_OK.Font = "Microsoft Sans Serif,10"
$btn_OK.add_Click($RunOpenSSLCommand)
$SSLCertificateSplit.controls.Add($btn_OK)

# CANCEL button
$btn_cancel = New-Object system.windows.Forms.Button
$btn_cancel.Text = "Cancel"
$btn_cancel.ForeColor = "#6d6d6d"
$btn_cancel.Width = 105
$btn_cancel.Height = 40
$btn_cancel.location = new-object system.drawing.point(534,315)
$btn_cancel.Font = "Microsoft Sans Serif,10"
$btn_cancel.add_Click({ $SSLCertificateSplit.Close() })
$SSLCertificateSplit.controls.Add($btn_cancel)


# Check if OpenSSL is installed and in System Paths, if not - show Warning message
#
# * Very simple method just by running the openssl command from current prompt location
# * TO DO:  Create a proper check
#
& cmd /c 'openssl version' 2> $null
if ($LASTEXITCODE -ne 0) {
   
    # Warning label
    $label_warning = New-Object system.windows.Forms.Label
    $label_warning.Text = "WARNING: OpenSSL NOT installed (Prerequisite!)"
    $label_warning.AutoSize = $true
    $label_warning.ForeColor = "#ff0000"
    $label_warning.Width = 25
    $label_warning.Height = 10
    $label_warning.location = new-object system.drawing.point(36,271)
    $label_warning.Font = "Microsoft Sans Serif,12,style=Bold"
    $SSLCertificateSplit.controls.Add($label_warning)

    # Download here label
    $label_download_here = New-Object system.windows.Forms.Label
    $label_download_here.Text = "Download it here: "
    $label_download_here.AutoSize = $true
    $label_download_here.Width = 20
    $label_download_here.Height = 10
    $label_download_here.location = new-object system.drawing.point(37,298)
    $label_download_here.Font = "Microsoft Sans Serif,8"
    $SSLCertificateSplit.controls.Add($label_download_here)

    # WIN-OpenSSL download link
    $label_ssl_link = New-Object system.windows.Forms.LinkLabel
    $label_ssl_link.Text = "https://slproweb.com/products/Win32OpenSSL.html"
    $label_ssl_link.AutoSize = $true
    $label_ssl_link.Width = 25
    $label_ssl_link.Height = 10
    $label_ssl_link.location = new-object system.drawing.point(150,299)
    $label_ssl_link.Font = "Microsoft Sans Serif,8"
    $label_ssl_link.LinkColor = "BLUE"
    $label_ssl_link.add_Click({[system.Diagnostics.Process]::start("https://slproweb.com/products/Win32OpenSSL.html")})
    $SSLCertificateSplit.controls.Add($label_ssl_link)

    }
   


#----------------------------------------------
#  2) Execute OpenSSL commands
#----------------------------------------------

# Define function
$RunOpenSSLCommand = {
    
    # Check if all 3 fields have been set/selected
    # Set Variables
    
    
    
    # If one of the 3 fields is empty
    If (
        [string]::IsNullOrEmpty($textBox_pfx_location.Text) -OR 
        [string]::IsNullOrEmpty($textBox_import_pass.Text) -OR 
        [string]::IsNullOrEmpty($textBox_export_folder.Text)
        ){

        # THEN => Show the Error Window
        [System.Windows.Forms.MessageBox]::Show("All fields are Required", 'WARNING')

    # When all 3 are filled in - Execute the OpenSSL Command
    }  Else {
      [System.Windows.Forms.MessageBox]::Show("TESTING: All worked and EXECUTED", 'WHOOP, WHOOP !')

      # After success - Close the window
      $SSLCertificateSplit.Close()
    } 

}



# Other Functions
# Define Function
$ShowAboutWindow = {
    [System.Windows.Forms.MessageBox]::Show("ABOUT: This is an about page !")
} # End About



#------------------------------------------------

# Finalise Form
[void]$SSLCertificateSplit.ShowDialog()
$SSLCertificateSplit.Dispose()
