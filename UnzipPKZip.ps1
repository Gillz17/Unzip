Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore,PresentationFramework

function Unzip{
    param(
        [Parameter(Mandatory=$true)] [string] $filePath
    )
    $7ZipPath = '"C:\Program Files\PKWARE\PKZIPC\PKZIPC.exe"'
    $username = $env:username
    $outputFolder = "C:\$username\Documents"
    Write-Output $filePath
    $unzipCommand = "& $7ZipPath -extract -overwrite=increment '$filePath' $outputFolder 2>&1"
    $outError = iex -Command $unzipCommand
	#Checks for any errors in the unzip process
	errorHandling $outError[8]
}

function UnzipPassword{
    param(
        [Parameter(Mandatory=$true)] [string] $filePath,
        [Parameter(Mandatory=$true)] [string] $pwd
    )
    $7ZipPath = '"C:\Program Files\PKWARE\PKZIPC\PKZIPC.exe"'
    $username = $env:username
    $outputFolder = "C:\$username\Documents"
    Write-Output $filePath
    $unzipCommand = "& $7ZipPath -extract -overwrite=increment -passphrase='$pwd' '$filePath' $outputFolder 2>&1"
    $outError = iex -Command $unzipCommand
	#Checks for any errors in the unzip process
	errorHandling $outError[8]
}

#Checks for errors during the unzipping process
function errorHandling{
    param(
        [Parameter(Mandatory=$true)] [String] $errCode
    )
    if($errCode -match "(W20)"){
		#Checks how many incorrect attempts
		if($errCounter -lt 3){
			[System.Windows.MessageBox]::Show("Wrong Password! Please try again.",'ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
			$errCounter++
			PwdRequest
		}else{
			#Too many attempts so we display a message and quit
			[System.Windows.MessageBox]::Show('Too many attempts! Please try again later','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
			exit
		}
	}elseif($errCode -match "E6"){
		[System.Windows.MessageBox]::Show('Not a valid Zip file!','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}elseif($errCode -match "E34" -or $errCode -match "W18"){
		[System.Windows.MessageBox]::Show('Cannot open Zip file! Format not supported.','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}elseif($errCode -match "E155"){
		[System.Windows.MessageBox]::Show('Too many files in the Zip file!','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}elseif($errCode -match "OK" -or $errCode -match "Inflating"){
		#Unzipping was a success
		[System.Windows.MessageBox]::Show('Extraction Complete!','Complete!',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)
	}else{
		#Unknown error, quiting the script
		[System.Windows.MessageBox]::Show("Error Extracting! Error code: `n'$errCode'",'ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
		exit
	}
}

function guiMenu{
	$menu = New-Object System.Windows.Forms.Form
	$menu.Text = 'Unzip with PKZip'
	$menu.Size = New-Object System.Drawing.Size(300,200)
	$menu.StartPosition = 'CenterScreen'
	
	$yesButton = New-Object System.Windows.Forms.Button
	$yesButton.Location = New-Object System.Drawing.Point(75,120)
	$yesButton.Size = New-Object System.Drawing.Size(75,25)
	$yesButton.Text = 'Yes'
	$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$menu.AcceptButton = $yesButton
	$menu.Controls.Add($yesButton)
	
	$noButton = New-Object System.Windows.Forms.Button
	$noButton.Location = New-Object System.Drawing.Point(150,120)
	$noButton.Size = New-Object System.Drawing.Size(75,25)
	$noButton.Text = 'No'
	$noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
	$menu.cancelButton = $noButton
	$menu.Controls.Add($noButton)

    $mLabel = New-Object System.Windows.Forms.Label
	$mLabel.Location = New-Object System.Drawing.Point(10,50)
	$mLabel.Size = New-Object System.Drawing.Size(280,20)
	$mLabel.Text = 'Do you have a password for this zip file?'
    $mLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
	$menu.Controls.Add($mLabel)
	
	$menu.Topmost = $true
	
	$result = $menu.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::OK){
		#The file has a password, call the password box
		PwdRequest
	}elseif ($result -eq [System.Windows.Forms.DialogResult]::No){
		#The file does not have a password, call the unzip method
		Unzip $file
	}else{
		#The user close the window, show an error
        [System.Windows.MessageBox]::Show('Error Extracting!','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
		exit
    }
}

function PwdRequest{
	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Password?'
	$form.Size = New-Object System.Drawing.Size(300,200)
	$form.StartPosition = 'CenterScreen'
	
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = New-Object System.Drawing.Point(75,120)
	$okButton.Size = New-Object System.Drawing.Size(75,25)
	$okButton.Text = 'OK'
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)

	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Location = New-Object System.Drawing.Point(150,120)
	$cancelButton.Size = New-Object System.Drawing.Size(75,25)
	$cancelButton.Text = 'Cancel'
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $cancelButton
	$form.Controls.Add($cancelButton)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10,20)
	$label.Size = New-Object System.Drawing.Size(300,20)
	$label.Text = 'Please enter the password in the space below:'
	$label.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Regular)
	$form.Controls.Add($label)

	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Point(10,60)
	$textBox.Size = New-Object System.Drawing.Size(260,20)
	$form.Controls.Add($textBox)

	$form.Topmost = $true
	
	$form.Add_Shown({$textBox.Select()})
	$result = $form.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::OK){
		$pwstring = $textBox.Text
		UnzipPassword $file $pwstring
	}else{
		[System.Windows.MessageBox]::Show('Error Extracting!','ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
		exit
	}
}
#Empty array to hold the filenames to unzip
#$files = @()

#Find if more than 9 files were passed
if($args.Count -ge 10){
    $tooManyFiles = 'Too many files!  We can only process 9 zip files at a time.'
    [System.Windows.MessageBox]::Show($tooManyFiles,'ERROR',[System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

#Put each item into the array of filepaths
foreach($item in $args){
    $file += $item
}

$errCounter = 1

guiMenu