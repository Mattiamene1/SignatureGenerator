# Define the domain
$domain = "mydomain.com"

# Define default Telephone
$defaultTelephone = "+39 051 0256394"

# Define paths to the original Template files
$rtfFilePath = "C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures\Template.rtf"
$htmlFilePath = "C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures\Template.htm"
$txtFilePath = "C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures\Template.txt"
$templateFilesFolder = "C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures\Template_files"

# Get user input
$name = Read-Host "Insert NAME"
$surname = Read-Host "Insert SURNAME (leave it blank if it is a Shared Mailbox)"
$department = Read-Host "Insert DEPARTMENT or ROLE"
$telephone = Read-Host "Insert TELEPHONE (leave it blank if you want to use the default one)"

# Construct filenames
$fileNameRtf = "$name`_$surname`_SIGNATURE.rtf"
$fileNameHtml = "$name`_$surname`_SIGNATURE.htm"
$fileNameTxt = "$name`_$surname`_SIGNATURE.txt"

# Construct full paths for the new files in the main directory - Outlook Signature Folder
$baseFolder = "C:\Users\Administrator\AppData\Roaming\Microsoft\Signatures"

$fullRtfPath = Join-Path -Path $baseFolder -ChildPath $fileNameRtf
$fullHtmlPath = Join-Path -Path $baseFolder -ChildPath $fileNameHtml
$fullTxtPath = Join-Path -Path $baseFolder -ChildPath $fileNameTxt

# Construct the email address: only name@doamin if surname is empty (For Shared Mailboxes)
if(!$surname){
	$email = ($name.ToLower() + $domain)
}else{
	$email = ($name.ToLower() + "." + $surname.ToLower() + $domain)
}

#If telephone is empty, set default number
if(!$telephone){
	$telephone = $defaultTelephone
}
# Create the new folder for copied files
$newTemplateFilesFolder = Join-Path -Path $baseFolder -ChildPath "$name`_$surname`_SIGNATURE_files"

# Create the new folder if it doesn't exist
if (-not (Test-Path -Path $newTemplateFilesFolder)) {
    New-Item -Path $newTemplateFilesFolder -ItemType Directory | Out-Null
}

# Function to process the file
function Process-File($filePath, $fullPath) {
    # Read the contents of the original file
    $content = Get-Content -Path $filePath | Out-String

    # Replace full-word placeholders with user input using word boundaries into the .trf file
    $content = $content -replace '\bNOME\b', $name
    $content = $content -replace '\bCOGNOME\b', $surname
    $content = $content -replace '\bREPARTO\b', $department
	$content = $content -replace '\bTEL\b', $telephone
    $content = $content -replace '\bMAIL\b', $email

    # Rename "Template" to "$Nome_$Cognome_FIRMA"
    $content = $content -replace 'Template', "$name`_$surname`_SIGNATURE"

    # Write the modified content back to the new file
    Set-Content -Path $fullPath -Value $content

    Write-Host "The file has been saved as '$fullPath' successfully."
}

# Process each original file and save to main directory
Process-File -filePath $rtfFilePath -fullPath $fullRtfPath
Process-File -filePath $htmlFilePath -fullPath $fullHtmlPath
Process-File -filePath $txtFilePath -fullPath $fullTxtPath

# Copy contents of Template_files to the new folder and rename
if (Test-Path -Path $templateFilesFolder) {
    # Get all files from the Template_files directory
    $files = Get-ChildItem -Path $templateFilesFolder
	$sourceTemplateFilesFolder = "C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures\Template_files\*"
	Copy-Item -Path $sourceTemplateFilesFolder -Destination $newTemplateFilesFolder
	Write-Host $sourceTemplateFilesFolder
	
    foreach ($file in $files) {
        # Create a new file path in the new folder
        $newFilePath = Join-Path -Path $newTemplateFilesFolder -ChildPath ($file.Name -replace 'Template', "$name`_$surname`_SIGNATURE")

        # Copy the file to the new folder with the modified name
        Copy-Item -Path $file.FullName -Destination $newFilePath

        Write-Host "Copied '$($file.Name)' to '$newFilePath'"
    }
	
} else {
    Write-Host "The Template_files folder does not exist."
}