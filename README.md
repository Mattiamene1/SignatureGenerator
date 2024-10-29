# Outlook Signature Generator
Outlook Signature Generator

# Preliminary Steps

### STEP 1: Build your .rtf signature using Word
First of all you need to build your .html file, then copy it into a Word file and save it as .rtf

### STEP 2: Save it from Outlook
Go to Outlook client → "new message" → Signature → Signatures... 
- New → Give the name "Template" → Copy in the body the content of your .rtf file → Fix the layout if necessary
- Try the Template signature sending new email
- Go in the Outlook Signature folder: Usually it is in the ```C:\Users\Administrator\AppData\Roaming\Microsoft\Signatures\```
- Here you should find 3 files (*Template.htm*, *Template.rtf*, *Template.txt*) and a folder called *Template_files*
  
> I leave my Template_Signatures folder

### STEP 3: Move these items to another folder
- Create the folder ```C:\Users\Administrator\AppData\Roaming\Microsoft\Template-Signatures```
- Move the 3 files and the folder into *Template-Signatures*

### STEP 4: Customize the script
In the first rows set:
- your domain (e.g. mydomain.com for the email address)
- your default telephone number
- 
### STEP 5: Lauch the script
Open Windows Powershell and type ```& .\signatureGenerator.ps1```.
> If you got some issues, type ```Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass``` and re-run the script

It will ask some customizations, basically will replace some keywords from the template and then move it into the Outlook signature folder
- *Insert NAME*: Set the name of the user, or the name of Shared Mailbox
- *Insert SURNAME*: Set the surname of the user, or leave it blank if it is a Shared Mailbox
- *Insert DEPARMENT*: Set the name of the department or the role
- *Insert TELEPHONE*: Set the telephone, orleave it blank if you want to use the default telephone)
  
> It will build the email with name.surname@domain

Then will be printed all the file moved and generated

### STEP 6: Use the signature
Open Outlook and look for the signature just created
> NOTE: You can build more than one signature, since the *name* field is different 

# Folder Config
![alt text](https://github.com/Mattiamene1/SignatureGenerator/blob/main/Img/FolderConfig.png)

# Results
![alt text](https://github.com/Mattiamene1/SignatureGenerator/blob/main/Img/Outlook.png)
