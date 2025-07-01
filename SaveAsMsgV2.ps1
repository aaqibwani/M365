# PowerShell Script to Save Outlook Emails as .msg Files

<#
.SYNOPSIS
    Saves selected Outlook emails or all emails from a specified folder to .msg format.

.DESCRIPTION
    This script connects to a running instance of Microsoft Outlook (or starts it if not running)
    and allows you to save emails in the Outlook Message Format (.msg).

    You have two primary modes of operation:
    1.  Save selected emails: If you have emails selected in Outlook when you run the script,
        it will save those specific emails.
    2.  Save emails from a specific folder: If no emails are selected, the script will open
        an Outlook GUI to allow the user to select the desired folder.

    The script will create a subfolder within the specified output directory for each
    date the emails were received, and save the .msg files inside those date-named folders.

.PARAMETER OutputFolderPath
    Specifies the root directory where the .msg files will be saved.
    A subfolder will be created for each email's received date.

.EXAMPLE
    # Saves selected emails to "C:\SavedEmails\MyProject"
    .\Save-OutlookEmailsAsMsg.ps1 -OutputFolderPath "C:\SavedEmails\MyProject"

.EXAMPLE
    # Runs the script and prompts for an output folder if not specified.
    # If no emails are selected in Outlook, it will open the Outlook folder selection GUI.
    .\Save-OutlookEmailsAsMsg.ps1

.NOTES
    - Requires Microsoft Outlook to be installed and accessible.
    - Uses the Outlook Application Object Model (COM).
    - Email filenames are generated based on the email's subject to avoid naming conflicts.
    - Special characters in subjects are replaced for valid filenames.
#>
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputFolderPath
)

function Sanitize-Filename {
    param(
        [string]$FileName
    )
    # Define invalid characters for filenames
    $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ""
    $regex = "[{0}]" -f ([RegEx]::Escape($invalidChars))
    # Replace invalid characters with an underscore
    return $FileName -replace $regex, "_"
}

Write-Host "Attempting to connect to Outlook..."

try {
    # Attempt to get a running instance of Outlook
    $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
}
catch {
    Write-Warning "Outlook is not running. Starting Outlook application..."
    # If Outlook is not running, create a new instance
    $outlook = New-Object -ComObject Outlook.Application
    # Make Outlook visible if it was just started (optional)
    # $outlook.Visible = $true
}

# Check if Outlook object was successfully created
if (-not $outlook) {
    Write-Error "Could not connect to or start Outlook. Please ensure Outlook is installed."
    exit 1
}

# Get MAPI namespace
$namespace = $outlook.GetNamespace("MAPI")

# Determine the emails to save
$itemsToSave = $null

# Check if any items are selected in the currently active explorer window
if ($outlook.ActiveExplorer -and $outlook.ActiveExplorer.Selection.Count -gt 0) {
    $itemsToSave = $outlook.ActiveExplorer.Selection
    Write-Host "Saving $($itemsToSave.Count) selected emails..."
} else {
    Write-Host "No emails selected in Outlook. Opening Outlook folder selection dialog..."

    # Use the PickFolder method to let the user select an Outlook folder via GUI
    try {
        $outlookFolder = $namespace.PickFolder()
    }
    catch {
        Write-Error "Error opening Outlook folder selection dialog: $_"
        Write-Error "Please ensure Outlook is running and you have permissions to access folders."
        exit 1
    }

    if ($outlookFolder) {
        # Filter for mail items (Class 43 is olMail)
        $itemsToSave = $outlookFolder.Items | Where-Object { $_.Class -eq 43 }
        Write-Host "Saving $($itemsToSave.Count) emails from selected folder '$($outlookFolder.FolderPath)'..."
    } else {
        Write-Host "No Outlook folder was selected. Exiting script."
        exit 0 # Exit gracefully if no folder is chosen
    }
}

# Prompt for output folder path if not provided as a parameter
if (-not $OutputFolderPath) {
    Write-Host "Opening folder selection dialog for output path..."
    # Load the System.Windows.Forms assembly to use FolderBrowserDialog
    Add-Type -AssemblyName System.Windows.Forms

    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the destination folder to save .msg files"
    $folderBrowser.ShowNewFolderButton = $true # Allow creating new folders

    # Set initial directory if a default or previous path is desired
    # $folderBrowser.SelectedPath = "C:\EmailBackups" # Uncomment and modify if you want a default path

    $dialogResult = $folderBrowser.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $OutputFolderPath = $folderBrowser.SelectedPath
        Write-Host "Selected output folder: $OutputFolderPath"
    } else {
        Write-Host "No output folder was selected. Exiting script."
        exit 0 # Exit gracefully if user cancels
    }
}

# Ensure the root output folder exists
if (-not (Test-Path $OutputFolderPath)) {
    Write-Host "Creating output folder: $OutputFolderPath"
    New-Item -Path $OutputFolderPath -ItemType Directory -Force | Out-Null
}

$savedCount = 0

if ($itemsToSave.Count -gt 0) {
    foreach ($item in $itemsToSave) {
        # Ensure it's a mail item
        if ($item.Class -eq 43) { # olMail = 43
            try {
                $subject = $item.Subject
                # Fallback for emails without a subject
                if ([string]::IsNullOrWhiteSpace($subject)) {
                    $subject = "No Subject - $(Get-Date -Format 'yyyyMMddHHmmss')"
                    Write-Verbose "Email has no subject. Using generated subject: '$subject'"
                }

                # Sanitize the subject for use as a filename
                $sanitizedSubject = Sanitize-Filename $subject

                # Get the received date to create subfolders
                $receivedDate = $null
                # Check if ReceivedTime property exists and is a valid DateTime
                if ($item.ReceivedTime -and ($item.ReceivedTime -as [DateTime])) {
                    $receivedDate = $item.ReceivedTime.ToString("yyyy-MM-dd")
                } else {
                    Write-Warning "ReceivedTime is null or invalid for email with subject: '$subject'. Using current date for folder."
                    $receivedDate = (Get-Date).ToString("yyyy-MM-dd") # Fallback to current date
                }
                $dateFolderPath = Join-Path $OutputFolderPath $receivedDate

                # Create the date-specific subfolder if it doesn't exist
                if (-not (Test-Path $dateFolderPath)) {
                    New-Item -Path $dateFolderPath -ItemType Directory -Force | Out-Null
                }

                # Construct the full file path
                $fileName = "$sanitizedSubject.msg"
                $fullFilePath = Join-Path $dateFolderPath $fileName

                # Add a counter for duplicate filenames in the same folder
                $counter = 1
                $originalFilePath = $fullFilePath
                while (Test-Path $fullFilePath) {
                    $fullFilePath = Join-Path $dateFolderPath "$sanitizedSubject ($counter).msg"
                    $counter++
                }

                # Save the email as MSG format
                $item.SaveAs($fullFilePath, [Microsoft.Office.Interop.Outlook.OlSaveAsType]::olMSG)
                Write-Host "Saved: $fullFilePath"
                $savedCount++
            }
            catch {
                $itemSubjectForError = if ([string]::IsNullOrWhiteSpace($item.Subject)) { "<No Subject>" } else { $item.Subject }
                Write-Error "Error saving email '$itemSubjectForError': $_"
            }
        }
    }
    Write-Host "`nScript completed. Successfully saved $savedCount emails."
} else {
    Write-Host "No emails found or selected to save."
}

# Clean up COM objects (important for preventing Outlook from staying open or memory leaks)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
Remove-Variable outlook, namespace -ErrorAction SilentlyContinue
