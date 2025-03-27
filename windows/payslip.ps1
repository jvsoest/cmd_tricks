function Invoke-SetProperty {
    # Auxiliar function to set properties. The SendUsingAccount property wouldn't be set in a different way
    param(
        [__ComObject] $Object,
        [String] $Property,
        $Value 
    )
    [Void] $Object.GetType().InvokeMember($Property,"SetProperty",$NULL,$Object,$Value)
}

function MakeEmail {
    param (
        [string]$AttachmentPath
    )

    # make function with one input argument as string
    function Get-FileNameFromPath {
        param (
            [string]$Path
        )
        # split the path by backslash and return the last element
        return $Path.Split("\")[-1]
    }


    # Resolve the full path (throws if file doesn't exist)
    try {
        $FullAttachmentPath = (Resolve-Path -Path $AttachmentPath).Path
    } catch {
        Write-Error "Attachment file not found: $AttachmentPath"
        exit
    }

    # read json file for employee info
    $employeeList = Get-Content -Path "C:\Users\j.vansoest\OneDrive - Medical Data Works B.V\employees.json" | ConvertFrom-Json

    # detect employee name from file name, it should match one of the employeeList names and fetch the email address
    $employeeName = $null
    $employeeEmail = $null
    foreach ($employee in $employeeList) {
        if ($FullAttachmentPath -match $employee.Name) {
            $employeeName = $employee.Name
            $employeeEmail = $employee.Email
            break
        }
    }

    if ($employeeName -eq $null) {
        Write-Error "Employee name not found in attachment file name: $FullAttachmentPath"
        exit
    }

    # Create Outlook COM object
    try {
        $Outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Error "Failed to create Outlook COM object. Is Outlook installed and configured properly?"
        exit
    }

    # Create a new mail item
    $mail = $Outlook.CreateItem(0)  # 0 = Mail Item

    $desiredFrom = "johan.vansoest@medicaldataworks.nl"

    # Loop through all accounts in the profile
    foreach ($account in $Outlook.Session.Accounts) {
        Write-Output "Checking account: $($account.SmtpAddress)"
        if ($account.SmtpAddress -eq $desiredFrom) {
            # $mail.SendUsingAccount = $account
            Invoke-SetProperty -Object $mail -Property "SendUsingAccount" -Value $account # from https://www.sqlservercentral.com/blogs/change-outlook-sender-mailbox-with-powershell-a-workaround
            Write-Output "Found account: $($account.SmtpAddress)"
            break
        }
    }

    # Set up the email properties
    $mail.Subject = "Pay slip current month"
    $mail.Body = "Dear $employeeName,`n`nPlease find attached your pay slip for the current month.`n`nKind regards,`nJohan"
    $mail.To = "$employeeEmail"

    # Check if the attachment path exists
    if (Test-Path $FullAttachmentPath) {
        $mail.Attachments.Add($FullAttachmentPath)
    } else {
        Write-Error "The specified attachment path does not exist: $FullAttachmentPath"
        exit
    }

    # Display the email (does not send)
    $mail.Display()
}

# loop over pdf files in folder and send email
$files = Get-ChildItem -Path ".\" -Filter "Loonstrook*.pdf"

# for every pdf file in the folder, call the MakeEmail function
foreach ($file in $files) {
    MakeEmail -AttachmentPath $file.FullName
}