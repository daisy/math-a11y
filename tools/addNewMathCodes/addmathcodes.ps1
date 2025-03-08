# Math autocorrect codes tool from the DAISY Consortium math-a11y intiative
# https://daisy.github.io/math-a11y/docs/ms-math/

# this tool makes it simple to add more math autocorrect codes to Microsoft Word

Function AddNewMathAutoCorrectEntries
{
    # Define new codes and symbols
    $autocorrectPairs = @(        @{shortcode = "\cents" ; symbol = [char]0x00a2 }
        @{shortcode = "\repeat" ; symbol = [char]0x00af }
        @{shortcode = "\repeating" ; symbol = [char]0x00af }
        @{shortcode = "\vinculum" ; symbol = [char]0x00af }
        @{shortcode = "\infinity" ; symbol = [char]0x221e }
        @{shortcode = "\2root" ; symbol = [char]0x221a }
        @{shortcode = "\3root" ; symbol = [char]0x221b }
        @{shortcode = "\4root" ; symbol = [char]0x221c }
        @{shortcode = "\comp" ; symbol = [char]0x2218 }
        @{shortcode = "\deg" ; symbol = [char]0x00b0 }
        @{shortcode = "\rad" ; symbol = [char]0x33ad }
        @{shortcode = "\join" ; symbol = [char]0x22c8 }
        @{shortcode = "\qed" ; symbol = [char]0x220e }
        @{shortcode = "\endproof" ; symbol = [char]0x220e }
        @{shortcode = "\circle" ; symbol = [char]0x25ef }
        @{shortcode = "\circledot" ; symbol = [char]0x2299 }
        @{shortcode = "\line" ; symbol = [char]0x20e1 }
        @{shortcode = "\seg" ; symbol = [char]0x00af }
        @{shortcode = "\measangle" ; symbol = [char]0x2221 }
        @{shortcode = "\rightangle" ; symbol = [char]0x221f }
        @{shortcode = "\triangle" ; symbol = [char]0x25b3 }
        @{shortcode = "\parallelogram" ; symbol = [char]0x25b1 }
        @{shortcode = "\notparallel" ; symbol = [char]0x2226 }
        @{shortcode = "\ray" ; symbol = [char]0x20d7 }
        @{shortcode = "\arc" ; symbol = [char]0x23dc }
        @{shortcode = "\nlt" ; symbol = [char]0x226e }
        @{shortcode = "\notlt" ; symbol = [char]0x226e }
        @{shortcode = "\ngt" ; symbol = [char]0x226f }
        @{shortcode = "\notgt" ; symbol = [char]0x226f }
        @{shortcode = "\nleq" ; symbol = [char]0x2270 }
        @{shortcode = "\notle" ; symbol = [char]0x2270 }
        @{shortcode = "\nge" ; symbol = [char]0x2271 }
        @{shortcode = "\notge" ; symbol = [char]0x2271 }
        @{shortcode = "\ngeq" ; symbol = [char]0x2271 }
        @{shortcode = "\notgeq" ; symbol = [char]0x2271 }
        @{shortcode = "\not" ; symbol = [char]0x00ac }
        @{shortcode = "\muchgreater" ; symbol = [char]0x226b }
        @{shortcode = "\muchless" ; symbol = [char]0x226a }
        @{shortcode = "\notapprox" ; symbol = [char]0x2249 }
        @{shortcode = "\notcong" ; symbol = [char]0x2247 }
        @{shortcode = "\doubleint" ; symbol = [char]0x222c }
        @{shortcode = "\tripleint" ; symbol = [char]0x222d }
        @{shortcode = "\dprime" ; symbol = [char]0x2033 }
        @{shortcode = "\doubleprime" ; symbol = [char]0x2033 }
        @{shortcode = "\tprime" ; symbol = [char]0x2034 }
        @{shortcode = "\tripleprime" ; symbol = [char]0x2034 }
        @{shortcode = "\qprime" ; symbol = [char]0x2057 }
        @{shortcode = "\quadprime" ; symbol = [char]0x2057 }
        @{shortcode = "\grad" ; symbol = [char]0x2207 }
        @{shortcode = "\laplace" ; symbol = [char]0x2206 }
        @{shortcode = "\union" ; symbol = [char]0x222a }
        @{shortcode = "\Union" ; symbol = [char]0x22c3 }
        @{shortcode = "\intersection" ; symbol = [char]0x2229 }
        @{shortcode = "\Intersection" ; symbol = [char]0x22c2 }
        @{shortcode = "\notsubset" ; symbol = [char]0x2284 }
        @{shortcode = "\notsuperset" ; symbol = [char]0x2285 }
        @{shortcode = "\notsubseteq" ; symbol = [char]0x2288 }
        @{shortcode = "\notsuperseteq" ; symbol = [char]0x2289 }
        @{shortcode = "\subsetnoteq" ; symbol = [char]0x228a }
        @{shortcode = "\supersetnoteq" ; symbol = [char]0x228b }
        @{shortcode = "\belongs" ; symbol = [char]0x2208 }
        @{shortcode = "\element" ; symbol = [char]0x2208 }
        @{shortcode = "\contains" ; symbol = [char]0x220b }
        @{shortcode = "\owns" ; symbol = [char]0x220b }
        @{shortcode = "\powerset" ; symbol = [char]0x2118 }
        @{shortcode = "\complement" ; symbol = [char]0x2201 }
        @{shortcode = "\divide" ; symbol = [char]0x2223 }
        @{shortcode = "\notdivide" ; symbol = [char]0x2224 }
        @{shortcode = "\and" ; symbol = [char]0x2227 }
        @{shortcode = "\land" ; symbol = [char]0x2227 }
        @{shortcode = "\or" ; symbol = [char]0x2228 }
        @{shortcode = "\lor" ; symbol = [char]0x2228 }
        @{shortcode = "\nand" ; symbol = [char]0x22bc }
        @{shortcode = "\nor" ; symbol = [char]0x22bd }
        @{shortcode = "\xor" ; symbol = [char]0x2295 }
        @{shortcode = "\xnor" ; symbol = [char]0x2299 }
        @{shortcode = "\proves" ; symbol = [char]0x22a2 }
        @{shortcode = "\tautology" ; symbol = [char]0x22a4 }
        @{shortcode = "\false" ; symbol = [char]0x22a5 }
        @{shortcode = "\contradiction" ; symbol = [char]0x22a5 }
        @{shortcode = "\implication" ; symbol = [char]0x2192 }
        @{shortcode = "\implies" ; symbol = [char]0x2192 }
        @{shortcode = "\biconditional" ; symbol = [char]0x2194 }
        @{shortcode = "\Implication" ; symbol = [char]0x21d2 }
        @{shortcode = "\Implies" ; symbol = [char]0x21d2 }
        @{shortcode = "\Biconditional" ; symbol = [char]0x21d4 }
        @{shortcode = "\forces" ; symbol = [char]0x22a9 }
        @{shortcode = "\entailment" ; symbol = [char]0x22a8 }
        @{shortcode = "\true" ; symbol = [char]0x22a8 }
        @{shortcode = "\foreach" ; symbol = [char]0x2200 }
        @{shortcode = "\forsome" ; symbol = [char]0x2203 }
        @{shortcode = "\stddev" ; symbol = [char]0x03c3 }
        @{shortcode = "\mean" ; symbol = [char]0x03bc }
        @{shortcode = "\corr" ; symbol = [char]0x03c1 }
        @{shortcode = "\expect" ; symbol = [char]0xd835 }
        @{shortcode = "\prob" ; symbol = [char]0x2119 }
        @{shortcode = "\kron" ; symbol = [char]0x2297 }
        @{shortcode = "\hadamard" ; symbol = [char]0x2299 }
        @{shortcode = "\adjoint" ; symbol = [char]0x2020 }
        @{shortcode = "\identity" ; symbol = [char]0xd835 }
        @{shortcode = "\directsum" ; symbol = [char]0x2295 }
    )
    
    $word = New-Object -ComObject word.application
    $word.visible = $false
    $entries = $word.OMathAutoCorrect.entries
    $Count = 0
    Foreach($e in $autocorrectPairs)
    {
    # Write-Output $e.shortcode
        Try
        { $entries.add($e.shortcode, $e.symbol) | out-null 
          $Count++ }
        Catch [system.exception]
        { “unable to add $($e.shortcode)” }
        } #end foreach
    $word.Quit()
    $word = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    [System.Windows.Forms.MessageBox]::Show("Just added $Count new math autocorrect codes. If Word is running you will need to quit and restart for the changes to take effect.")
} #End function Set-MathAutoCorrectEntries

function Backup-MathAutoCorrectFile {
    $sourcePath = "$env:APPDATA\Microsoft\Office\mso0127.acl"
    $backupPath = ".\mathautocorrect.backup"
    
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $backupPath -Force
        [System.Windows.Forms.MessageBox]::Show("Backup created successfully: $backupPath")
    } else {
        [System.Windows.Forms.MessageBox]::Show("There was a problem. The math autoCorrect file was not backed up.")
    }
}

function Restore-MathAutoCorrectFile {
    $backupPath = ".\mathautocorrect.backup"
    $destinationPath = "$env:APPDATA\Microsoft\Office\mso0127.acl"
    
    if (Test-Path $backupPath) {
        Copy-Item -Path $backupPath -Destination $destinationPath -Force
        [System.Windows.Forms.MessageBox]::Show("Backup restored successfully. If Word is running you will need to quit and restart for changes to take effect.")
    } else {
        [System.Windows.Forms.MessageBox]::Show("There was a problem. The math autoCorrect file was not restored.")
    }
}

$scriptName = "addmathcodes.exe"
$runningInstances = Get-WmiObject Win32_Process | Where-Object { $_.CommandLine -match $scriptName }

if ($runningInstances.Count -gt 1) {
    # Another instance of the script is already running
    exit
}

# Function to create the Documentation Dialog
function Show-DocumentationDialog {
    $docForm = New-Object System.Windows.Forms.Form
    $docForm.Text = "Docs and Links"
    $docForm.Size = New-Object System.Drawing.Size(300,250)
    $docForm.StartPosition = "CenterParent"
    $docForm.KeyPreview = $true

    # Close documentation dialog on Escape key
    $docForm.Add_KeyDown({
        if ($_.KeyCode -eq "Escape") {
            $docForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $docForm.Close()
        }
    })

    # Create buttons for the Docs and links dialog
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Visit Accessible Math in Microsoft Word project webpage"
    $button.Size = New-Object System.Drawing.Size(250, 40)
    $button.Location = New-Object System.Drawing.Point(20, 10)
    $button.Add_Click({ Invoke-Expression "Start-Process 'https://daisy.github.io/math-a11y/docs/ms-math/'" })
    $docform.Controls.Add($button)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Visit webpage with commonly used math autocorrect codes"
    $button.Size = New-Object System.Drawing.Size(250, 40)
    $button.Location = New-Object System.Drawing.Point(20, 60)
    $button.Add_Click({ Invoke-Expression "Start-Process 'https://daisy.org/msmathcodes'" })
    $docform.Controls.Add($button)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Visit webpage with new math autocorrect codes"
    $button.Size = New-Object System.Drawing.Size(250, 40)
    $button.Location = New-Object System.Drawing.Point(20, 110)
    $button.Add_Click({ Invoke-Expression "Start-Process 'https://github.com/daisy/math-a11y/wiki/Proposed-New-Math-AutoCorrect-Codes-for-Commonly-Used-Symbols'" })
    $docform.Controls.Add($button)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Return to main menu"
    $button.Size = New-Object System.Drawing.Size(250, 40)
    $button.Location = New-Object System.Drawing.Point(20, 160)
    $button.Add_Click({
        $docForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $docForm.Close()
    })
    $docform.Controls.Add($button)

    # Show the documentation dialog as a modal dialog
    $docForm.ShowDialog()
}

Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "DAISY math-a11y working group"
$form.Size = New-Object System.Drawing.Size(360, 340)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.KeyPreview = $true  # Enable keyboard events

# Add an event to close the form when Escape is pressed
$form.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") { $form.Close() }
})

# Create a label (test message before buttons)
$label1 = New-Object System.Windows.Forms.Label
$label1.Text = "This utility is provided in good faith. However, use at your own risk!. If you are unsure, exit with no changes."
$label1.Size = New-Object System.Drawing.Size(300, 40)
$label1.Location = New-Object System.Drawing.Point(20, 10)
$form.Controls.Add($label1)

# Create buttons
$button1 = New-Object System.Windows.Forms.Button
$button1.Text = "Documentation and links menu"
$button1.Size = New-Object System.Drawing.Size(300, 40)
$button1.Location = New-Object System.Drawing.Point(20, 50)
$button1.Add_Click({ Show-DocumentationDialog })
$form.Controls.Add($button1)

$button2 = New-Object System.Windows.Forms.Button
$button2.Text = "Backup Word autocorrect file to the current folder"
$button2.Size = New-Object System.Drawing.Size(300, 40)
$button2.Location = New-Object System.Drawing.Point(20, 100)
$button2.Add_Click({ Backup-MathAutoCorrectFile })
$form.Controls.Add($button2)

$button3 = New-Object System.Windows.Forms.Button
$button3.Text = "Add new math autocorrect codes to Microsoft Word"
$button3.Size = New-Object System.Drawing.Size(300, 40)
$button3.Location = New-Object System.Drawing.Point(20, 150)
# $button3.Add_Click({ AddNewMathAutoCorrectEntries })
$button3.Add_Click({ 
    $Result = [System.Windows.Forms.MessageBox]::Show("Do you want to continue? This will take just a few seconds.", "Confirmation", 1)
    if ($Result -eq "OK") {
        AddNewMathAutoCorrectEntries } 
    })
$form.Controls.Add($button3)

$button4 = New-Object System.Windows.Forms.Button
$button4.Text = "Restore backup of Word autocorrect file"
$button4.Size = New-Object System.Drawing.Size(300, 40)
$button4.Location = New-Object System.Drawing.Point(20, 200)
$button4.Add_Click({ Restore-MathAutoCorrectFile
})
$form.Controls.Add($button4)

# Create an "Exit" button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Size = New-Object System.Drawing.Size(300, 40)
$exitButton.Location = New-Object System.Drawing.Point(20, 250)
$exitButton.Add_Click({ $form.Close() })
$form.Controls.Add($exitButton)

# Create a label (test message after buttons)
# $label2 = New-Object System.Windows.Forms.Label
# $label2.Text = "This is a spare message after the buttons..."
# $label2.AutoSize = $true
# $label2.Location = New-Object System.Drawing.Point(20, 400)
# $form.Controls.Add($label2)

# Show the form
$form.ShowDialog() | Out-Null
