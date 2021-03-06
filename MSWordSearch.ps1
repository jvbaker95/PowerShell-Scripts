#Hides PowerShell Console on creation if called.
#from GUI-Functions.psm1.
Function Hide-PSWindow {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}

#This will create a file prompt.
#from GUI-Functions.psm1.
Function Create-FilePrompt {
    Param(
        [Parameter(Mandatory=$false)][Object]$InitialDirectory=$env:HOMEDRIVE,
        [Parameter(Mandatory=$false)][Object]$Title="Open",
        [Parameter(Mandatory=$false)][ValidateSet("Open","Save")][String]$PromptType="Open",
        [Parameter(Mandatory=$false)][Switch]$EnableMultiSelect=$false,
        [Parameter(Mandatory=$false)][Object]$FileFilter="All files (*.*)|*.*"
    )
    switch ($PromptType) {
        ("Open") {
            $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
        }
        ("Save") {
            $Title = "Save"
            $FileDialog = New-Object System.Windows.Forms.SaveFileDialog
        }
    }
    $FileDialog.Multiselect = $EnableMultiSelect
    $FileDialog.InitialDirectory = $InitialDirectory
    $FileDialog.Filter = $FileFilter
    $FileDialog.Title = $Title
    $UserInput = $FileDialog.ShowDialog() 

    switch ($UserInput) {

        ("OK") {
            if ($EnableMultiSelect) {
                return $FileDialog.FileNames
            }
            else {
                return $FileDialog.FileName
            }
        }

        ("Cancel") {
            return $UserInput
        }
    }
}

#This will create form boxes where a target user can enter input in a UI-Box.
#from GUI-Functions.psm1.
Function Create-FormBox {
    Param(
        [Parameter(Mandatory=$false)][Object]$Title=" ",
        [Parameter(Mandatory=$true)][Object]$Message
    )
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    return [Microsoft.VisualBasic.Interaction]::InputBox($Message, $Title)
}

#This will create a message box that relays information to the user in a friendly manner.
#from GUI-Functions.psm1.
Function Create-MessageBox {
    param(
        [Parameter(Mandatory=$true)][Object]$Message,
        [Parameter(Mandatory=$false)][Object]$Title="",
        [Parameter(Mandatory=$false)][ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNoCancel","RetryCancel")]
            [Object]$ButtonOptions="OK",
        [Parameter(Mandatory=$false)][ValidateSet("None","Hand","Error","Stop","Question","Exclamation","Warning","Asterisk","Information")]
            [Object]$Icon="None"
    )
    Return [System.Windows.Forms.MessageBox]::Show($Message,$Title,$ButtonOptions,$Icon)
}

# This was a GUI built using POSHGUI with some modifications.
Function WordSearchMenu {
    
    # Get all existing winword instances and keep them in memory; we will not exit those.
    $ExistingWinwordInstances = Get-Process Winword | Select -ExpandProperty Id
    
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '400,400'
    $Form.text                       = "MSWord Text Searcher"
    $Form.TopMost                    = $false

    $T_Message                       = New-Object system.Windows.Forms.TextBox
    $T_Message.multiline             = $true
    $T_Message.text                  = "Enter the string you`'re looking for:"
    $T_Message.width                 = 260
    $T_Message.height                = 30
    $T_Message.location              = New-Object System.Drawing.Point(21.5,21)
    $T_Message.Font                  = 'Microsoft Sans Serif,10'
    $T_Message.ReadOnly              = $true
    $T_Message.BorderStyle           = "None"

    $f_FileList                      = New-Object system.Windows.Forms.DataGridView
    $f_FileList.width                = 260
    $f_FileList.height               = 240
    $f_FileList.ColumnCount          = 1
    $f_FileList.ColumnHeadersVisible = $true
    $f_FileList.Columns[0].Name      = "File Name"
    $f_FileList.location             = New-Object System.Drawing.Point(22,100)

    $M_Word_Search                   = New-Object system.Windows.Forms.TextBox
    $M_Word_Search.multiline         = $false
    $M_Word_Search.width             = 260
    $M_Word_Search.height            = 20
    $M_Word_Search.location          = New-Object System.Drawing.Point(22,68)
    $M_Word_Search.Font              = 'Microsoft Sans Serif,10'
    $M_Word_Search.text = ""

    $A_Add_File                      = New-Object system.Windows.Forms.Button
    $A_Add_File.text                 = "Add Files..."
    $A_Add_File.width                = 92
    $A_Add_File.height               = 30
    $A_Add_File.Anchor               = 'top,right'
    $A_Add_File.location             = New-Object System.Drawing.Point(294,125)
    $A_Add_File.Font                 = 'Microsoft Sans Serif,10'

    $B_Quit                          = New-Object system.Windows.Forms.Button
    $B_Quit.text                     = "Exit"
    $B_Quit.width                    = 92
    $B_Quit.height                   = 30
    $B_Quit.Anchor                   = 'right,bottom'
    $B_Quit.location                 = New-Object System.Drawing.Point(294,357)
    $B_Quit.Font                     = 'Microsoft Sans Serif,10'

    $B_Remove_File                   = New-Object system.Windows.Forms.Button
    $B_Remove_File.text              = "Clear Files..."
    $B_Remove_File.width             = 92
    $B_Remove_File.height            = 30
    $B_Remove_File.Anchor            = 'right'
    $B_Remove_File.location          = New-Object System.Drawing.Point(294,172)
    $B_Remove_File.Font              = 'Microsoft Sans Serif,10'

    $B_Execute                       = New-Object system.Windows.Forms.Button
    $B_Execute.text                  = "Begin Search"
    $B_Execute.width                 = 106
    $B_Execute.height                = 30
    $B_Execute.Anchor                = 'bottom,left'
    $B_Execute.location              = New-Object System.Drawing.Point(21,357)
    $B_Execute.Font                  = 'Microsoft Sans Serif,10'

    $Form.controls.AddRange(@($M_Word_Search,$A_Add_File,$B_Quit,$B_Remove_File,$B_Execute,$f_FileList,$T_Message))

    # Add click functions for each button; Add_file uses a file prompt that filters for exclusively .docx's and .docs since the other document formats have not been tested.
    $A_Add_File.Add_Click({ (Create-FilePrompt -EnableMultiSelect -FileFilter "All Word Documents (*.docx; *.doc)|*.doc?") | Where-Object {$_ -ne "Cancel"} | Foreach {$F_FileList.rows.add($_)} })
    $B_Remove_File.Add_Click({ $f_FileList.rows.Clear() })
    $B_Quit.Add_Click({ $Form.close() })
    $Form.Add_Click({  })
    $Form.Add_Activated({  })
    $B_Execute.Add_Click({ Get-StringInWord -Word $M_Word_Search.Text -Files $f_FileList.Rows.Cells.FormattedValue }) # The main engine - this actually does the word search.
    
    $Listen = $Form.ShowDialog()

    # Keep the menu showing as long as the user doesn't voluntarily or force-exits the program.
    while ($Listen -ne "Cancel") {
        $Listen = $Form.ShowDialog()
    }

    # If the user opts to finally close the program, exit out of all Word instances; future versions of this will kill only the instances in question.
    # Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue
    
    return $Form
}

Function Get-StringInWord {
    param(
        [String]$Word=$(Create-FormBox -Message "What word are you looking for?" -Title "Enter word" ),
        [Object[]]$Files
    ) 

    # A last-validation step to make sure files are valid formats.
    $validFiles = $Files | Where-Object { ( $_.contains(".docx") ) -or ( $_.contains(".doc") ) }

    if ( ($validFiles.Count -eq 0) -or ($Word -eq "") ) {
        Create-MessageBox -Message "Can't search for nothing!`n`nClick OK to return to main menu." -Title "Null Exception!" -ButtonOptions OK -Icon Error
        return
    }      

    $List = [System.Collections.ArrayList]::new($files.Count)


    # This opens up a word instance under the hood and opens each document separately and searches for the requested info.
    # If the word is found (not case-sensitive), then mark it positive, and negative otherwise, then add it to an arraylist for 
    # viewing using Out-Gridview.
    $WordApplication = New-Object -ComObject word.application
    $WordApplication.visible = $false

    ForEach ($file in $validFiles) {
        $document = $WordApplication.documents.open($file,$false,$true)
        $range = $document.content

        # Searches are not case-sensitive.
        $wordFound = $range.find.execute($word)

        if ($wordFound) {
            $List.add( $(New-Object -TypeName PSObject -Property @{"FileName"=$File; "Word Exists"=$true} ) )

        }
        else {
            $List.add( $(New-Object -TypeName PSObject -Property @{"FileName"=$File; "Word Exists"=$false} ) )
        }
    }
    $document.close()
    $WordApplication.quit()

    return $List | Out-GridView -OutputMode Multiple -Title "Results for '$word'"
}

Hide-PSWindow | Out-Null

WordSearchMenu | Out-Null
