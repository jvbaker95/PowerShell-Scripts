#Set-ExecutionPolicy RemoteSigned

# Hide PowerShell Console
function Hide-PSWindow {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}

Function Create-FormBox {
    Param([String]$Title,[String]$Message)
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $response = [Microsoft.VisualBasic.Interaction]::InputBox($Message, $Title)
    return $response
}

Hide-PSWindow
$Time = [Int32](Create-FormBox -Title "Shutdown Scheduler" -Message "Enter the number of minutes before shutdown.`n`nType 0 or click Cancel to cancel an existing Shutdown Request.")

if ($time -eq 0) {
    Shutdown -a
    Exit
}

if ($Time.GetType().Name.equals("Int32")) {
    Shutdown -t (60*$Time) -s 
}
else {
    Exit
}
