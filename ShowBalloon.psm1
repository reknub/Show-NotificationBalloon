<#

.Synopsis
   Displays a Windows Notification Balloon in the System Tray

.DESCRIPTION
   A CMDlet to Display a Windows Notification Balloon for use in Powershell Scripts

.EXAMPLE
   Show-NotificationBalloon -Message "This message will appear" -Title "The Message has a Title" -TimeOut 600 -TypeOfMessage Info -TypeOfSysTrayIcon Question
   
   Description
   -----------
   Displays a notification bubble in the System Tray

.NOTES
   Author: Richard Bunker
   Version History:-
   1.0 - 25/02/2016 - Richard Bunker

.FUNCTIONALITY
   Enables the use of Windows Notification Balloons in PowerShell

#>

$Script:ErrorActionPreference = "Stop"

Function Show-NotificationBalloon {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [String]
        $Message,
        [Parameter(Mandatory=$false, Position=1)]
        [String]
        $Title="",
        [Parameter(Mandatory=$false, Position=2)]
        [Int]
        $TimeOut=600,
        [Parameter(Mandatory=$false, Position=3)]
        [ValidateSet("Application","Error","Warning","Information","Question","Shield")]
        [String]
        $TypeOfSysTrayIcon="Information",
        [Parameter(Mandatory=$false, Position=4)]
        [ValidateSet("None", "Info","Warning","Error")]
        [String]
        $TypeOfMessage="Info"
    )

    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    
    Remove-Event BalloonClicked_event -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier BalloonClicked_event -ErrorAction SilentlyContinue
    Remove-Event BalloonClosed_event -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier BalloonClosed_event -ErrorAction SilentlyContinue

    # Create a new Balloon
    $Global:Notification=New-Object System.Windows.Forms.NotifyIcon
    $Global:Notification.Icon = [System.Drawing.SystemIcons]::$TypeOfSysTrayIcon # SysTray Icon
    
    #Setup Balloon Title and Message
    $Global:Notification.BalloonTipTitle=$Title
    $Global:Notification.BalloonTipText=$Message
    $Global:Notification.BalloonTipIcon=$TypeOfMessage

    # Register Close Events
    Register-ObjectEvent $Global:Notification BalloonTipClicked BalloonClicked_event -Action {$Global:Notification.Visible=$false} | Out-Null
    Register-ObjectEvent $Global:Notification BalloonTipClosed BalloonClosed_event -Action {$Global:Notification.Visible=$false} | Out-Null

    # Display Balloon
    $Global:Notification.Visible=$true
    $Global:Notification.ShowBalloonTip($TimeOut)

}

# Export Module Functions:
Export-ModuleMember -Function Show-NotificationBalloon