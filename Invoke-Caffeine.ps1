function Invoke-Caffeine {
<#
.SYNOPSIS
    Prevents computer from going to sleep.
.DESCRIPTION
    Invoke-Caffeine keeps computer awake and prevents the screensaver from becoming active.
    Every few minutes a harmless keystroke is simulated. This script will close the Powershell
    window but Caffeine will continue running in the system tray. Caffeine is represented with 
    a windows cmd icon in the systray. Clicking on this icon will display a control menu to 
    pause, start, or exit the script.
.EXAMPLE
    This script calls the function 'Invoke-Caffeine' at its end starting Caffeine immediately.
#>
    [CmdletBinding()]
    [Alias()]
    param (
    )

    process {
        [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null
        [System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | Out-Null

        $sysTrayIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\cmd.exe")    
        $sysTrayTool = New-Object System.Windows.Forms.NotifyIcon
        $sysTrayTool.Text = "caffeine"
        $sysTrayTool.Icon = $sysTrayIcon
        $sysTrayTool.Visible = $true
        $startMenu = New-Object System.Windows.Forms.MenuItem
        $startMenu.Enabled = $false
        $startMenu.Text = "Start"
        $pauseMenu = New-Object System.Windows.Forms.MenuItem
        $pauseMenu.Enabled = $true
        $pauseMenu.Text = "Pause"
        $exitMenu = New-Object System.Windows.Forms.MenuItem
        $exitMenu.Text = "Exit"
        $contextMenu = New-Object System.Windows.Forms.ContextMenu
        $sysTrayTool.ContextMenu = $contextMenu
        $sysTrayTool.ContextMenu.MenuItems.AddRange($startMenu)
        $sysTrayTool.ContextMenu.MenuItems.AddRange($pauseMenu)
        $sysTrayTool.ContextMenu.MenuItems.AddRange($exitMenu)

        $sysTrayTool.Add_Click({
            if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
                $sysTrayTool.GetType().GetMethod("ShowContextMenu",[System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic).Invoke($sysTrayTool,$null)
            }
        })
        $startMenu.Add_Click({
            $pauseMenu.Enabled = $true
            $startMenu.Enabled = $false
            Start-Job -ScriptBlock $sendKeystroke -Name "caffeine"
        })
        $pauseMenu.Add_Click({
            $pauseMenu.Enabled = $false
            $startMenu.Enabled = $true
            Stop-Job -Name "caffeine"
        })
        $exitMenu.Add_Click({
            $sysTrayIcon.Visible = $false
            Stop-Job -Name "caffeine"
            Stop-Process $pid
            $psWindow.close()
        })


        # Close the powershell window but continue running the script
        $psWindow = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
        $asyncWindow = Add-Type -MemberDefinition $psWindow -Name Win32ShowWindowAsync -Namespace Win32Functions -PassThru
        $null = $asyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

        # This context provides stability when restarting script from the systray
        $appContext = New-Object System.Windows.Forms.ApplicationContext
        [void][System.Windows.Forms.Application]::Run($appContext)
        [System.GC]::Collect()


        $sendKeystroke = {
            while (1) {
                $wsh = New-Object -ComObject WScript.Shell
                $wsh.SendKeys('+{F14}')
                Start-Sleep -Seconds 180
            }
        }

        Start-Job -ScriptBlock $sendKeystroke -Name "caffeine"

    }
}

Invoke-Caffeine
