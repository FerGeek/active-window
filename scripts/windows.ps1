[CmdletBinding()]
Param(
[string]$n,[string]$interval
)
Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class UserWindows {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
"@
Add-Type -AssemblyName System.Windows.Forms
try {
	while($n -ne 0){
		$ActiveHandle = [UserWindows]::GetForegroundWindow()
		$Process = Get-Process | ? {$_.MainWindowHandle -eq $activeHandle}
    $X = [System.Windows.Forms.Cursor]::Position.X
    $Y = [System.Windows.Forms.Cursor]::Position.Y
    $string =  $Process | Select ProcessName, @{Name="AppTitle";Expression= {($_.MainWindowTitle)}}
    $string = "$string;$X;$Y"
		Write-Host -NoNewline $string
		Start-Sleep -m $interval
		If ($n -gt 0) {$n-=1}
	}
} catch {
 Write-Error "Failed to get active Window details. More Info: $_"
}
