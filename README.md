# How To Load Exchange Management Shell into PowerShell ISE

I initially wrote a Technet article about this [here](https://learn.microsoft.com/en-us/archive/blogs/samdrey/how-to-load-exchange-management-shell-into-powershell-ise-2), but the formatting has been messed up a bit since it was moved to learn.microsoft.com. Since I cannot modify my initial article, I'll just put the quick steps here.


![image](https://user-images.githubusercontent.com/33433229/235512180-654bf049-c40a-4be3-b2ea-3661bd3e34d1.png)


## Step 1 - Force create the PowerShell profile path if it's not here

Just run the below in either a PowerShell ISE console or a plain PowerShell text regular console:

```powershell
if (!(Test-Path -Path $PROFILE ))
{ New-Item -Type File -Path $PROFILE -Force }
```

![image](https://user-images.githubusercontent.com/33433229/235512208-cc4f8bfc-e6b0-4818-8a47-79a3349c5dff.png)



Then, from a **PowerShell ISE** console, run the editor by executing this - you can run it from the Script pane or from the execution pane

```powershell
psEdit $profile
```

![image](https://user-images.githubusercontent.com/33433229/235512250-b739f547-6cd8-44b5-8b8a-850a7a821dab.png)



## Step 2 - Paste the PowerShell Exchange Management Console loading code

Paste the below code, and save:


```powershell
$StopWatch = [System.Diagnostics.StopWatch]::StartNew()

Function Test-Command ($Command)
{
    Try
    {
        Get-command $command -ErrorAction Stop
        Return $True
    }
    Catch [System.SystemException]
    {
        Return $False
    }
}

IF (Test-Command "Get-Mailbox") {Write-Host "Exchange cmdlets already present"}
Else {
    $CallEMS = ". '$env:ExchangeInstallPath\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell "
    Invoke-Expression $CallEMS

$stopwatch.Stop()  
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."  
Write-Host $msg  
$msg = $null  
$StopWatch = $null
}

```

![image](https://user-images.githubusercontent.com/33433229/235512304-26ff42dc-a3fa-4c4e-ac5e-8d0075500942.png)


## Step 3 - Close and relaunch your ISE console

![image](https://user-images.githubusercontent.com/33433229/235512325-7a9ee46a-fb73-4ad4-bc94-d627e1a6fda8.png)

