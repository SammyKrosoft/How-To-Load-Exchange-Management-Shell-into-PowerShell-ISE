# How To Load Exchange Management Shell into PowerShell ISE

I initially wrote a Technet article about this [here](https://learn.microsoft.com/en-us/archive/blogs/samdrey/how-to-load-exchange-management-shell-into-powershell-ise-2), but the formatting has been messed up a bit since it was moved to learn.microsoft.com. Since I cannot modify my initial article, I'll just put the quick steps here.

Basically what we do here is that we just copy the Exchange Management Shell "traditional" shortcut definition on the ISE profile file. In the below script, we also test that Exchange cmdlets are not already present before loading the Exchange Management cmdlets – for more information about ISE profile, see this link.

What I find convenient with ISE is that you can use intellisense and color highlight when editing your Exchange PowerShell Management scripts, and also copy/paste color-coded Exchange instructions for your blogging or documentation purposes, which I did for my last posts.

Also, with Powershell ISE, you have an action pane that drills down all the cmdlets that are loaded with your current session – loading the Exchange Management Shell within ISE will also provide you all the available Exchange (2010, 2013, 2016) cmdlets available for Exchange management. Note that this "action pane" in ISE also loads all O365 and/or Azure cmdlets if you import a Powershell Session where you are connected to your O365/Azure tenant.

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

This will look like the below:

![image](https://user-images.githubusercontent.com/33433229/235512304-26ff42dc-a3fa-4c4e-ac5e-8d0075500942.png)

Save the file.


## Step 3 - Close and relaunch your ISE console

You can either launch the script you just pasted, or just close and re-open the ISE console.

![image](https://user-images.githubusercontent.com/33433229/235512325-7a9ee46a-fb73-4ad4-bc94-d627e1a6fda8.png)

