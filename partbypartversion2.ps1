# Input parameters
param (
    [string]$inputFilePath = ".\Input_VMList.xlsx",
    [string]$outputFilePath = ".\VM_Assessment_Results.xlsx",
    [ValidateSet("Excel", "TXT")]
    [string]$outputFormat = "Excel"
)

# Load required modules
Import-Module ImportExcel
Import-Module Posh-SSH

# Start log
$logFilePath = ".\VM_Assessment_Log.txt"
Start-Transcript -Path $logFilePath -Append

# Validate input file
function Validate-InputFile {
    param (
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Write-Host "Input file not found: $Path" -ForegroundColor Red
        Stop-Transcript
        exit
    }
    $requiredColumns = @("VMIPAddress", "VMOSType", "VMUsername", "VMPassword")
    $sheet = Import-Excel -Path $Path
    $missingColumns = $requiredColumns | Where-Object { $_ -notin $sheet[0].PSObject.Properties.Name }
    if ($missingColumns) {
        Write-Host "Missing required columns in input file: $($missingColumns -join ', ')" -ForegroundColor Red
        Stop-Transcript
        exit
    }
    return $sheet
}

# Read and validate input
$vmList = Validate-InputFile -Path $inputFilePath
$results = @()

foreach ($vmEntry in $vmList) {
    $vmIP = $vmEntry.VMIPAddress
    $vmOSType = $vmEntry.VMOSType
    $vmUsername = $vmEntry.VMUsername
    $vmPassword = $vmEntry.VMPassword | ConvertTo-SecureString -AsPlainText -Force
    $vmCreds = New-Object System.Management.Automation.PSCredential($vmUsername, $vmPassword)

    Write-Host "Processing VM with IP: $vmIP, OS Type: $vmOSType" -ForegroundColor Cyan
    
    $result = [PSCustomObject]@{
        VMIPAddress       = $vmIP
        VMName            = "N/A"
        OSType            = $vmOSType
        PowerState        = "N/A"
        Databases         = "N/A"
        WebServices       = "N/A"
        IISInstalled      = "No"
        IISWebSites       = "N/A"
        JavaVersions      = "N/A"
        JavaLocations     = "N/A"
        SpringBootApps    = "N/A"
        HardcodedIPs      = "N/A"
        MountedDrives     = "N/A"
        AdditionalDetails = "N/A"
        NetworkAdapters   = "N/A"
    IPv4Addresses     = "N/A"
    IPv6Addresses     = "N/A"
    DHCPEnabled       = "N/A"
    DHCPServer        = "N/A"
    DNSServers        = "N/A"
    Gateways          = "N/A"
    MACAddress        = "N/A"
    # InstalledSoftware = $systemInfo.InstalledSoftware
    # RunningServices   = $systemInfo.RunningServices
    # FirewallStatus    = $systemInfo.FirewallStatus
    LastBootTime      = $systemInfo.LastBootTime
    ConnectionError   = $null
      
        
    }

    if ($vmOSType -eq "Windows") {
        try {
            if (-not (Test-Connection -ComputerName $vmIP -Count 2 -Quiet)) {
                $result.ConnectionError = "VM not reachable via ping"
                $results += $result
                continue
            }

            # WinRM setup from control VM
            Get-Service WinRM -ErrorAction SilentlyContinue | Start-Service
            Enable-PSRemoting -Force -ErrorAction SilentlyContinue
            $currentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
            if (-not $currentTrusted.Split(",") -contains $vmIP) {
                if ([string]::IsNullOrWhiteSpace($currentTrusted)) {
                    Set-Item WSMan:\localhost\Client\TrustedHosts -Value $vmIP -Force
                } else {
                    Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$currentTrusted,$vmIP" -Force
                }
            }
            winrm quickconfig -Force | Out-Null

            # Wait a moment before session creation
            Start-Sleep -Seconds 5

            # Try remote session
            $session = New-PSSession -ComputerName $vmIP -Credential $vmCreds -ErrorAction Stop

            $systemInfo = Invoke-Command -Session $session -ScriptBlock {
                $services = Get-Service
                $processes = Get-Process
                $detectedDbs = @()
                $detectedWeb = @()
                $consoleLogs = @()
                $javaVersions = @()
                $javaLocations = @()
                $springBootApps = @()
                $iisSites = @()
                $iisInstalled = $false
                
                # Check IIS installation and sites
                try {
                    $iisFeature = Get-WindowsFeature -Name "Web-Server" -ErrorAction SilentlyContinue
                    if ($iisFeature -and $iisFeature.Installed) {
                        $iisInstalled = $true
                        $consoleLogs += "IIS is installed"
                        
                        # Check if WebAdministration module is available
                        if (Get-Module -ListAvailable -Name WebAdministration) {
                            Import-Module WebAdministration
                            $sites = Get-ChildItem IIS:\Sites -ErrorAction SilentlyContinue
                            foreach ($site in $sites) {
                                $iisSites += "$($site.Name) (State: $($site.State), Bindings: $($site.Bindings.Collection.bindingInformation -join '; ')"
                            }
                            $consoleLogs += "Found $($sites.Count) IIS sites"
                        } else {
                            $consoleLogs += "WebAdministration module not available - cannot enumerate IIS sites"
                        }
                    }
                } catch {
                    $consoleLogs += "IIS check failed: $_"
                }
                
                # Check Apache
                $apacheServices = $services | Where-Object { $_.Name -like "*apache*" -or $_.DisplayName -like "*apache*" }
                $apacheProcesses = $processes | Where-Object { $_.ProcessName -eq "httpd" -or $_.ProcessName -eq "apache" }
                if ($apacheServices -or $apacheProcesses) {
                    $detectedWeb += "Apache"
                    $consoleLogs += "Apache detected"
                }
                
                # Check Nginx
                $nginxServices = $services | Where-Object { $_.Name -like "*nginx*" -or $_.DisplayName -like "*nginx*" }
                $nginxProcesses = $processes | Where-Object { $_.ProcessName -eq "nginx" }
                if ($nginxServices -or $nginxProcesses) {
                    $detectedWeb += "Nginx"
                    $consoleLogs += "Nginx detected"
                }
                
                # If IIS is installed with sites, add to web services
                if ($iisInstalled) {
                    $detectedWeb += "IIS"
                    if ($iisSites.Count -gt 0) {
                        $detectedWeb += "IIS ($($iisSites.Count) sites"
                    }
                }
                
                # Helper to extract version from service name or display name
                function Extract-Version($text) {
                    if ($text -match "(\d{2,})") {
                        $raw = $matches[1]
                        if ($raw.Length -eq 3) {
                            return "$($raw.Substring(0,1)).$($raw.Substring(1))"
                        } elseif ($raw.Length -eq 4) {
                            return "$($raw.Substring(0,2)).$($raw.Substring(2))"
                        } else {
                            return $raw
                        }
                    }
                    return "Unknown"
                }
                
                # Check Java installations
                $javaPaths = @()
                
                # Check Java in Program Files
                $javaPaths += Get-ChildItem "C:\Program Files\Java" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "jdk|jre" }
                $javaPaths += Get-ChildItem "C:\Program Files (x86)\Java" -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "jdk|jre" }
                
                # Check Java in system PATH
                $pathDirs = $env:PATH -split ';'
                foreach ($dir in $pathDirs) {
                    if ($dir -like "*java*" -and (Test-Path $dir)) {
                        $javaExe = Get-ChildItem $dir -Filter "java.exe" -ErrorAction SilentlyContinue
                        if ($javaExe) {
                            $javaPaths += $javaExe.Directory.Parent
                        }
                    }
                }
                
                # Get unique Java installations
                $uniqueJavaPaths = $javaPaths | Sort-Object FullName -Unique
                
                foreach ($javaPath in $uniqueJavaPaths) {
                    $javaExePath = Join-Path $javaPath.FullName "bin\java.exe"
                    if (Test-Path $javaExePath) {
                        try {
                            $versionOutput = & $javaExePath -version 2>&1
                            $versionString = $versionOutput | Where-Object { $_ -match "version" } | Select-Object -First 1
                            if ($versionString -match '"(\d+(?:\.\d+)*)(?:_\d+)?') {
                                $version = $matches[1]
                                $javaVersions += $version
                                $javaLocations += $javaPath.FullName
                                $consoleLogs += "Java detected - Version: $version, Location: $($javaPath.FullName)"
                            }
                        } catch {
                            $consoleLogs += "Java detected but version check failed at: $($javaPath.FullName)"
                        }
                    }
                }
                
                # Check for Spring Boot applications
                $javaProcesses = $processes | Where-Object { $_.ProcessName -eq "java" }
                foreach ($proc in $javaProcesses) {
                    try {
                        $commandLine = (Get-WmiObject Win32_Process -Filter "ProcessId = $($proc.Id)").CommandLine
                        if ($commandLine -match "spring-boot") {
                            $appName = "Unknown"
                            if ($commandLine -match "-Dspring.application.name=([^\s]+)") {
                                $appName = $matches[1]
                            }
                            $springBootApps += $appName
                            $consoleLogs += "Spring Boot application detected: $appName"
                        }
                    } catch {
                        $consoleLogs += "Could not inspect Java process $($proc.Id) for Spring Boot"
                    }
                }
                
                # Add Spring Boot to web services if found
                if ($springBootApps.Count -gt 0) {
                    $detectedWeb += "Spring Boot ($($springBootApps -join ', '))"
                }
                
                # PostgreSQL
                $postgres = $services | Where-Object { $_.Name -like "*postgres*" }
                foreach ($svc in $postgres) {
                    $ver = "Unknown"
                    if ($svc.Name -match "-(\d+)$") {
                        $ver = $matches[1]
                    }
                    $detectedDbs += "PostgreSQL ($ver)"
                    $consoleLogs += "PostgreSQL detected - Version: $ver"
                }

                # Get mounted drives information
# $mountedDrives = @()
# try {
#     $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -ne $null }
#     foreach ($drive in $drives) {
#         $driveInfo = "$($drive.Name): ($($drive.DisplayRoot))"
#         if ($drive.Free -and $drive.Used) {
#             $totalSize = [math]::Round(($drive.Free + $drive.Used) / 1GB, 2)
#             $freeSpace = [math]::Round($drive.Free / 1GB, 2)
#             $driveInfo += " - $freeSpace GB free of $totalSize GB"
#         }
#         $mountedDrives += $driveInfo
#     }
# }
# catch {
#     $mountedDrives += "Error retrieving drive info: $_"
# }

# In the remote session script block:
$driveDetails = @()
try {
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root }
    foreach ($drive in $drives) {
        try {
            $driveInfo = [PSCustomObject]@{
                Letter      = $drive.Name + ":"
                Root        = $drive.Root
                Type        = $drive.Description
                TotalGB     = if ($drive.Free -ne $null -and $drive.Used -ne $null) { 
                                [math]::Round(($drive.Free + $drive.Used)/1GB, 2) 
                              } else { "N/A" }
                FreeGB      = if ($drive.Free) { [math]::Round($drive.Free/1GB, 2) } else { "N/A" }
                UsedGB      = if ($drive.Used) { [math]::Round($drive.Used/1GB, 2) } else { "N/A" }
                FreePercent = if ($drive.Free -ne $null -and $drive.Used -ne $null) {
                                [math]::Round(($drive.Free/($drive.Free + $drive.Used))*100, 2)
                              } else { "N/A" }
            }
            $driveDetails += $driveInfo
        } catch {
            $driveDetails += [PSCustomObject]@{
                Letter = $drive.Name + ":"
                Root   = $drive.Root
                Error  = "Error collecting details"
            }
        }
    }
} catch {
    $driveDetails += [PSCustomObject]@{ Error = "Failed to enumerate drives: $_" }
}

$installedSoftware = try {
    $software = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
                                "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
                Where-Object { $_.DisplayName } |
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
                Sort-Object DisplayName
    $software | ForEach-Object {
        "$($_.DisplayName) v$($_.DisplayVersion) ($($_.Publisher))"
    } -join " | "
} catch { "Error collecting software: $_" }

# $runningServices = try {
#     Get-Service | Where-Object { $_.Status -eq 'Running' } | 
#     Select-Object DisplayName, Name, StartType |
#     Sort-Object DisplayName |
#     ForEach-Object {
#         "$($_.DisplayName) ($($_.Name)) [Start: $($_.StartType)]"
#     } -join " | "
# } catch { "Error collecting services: $_" }
# $scheduledTasks = try {
#     Get-ScheduledTask | Where-Object { $_.State -eq 'Ready' } |
#     Select-Object TaskName, TaskPath |
#     ForEach-Object {
#         "$($_.TaskPath)$($_.TaskName)"
#     } -join " | "
# } catch { "Error collecting tasks: $_" }
# $firewallStatus = try {
#     $fw = Get-NetFirewallProfile
#     "Domain: $($fw.Domain.Enabled) | Private: $($fw.Private.Enabled) | Public: $($fw.Public.Enabled)"
# } catch { "Error checking firewall: $_" }
$lastBoot = try {
    (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
} catch { "N/A" }

# Format for output
$mountedDrivesOutput = $driveDetails | ForEach-Object {
    if ($_.Error) {
        "$($_.Letter) - $($_.Error)"
    } else {
        "$($_.Letter) | $($_.Root) | $($_.Type) | " +
        "Size: $($_.TotalGB)GB | Free: $($_.FreeGB)GB ($($_.FreePercent)%)"
    }
}
                
                # MySQL
                $mysqlServices = $services | Where-Object { $_.Name -like "*mysql*" }
                foreach ($svc in $mysqlServices) {
                    if ($svc.Name -match "MySQL(\d{2})") {
                        $raw = $matches[1]
                        switch ($raw) {
                            {$_ -ge 80} { $version = "8.0.$($raw - 80)"; break }
                            {$_ -ge 70 -and $_ -lt 80} { $version = "5.7.$($raw - 70)"; break }
                            {$_ -ge 60 -and $_ -lt 70} { $version = "5.6.$($raw - 60)"; break }
                            {$_ -ge 50 -and $_ -lt 60} { $version = "5.5.$($raw - 50)"; break }
                            default { $version = "Unknown" }
                        }
                        $detectedDbs += "MySQL $version"
                        $consoleLogs += "MySQL detected - Version: $version"
                    } else {
                        $detectedDbs += "MySQL"
                        $consoleLogs += "MySQL detected - Version not parsed"
                    }
                }
                
                # MongoDB
                $mongo = $services | Where-Object { $_.Name -like "*mongo*" }
                foreach ($svc in $mongo) {
                    $ver = Extract-Version $svc.Name
                    $detectedDbs += "MongoDB ($ver)"
                    $consoleLogs += "MongoDB detected - Version: $ver"
                }
                
                # Oracle
                $oracle = $services | Where-Object { $_.Name -like "*oracle*" }
                foreach ($svc in $oracle) {
                    $ver = Extract-Version $svc.Name
                    $detectedDbs += "Oracle DB ($ver)"
                    $consoleLogs += "Oracle DB detected - Version: $ver"
                }

                # Tomcat
                $tomcatServices = $services | Where-Object { $_.DisplayName -like "*Tomcat*" -or $_.Name -like "*Tomcat*" }
                foreach ($svc in $tomcatServices) {
                    $ver = "Unknown"
                    if ($svc.DisplayName -match "Tomcat(?:\s*|\D)?(\d{1,2})") {
                        $ver = $matches[1]
                    } elseif ($svc.Name -match "Tomcat(?:\s*|\D)?(\d{1,2})") {
                        $ver = $matches[1]
                    }

                    if ($ver -ne "Unknown") {
                        $detectedWeb += "Tomcat $ver"
                        $consoleLogs += "Tomcat detected - Version: $ver"
                    } else {
                        $detectedWeb += "Tomcat"
                        $consoleLogs += "Tomcat detected - Version not parsed"
                    }
                }

         
# Get network adapter information
$netAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' }
$networkInfo = @()

foreach ($adapter in $netAdapters) {
    $adapterInfo = [PSCustomObject]@{
        AdapterName    = $adapter.Name
        InterfaceDesc  = $adapter.InterfaceDescription
        MACAddress     = $adapter.MacAddress
        Status         = $adapter.Status
        IPv4Addresses = "Not Available"
        IPv6Addresses = "Not Available"
        DHCPEnabled   = $false
        DHCPServer    = "Not Available"
        DNSServers    = "Not Available"
        Gateways      = "Not Available"
    }

    try {
        # First verify the interface exists
        $null = Get-NetIPInterface -InterfaceIndex $adapter.InterfaceIndex -ErrorAction Stop
        $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -ErrorAction Stop

        # Update properties only if we successfully got the IP config
        $adapterInfo.IPv4Addresses = if ($ipConfig.IPv4Address.IPAddress) { $ipConfig.IPv4Address.IPAddress -join ', ' } else { "Not Available" }
        $adapterInfo.IPv6Addresses = if ($ipConfig.IPv6Address.IPAddress) { $ipConfig.IPv6Address.IPAddress -join ', ' } else { "Not Available" }
        
        try { $adapterInfo.DHCPEnabled = ($ipConfig.NetIPv4Interface.DHCP -eq 'Enabled') } catch { $adapterInfo.DHCPEnabled = $false }
        try { $adapterInfo.DHCPServer = $ipConfig.NetIPv4Interface.DHCPServer } catch { $adapterInfo.DHCPServer = "Not Available" }
        try { $adapterInfo.DNSServers = $ipConfig.DNSServer.ServerAddresses -join ', ' } catch { $adapterInfo.DNSServers = "Not Available" }
        try { $adapterInfo.Gateways = $ipConfig.IPv4DefaultGateway.NextHop -join ', ' } catch { $adapterInfo.Gateways = "Not Available" }
    }
    catch [Microsoft.PowerShell.Cmdletization.Cim.CimJobException] {
        # Skip interfaces that don't exist
        continue
    }
    catch {
        # Keep the default "Not Available" values we set initially
    }

    $networkInfo += $adapterInfo
}

# If no network adapters were found (either none exist or all failed)
if ($networkInfo.Count -eq 0) {
    $networkInfo = [PSCustomObject]@{
        AdapterName    = "No active adapters found"
        InterfaceDesc  = "N/A"
        MACAddress     = "N/A"
        Status         = "N/A"
        IPv4Addresses = "Not Available"
        IPv6Addresses = "Not Available"
        DHCPEnabled   = $false
        DHCPServer    = "Not Available"
        DNSServers    = "Not Available"
        Gateways      = "Not Available"
    }
}


                
                return [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    OSVersion    = (Get-CimInstance Win32_OperatingSystem).Caption
                    Databases    = if ($detectedDbs.Count -gt 0) { $detectedDbs -join ", " } else { "None detected" }
                    WebServices  = if ($detectedWeb.Count -gt 0) { $detectedWeb -join ", " } else { "None detected" }
                    IISInstalled = $iisInstalled
                    IISWebSites  = if ($iisSites.Count -gt 0) { $iisSites -join " | " } else { "None detected" }
                    JavaVersions = if ($javaVersions.Count -gt 0) { $javaVersions -join ", " } else { "None detected" }
                    JavaLocations = if ($javaLocations.Count -gt 0) { $javaLocations -join " | " } else { "None detected" }
                    SpringBootApps = if ($springBootApps.Count -gt 0) { $springBootApps -join ", " } else { "None detected" }
                    ConsoleLogs  = $consoleLogs
                    MountedDrives     = if ($mountedDrivesOutput) { $mountedDrivesOutput -join "`n" } else { "None" }
                    # InstalledSoftware = $installedSoftware
                    # RunningServices   = $runningServices
                    # FirewallStatus    = $firewallStatus
                    LastBootTime      = $lastBoot
                    NetworkAdapters = if ($networkInfo.Count -gt 0) { 
                        ($networkInfo | ForEach-Object {
                            "Adapter: $($_.AdapterName) | " +
                            "IPv4: $($_.IPv4Addresses) | " +
                            "IPv6: $($_.IPv6Addresses) | " +
                            "DHCP: $(if ($_.DHCPEnabled) { 'Yes' } else { 'No' }) | " +
                            "DNS: $($_.DNSServers) | " +
                            "Gateway: $($_.Gateways) | " +
                            "MAC: $($_.MACAddress)"
                        }) -join "`n"
                      } else { 
                        "Not Available" 
                      }
    IPv4Addresses   = if ($networkInfo.Count -gt 0) { 
                        $ipv4 = $networkInfo.IPv4Addresses | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($ipv4) { $ipv4 -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
    IPv6Addresses   = if ($networkInfo.Count -gt 0) { 
                        $ipv6 = $networkInfo.IPv6Addresses | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($ipv6) { $ipv6 -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
    DHCPEnabled     = if ($networkInfo.Count -gt 0) { 
                        if (($networkInfo | Where-Object { $_.DHCPEnabled }).Count -gt 0) { 
                            "Yes" 
                        } else { 
                            "No" 
                        }
                      } else { 
                        "Not Available" 
                      }
    DHCPServer     = if ($networkInfo.Count -gt 0) { 
                        $dhcp = $networkInfo.DHCPServer | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($dhcp) { $dhcp -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
    DNSServers     = if ($networkInfo.Count -gt 0) { 
                        $dns = $networkInfo.DNSServers | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($dns) { $dns -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
    Gateways       = if ($networkInfo.Count -gt 0) { 
                        $gw = $networkInfo.Gateways | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($gw) { $gw -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
    MACAddress     = if ($networkInfo.Count -gt 0) { 
                        $mac = $networkInfo.MACAddress | Where-Object { $_ -ne "Not Available" } | Select-Object -Unique
                        if ($mac) { $mac -join ', ' } else { "Not Available" }
                      } else { 
                        "Not Available" 
                      }
                    }            
}
                    
                
            

            $result.VMName = $systemInfo.ComputerName
            $result.PowerState = "Online"
            $result.AdditionalDetails = "Connected via PS Remoting after WinRM setup"
            $result.Databases = $systemInfo.Databases
            $result.WebServices = $systemInfo.WebServices
            $result.IISInstalled = if ($systemInfo.IISInstalled) { "Yes" } else { "No" }
            $result.IISWebSites = $systemInfo.IISWebSites
            $result.JavaVersions = $systemInfo.JavaVersions
            $result.JavaLocations = $systemInfo.JavaLocations
            $result.SpringBootApps = $systemInfo.SpringBootApps
            $result.NetworkAdapters = $systemInfo.NetworkAdapters
$result.IPv4Addresses = $systemInfo.IPv4Addresses
$result.IPv6Addresses = $systemInfo.IPv6Addresses
$result.DHCPEnabled = $systemInfo.DHCPEnabled
$result.DHCPServer = $systemInfo.DHCPServer
$result.DNSServers = $systemInfo.DNSServers
$result.Gateways = $systemInfo.Gateways
$result.MACAddress = $systemInfo.MACAddress
$result.MountedDrives = $systemInfo.MountedDrives
# $result.InstalledSoftware = $systemInfo.InstalledSoftware
# $result.RunningServices = $systemInfo.RunningServices
# $result.FirewallStatus = $systemInfo.FirewallStatus
$result.LastBootTime = $systemInfo.LastBootTime
            
            if ($systemInfo.ConsoleLogs) {
                foreach ($line in $systemInfo.ConsoleLogs) {
                    Write-Host "$($vmIP): $line" -ForegroundColor Yellow
                }
            }

        } catch {
            $result.ConnectionError = "Failed to connect via PowerShell Remoting after setup: $_"
        }
    }

    $results += $result
}

# Output results
if ($outputFormat -eq "Excel") {
    $results | Export-Excel -Path $outputFilePath -AutoSize -WorksheetName "VM Assessment" -TableName "VMAssessment" -FreezeTopRow
} elseif ($outputFormat -eq "TXT") {
    $results | Out-File -FilePath $outputFilePath
}

Write-Host "`nProcessing complete. Results saved to $outputFilePath" -ForegroundColor Green
Stop-Transcript