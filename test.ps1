class TcAutomationContext {
  $Dte
  $Solution
  $SystemManager
  $ConfigManager
  $PlcProject
  $PlcIecProject
  $RemoteManager
  $AutomationSettings
  $ErrorList
  $CreationTime

  TcAutomationContext() {
      $this.Init($true, $false)
  }

  TcAutomationContext([bool] $SuppressUI, [bool] $MainWindowVisible) {
      $this.Init($SuppressUI, $MainWindowVisible)
  }

  hidden [void] Init([bool] $suppressUI, [bool] $mainWindowVisible) {
    $this.CreationTime = Get-Date
    write-header "TcAutomationContext started at $($this.CreationTime.ToString('HH\:mm\:ss\.fff'))"

    Add-TcMessageFilter
    Import-EnvDTE80

    try {
      write-msg "Creating TcXaeShell.DTE.15.0 COM object..."
      # Wait to avoid exception "Das aufgerufene Objekt wurde von den Clients getrennt"
      Start-Sleep -Seconds 3
      $this.Dte = New-Object -ComObject TcXaeShell.DTE.15.0 -ErrorAction Stop
    }
    catch {
      throw "Failed to create TcXaeShell DTE COM object: $($_.Exception.Message). Ensure TwinCAT XAE is installed and correct version registered."
    }
    
    try {
      $this.Dte.SuppressUI = $suppressUI
      $this.Dte.MainWindow.Visible = $mainWindowVisible
    } catch {
      throw "Failed to set UI properties: $($_.Exception.Message)"
    }

    $this.RemoteManager = $this.Dte.GetObject("TcRemoteManager")
    $this.AutomationSettings = $this.Dte.GetObject("TcAutomationSettings")
  }

  [void] OpenSolution([string] $PathToSolution) {
    if (-not (Test-Path -LiteralPath $PathToSolution)) {
      throw "Solution file not found at path: $PathToSolution"
    }

    try {
      write-msg "Opening solution at path: $PathToSolution. This may take a while..."
      $this.Solution = $this.Dte.Solution
      $this.Solution.Open($PathToSolution)
    }
    catch {
      throw "Failed to open solution at '$PathToSolution': $($_.Exception.Message)"
    }

    write-msg "Solution opened successfully."
  }

  [void] SetRemoteVersionFromTsProj([string] $PathToTsproj) {
    $version = Get-TcProjectVersion -PathToTsproj $PathToTsproj
    $this.SetRemoteVersion($version)
  }

  [void] SetRemoteVersion([string] $Version) {
    if (-not $Version -or ($Version.Trim() -eq '')) {
      throw "TwinCAT Version must not be empty."
    }

    write-msg "Setting TwinCAT RemoteManager version to: $Version"
    $this.RemoteManager.Version = $Version

    if ($this.RemoteManager.Version -ne $Version) {
      throw "Requested TwinCAT version '$Version' could not be activated (RemoteManager reports '$($this.RemoteManager.Version)')"
    }
    else {
      write-msg "TwinCAT version successfully set via Remote Manager"
    }
  }

  [void] SetSilentMode([bool] $SilentMode) {
    if (-not $this.AutomationSettings) {
      throw "AutomationSettings not available."
    }

    write-msg "Setting TwinCAT AutomationSettings SilentMode to: $SilentMode"
    $this.AutomationSettings.SilentMode = $SilentMode
  }

  [void] InitializeProjectById([int] $ProjectId) {
    $project = $this.Solution.Projects.Item($ProjectId)

    if (-not $project) {
      throw "Project with ID '$ProjectId' not found in solution."
    }

    try {
      $this.SystemManager = $project.Object
      $this.ConfigManager = $this.SystemManager.ConfigurationManager
    } catch {
      throw "Failed to initialize SystemManager for project id '$ProjectId': $($_.Exception.Message)"
    }

    write-msg "Project '$($project.Name)' initialized successfully."
  }

  [void] InitializeProjectByName([string] $ProjectName) {
    $project = $this.Solution.Projects | Where-Object { $_.Name -eq $ProjectName }

    if (-not $project) {
      throw "Project with name '$ProjectName' not found in solution."
    }

    try {
      $this.SystemManager = $project.Object
      $this.ConfigManager = $this.SystemManager.ConfigurationManager
    } catch {
      throw "Failed to initialize SystemManager for project '$ProjectName': $($_.Exception.Message)"
    }

    write-msg "Project '$ProjectName' initialized successfully."
  }

  [void] LoadPlcProject([string] $PlcProjectName) {
    $plcProjectLocator = "TIPC^" + $PlcProjectName
    $this.PlcProject = $this.SystemManager.LookupTreeItem($plcProjectLocator)

    if (-not $this.PlcProject) {
      throw "PLC Project with name '$PlcProjectName' not found."
    }

    write-msg "PLC Project '$PlcProjectName' loaded successfully."
  }

  [void] LoadPlcIecProject([string] $PlcProjectName) {
    $plcIecProjectLocator = "TIPC^" + $PlcProjectName + "^" + $PlcProjectName + " Project"
    $this.PlcIecProject = $this.SystemManager.LookupTreeItem($plcIecProjectLocator)

    if (-not $this.PlcIecProject) {
      throw "PLC IEC Project with name '$PlcProjectName' not found."
    }

    write-msg "PLC IEC Project '$PlcProjectName Project' loaded successfully."
  }

  [void] SetTargetPlatform([string] $PlatformName) {
    $this.ConfigManager.ActiveTargetPlatform = $PlatformName

    if ($this.ConfigManager.ActiveTargetPlatform -ne $PlatformName) {
      throw "Failed to set target platform to '$PlatformName'. Current platform is '$($this.ConfigManager.ActiveTargetPlatform)'."
    }
    else {
      write-msg "Target platform set to '$($this.ConfigManager.ActiveTargetPlatform)'."
    }
  }

  [void] GenerateBootProject() {
    $this.GenerateBootProject($true, $false)
  }

  [void] GenerateBootProject([bool] $AutoStart, [bool] $Activate) {
    if (-not $this.PlcProject) {
      throw "PLC Project not loaded."
    }

    write-msg "Generating boot project. Autostart: $AutoStart, Activate: $Activate"
    try {
      $this.PlcProject.BootProjectAutostart = $AutoStart
      $this.PlcProject.GenerateBootProject($Activate)
    }
    catch {
      throw "Failed to generate boot project: $($_.Exception.Message)"
    }

    write-msg "Boot project generated successfully."
  }

  [void] SaveAsLibrary([string] $PathToLibraryFile, [bool] $Install) {
    if (-not $this.PlcIecProject) {
      throw "PLC IEC Project not loaded."
    }

    if (Test-Path $PathToLibraryFile -PathType leaf) {
      throw "Library file already exists at path: $PathToLibraryFile"
    }

    write-msg "Saving PLC Project as library to '$PathToLibraryFile'. Install: $Install"
    try {
      $this.PlcIecProject.SaveAsLibrary($PathToLibraryFile, $Install)

      if (-not (Test-Path $PathToLibraryFile -PathType leaf)) {
        throw "Library file was not created at path: $PathToLibraryFile"
      }
    }
    catch {
      throw "Failed to save PLC Project as library: $($_.Exception.Message)"
    }

    write-msg "PLC Project saved as library successfully."
  }

  [bool] HandleErrorList() {
    return $this.HandleErrorList($false)
  }

  [bool] HandleErrorList([bool] $Verbose) {
    $this.ErrorList = Get-TcErrorList -Dte $this.Dte

    write-header "TwinCAT Error List:"

    if (-not $this.ErrorList -or $this.ErrorList.Count -eq 0) {
      write-msg "No Messages in Error List."
    }

    $containsErrors = $false

    foreach ($err in $this.ErrorList) {
      if ($err.Level -eq 'ERROR') { $containsErrors = $true }
      
      if (-not $Verbose -and $err.Level -ne 'ERROR') {
        continue
      }

      $msg = "$($err.Level): $($err.Description)"
      $hasFile = $err.FileName -and ($err.FileName.ToString().Trim() -ne '')
      if ($hasFile) {
        $msg += ":$($err.FileName):$($err.Line)"
      }
      Write-Host "  $msg"
    }

    $line = "=" * 60
    Write-Host ""
    Write-Host $line

    return $containsErrors
  }

  [void] Dispose() {
    write-msg "Quitting TcXaeShell DTE instance..."
    $this.Dte.Quit()
    Remove-TcMessageFilter

    $closeTime = Get-Date
    $elapsed = $closeTime - $this.CreationTime

    write-header "Total execution time: $($elapsed.ToString('hh\:mm\:ss\.fff'))"  
  }
}

function write-msg ($text) { Write-Host "[$(Get-Date -FORMAT G)] $text" }

function write-header {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string] $Text
  )
  $line = "=" * 60

  Write-Host ""
  Write-Host $line
  Write-Host "    $Text"
  Write-Host $line
  Write-Host ""
}

function Write-TcError {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)] [string] $Message
  )
  $Message = "#   [ERROR] $Message   #"
  $length = $Message.Length
  $line = "#" * $length
  Write-Host ""
  Write-Host $line -ForegroundColor Red
  Write-Host $Message -ForegroundColor Red
  Write-Host $line -ForegroundColor Red
  Write-Host ""
}

function Add-TcMessageFilter {
  [CmdletBinding()] param()
  write-msg "Checking if EnvDteUtils.MessageFilter is already loaded in current AppDomain"
  if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetTypes().FullName -contains 'EnvDteUtils.MessageFilter' })) {
    write-msg "EnvDteUtils.MessageFilter not found, loading from source file"
    $csPath = Join-Path -Path $PSScriptRoot -ChildPath 'EnvDteMessageFilter.cs'
    
    if (-not (Test-Path -LiteralPath $csPath)) { 
      write-msg "ERROR: EnvDteMessageFilter.cs not found at $csPath"
      throw "EnvDteMessageFilter.cs not found at $csPath" 
    }
    
    Add-Type -Path $csPath -ErrorAction Stop
  }
  else {
    write-msg "EnvDteUtils.MessageFilter already loaded in AppDomain"
  }
  
  write-msg "Registering message filter"
  [EnvDteUtils.MessageFilter]::Register()
}

function Import-EnvDTE80 {
  [CmdletBinding()] param()
  $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EnvDTE80' }
  if ($alreadyLoaded) { return }

  $defaultPath = "C:\Program Files (x86)\Beckhoff\TcXaeShell\Common7\IDE\PublicAssemblies\envdte80.dll"
  if (-not (Test-Path -LiteralPath $defaultPath -PathType Leaf)) {
    throw "EnvDTE80 assembly not found. Checked default path '$defaultPath'."
  }
  Add-Type -Path $defaultPath
}

function Remove-TcMessageFilter {
  [CmdletBinding()] param()
  if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetTypes().FullName -contains 'EnvDteUtils.MessageFilter' }) {
    write-msg "Revoking message filter"
    [EnvDteUtils.MessageFilter]::Revoke()
  }
}

function Get-TcProjectVersion {
  [CmdletBinding()] 
  param(
    [Parameter(Mandatory, Position = 0)] [string]$PathToTsproj
  )
  write-msg "Reading TwinCAT version"
  if (-not (Test-Path -LiteralPath $PathToTsproj -PathType Leaf)) {
    throw "TwinCAT project file not found: $PathToTsproj"
  }
  try {
    Write-Verbose "Reading TwinCAT project file: $PathToTsproj"
    $xml = [xml](Get-Content -LiteralPath $PathToTsproj -ErrorAction Stop)
  }
  catch {
    throw "Failed to load XML from $PathToTsproj : $($_.Exception.Message)"
  }
  if (-not $xml.TcSmProject) {
    throw "Missing TcSmProject root element in $PathToTsproj"
  }
  $version = $xml.TcSmProject.TcVersion
  if (-not $version) {
    throw "No <TcVersion> element found in $PathToTsproj"
  }
  write-msg ("Found version: " + $version.Trim())
  return $version.Trim()
}

function Get-TcErrorList {
  param(
    [Parameter(Mandatory)][object] $Dte
  )
  $toolWindowsGetter = [EnvDTE80.DTE2].GetProperty('ToolWindows').GetGetMethod()
  $errorItemsGetter = [EnvDTE80.ErrorList].GetProperty('ErrorItems').GetGetMethod()
  $errorItemGetter = [EnvDTE80.ErrorItems].GetMethod('Item')
  $errorCountGetter = [EnvDTE80.ErrorItems].GetProperty('Count').GetGetMethod()
  $descriptionGetter = [EnvDTE80.ErrorItem].GetProperty('Description').GetGetMethod()
  $errorLevelGetter = [EnvDTE80.ErrorItem].GetProperty('ErrorLevel').GetGetMethod()
  $projectGetter = [EnvDTE80.ErrorItem].GetProperty('Project').GetGetMethod()
  $fileNameGetter = [EnvDTE80.ErrorItem].GetProperty('FileName').GetGetMethod()
  $lineGetter = [EnvDTE80.ErrorItem].GetProperty('Line').GetGetMethod()

  $toolWindows = $toolWindowsGetter.Invoke($Dte, @())
  $errorItems = $errorItemsGetter.Invoke($toolWindows.ErrorList, @())
  $errorCount = $errorCountGetter.Invoke($errorItems, @())

  $list = @()
  for ($i = 1; $i -le $errorCount; $i++) {
    $item = $errorItemGetter.Invoke($errorItems, @($i))
    $rawLevel = $errorLevelGetter.Invoke($item, @())
    $description = $descriptionGetter.Invoke($item, @())
    $project = $projectGetter.Invoke($item, @())
    $fileName = $fileNameGetter.Invoke($item, @())
    $line = $lineGetter.Invoke($item, @())
    $level = switch ($rawLevel) {
      'vsBuildErrorLevelHigh' { 'ERROR' }
      'vsBuildErrorLevelMedium' { 'WARNING' }
      'vsBuildErrorLevelLow' { 'MESSAGE' }
      Default { 'ERROR_LEVEL_UNKNOWN' }
    }
    $list += [pscustomobject]@{
      FileName    = $fileName
      Line        = $line
      Level       = $level
      Description = $description
      Project     = $project
    }
  }
  return $list
}

function Use-TcAutomationContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ScriptBlock] $ScriptBlock,
        [bool] $SuppressUI = $true,
        [bool] $MainWindowVisible = $false
    )
    $ExitCode = 0
    $ctx = [TcAutomationContext]::new()

    try {
      $result =  & $ScriptBlock $ctx
      if ($result -is [int]) { $ExitCode = $result } else { $ExitCode = 0 }
    }
    catch {
      Write-TcError $_.Exception.Message
      $ExitCode = 999
    }
    finally {
      $ctx.Dispose()
    }

    return $ExitCode
}
