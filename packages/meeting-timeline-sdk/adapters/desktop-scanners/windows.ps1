$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Get-SafeProperty {
  param(
    [System.Windows.Automation.AutomationElement]$Element,
    [System.Windows.Automation.AutomationProperty]$Property
  )
  try {
    $value = $Element.GetCurrentPropertyValue($Property, $true)
    if ($value -eq [System.Windows.Automation.AutomationElement]::NotSupported) { return $null }
    return $value
  } catch {
    return $null
  }
}

function Convert-Control {
  param([System.Windows.Automation.AutomationElement]$Element)
  $controlType = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::ControlTypeProperty)
  [ordered]@{
    role = if ($controlType) { $controlType.ProgrammaticName } else { $null }
    name = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::NameProperty)
    description = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::HelpTextProperty)
    automation_id = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::AutomationIdProperty)
    enabled = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::IsEnabledProperty)
    focused = Get-SafeProperty $Element ([System.Windows.Automation.AutomationElement]::HasKeyboardFocusProperty)
  }
}

$targetNames = @('ms-teams', 'msteams', 'teams', 'zoom', 'zoom.us')
$applications = @()
$errors = @()

foreach ($process in Get-Process -ErrorAction SilentlyContinue) {
  if ($targetNames -notcontains $process.ProcessName.ToLowerInvariant()) { continue }
  $windows = @()
  try {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition(
      [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
      $process.Id
    )
    $elements = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $condition)
    foreach ($window in $elements) {
      $controls = @()
      $descendants = $window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition
      )
      $limit = [Math]::Min($descendants.Count, 1200)
      for ($index = 0; $index -lt $limit; $index += 1) {
        $control = Convert-Control $descendants.Item($index)
        if ($control.name -or $control.description -or $control.automation_id -or $control.role -match 'Button|Text|Menu|Radio|Check') {
          $controls += $control
        }
      }
      $windows += [ordered]@{
        title = Get-SafeProperty $window ([System.Windows.Automation.AutomationElement]::NameProperty)
        role = 'Window'
        focused = Get-SafeProperty $window ([System.Windows.Automation.AutomationElement]::HasKeyboardFocusProperty)
        visible = $true
        controls = $controls
        scanned_node_count = $limit
      }
    }
  } catch {
    $errors += [ordered]@{
      application = $process.ProcessName
      code = 'ui_automation_scan_failed'
      message = $_.Exception.Message
    }
  }
  $applications += [ordered]@{
    name = $process.ProcessName
    process_name = $process.ProcessName
    pid = $process.Id
    frontmost = $false
    visible = $true
    window_title = $process.MainWindowTitle
    windows = $windows
  }
}

[ordered]@{
  scanner = 'windows_ui_automation'
  accessibility_trusted = $true
  applications = $applications
  errors = $errors
} | ConvertTo-Json -Depth 12 -Compress
