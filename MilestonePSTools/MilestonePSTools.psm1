# Copyright 2025 Milestone Systems A/S
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Text.RegularExpressions
using namespace System.Windows.Forms
using namespace MilestonePSTools
using namespace MilestonePSTools.Utility
using namespace VideoOS.Platform.ConfigurationItems
Import-Module "$PSScriptRoot\bin\MilestonePSTools.dll"

enum VmsTaskState {
    Completed
    Error
    Idle
    InProgress
    Success
    Unknown
}

class VmsTaskResult {
    [int] $Progress
    [string] $Path
    [string] $ErrorCode
    [string] $ErrorText
    [VmsTaskState] $State

    VmsTaskResult () {}

    VmsTaskResult([VideoOS.ConfigurationApi.ClientService.ConfigurationItem] $InvokeItem) {
        foreach ($p in $InvokeItem.Properties) {
            try {
                switch ($p.ValueType) {
                    'Progress' {
                        $this.($p.Key) = [int]$p.Value
                    }
                    'Tick' {
                        $this.($p.Key) = [bool]::Parse($p.Value)
                    }
                    default {
                        $this.($p.Key) = $p.Value
                    }
                }
            } catch {
                if ($p -in 'Progress', 'Path', 'ErrorCode', 'ErrorText', 'State' ) {
                    throw
                }
            }

        }
    }
}

class VmsHardwareScanResult : VmsTaskResult {
    [uri]    $HardwareAddress
    [string] $UserName
    [string] $Password
    [bool]   $MacAddressExistsGlobal
    [bool]   $MacAddressExistsLocal
    [bool]   $HardwareScanValidated
    [string] $MacAddress
    [string] $HardwareDriverPath

    # Property hidden so that this type can be cleanly exported to CSV or something
    # without adding a column with a complex object in it.
    hidden [VideoOS.Platform.ConfigurationItems.RecordingServer] $RecordingServer

    VmsHardwareScanResult() {}

    VmsHardwareScanResult([VideoOS.ConfigurationApi.ClientService.ConfigurationItem] $InvokeItem) {
        $members = ($this.GetType().GetMembers() | Where-Object MemberType -EQ 'Property').Name
        foreach ($p in $InvokeItem.Properties) {
            if ($p.Key -notin $members) {
                continue
            }
            switch ($p.ValueType) {
                'Progress' {
                    $this.($p.Key) = [int]$p.Value
                }
                'Tick' {
                    $this.($p.Key) = [bool]::Parse($p.Value)
                }
                default {
                    $this.($p.Key) = $p.Value
                }
            }
        }
    }
}

# Contains the output from the script passed to LocalJobRunner.AddJob, in addition to any errors thrown in the script if present.
class LocalJobResult {
    [object[]] $Output
    [System.Management.Automation.ErrorRecord[]] $Errors
}

# Contains the IAsyncResult object returned by PowerShell.BeginInvoke() as well as the PowerShell instance we need to
class LocalJob {
    [System.Management.Automation.PowerShell] $PowerShell
    [System.IAsyncResult] $Result
}

# Centralizes the complexity of running multiple commands/scripts at a time and receiving the results, including errors, when they complete.
class LocalJobRunner : IDisposable {
    hidden [System.Management.Automation.Runspaces.RunspacePool] $RunspacePool
    hidden [System.Collections.Generic.List[LocalJob]] $Jobs
    [timespan] $JobPollingInterval = (New-TimeSpan -Seconds 1)
    [string[]] $Modules = @()

    # Default constructor creates an underlying runspace pool with a max size matching the number of processors
    LocalJobRunner () {
        $this.Initialize($env:NUMBER_OF_PROCESSORS)
    }

    LocalJobRunner ([string[]]$Modules) {
        $this.Modules = $Modules
        $this.Initialize($env:NUMBER_OF_PROCESSORS)
    }

    # Optionally you may manually specify a max size for the underlying runspace pool.
    LocalJobRunner ([int]$MaxSize) {
        $this.Initialize($MaxSize)
    }

    hidden [void] Initialize([int]$MaxSize) {
        $this.Jobs = New-Object System.Collections.Generic.List[LocalJob]
        $iss = [initialsessionstate]::CreateDefault()
        if ($this.Modules.Count -gt 0) {
            $iss.ImportPSModule($this.Modules)
        }
        $this.RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxSize, $iss, (Get-Host))
        $this.RunspacePool.Open()
    }

    # Accepts a scriptblock and a set of parameters. A new powewershell instance will be created, attached to a runspacepool, and the results can be collected later in a call to ReceiveJobs.
    [LocalJob] AddJob([scriptblock]$scriptblock, [hashtable]$parameters) {
        $parameters = if ($null -eq $parameters) { $parameters = @{} } else { $parameters }
        $shell = [powershell]::Create()
        $shell.RunspacePool = $this.RunspacePool
        $asyncResult = $shell.AddScript($scriptblock).AddParameters($parameters).BeginInvoke()
        $job = [LocalJob]@{
            PowerShell = $shell
            Result     = $asyncResult
        }
        $this.Jobs.Add($job)
        return $job
    }

    # Returns the output from specific jobs
    [LocalJobResult[]] ReceiveJobs([LocalJob[]]$localJobs) {
        $completedJobs = $localJobs | Where-Object { $_.Result.IsCompleted }
        $completedJobs | ForEach-Object { $this.Jobs.Remove($_) }
        $results = $completedJobs | ForEach-Object {
            [LocalJobResult]@{
                Output = $_.PowerShell.EndInvoke($_.Result)
                Errors = $_.PowerShell.Streams.Error
            }

            $_.PowerShell.Dispose()
        }
        return $results
    }

    # Returns the output from any completed jobs in an object that also includes any errors if present.
    [LocalJobResult[]] ReceiveJobs() {
        return $this.ReceiveJobs($this.Jobs)
    }

    # Block until all jobs have completed. The list of jobs will be polled on an interval of JobPollingInterval, which is 1 second by default.
    [void] Wait() {
        $this.Wait($this.Jobs)
    }

    # Block until all jobs have completed. The list of jobs will be polled on an interval of JobPollingInterval, which is 1 second by default.
    [void] Wait([LocalJob[]]$jobList) {
        while ($jobList.Result.IsCompleted -contains $false) {
            Start-Sleep -Seconds $this.JobPollingInterval.TotalSeconds
        }
    }

    # Returns $true if there are any jobs available to be received using ReceiveJobs. Use to implement your own polling strategy instead of using Wait.
    [bool] HasPendingJobs() {
        return ($this.Jobs.Count -gt 0)
    }

    # Make sure to dispose of this class so that the underlying runspace pool gets disposed.
    [void] Dispose() {
        $this.Jobs.Clear()
        $this.RunspacePool.Close()
        $this.RunspacePool.Dispose()
    }
}

class VmsCameraStreamConfig {
    [string] $Name
    [string] $DisplayName
    [bool] $Enabled
    [bool] $LiveDefault
    [string] $LiveMode
    [bool] $PlaybackDefault
    [bool] $Recorded
    [string] $RecordingTrack
    [bool] $UseEdge
    [guid] $StreamReferenceId
    [hashtable] $Settings
    [hashtable] $ValueTypeInfo
    hidden [VideoOS.Platform.ConfigurationItems.Camera] $Camera
    hidden [bool] $UseRawValues
    hidden [System.Collections.Generic.Dictionary[string, string]] $RecordToValues

    [void] Update() {
        $this.Camera.DeviceDriverSettingsFolder.ClearChildrenCache()
        $this.Camera.StreamFolder.ClearChildrenCache()
        $deviceDriverSettings = $this.Camera.DeviceDriverSettingsFolder.DeviceDriverSettings[0]
        $streamUsages = $this.Camera.StreamFolder.Streams[0]

        $stream = $deviceDriverSettings.StreamChildItems | Where-Object DisplayName -EQ $this.Name
        $streamUsage = $streamUsages.StreamUsageChildItems | Where-Object {
            $_.StreamReferenceId -eq $_.StreamReferenceIdValues[$stream.DisplayName]
        }
        if ($streamUsage) {
            $this.RecordToValues = $streamUsage.RecordToValues
        }
        $this.DisplayName = $streamUsage.Name
        $this.Enabled = $null -ne $streamUsage

        $this.LiveDefault = $streamUsage.LiveDefault
        $this.LiveMode = $streamUsage.LiveMode

        # StreamUsageChildItem.Record is true only for the primary recording track. Or for the recorded track on 2023 R1 and older.
        # It will be false for the secondary recording track on 2023 R2 and later.
        $this.Recorded = $streamUsage.Record -or ($streamUsage.RecordToValues.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($streamUsage.RecordTo))
        $this.RecordingTrack = $streamUsage.RecordTo
        $this.PlaybackDefault = if ($streamUsage.RecordToValues.Count -gt 0) { $streamUsage.DefaultPlayback } else { $streamUsage.Record -eq $true }
        $this.UseEdge = $streamUsage.UseEdge
        $this.StreamReferenceId = if ($streamUsages.StreamUsageChildItems.Count -gt 0) { $streamUsages.StreamUsageChildItems[0].StreamReferenceIdValues[$this.Name] } else { [guid]::Empty }
        $parsedSettings = $stream | ConvertFrom-ConfigChildItem -RawValues:($this.UseRawValues)
        $this.Settings = $parsedSettings.Properties.Clone()
        $this.ValueTypeInfo = $parsedSettings.ValueTypeInfo.Clone()
    }

    [string] GetRecordingTrackName() {
        if ($this.RecordToValues.Count) {
            return ($this.RecordToValues.GetEnumerator() | Where-Object Value -EQ $this.RecordingTrack).Key
        } elseif ($this.Recorded) {
            return 'Primary recording'
        } else {
            return 'No recording'
        }
    }
}

class VmsStreamDeviceStatus : VideoOS.Platform.SDK.Proxy.Status2.MediaStreamDeviceStatusBase {
    [string] $DeviceName
    [string] $DeviceType
    [string] $RecorderName
    [guid]   $RecorderId
    [bool]   $Motion

    VmsStreamDeviceStatus () {}
    VmsStreamDeviceStatus ([VideoOS.Platform.SDK.Proxy.Status2.MediaStreamDeviceStatusBase]$status) {
        $this.DbMoveInProgress = $status.DbMoveInProgress
        $this.DbRepairInProgress = $status.DbRepairInProgress
        if ($null -ne $status.DeviceId) {
            $this.DeviceId = $status.DeviceId
        }
        $this.Enabled = $status.Enabled
        $this.Error = $status.Error
        $this.ErrorNoConnection = $status.ErrorNoConnection
        $this.ErrorNotLicensed = $status.ErrorNotLicensed
        $this.ErrorOverflow = $status.ErrorOverflow
        $this.ErrorWritingGop = $status.ErrorWritingGop
        $this.IsChange = $status.IsChange
        $this.Recording = $status.Recording
        $this.Started = $status.Started
        if ($null -ne $status.Time) {
            $this.Time = $status.Time
        }
        if ($null -ne $status.Motion) {
            $this.Motion = $status.Motion
        }
    }
}

enum ViewItemImageQuality {
    Full = 100
    SuperHigh = 101
    High = 102
    Medium = 103
    Low = 104
}

enum ViewItemPtzMode {
    Default
    ClickToCenter
    VirtualJoystick
}

class VmsCameraViewItemProperties {
    # These represent the default XProtect Smart Client camera view item properties
    [guid]   $Id = [guid]::NewGuid()
    [guid]   $SmartClientId = [guid]::NewGuid()
    [guid]   $CameraId = [guid]::Empty
    [string] $CameraName = [string]::Empty
    [nullable[int]] $Shortcut = $null
    [guid]   $LiveStreamId = [guid]::Empty
    [ValidateRange(100, 104)]
    [int]    $ImageQuality = [ViewItemImageQuality]::Full
    [int]    $Framerate = 0
    [bool]   $MaintainImageAspectRatio = $true
    [bool]   $UseDefaultDisplaySettings = $true
    [bool]   $ShowTitleBar = $true
    [bool]   $KeepImageQualityWhenMaximized = $false
    [bool]   $UpdateOnMotionOnly = $false
    [bool]   $SoundOnMotion = $false
    [bool]   $SoundOnEvent = $false
    [int]    $SmartSearchGridWidth = 0
    [int]    $SmartSearchGridHeight = 0
    [string] $SmartSearchGridMask = [string]::Empty
    [ValidateRange(0, 2)]
    [int]    $PointAndClickMode = [ViewItemPtzMode]::Default
}

class VmsViewGroupAcl {
    [VideoOS.Platform.ConfigurationItems.Role] $Role
    [string] $Path
    [hashtable] $SecurityAttributes
}

class SecurityNamespaceTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($null -eq $inputData -or $inputData.Count -eq 0) { return [guid]::Empty }
        if ($inputData -is [guid] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [guid])) {
            return $inputData
        }
        if ($inputData.SecurityNamespace) {
            $inputData = $inputData.SecurityNamespace
        }
        if ($inputData -is [string] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [string])) {
            $securityNamespaces = Get-SecurityNamespaceValues
            $result = [string[]]@()
            foreach ($value in $inputData) {
                $id = [guid]::Empty
                if (-not [guid]::TryParse($value, [ref]$id)) {
                    try {
                        $id = if ($securityNamespaces.SecurityNamespacesByName.ContainsKey($value)) { $securityNamespaces.SecurityNamespacesByName[$value] } else { $value }
                    } catch {
                        $id = $value
                    }
                    $result += $id
                } else {
                    $result += $id
                }
            }
            if ($result.Count -eq 0) {
                throw 'No matching SecurityNamespace(s) found.'
            }
            if ($inputData -is [string]) {
                return $result[0]
            }
            return $result
        }
        throw "Unexpected type '$($inputData.GetType().FullName)'"
    }

    [string] ToString() {
        return '[SecurityNamespaceTransformAttribute()]'
    }
}

class TimeProfileNameTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($inputData -is [VideoOS.Platform.ConfigurationItems.TimeProfile] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [VideoOS.Platform.ConfigurationItems.TimeProfile])) {
            return $inputData
        }
        try {
            if ($inputData.TimeProfile) {
                $inputData = $inputData.TimeProfile
            }
            if ($inputData -is [string] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [string])) {
                $items = $inputData | ForEach-Object {
                    if ($_ -eq 'Always') {
                        @(
                            $always = [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]@{
                                DisplayName  = 'Always'
                                ItemCategory = 'Item'
                                ItemType     = 'TimeProfile'
                                Path         = 'TimeProfile[11111111-1111-1111-1111-111111111111]'
                                ParentPath   = '/TimeProfileFolder'
                            }
                            [VideoOS.Platform.ConfigurationItems.TimeProfile]::new((Get-VmsManagementServer).ServerId, $always)
                        )
                    } elseif ($_ -eq 'Default') {
                        @(
                            $default = [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]@{
                                DisplayName  = 'Default'
                                ItemCategory = 'Item'
                                ItemType     = 'TimeProfile'
                                Path         = 'TimeProfile[00000000-0000-0000-0000-000000000000]'
                                ParentPath   = '/TimeProfileFolder'
                            }
                            [VideoOS.Platform.ConfigurationItems.TimeProfile]::new((Get-VmsManagementServer).ServerId, $default)
                        )
                    } else {
                        (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | Where-Object Name -EQ $_
                    }
                }
                if ($items.Count -eq 0) {
                    throw 'No matching TimeProfile(s) found.'
                }
                if ($inputData -is [string]) {
                    return $items[0]
                } else {
                    return $items
                }
            } else {
                throw "Unexpected type '$($inputData.GetType().FullName)'"
            }
        } catch {
            throw $_.Exception
        }
    }

    [string] ToString() {
        return '[TimeProfileNameTransformAttribute()]'
    }
}

class StorageNameTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($inputData -is [VideoOS.Platform.ConfigurationItems.Storage] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [VideoOS.Platform.ConfigurationItems.Storage])) {
            return $inputData
        }
        try {
            if ($inputData.Storage) {
                $inputData = $inputData.Storage
            }
            if ($inputData -is [string] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [string])) {
                $items = $inputData | ForEach-Object {
                    Get-VmsRecordingServer | Get-VmsStorage | Where-Object Name -EQ $_
                }
                if ($items.Count -eq 0) {
                    throw 'No matching storage(s) found.'
                }
                return $items
            } else {
                throw "Unexpected type '$($inputData.GetType().FullName)'"
            }
        } catch {
            throw $_.Exception
        }
    }

    [string] ToString() {
        return '[StorageNameTransformAttribute()]'
    }
}

class BooleanTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($inputData -is [bool]) {
            return $inputData
        } elseif ($inputData -is [string]) {
            return [bool]::Parse($inputData)
        } elseif ($inputData -is [int]) {
            return [bool]$inputData
        } elseif ($inputData -is [VideoOS.ConfigurationApi.ClientService.EnablePropertyInfo]) {
            return $inputData.Enabled
        }
        throw "Unexpected type '$($inputData.GetType().FullName)'"
    }

    [string] ToString() {
        return '[BooleanTransformAttribute()]'
    }
}

class ReplaceHardwareTaskInfo {
    [string]
    $HardwareName

    [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]
    $HardwarePath

    [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]
    $RecorderPath

    [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
    $Task
}

class HardwareDriverTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        $driversById = @{}
        $driversByName = @{}
        $rec = $null
        return ($inputData | ForEach-Object {
                $obj = $_
                if ($obj -is [VideoOS.Platform.ConfigurationItems.HardwareDriver]) {
                    $obj
                    return
                }

                if ($driversById.Count -eq 0) {
                    $rec = Get-VmsRecordingServer | Select-Object -First 1
                    $rec | Get-VmsHardwareDriver | ForEach-Object {
                        $driversById[$_.Number] = $_
                        $driversByName[$_.Name] = $_
                    }
                }
                switch ($obj.GetType()) {
                ([int]) {
                        if (-not $driversById.ContainsKey($obj)) {
                            throw [VideoOS.Platform.PathNotFoundMIPException]::new('Hardware driver with ID {0} not found on recording server "{1}".' -f $obj, $_)
                        }
                        $driversById[$obj]
                    }

                ([string]) {
                        $driversByName[$obj]
                    }

                    default {
                        throw [System.InvalidOperationException]::new("Unable to transform object of type $($_.FullName) to type VideoOS.Platform.ConfigurationItems.HardwareDriver")
                    }
                }
            })
    }

    [string] ToString() {
        return '[HardwareDriverTransformAttribute()]'
    }
}

class SecureStringTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        return ($inputData | ForEach-Object {
                $obj = $_
                if ($obj -as [securestring]) {
                    $obj
                    return
                }
                if ($null -eq $obj -or $obj -isnot [string]) {
                    throw 'Expected object of type SecureString or String.'
                }
                $obj | ConvertTo-SecureString -AsPlainText -Force
            })
    }

    [string] ToString() {
        return '[SecureStringTransformAttribute()]'
    }
}

class BoolTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        return ($inputData | ForEach-Object {
                $obj = $_
                if ($obj -is [bool]) {
                    $obj
                    return
                }
                if ($null -eq $obj -or -not [bool]::TryParse($obj, [ref]$obj)) {
                    throw "Failed to parse '$obj' as [bool]"
                }
                $obj
            })
    }

    [string] ToString() {
        return '[BoolTransformAttribute()]'
    }
}

class ClaimTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($inputData -is [VideoOS.Platform.ConfigurationItems.ClaimChildItem] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [VideoOS.Platform.ConfigurationItems.ClaimChildItem])) {
            return $inputData
        }
        try {
            if ($inputData.Claim) {
                $inputData = $inputData.Claim
            }
            if ($inputData -is [string] -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is [string])) {
                $items = Get-VmsLoginProvider | Where-Object { $_.Name -eq $inputData -or $_.Id -eq $inputData }
                if ($inputData -is [string]) {
                    return $items[0]
                }
                return $items
            } else {
                throw "Unexpected type '$($inputData.GetType().FullName)'"
            }
        } catch {
            throw $_.Exception
        }
    }

    [string] ToString() {
        return '[LoginProviderTransformAttribute()]'
    }
}

class PropertyCollectionTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        if ($inputData -is [System.Collections.IDictionary]) {
            return $inputData
        }
        try {
            $hashtable = @{}
            $inputData.GetEnumerator() | ForEach-Object {
                if ($null -eq ($_ | Get-Member -Name Key) -or $null -eq ($_ | Get-Member -Name Value)) {
                    throw 'Key and Value properties most both be present in a property collection.'
                }
                $hashtable[$_.Key] = $_.Value
            }
            return $hashtable
        } catch {
            throw $_.Exception
        }
    }

    [string] ToString() {
        return '[PropertyCollectionTransformAttribute()]'
    }
}

class RuleNameTransformAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics]$engineIntrinsics, [object] $inputData) {
        $expectedType = [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $itemType = 'Rule'

        if ($inputData -is $expectedType -or ($inputData -is [system.collections.ienumerable] -and $inputData[0] -is $expectedType)) {
            return $inputData
        }
        try {
            $items = $inputData | ForEach-Object {
                $stringValue = $_.ToString() -replace "^$ItemType\[(.+)\](?:/.+)?", '$1'
                $id = [guid]::Empty
                if ([guid]::TryParse($stringValue, [ref]$id)) {
                    Get-VmsRule | Where-Object Path -Match $stringValue
                } else {
                    Get-VmsRule | Where-Object DisplayName -EQ $stringValue
                }
            }
            if ($null -eq $items) {
                throw ([System.Management.Automation.ItemNotFoundException]::new("$itemType '$($inputData)' not found."))
            }
            return $items
        } catch {
            throw $_.Exception
        }
    }

    [string] ToString() {
        return '[RuleNameTransformAttribute()]'
    }
}

function BuildGroupsOfGivenSize {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [object[]]
        $InputObject,

        [Parameter(Mandatory, Position = 0)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $GroupSize,

        [Parameter()]
        [object]
        $EmptyItem = $null,

        [Parameter()]
        [switch]
        $TrimLastGroup

    )

    begin {
        $allObjects = [collections.generic.list[object]]::new()
        $groupOfGroups = [collections.generic.list[[collections.generic.list[object]]]]::new()
    }
    
    process {
        foreach ($obj in $InputObject) {
            $allObjects.Add($obj)
        }
    }
    
    end {
        $index = 0
        do {
            $group = [collections.generic.list[object]]::new()
            for ($i = 0; $i -lt $GroupSize; $i++) {
                $pos = $index + $i
                if ($pos -lt $allObjects.Count) {
                    $group.Add($allObjects[$pos])
                } elseif (!$TrimLastGroup) {
                    $group.Add($EmptyItem)
                }
            }
            $groupOfGroups.Add($group)
            $index += $GroupSize
        } while ($index -lt $allObjects.Count)
        $groupOfGroups
    }
}

function Complete-SimpleArgument {
    <#
    .SYNOPSIS
    Implements a simple argument-completer.
    .DESCRIPTION
    This cmdlet is a helper function that implements a basic argument completer
    which matches the $wordToComplete against a set of values that can be
    supplied in the form of a string array, or produced by a scriptblock you
    provide to the function.
    .PARAMETER Arguments
    The original $args array passed from Register-ArgumentCompleter into the
    scriptblock.
    .PARAMETER ValueSet
    An array of strings representing the valid values for completion.
    .PARAMETER Completer
    A scriptblock which produces an array of strings representing the valid values for completion.
    .EXAMPLE
    Register-ArgumentCompleter -CommandName Get-VmsRole -ParameterName Name -ScriptBlock {
        Complete-SimpleArgument $args {(Get-VmsManagementServer).RoleFolder.Roles.Name}
    }
    Registers an argument completer for the Name parameter on the Get-VmsRole
    command. Complete-SimpleArgument cmdlet receives the $args array, and a
    simple scriptblock which returns the names of all roles in the VMS.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [object[]]
        $Arguments,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ValuesFromArray')]
        [string[]]
        $ValueSet,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ValuesFromScriptBlock')]
        [scriptblock]
        $Completer
    )

    process {
        # Get ValueSet from scriptblock if provided, otherwise use $ValueSet.
        if ($PSCmdlet.ParameterSetName -eq 'ValuesFromScriptBlock') {
            $ValueSet = $Completer.Invoke($Arguments)
        }

        # Trim single/double quotes off of beginning of word if present. If no
        # characters have been provided, set the word to "*" for wildcard matching.
        if ([string]::IsNullOrWhiteSpace($Arguments[2])) {
            $wordToComplete = '*'
        } else {
            $wordToComplete = $Arguments[2].Trim('''').Trim('"')
        }

        # Return matching values from ValueSet.
        $ValueSet | Foreach-Object {
            if ($_ -like "$wordToComplete*") {
                if ($_ -like '* *') {
                    "'$_'"
                } else {
                    $_
                }
            }
        }
    }
}


class VmsConfigChildItemSettings {
    [string]    $Name
    [hashtable] $Properties
    [hashtable] $ValueTypeInfo
}

function ConvertFrom-ConfigChildItem {
    [CmdletBinding()]
    [OutputType([VmsConfigChildItemSettings])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationChildItem]
        $InputObject,

        [Parameter()]
        [switch]
        $RawValues
    )

    process {
        # When we look up display values for raw values, sometimes
        # the raw value matches the value of a valuetypeinfo property
        # like MinValue or MaxValue. We don't want to display "MinValue"
        # as the display value for a setting, so this list of valuetypeinfo
        # entry names should be ignored.
        $ignoredNames = 'MinValue', 'MaxValue', 'StepValue'
        $properties = @{}
        $valueTypeInfos = @{}
        foreach ($key in $InputObject.Properties.Keys) {
            # Sometimes the Keys are the same as KeyFullName and other times
            # they are short, easy to read names. So just in case, we'll test
            # the key by splitting it and seeing how many parts there are. A
            # KeysFullName value looks like 'device:0.0/RecorderMode/75f374ab-8dd2-4fd0-b8f5-155fa730702c'
            $keyParts = $key -split '/', 3
            $keyName = if ($keyParts.Count -gt 1) { $keyParts[1] } else { $key }

            $value = $InputObject.Properties.GetValue($key)
            $valueTypeInfo = $InputObject.Properties.GetValueTypeInfoCollection($key)

            if (-not $RawValues) {
                <#
                  Unless -RawValues was used, we'll check to see if there's a
                  display name available for the value for the current setting.
                  If a ValueTypeInfo entry has a Value matching the raw value,
                  and the Name of that value isn't one of the internal names we
                  want to ignore, we'll replace $value with the ValueTypeInfo
                  Name. Here's a reference ValueTypeInfo table for RecorderMode:

                  TranslationId                        Name       Value
                  -------------                        ----       -----
                  b9f5c797-ebbf-55ad-ccdd-8539a65a0241 Disabled   0
                  535863a8-2f16-3709-557e-59e2eb8139a7 Continuous 1
                  8226588f-03da-49b8-57e5-ddf8c508dd2d Motion     2

                  So if the raw value of RecorderMode is 0, we would return
                  "Disabled" unless the -RawValues switch is used.
                #>

                $friendlyValue = ($valueTypeInfo | Select-Object | Where-Object {
                        $_.Value -eq $value -and $_.Name -notin $ignoredNames
                    }).Name
                if (-not [string]::IsNullOrWhiteSpace($friendlyValue)) {
                    $value = $friendlyValue
                }
            }

            $properties[$keyName] = $value
            $valueTypeInfos[$keyName] = $valueTypeInfo
        }

        [VmsConfigChildItemSettings]@{
            Name          = $InputObject.DisplayName
            Properties    = $properties
            ValueTypeInfo = $valueTypeInfos
        }
    }
}


function ConvertFrom-StreamUsage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.StreamUsageChildItem]
        $StreamUsage
    )

    process {
        $streamName = $StreamUsage.StreamReferenceIdValues.Keys | Where-Object {
            $StreamUsage.StreamReferenceIdValues.$_ -eq $StreamUsage.StreamReferenceId
        }
        Write-Output $streamName
    }
}

function ConvertTo-ConfigItemPath {
    [CmdletBinding()]
    [OutputType([videoos.platform.proxy.ConfigApi.ConfigurationItemPath])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Path
    )

    process {
        foreach ($p in $Path) {
            try {
                [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]::new($p)
            } catch {
                Write-Error -Message "The value '$p' is not a recognized configuration item path format." -Exception $_.Exception
            }
        }
    }
}


function ConvertTo-PSCredential {
    [CmdletBinding()]
    [OutputType([pscredential])]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [System.Net.NetworkCredential]
        $NetworkCredential
    )
        
    process {
        if ([string]::IsNullOrWhiteSpace($NetworkCredential.UserName)) {
            Write-Error 'NetworkCredential username is empty. This usually means the credential is the default network credential and this cannot be converted to a pscredential.'
            return
        }
        $sb = [text.stringbuilder]::new()
        if (-not [string]::IsNullOrWhiteSpace($NetworkCredential.Domain)) {
            [void]$sb.Append("$($NetworkCredential.Domain)\")
        }
        [void]$sb.Append($NetworkCredential.UserName)
        [pscredential]::new($sb.ToString(), $NetworkCredential.SecurePassword)
    }
}

function ConvertTo-Sid {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]
        $AccountName,

        [Parameter()]
        [string]
        $Domain
    )

    process {
        try {
            if ($AccountName -match '^\[BASIC\]\\(?<username>.+)$') {
                $sid = (Get-VmsManagementServer).BasicUserFolder.BasicUsers | Where-Object Name -eq $Matches.username | Select-Object -ExpandProperty Sid
                if ($sid) {
                    $sid
                } else {
                    throw "No basic user found matching '$AccountName'"
                }
            } else {
                [System.Security.Principal.NTAccount]::new($Domain, $AccountName).Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
        } catch [System.Security.Principal.IdentityNotMappedException] {
            Write-Error -ErrorRecord $_
        }
    }
}


function ConvertTo-StringFromSecureString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [securestring]
        $SecureString
    )

    process {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        try {
            [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        } finally {
            [System.Runtime.InteropServices.Marshal]::FreeBSTR($bstr)
        }
    }
}


function ConvertTo-Uri {
    <#
    .SYNOPSIS
    Accepts an IPv4 or IPv6 address and converts it to an http or https URI

    .DESCRIPTION
    Accepts an IPv4 or IPv6 address and converts it to an http or https URI. IPv6 addresses need to
    be wrapped in square brackets when used in a URI. This function is used to help normalize data
    into an expected URI format.

    .PARAMETER IPAddress
    Specifies an IPAddress object of either Internetwork or InternetworkV6.

    .PARAMETER UseHttps
    Specifies whether the resulting URI should use https as the scheme instead of http.

    .PARAMETER HttpPort
    Specifies an alternate port to override the default http/https ports.

    .EXAMPLE
    '192.168.1.1' | ConvertTo-Uri
    #>
    [CmdletBinding()]
    [OutputType([uri])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [IPAddress]
        $IPAddress,

        [Parameter()]
        [switch]
        $UseHttps,

        [Parameter()]
        [int]
        $HttpPort = 80
    )

    process {
        $builder = [uribuilder]::new()
        $builder.Scheme = if ($UseHttps) { 'https' } else { 'http' }
        $builder.Host = if ($IPAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
            "[$IPAddress]"
        }
        else {
            $IPAddress
        }
        $builder.Port = $HttpPort
        Write-Output $builder.Uri
    }
}


function ConvertTo-Webhook {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $InputObject
    )

    process {
        try {
            [MilestonePSTools.Webhook]$InputObject
        } catch {
            Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $InputObject
        }
    }
}

function Copy-ConfigurationItem {
    [CmdletBinding()]
    param (
        [parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [pscustomobject]
        $InputObject,
        [parameter(Mandatory, Position = 1)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $DestinationItem
    )

    process {
        if (!$DestinationItem.ChildrenFilled) {
            Write-Verbose "$($DestinationItem.DisplayName) has not been retrieved recursively. Retrieving child items now."
            $DestinationItem = $DestinationItem | Get-ConfigurationItem -Recurse -Sort
        }

        $srcStack = New-Object -TypeName System.Collections.Stack
        $srcStack.Push($InputObject)
        $dstStack = New-Object -TypeName System.Collections.Stack
        $dstStack.Push($DestinationItem)

        Write-Verbose "Configuring $($DestinationItem.DisplayName) ($($DestinationItem.Path))"
        while ($dstStack.Count -gt 0) {
            $dirty = $false
            $src = $srcStack.Pop()
            $dst = $dstStack.Pop()

            if (($src.ItemCategory -ne $dst.ItemCategory) -or ($src.ItemType -ne $dst.ItemType)) {
                Write-Error "Source and Destination ConfigurationItems are different"
                return
            }

            if ($src.EnableProperty.Enabled -ne $dst.EnableProperty.Enabled) {
                Write-Verbose "$(if ($src.EnableProperty.Enabled) { "Enabling"} else { "Disabling" }) $($dst.DisplayName)"
                $dst.EnableProperty.Enabled = $src.EnableProperty.Enabled
                $dirty = $true
            }

            $srcChan = $src.Properties | Where-Object { $_.Key -eq "Channel"} | Select-Object -ExpandProperty Value
            $dstChan = $dst.Properties | Where-Object { $_.Key -eq "Channel"} | Select-Object -ExpandProperty Value
            if ($srcChan -ne $dstChan) {
                Write-Error "Sorting mismatch between source and destination configuration."
                return
            }

            foreach ($srcProp in $src.Properties) {
                $dstProp = $dst.Properties | Where-Object Key -eq $srcProp.Key
                if ($null -eq $dstProp) {
                    Write-Verbose "Key '$($srcProp.Key)' not found on $($dst.Path)"
                    Write-Verbose "Available keys`r`n$($dst.Properties | Select-Object Key, Value | Format-Table)"
                    continue
                }
                if (!$srcProp.IsSettable -or $srcProp.ValueType -eq 'PathList' -or $srcProp.ValueType -eq 'Path') { continue }
                if ($srcProp.Value -ne $dstProp.Value) {
                    Write-Verbose "Changing $($dstProp.DisplayName) to $($srcProp.Value) on $($dst.Path)"
                    $dstProp.Value = $srcProp.Value
                    $dirty = $true
                }
            }
            if ($dirty) {
                if ($dst.ItemCategory -eq "ChildItem") {
                    $result = $lastParent | Set-ConfigurationItem
                } else {
                    $result = $dst | Set-ConfigurationItem
                }

                if (!$result.ValidatedOk) {
                    foreach ($errorResult in $result.ErrorResults) {
                        Write-Error $errorResult.ErrorText
                    }
                }
            }

            if ($src.Children.Count -eq $dst.Children.Count -and $src.Children.Count -gt 0) {
                foreach ($child in $src.Children) {
                    $srcStack.Push($child)
                }
                foreach ($child in $dst.Children) {
                    $dstStack.Push($child)
                }
                if ($dst.ItemCategory -eq "Item") {
                    $lastParent = $dst
                }
            } elseif ($src.Children.Count -ne 0) {
                Write-Warning "Number of child items is not equal on $($src.DisplayName)"
            }
        }
    }
}

function Copy-ViewGroupFromJson {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [pscustomobject]
        $Source,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $NewName,

        [Parameter()]
        [ValidateNotNull()]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $ParentViewGroup
    )

    process {
        if ($MyInvocation.BoundParameters.ContainsKey('NewName')) {
            ($source.Properties | Where-Object Key -eq 'Name').Value = $NewName
        }

        ##
        ## Clean duplicate views in export caused by config api bug
        ##

        $groups = [system.collections.generic.queue[pscustomobject]]::new()
        $groups.Enqueue($source)
        $views = [system.collections.generic.list[pscustomobject]]::new()
        while ($groups.Count -gt 0) {
            $group = $groups.Dequeue()
            $views.Clear()
            foreach ($v in ($group.Children | Where-Object ItemType -eq 'ViewFolder').Children) {
                if ($v.Path -notin (($group.Children | Where-Object ItemType -eq 'ViewGroupFolder').Children.Children | Where-Object ItemType -eq 'ViewFolder').Children.Path) {
                    $views.Add($v)
                } else {
                    Write-Verbose "Skipping duplicate view"
                }
            }
            if ($null -ne ($group.Children | Where-Object ItemType -eq 'ViewFolder').Children) {
                ($group.Children | Where-Object ItemType -eq 'ViewFolder').Children = $views.ToArray()
            }
            foreach ($childGroup in ($group.Children | Where-Object ItemType -eq 'ViewGroupFolder').Children) {
                $groups.Enqueue($childGroup)
            }
        }


        $rootFolder = Get-ConfigurationItem -Path /ViewGroupFolder
        if ($null -ne $ParentViewGroup) {
            $rootFolder = $ParentViewGroup.ViewGroupFolder | Get-ConfigurationItem
        }
        $newViewGroup = $null
        $stack = [System.Collections.Generic.Stack[pscustomobject]]::new()
        $stack.Push(([pscustomobject]@{ Folder = $rootFolder; Group = $source }))
        while ($stack.Count -gt 0) {
            $entry = $stack.Pop()
            $parentFolder = $entry.Folder
            $srcGroup = $entry.Group

            ##
            ## Create matching ViewGroup
            ##
            $invokeInfo = $parentFolder | Invoke-Method -MethodId 'AddViewGroup'
            foreach ($key in ($srcGroup.Properties | Where-Object IsSettable).Key) {
                $value = ($srcGroup.Properties | Where-Object Key -eq $key).Value
                ($invokeInfo.Properties | Where-Object Key -eq $key).Value = $value
            }
            $invokeResult = $invokeInfo | Invoke-Method -MethodId 'AddViewGroup'
            $props = ConvertPropertiesToHashtable -Properties $invokeResult.Properties
            if ($props.State.Value -ne 'Success') {
                Write-Error $props.ErrorText
            }
            $newViewFolder = Get-ConfigurationItem -Path "$($props.Path.Value)/ViewFolder"
            $newViewGroupFolder = Get-ConfigurationItem -Path "$($props.Path.Value)/ViewGroupFolder"
            if ($null -eq $newViewGroup) {
                $serverId = (Get-VmsManagementServer).ServerId
                $newViewGroup = [VideoOS.Platform.ConfigurationItems.ViewGroup]::new($serverId, $props.Path.Value)
            }

            ##
            ## Create all child views of the current view group
            ##
            foreach ($srcView in ($srcGroup.Children | Where-Object ItemType -eq ViewFolder).Children) {
                # Create new view based on srcView layout
                $invokeInfo = $newViewFolder | Invoke-Method -MethodId 'AddView'
                foreach ($key in ($invokeInfo.Properties | Where-Object IsSettable).Key) {
                    $value = ($srcView.Properties | Where-Object Key -eq $key).Value
                    ($invokeInfo.Properties | Where-Object Key -eq $key).Value = $value
                }
                $newView = $invokeInfo | Invoke-Method -MethodId 'AddView'

                # Rename view and update any other settable values
                foreach ($key in ($newView.Properties | Where-Object IsSettable).Key) {
                    $value = ($srcView.Properties | Where-Object Key -eq $key).Value
                    ($newView.Properties | Where-Object Key -eq $key).Value = $value
                }

                # Update all viewitems of new view to match srcView
                for ($i = 0; $i -lt $newView.Children.Count; $i++) {
                    foreach ($key in ($newView.Children[$i].Properties | Where-Object IsSettable).Key) {
                        $value = ($srcView.Children[$i].Properties | Where-Object Key -eq $key).Value
                        ($newView.Children[$i].Properties | Where-Object Key -eq $key).Value = $value
                    }
                }

                # Save changes to new view
                $invokeResult = $newView | Invoke-Method -MethodId 'AddView'
                $props = ConvertPropertiesToHashtable -Properties $invokeResult.Properties
                if ($props.State.Value -ne 'Success') {
                    Write-Error $props.ErrorText
                }
            }

            ##
            ## Get the new child ViewGroupFolder, and add all child view groups from the JSON object to the stack
            ##
            foreach ($childViewGroup in ($srcGroup.Children | Where-Object ItemType -eq ViewGroupFolder).Children) {
                $stack.Push(([pscustomobject]@{ Folder = $newViewGroupFolder; Group = $childViewGroup }))
            }
        }

        if ($null -ne $newViewGroup) {
            Write-Output $newViewGroup
        }
    }
}

function ConvertPropertiesToHashtable {
    param([VideoOS.ConfigurationApi.ClientService.Property[]]$Properties)

    $props = @{}
    foreach ($prop in $Properties) {
        $props[$prop.Key] = $prop
    }
    Write-Output $props
}


function ExecuteWithRetry {
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [scriptblock]
        $ScriptBlock,

        [Parameter()]
        [object[]]
        $ArgumentList = [object[]]::new(0),

        [Parameter()]
        [int]
        $Attempts = 2,

        [Parameter()]
        [switch]
        $ClearVmsCache
    )

    process {
        do {
            try {
                $ScriptBlock.Invoke($ArgumentList)
                break
            } catch {
                if ($Attempts -gt 1) {
                    Write-Verbose "ExecuteWithRetry: Failed with $_"
                    if ($ClearVmsCache) {
                        Clear-VmsCache
                    }
                    Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
                    continue
                }
                throw
            }
        } while ((--$Attempts) -gt 0)
    }
}


class CidrInfo {
    [string] $Cidr
    [IPAddress] $Address
    [int] $Mask

    [IPAddress] $Start
    [IPAddress] $End
    [IPAddress] $SubnetMask
    [IPAddress] $HostMask

    [int] $TotalAddressCount
    [int] $HostAddressCount

    CidrInfo([string] $Cidr) {
        [System.Net.IPAddress]$this.Address, [int]$this.Mask = $Cidr -split '/'
        if ($this.Address.AddressFamily -notin @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
            throw "CidrInfo is not compatible with AddressFamily $($this.Address.AddressFamily). Expected InterNetwork or InterNetworkV6."
        }
        $min, $max = if ($this.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) { 0, 32 } else { 0, 128 }
        if ($this.Mask -lt $min -or $this.Mask -gt $max) {
            throw "CIDR mask value out of range. Expected a value between $min and $max for AddressFamily $($this.Address.AddressFamily)"
        }
        $hostMaskLength = $max - $this.Mask
        $this.Cidr = $Cidr
        $this.TotalAddressCount = [math]::pow(2, $hostMaskLength)
        # RFC 3021 support is assumed. When the range supports only two hosts, RFC 3021 defines it usable for point-to-point communications but not all systems support this.
        $this.HostAddressCount = if ($hostMaskLength -eq 0) { 1 } elseif ($hostMaskLength -eq 1) { 2 } else { $this.TotalAddressCount - 2 }

        $addressBytes = $this.Address.GetAddressBytes()
        $netMaskBytes = [byte[]]::new($addressBytes.Count)
        $hostMaskBytes = [byte[]]::new($addressBytes.Count)
        $bitCounter = 0
        for ($octet = 0; $octet -lt $addressBytes.Count; $octet++) {
            for ($bit = 0; $bit -lt 8; $bit++) {
                $bitCounter += 1
                $bitValue = 0
                if ($bitCounter -le $this.Mask) {
                    $bitValue = 1
                }
                $netMaskBytes[$octet] = $netMaskBytes[$octet] -bor ( $bitValue -shl ( 7 - $bit ) )
                $hostMaskBytes[$octet] = $netMaskBytes[$octet] -bxor 255
            }
        }
        $this.SubnetMask = [ipaddress]::new($netMaskBytes)
        $this.HostMask = [IPAddress]::new($hostMaskBytes)

        $startBytes = [byte[]]::new($addressBytes.Count)
        $endBytes = [byte[]]::new($addressBytes.Count)
        for ($octet = 0; $octet -lt $addressBytes.Count; $octet++) {
            $startBytes[$octet] = $addressBytes[$octet] -band $netMaskBytes[$octet]
            $endBytes[$octet] = $addressBytes[$octet] -bor $hostMaskBytes[$octet]
        }
        $this.Start = [IPAddress]::new($startBytes)
        $this.End = [IPAddress]::new($endBytes)
    }
}

function Expand-IPRange {
    <#
    .SYNOPSIS
    Expands a start and end IP address or a CIDR notation into an array of IP addresses within the given range.

    .DESCRIPTION
    Accepts start and end IP addresses in the form of IPv4 or IPv6 addresses, and returns each IP
    address falling within the range including the Start and End values.

    The Start and End IP addresses must be in the same address family (IPv4 or IPv6) and if the
    addresses are IPv6, they must have the same scope ID.

    .PARAMETER Start
    Specifies the first IP address in the range to be expanded.

    .PARAMETER End
    Specifies the last IP address in the range to be expanded. Must be greater than or equal to Start.

    .PARAMETER Cidr
    Specifies an IP address range in CIDR notation. Example: 192.168.0.0/23 represents 192.168.0.0-192.168.1.255.

    .PARAMETER AsString
    Specifies that each IP address in the range should be returned as a string instead of an [IPAddress] object.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start 192.168.1.1 -End 192.168.2.255
    Returns 511 IPv4 IPAddress objects.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start fe80::5566:e22e:3f34:5a0f -End fe80::5566:e22e:3f34:5a16
    Returns 8 IPv6 IPAddress objects.

    .EXAMPLE
    PS C:\> Expand-IPRange -Start 10.1.1.100 -End 10.1.10.50 -AsString
    Returns 2255 IPv4 addresses as strings.

    .EXAMPLE
    PS C:\> Expand-IPRange -Cidr 172.16.16.0/23
    Returns IPv4 IPAddress objects from 172.16.16.0 to 172.16.17.255.
    #>
    [CmdletBinding(DefaultParameterSetName = 'FromRange')]
    [OutputType([System.Net.IPAddress], [string])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'FromRange')]
        [ValidateScript({
            if ($_.AddressFamily -in @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
                return $true
            }
            throw "Start IPAddress is from AddressFamily '$($_.AddressFamily)'. Expected InterNetwork or InterNetworkV6."
        })]
        [System.Net.IPAddress]
        $Start,

        [Parameter(Mandatory, ParameterSetName = 'FromRange')]
        [ValidateScript({
            if ($_.AddressFamily -in @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6)) {
                return $true
            }
            throw "Start IPAddress is from AddressFamily '$($_.AddressFamily)'. Expected InterNetwork or InterNetworkV6."
        })]
        [System.Net.IPAddress]
        $End,

        [Parameter(Mandatory, ParameterSetName = 'FromCidr')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Cidr,

        [Parameter()]
        [switch]
        $AsString
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'FromCidr') {
            $cidrInfo = [CidrInfo]$Cidr
            $Start = $cidrInfo.Start
            $End = $cidrInfo.End
        }

        if (-not $Start.AddressFamily.Equals($End.AddressFamily)) {
            throw 'Expand-IPRange received Start and End addresses from different IP address families (IPv4 and IPv6). Both addresses must be of the same IP address family.'
        }

        if ($Start.ScopeId -ne $End.ScopeId) {
            throw 'Expand-IPRange received IPv6 Start and End addresses with different ScopeID values. The ScopeID values must be identical.'
        }

        # Assert that the End IP is greater than or equal to the Start IP.
        $startBytes = $Start.GetAddressBytes()
        $endBytes = $End.GetAddressBytes()
        for ($i = 0; $i -lt $startBytes.Length; $i++) {
            if ($endBytes[$i] -lt $startBytes[$i]) {
                throw 'Expand-IPRange must receive an End IPAddress which is greater than or equal to the Start IPAddress'
            }
            if ($endBytes[$i] -gt $startBytes[$i]) {
                # We can break early if a higher-order byte from the End address is greater than the matching byte of the Start address
                break
            }
        }

        $current = $Start
        while ($true) {
            if ($AsString) {
                Write-Output $current.ToString()
            }
            else {
                Write-Output $current
            }

            if ($current.Equals($End)) {
                break
            }

            $bytes = $current.GetAddressBytes()
            for ($i = $bytes.Length - 1; $i -ge 0; $i--) {
                if ($bytes[$i] -lt 255) {
                    $bytes[$i] += 1
                    break
                }
                $bytes[$i] = 0
            }
            if ($null -ne $current.ScopeId) {
                $current = [System.Net.IPAddress]::new($bytes, $current.ScopeId)
            }
            else {
                $current = [System.Net.IPAddress]::new($bytes)
            }
        }
    }
}


function ExportHardwareCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [VideoOS.Platform.ConfigurationItems.Hardware[]]
        $Hardware,

        [Parameter()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output')]
        [string[]]
        $DeviceType = @('Camera'),

        [Parameter()]
        [ValidateSet('All', 'Enabled', 'Disabled')]
        [string]
        $EnableFilter = 'Enabled'
    )

    process {
        $recorders = @{}
        $storage = @{}
        $deviceGroupsById = @{}
        $DeviceType | ForEach-Object {
            Get-VmsDeviceGroup -Type $_ -Recurse | ForEach-Object {
                $group = $_
                $groupPath = $group | Resolve-VmsDeviceGroupPath -NoTypePrefix
                foreach ($device in $group | Get-VmsDeviceGroupMember -EnableFilter $EnableFilter) {
                    if (-not $deviceGroupsById.ContainsKey($device.Id)) {
                        $deviceGroupsById[$device.Id] = [collections.generic.list[string]]::new()
                    }
                    $deviceGroupsById[$device.Id].Add($groupPath)
                }
            }
        }
        foreach ($hw in $Hardware) {
            if (-not $recorders.ContainsKey($hw.ParentItemPath)) {
                $recorders[$hw.ParentItemPath] = $hw | Get-VmsParentItem
            }
            $recorder = $recorders[$hw.ParentItemPath]
            
            try {
                $password = $hw | Get-VmsHardwarePassword
                $driver = $hw | Get-VmsHardwareDriver
            } catch {
                $password = $null
                $driver = $null
            }
            
            $splat = @{
                Type         = $DeviceType
                EnableFilter = $EnableFilter
            }
            foreach ($device in $hw | Get-VmsDevice @splat) {
                if ($null -ne $device.RecordingStorage -and -not $storage.ContainsKey($device.RecordingStorage)) {
                    $storage[$device.RecordingStorage] = $recorder | Get-VmsStorage | Where-Object Path -eq $device.RecordingStorage
                }
                $storageName = if ($device.RecordingStorage) { $storage[$device.RecordingStorage].Name } else { $null }
                $coordinates = if ($device.GisPoint -ne 'POINT EMPTY') { $device.GisPoint | ConvertFrom-GisPoint } else { $null }
                [pscustomobject]@{
                    DeviceType      = ($device.Path -split '\[' | Select-Object -First 1) -replace 'Event$'
                    Name            = $device.Name
                    Address         = $hw.Address
                    Channel         = $device.Channel
                    UserName        = $hw.UserName
                    Password        = $password
                    DriverNumber    = $driver.Number
                    DriverGroup    = $driver.GroupName
                    RecordingServer = $recorder.Name
                    Enabled         = $device.Enabled
                    HardwareName    = $hw.Name
                    StorageName     = $storageName
                    Coordinates     = $coordinates
                    DeviceGroups    = $deviceGroupsById[$device.Id] -join ';'
                }
            }
        }
    }
}

function ExportVmsLoginSettings {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param ()
    
    process {
        $settings = Get-LoginSettings | Where-Object Guid -EQ ([milestonepstools.connection.milestoneconnection]::Instance.MainSite).FQID.ObjectId
        $vmsProfile = @{
            ServerAddress     = $settings.Uri
            Credential        = $settings.NetworkCredential | ConvertTo-PSCredential -ErrorAction SilentlyContinue
            BasicUser         = $settings.IsBasicUser
            SecureOnly        = $settings.SecureOnly
            IncludeChildSites = [milestonepstools.connection.milestoneconnection]::Instance.IncludeChildSites
            AcceptEula        = $true
        }
        if ($null -eq $vmsProfile.Credential) {
            $vmsProfile.Remove('Credential')
        }
        $vmsProfile
    }
}

function FillChildren {
    [CmdletBinding()]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $ConfigurationItem,

        [Parameter()]
        [int]
        $Depth = 1
    )

    process {
        $stack = New-Object System.Collections.Generic.Stack[VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $stack.Push($ConfigurationItem)
        while ($stack.Count -gt 0) {
            $Depth = $Depth - 1
            $item = $stack.Pop()
            $item.Children = $item | Get-ConfigurationItem -ChildItems
            $item.ChildrenFilled = $true
            if ($Depth -gt 0) {
                $item.Children | Foreach-Object {
                    $stack.Push($_)
                }
            }
        }
        Write-Output $ConfigurationItem
    }
}


function Find-XProtectDeviceDialog {
    [CmdletBinding()]
    [RequiresInteractiveSession()]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        Add-Type -AssemblyName PresentationFramework
        $xaml = [xml]@"
        <Window
                xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:local="clr-namespace:Search_XProtect"
                Title="Search XProtect" Height="500" Width="800"
                FocusManager.FocusedElement="{Binding ElementName=cboItemType}">
            <Grid>
                <GroupBox Name="gboAdvanced" Header="Advanced Parameters" HorizontalAlignment="Left" Height="94" Margin="506,53,0,0" VerticalAlignment="Top" Width="243"/>
                <Label Name="lblItemType" Content="Item Type" HorizontalAlignment="Left" Margin="57,22,0,0" VerticalAlignment="Top"/>
                <ComboBox Name="cboItemType" HorizontalAlignment="Left" Margin="124,25,0,0" VerticalAlignment="Top" Width="120" TabIndex="0">
                    <ComboBoxItem Content="Camera" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Hardware" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="InputEvent" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Metadata" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Microphone" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Output" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Speaker" HorizontalAlignment="Left" Width="118"/>
                </ComboBox>
                <Label Name="lblName" Content="Name" HorizontalAlignment="Left" Margin="77,53,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <Label Name="lblPropertyName" Content="Property Name" HorizontalAlignment="Left" Margin="519,80,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <ComboBox Name="cboPropertyName" HorizontalAlignment="Left" Margin="614,84,0,0" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="5"/>
                <TextBox Name="txtName" HorizontalAlignment="Left" Height="23" Margin="124,56,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="187" IsEnabled="False" TabIndex="1"/>
                <Button Name="btnSearch" Content="Search" HorizontalAlignment="Left" Margin="306,154,0,0" VerticalAlignment="Top" Width="75" TabIndex="7" IsEnabled="False"/>
                <DataGrid Name="dgrResults" HorizontalAlignment="Left" Height="207" Margin="36,202,0,0" VerticalAlignment="Top" Width="719" IsReadOnly="True"/>
                <Label Name="lblAddress" Content="IP Address" HorizontalAlignment="Left" Margin="53,84,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtAddress" HorizontalAlignment="Left" Height="23" Margin="124,87,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="2"/>
                <Label Name="lblEnabledFilter" Content="Enabled/Disabled" HorizontalAlignment="Left" Margin="506,22,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <ComboBox Name="cboEnabledFilter" HorizontalAlignment="Left" Margin="614,26,0,0" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="4">
                    <ComboBoxItem Content="Enabled" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Content="Disabled" HorizontalAlignment="Left" Width="118"/>
                    <ComboBoxItem Name="cbiEnabledAll" Content="All" HorizontalAlignment="Left" Width="118" IsSelected="True"/>
                </ComboBox>
                <Label Name="lblMACAddress" Content="MAC Address" HorizontalAlignment="Left" Margin="37,115,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtMACAddress" HorizontalAlignment="Left" Height="23" Margin="124,118,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="3"/>
                <Label Name="lblPropertyValue" Content="Property Value" HorizontalAlignment="Left" Margin="522,108,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox Name="txtPropertyValue" HorizontalAlignment="Left" Height="23" Margin="614,111,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" IsEnabled="False" TabIndex="6"/>
                <Button Name="btnExportCSV" Content="Export CSV" HorizontalAlignment="Left" Margin="680,429,0,0" VerticalAlignment="Top" Width="75" TabIndex="9" IsEnabled="False"/>
                <Label Name="lblNoResults" Content="No results found!" HorizontalAlignment="Left" Margin="345,175,0,0" VerticalAlignment="Top" Foreground="Red" Visibility="Hidden"/>
                <Button Name="btnResetForm" Content="Reset Form" HorizontalAlignment="Left" Margin="414,154,0,0" VerticalAlignment="Top" Width="75" TabIndex="8"/>
                <Label Name="lblTotalResults" Content="Total Results:" HorizontalAlignment="Left" Margin="32,423,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
                <TextBox Name="txtTotalResults" HorizontalAlignment="Left" Height="23" Margin="120,427,0,0" VerticalAlignment="Top" Width="53" IsEnabled="False"/>
                <Label Name="lblPropertyNameBlank" Content="Property Name cannot be blank if Property&#xD;&#xA;Value has an entry." HorizontalAlignment="Left" Margin="507,152,0,0" VerticalAlignment="Top" Foreground="Red" Width="248" Height="45" Visibility="Hidden"/>
                <Label Name="lblPropertyValueBlank" Content="Property Value cannot be blank if Property&#xA;Name has a selection." HorizontalAlignment="Left" Margin="507,152,0,0" VerticalAlignment="Top" Foreground="Red" Width="248" Height="45" Visibility="Hidden"/>
            </Grid>
        </Window>
"@

        function Clear-Results {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Private function.')]
            param()
            $var_dgrResults.Columns.Clear()
            $var_dgrResults.Items.Clear()
            $var_txtTotalResults.Clear()
            $var_lblNoResults.Visibility = "Hidden"
            $var_lblPropertyNameBlank.Visibility = "Hidden"
            $var_lblPropertyValueBlank.Visibility = "Hidden"
        }

        $reader = [system.xml.xmlnodereader]::new($xaml)
        $window = [windows.markup.xamlreader]::Load($reader)
        $searchResults = $null

        # Create variables based on form control names.
        # Variable will be named as 'var_<control name>'
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
            try {
                Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
            } catch {
                throw
            }
        }

        $iconBase64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAMMOAADDDgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAA2pkAgNqZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAADamQCA2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgNqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkA/9qZAP/amQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAP/amQD/2pkAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQD/2pkA/9qZAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAANqZAIDamQCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA//5////8P///+B////AP///gB///wAP//4AB//8AAP/+AAB//AAAP/gAAB/wAAAP4AAAB8AAAAOAAAABAAAAAAAAAACAAAABwAAAA+AAAAfwAAAP+AAAH/wAAD/+AAB//wAA//+AAf//wAP//+AH///wD///+B////w////+f/8="
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $window.Icon = $iconBytes

        $assembly = [System.Reflection.Assembly]::GetAssembly([VideoOS.Platform.ConfigurationItems.Hardware])

        $excludedItems = "Folder|Path|Icon|Enabled|DisplayName|RecordingFramerate|ItemCategory|Wrapper|Address|Channel"

        $var_cboItemType.Add_SelectionChanged( {
                [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification='This is an event handler.')]
                param($sender, $e)
                $itemType = $e.AddedItems[0].Content

                $var_cboPropertyName.Items.Clear()
                $var_dgrResults.Columns.Clear()
                $var_dgrResults.Items.Clear()
                $var_txtTotalResults.Clear()
                $var_txtPropertyValue.Clear()
                $var_lblNoResults.Visibility = "Hidden"
                $var_lblPropertyNameBlank.Visibility = "Hidden"
                $var_lblPropertyValueBlank.Visibility = "Hidden"

                $properties = ($assembly.GetType("VideoOS.Platform.ConfigurationItems.$itemType").DeclaredProperties | Where-Object { $_.PropertyType.Name -eq 'String' }).Name + ([VideoOS.Platform.ConfigurationItems.IConfigurationChildItem].DeclaredProperties | Where-Object { $_.PropertyType.Name -eq 'String' }).Name | Where-Object { $_ -notmatch $excludedItems }
                foreach ($property in $properties) {
                    $newComboboxItem = [System.Windows.Controls.ComboBoxItem]::new()
                    $newComboboxItem.AddChild($property)
                    $var_cboPropertyName.Items.Add($newComboboxItem)
                }

                $sortDescription = [System.ComponentModel.SortDescription]::new("Content", "Ascending")
                $var_cboPropertyName.Items.SortDescriptions.Add($sortDescription)

                $var_cboEnabledFilter.IsEnabled = $true
                $var_lblEnabledFilter.IsEnabled = $true
                $var_cboPropertyName.IsEnabled = $true
                $var_lblPropertyName.IsEnabled = $true
                $var_txtPropertyValue.IsEnabled = $true
                $var_lblPropertyValue.IsEnabled = $true
                $var_txtName.IsEnabled = $true
                $var_lblName.IsEnabled = $true
                $var_btnSearch.IsEnabled = $true

                if ($itemType -eq "Hardware") {
                    $var_txtAddress.IsEnabled = $true
                    $var_lblAddress.IsEnabled = $true
                    $var_txtMACAddress.IsEnabled = $true
                    $var_lblMACAddress.IsEnabled = $true
                } else {
                    $var_txtAddress.IsEnabled = $false
                    $var_txtAddress.Clear()
                    $var_lblAddress.IsEnabled = $false
                    $var_txtMACAddress.IsEnabled = $false
                    $var_txtMACAddress.Clear()
                    $var_lblMACAddress.IsEnabled = $false
                }
            })

        $var_txtName.Add_TextChanged( {
                Clear-Results
            })

        $var_txtAddress.Add_TextChanged( {
                Clear-Results
            })

        $var_txtMACAddress.Add_TextChanged( {
                Clear-Results
            })

        $var_cboEnabledFilter.Add_SelectionChanged( {
                Clear-Results
            })

        $var_cboPropertyName.Add_SelectionChanged( {
                Clear-Results
            })

        $var_txtPropertyValue.Add_TextChanged( {
                Clear-Results
            })

        $var_btnSearch.Add_Click( {
                if (-not [string]::IsNullOrEmpty($var_cboPropertyName.Text) -and [string]::IsNullOrEmpty($var_txtPropertyValue.Text)) {
                    $var_lblPropertyValueBlank.Visibility = "Visible"
                    Return
                } elseif ([string]::IsNullOrEmpty($var_cboPropertyName.Text) -and -not [string]::IsNullOrEmpty($var_txtPropertyValue.Text)) {
                    $var_lblPropertyNameBlank.Visibility = "Visible"
                    Return
                }

                $script:searchResults = Find-XProtectDeviceSearch -ItemType $var_cboItemType.Text -Name $var_txtName.Text -Address $var_txtAddress.Text -MAC $var_txtMACAddress.Text -Enabled $var_cboEnabledFilter.Text -PropertyName $var_cboPropertyName.Text -PropertyValue $var_txtPropertyValue.Text
                if ($null -ne $script:searchResults) {
                    $var_btnExportCSV.IsEnabled = $true
                } else {
                    $var_btnExportCSV.IsEnabled = $false
                }
            })

        $var_btnExportCSV.Add_Click( {
                $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
                $saveDialog.Title = "Save As CSV"
                $saveDialog.Filter = "Comma delimited (*.csv)|*.csv"

                $saveAs = $saveDialog.ShowDialog()

                if ($saveAs -eq $true) {
                    $script:searchResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
                }
            })

        $var_btnResetForm.Add_Click( {
                $var_dgrResults.Columns.Clear()
                $var_dgrResults.Items.Clear()
                $var_cboItemType.SelectedItem = $null
                $var_cboEnabledFilter.IsEnabled = $false
                $var_lblEnabledFilter.IsEnabled = $false
                $var_cbiEnabledAll.IsSelected = $true
                $var_cboPropertyName.IsEnabled = $false
                $var_cboPropertyName.Items.Clear()
                $var_lblPropertyName.IsEnabled = $false
                $var_txtPropertyValue.IsEnabled = $false
                $var_txtPropertyValue.Clear()
                $var_lblPropertyValue.IsEnabled = $false
                $var_txtName.IsEnabled = $false
                $var_txtName.Clear()
                $var_lblName.IsEnabled = $false
                $var_btnSearch.IsEnabled = $false
                $var_btnExportCSV.IsEnabled = $false
                $var_txtAddress.IsEnabled = $false
                $var_txtAddress.Clear()
                $var_lblAddress.IsEnabled = $false
                $var_txtMACAddress.IsEnabled = $false
                $var_txtMACAddress.Clear()
                $var_lblMACAddress.IsEnabled = $false
                $var_txtTotalResults.Clear()
                $var_lblNoResults.Visibility = "Hidden"
                $var_lblPropertyNameBlank.Visibility = "Hidden"
                $var_lblPropertyValueBlank.Visibility = "Hidden"
            })

        $null = $window.ShowDialog()
    }
}

function Find-XProtectDeviceSearch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ItemType,
        [Parameter(Mandatory = $false)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$Address,
        [Parameter(Mandatory = $false)]
        [string]$MAC,
        [Parameter(Mandatory = $false)]
        [string]$Enabled,
        [Parameter(Mandatory = $false)]
        [string]$PropertyName,
        [Parameter(Mandatory = $false)]
        [string]$PropertyValue
    )

    process {
        $var_dgrResults.Columns.Clear()
        $var_dgrResults.Items.Clear()
        $var_lblNoResults.Visibility = "Hidden"
        $var_lblPropertyNameBlank.Visibility = "Hidden"
        $var_lblPropertyValueBlank.Visibility = "Hidden"

        if ([string]::IsNullOrEmpty($PropertyName) -or [string]::IsNullOrEmpty($PropertyValue)) {
            $PropertyName = "Id"
            $PropertyValue = $null
        }

        if ($ItemType -eq "Hardware" -and $null -eq [string]::IsNullOrEmpty($MAC)) {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -MacAddress $MAC -EnableFilter $Enabled -Properties @{Name = $Name; Address = $Address; $PropertyName = $PropertyValue })
        } elseif ($ItemType -eq "Hardware" -and $null -ne [string]::IsNullOrEmpty($MAC)) {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -EnableFilter $Enabled -Properties @{Name = $Name; Address = $Address; $PropertyName = $PropertyValue })
        } else {
            $results = [array](Find-XProtectDevice -ItemType $ItemType -EnableFilter $Enabled -Properties @{Name = $Name; $PropertyName = $PropertyValue })
        }

        if ($null -ne $results) {
            $columnNames = $results[0].PsObject.Properties | ForEach-Object { $_.Name }
        } else {
            $var_lblNoResults.Visibility = "Visible"
        }

        foreach ($columnName in $columnNames) {
            $newColumn = [System.Windows.Controls.DataGridTextColumn]::new()
            $newColumn.Header = $columnName
            $newColumn.Binding = New-Object System.Windows.Data.Binding($columnName)
            $newColumn.Width = "SizeToCells"
            $var_dgrResults.Columns.Add($newColumn)
        }

        if ($ItemType -eq "Hardware") {
            foreach ($result in $results) {
                $var_dgrResults.AddChild([pscustomobject]@{Hardware = $result.Hardware; RecordingServer = $result.RecordingServer })
            }
        } else {
            foreach ($result in $results) {
                $var_dgrResults.AddChild([pscustomobject]@{$columnNames[0] = $result.((Get-Variable -Name columnNames).Value[0]); Hardware = $result.Hardware; RecordingServer = $result.RecordingServer })
            }
        }

        $var_txtTotalResults.Text = $results.count
    }
    end {
        return $results
    }
}


function Get-DevicesByRecorder {
    <#
    .SYNOPSIS
        Gets all enabled cameras in a hashtable indexed by recording server id.
    .DESCRIPTION
        This cmdlet quickly returns a hashtable where the keys are recording
        server ID's and the values are lists of "VideoOS.Platform.Item" objects.

        The cmdlet will complete much quicker than if we were to use
        Get-VmsRecordingServer | Get-VmsCamera, because it does not rely on the
        configuration API at all. Instead, it has the same functionality as
        XProtect Smart Client where the command "sees" only the devices that are enabled
        and loaded by the Recording Server.
    .EXAMPLE
        Get-CamerasByRecorder
        Name                           Value
        ----                           -----
        bb82b2cd-0bb9-4c88-9cb8-128... {Canon VB-M40 (192.168.101.64) - Camera 1}
        f9dc2bcd-faea-4138-bf5a-32c... {Axis P1375 (10.1.77.178) - Camera 1, Test Cam}

        This is what the output would look like on a small system.
    .OUTPUTS
        [hashtable]
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        [Parameter()]
        [Alias('Kind')]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', IgnoreCase = $false)]
        [string[]]
        $DeviceType = 'Camera'
    )

    process {
        $config = [videoos.platform.configuration]::Instance
        $serverKind = [VideoOS.Platform.Kind]::Server
        $selectedKinds = @(($DeviceType | ForEach-Object { [VideoOS.Platform.Kind]::$_ }))
        $systemHierarchy = [VideoOS.Platform.ItemHierarchy]::SystemDefined

        $stack = [Collections.Generic.Stack[VideoOS.Platform.Item]]::new()
        $rootItems = $config.GetItems($systemHierarchy)
        foreach ($mgmtSrv in $rootItems | Where-Object { $_.FQID.Kind -eq $serverKind }) {
            foreach ($recorder in $mgmtSrv.GetChildren()) {
                if ($recorder.FQID.Kind -eq $serverKind -and ($RecordingServerId.Count -eq 0 -or $recorder.FQID.ObjectId -in $RecordingServerId)) {
                    $stack.Push($recorder)
                }
            }
        }

        $result = @{}
        $lastServerId = $null
        while ($stack.Count -gt 0) {
            $item = $stack.Pop()
            if ($item.FQID.Kind -eq $serverKind) {
                $lastServerId = $item.FQID.ObjectId
                $result.$lastServerId = [Collections.Generic.List[VideoOS.Platform.Item]]::new()
            } elseif ($item.FQID.Kind -in $selectedKinds -and $item.FQID.FolderType -eq 'No') {
                $result.$lastServerId.Add($item)
                continue
            }

            if ($item.HasChildren -ne 'No' -and ($item.FQID.Kind -eq $serverKind -or $item.FQID.Kind -in $selectedKinds)) {
                foreach ($child in $item.GetChildren()) {
                    if ($child.FQID.Kind -in $selectedKinds) {
                        $stack.Push($child)
                    }
                }
            }
        }
        Write-Output $result
    }
}


function Get-HttpSslCertThumbprint {
    <#
    .SYNOPSIS
        Gets the certificate thumbprint from the sslcert binding information put by netsh http show sslcert ipport=$IPPort
    .DESCRIPTION
        Gets the certificate thumbprint from the sslcert binding information put by netsh http show sslcert ipport=$IPPort.
        Returns $null if no binding is present for the given ip:port value.
    .PARAMETER IPPort
        The ip:port string representing the binding to retrieve the thumbprint from.
    .EXAMPLE
        Get-HttpSslCertThumbprint 0.0.0.0:8082
        Gets the sslcert thumbprint for the binding found matching 0.0.0.0:8082 which is the default HTTPS IP and Port for
        XProtect Mobile Server. The value '0.0.0.0' represents 'all interfaces' and 8082 is the default https port.
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]
        $IPPort
    )
    process {
        $netshOutput = [string](netsh.exe http show sslcert ipport=$IPPort)

        if (!$netshOutput.Contains('Certificate Hash')) {
            Write-Error "No SSL certificate binding found for $ipPort"
            return
        }

        if ($netshOutput -match "Certificate Hash\s+:\s+(\w+)\s+") {
            $Matches[1]
        } else {
            Write-Error "Certificate Hash not found for $ipPort"
        }
    }
}

function Get-ProcessOutput
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $FilePath,
        [Parameter()]
        [string[]]
        $ArgumentList
    )
    
    process {
        try {
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo.UseShellExecute = $false
            $process.StartInfo.RedirectStandardOutput = $true
            $process.StartInfo.RedirectStandardError = $true
            $process.StartInfo.FileName = $FilePath
            $process.StartInfo.CreateNoWindow = $true

            if($ArgumentList) { $process.StartInfo.Arguments = $ArgumentList }
            Write-Verbose "Executing $($FilePath) with the following arguments: $([string]::Join(' ', $ArgumentList))"
            $null = $process.Start()
    
            [pscustomobject]@{
                StandardOutput = $process.StandardOutput.ReadToEnd()
                StandardError = $process.StandardError.ReadToEnd()
                ExitCode = $process.ExitCode
            }
        }
        finally {
            $process.Dispose()
        }
        
    }
}

function Get-SecurityNamespaceValues {
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Private function.')]
    param ()
    
    process {
        if (-not [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache.ContainsKey('SecurityNamespaceValues')) {
            [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespacesById'] = [Collections.Generic.Dictionary[[string], [string]]]::new()
            
            if (($r = (Get-VmsManagementServer).RoleFolder.Roles | Where-Object RoleType -EQ 'UserDefined' | Select-Object -First 1)) {
                $task = $r.ChangeOverallSecurityPermissions()
                [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespaceValues'] = $task.SecurityNamespaceValues
                $task.SecurityNamespaceValues.GetEnumerator() | ForEach-Object {
                    [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespacesById'][$_.Value] = $_.Key
                }
            } else {
                [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespaceValues'] = [Collections.Generic.Dictionary[[string], [string]]]::new()
            }
        }
        [pscustomobject]@{
            SecurityNamespacesByName = [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespaceValues']
            SecurityNamespacesById   = [MilestonePSTools.Connection.MilestoneConnection]::Instance.Cache['SecurityNamespacesById']
        }
    }
}

function GetCodecValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("Codec")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("Codec") | Where-Object Value -eq $res).Name
        return
    }
}

function GetFpsValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("FPS")
    if ($null -ne $res) {
        $val = ($Stream.Properties.GetValueTypeInfoCollection("FPS") | Where-Object Value -eq $res).Name
        if ($null -eq $val) {
            $res
        }
        else {
            $val
        }
        return
    }

    $res = $Stream.Properties.GetValue("Framerate")
    if ($null -ne $res) {
        $val = ($Stream.Properties.GetValueTypeInfoCollection("Framerate") | Where-Object Value -eq $res).Name
        if ($null -eq $val) {
            $res
        }
        else {
            $val
        }
        return
    }
}

function GetResolutionValueFromStream {
    param([VideoOS.Platform.ConfigurationItems.StreamChildItem]$Stream)

    $res = $Stream.Properties.GetValue("StreamProperty")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("StreamProperty") | Where-Object Value -eq $res).Name
        return
    }

    $res = $Stream.Properties.GetValue("Resolution")
    if ($null -ne $res) {
        ($Stream.Properties.GetValueTypeInfoCollection("Resolution") | Where-Object Value -eq $res).Name
        return
    }
}

function GetVmsConnectionProfile {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name = 'default',

        [Parameter(ParameterSetName = 'All')]
        [switch]
        $All
    )

    begin {
        if (-not (Test-Path -Path (GetVmsConnectionProfilePath))) {
            @{} | Export-Clixml -Path (GetVmsConnectionProfilePath)
        }
        $vmsProfiles = (Import-Clixml -Path (GetVmsConnectionProfilePath)) -as [hashtable]
    }

    process {
        if ($All) {
            $vmsProfiles
        } elseif ($vmsProfiles.ContainsKey($Name)) {
            $vmsProfiles[$Name]
        }
    }
}

function GetVmsConnectionProfilePath {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    process {
        Join-Path -Path (NewVmsAppDataPath) -ChildPath 'credentials.xml'
    }
}

function HandleValidateResultException {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [System.Management.Automation.ErrorRecord]
        $ErrorRecord,

        [Parameter()]
        [object]
        $TargetObject,

        [Parameter()]
        [string]
        $ItemName = 'Not set'
    )

    process {
        # This function makes it easier to understand the reason for the validation error, but it hides the original
        # ScriptStackTrace so the error looks like it originates in the original "catch" block instead of the line
        # inside the "try" block triggering the exception. Writing the original stacktrace out to the debug stream
        # is a compromise for being unable to throw the original exception and stacktrace with a better exception
        # message.
        Write-Debug -Message "Original ScriptStackTrace:`n$($ErrorRecord.ScriptStackTrace)"

        $lastCommand = (Get-PSCallStack)[1]
        $origin = $lastCommand.Command
        $exception = $ErrorRecord.Exception
        $validateResult = $exception.ValidateResult
        if (-not $MyInvocation.BoundParameters.ContainsKey('ItemName') -and -not [string]::IsNullOrWhiteSpace($validateResult.ResultItem.DisplayName)) {
            $ItemName = $validateResult.ResultItem.DisplayName
        }
        foreach ($errorResult in $validateResult.ErrorResults) {
            $errorParams = @{
                Message           = '{0}: Invalid value for property "{1}" on {2}. ErrorText = "{3}". Origin = {4}' -f $errorResult.ErrorTextId, $errorResult.ErrorProperty, $ItemName, $errorResult.ErrorText, $origin
                Exception         = $Exception
                Category          = 'InvalidData'
                RecommendedAction = 'Review the invalid property value and try again.'
            }
            if ($TargetObject) {
                $errorParams.TargetObject = $TargetObject
            }
            Write-Error @errorParams
        }
    }
}


function ImportHardwareCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]
        $Path,
        
        [Parameter()]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter()]
        [pscredential[]]
        $Credential,

        [Parameter()]
        [switch]
        $UpdateExisting,

        [Parameter()]
        [char]
        $Delimiter = ','
    )

    process {
        $progress = @{
            Id               = 1
            Activity         = 'Import hardware from CSV'
            Status           = 'Loading CSV file'
            PercentComplete  = 0
            CurrentOperation = ''
        }
        Write-Progress @progress
        # Read CSV file, perform basic validation, and normalize records
        $rows = Import-Csv -LiteralPath $Path -Delimiter $Delimiter
        if ($RecordingServer) {
            $rows | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.RecordingServer)) {
                    $_.RecordingServer = $RecordingServer.Name
                }
            }
        }
        $records = [pscustomobject[]](ValidateHardwareCsvRows -Rows $rows)
        $recordsProcessed = 0

        # Set RecordingServer property on all records to match RecordingServer parameter if provided.
        # Warn user that the RecordingServer from the CSV, if present, will be ignored.
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecordingServer')) {
            for ($i = 0; $i -lt $records.Count; $i++) {
                if (-not [string]::IsNullOrWhiteSpace($records[$i].RecordingServer) -and $records[$i].RecordingServer -ne $RecordingServer.Name) {
                    Write-Warning "Ignoring RecordingServer value in row $($i + 1) in favor of `"$($RecordingServer.Name)`""
                }
                $records[$i].RecordingServer = $RecordingServer.Name
            }
        }

        # Check if there are any duplicate device entries
        $duplicateDevices = $records | Group-Object { "$($_.RecordingServer).$($_.Address).$($_.DeviceType).$($_.Channel)" } | Where-Object Count -GT 1
        if ($duplicateDevices) {
            Write-Error -Message 'Duplicate device records found. Please ensure there are no rows with identical values for RecordingServer, Address, DeviceType, and Channel.' -TargetObject $duplicateDevices -Category InvalidData -ErrorId 'DuplicateDeviceRecord'
            return
        }

        $recordsByRecorder = $records | Group-Object RecordingServer
        $recorders = @{}
        Get-VmsRecordingServer | Where-Object Name -In $recordsByRecorder.Name | ForEach-Object {
            $recorders[$_.Name] = $_
        }
        try {
            foreach ($recorderGroup in $recordsByRecorder) {
                $progress.Status = "Processing $($recorderGroup.Count) records for recording server $($recorderGroup.Name)"
                Write-Progress @progress
                # Abort if no Recording Server was specified in CSV file or in RecordingServer argument.
                if ([string]::IsNullOrWhiteSpace($recorderGroup.Name)) {
                    $recorderGroup.Group | ForEach-Object {
                        $_.Result += 'RecordingServer not specified.'
                    }
                    Write-Error -Message "RecordingServer not specified. Specify the destination recording server using the RecordingServer parameter, or add a RecordingServer column to your CSV file with the display name of the destination recording server. This affects $($recorderGroup.Count) rows in the file `"$Path`"."
                    continue
                }
    
                # Abort if the specified recording server is not found
                $recorder = $recorders[$recorderGroup.Name]
                if ($null -eq $recorder) {
                    $recorderGroup.Group | ForEach-Object {
                        $_.Result += 'RecordingServer not found.'
                    }
                    Write-Error -Message "RecordingServer with display name `"$($recorderGroup.Name)`" not found. This affects $($recorderGroup.Count) rows in the file `"$Path`"."
                    continue
                }
    
                # Check for unrecognized StorageName values
                $existingStorage = @{}
                $recorder | Get-VmsStorage | ForEach-Object {
                    $existingStorage[$_.Name] = $_
                }
                foreach ($storageGroup in $recorderGroup.Group | Group-Object StorageName | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) }) {
                    if (!$existingStorage.ContainsKey($storageGroup.Name)) {
                        Write-Error -Message "Storage with display name `"$($storageGroup.Name)`" not found. This affects $($storageGroup.Count) rows in the file `"$Path`"."
                    }
                }
    
                $progress.CurrentOperation = "Getting existing hardware and available drivers"
                Write-Progress @progress
    
                # Cache hardware already present
                $recorder.HardwareFolder.ClearChildrenCache()
                $existingHardware = @{}
                $recorder | Get-VmsHardware | ForEach-Object {
                    $existingHardware[([uribuilder]$_.Address).Uri.GetComponents([uricomponents]::SchemeAndServer, [uriformat]::SafeUnescaped)] = $_
                }
    
                $availableDrivers = $recorder | Get-VmsHardwareDriver -ErrorAction Stop
                $progress.CurrentOperation = ''
                Write-Progress @progress

                # Only perform express scan if at least one row for this recording server lacks a DriverNumber value
                $executeExpressScan = $null -ne ($recorderGroup.Group | Where-Object { $_.DriverNumber -eq 0 -and !$existingHardware.Contains($_.Address)})
                $expressScanResults = @{}
                if ($executeExpressScan) {
                    $progress.CurrentOperation = "Running Start-VmsHardwareScan -Express"
                    Write-Progress @progress
                    $expressScanSplat = @{
                        Express               = $true
                        UseDefaultCredentials = $true
                    }
    
                    # Build a group of credentials from the Credential parameter if provided, and a collection of unique
                    # credentials provided for devices on this recorder in the CSV file.
                    $expressCredentials = [collections.generic.list[pscredential]]::new()
                    $Credential | ForEach-Object {
                        if ($null -ne $_) {
                            $expressCredentials.Add($_)
                        }
                    }
                    $recorderGroup.Group | Where-Object {
                        ![string]::IsNullOrWhiteSpace($_.UserName) -and ![string]::IsNullOrWhiteSpace($_.Password)
                    } | Group-Object { 
                        '{0}:{1}' -f $_.UserName, $_.Password 
                    } | ForEach-Object {
                        $expressCredentials.Add([pscredential]::new($_.Group[0].UserName, ($_.Group[0].Password | ConvertTo-SecureString -AsPlainText -Force)))
                    }
                    if ($expressCredentials.Count -gt 0) {
                        $expressScanSplat.Credential = $expressCredentials
                    }
                    $recorder | Start-VmsHardwareScan @expressScanSplat -Verbose:$false -ErrorAction SilentlyContinue | Where-Object HardwareScanValidated | ForEach-Object {
                        $uri = ([uribuilder]$_.HardwareAddress).Uri.GetComponents([uricomponents]::SchemeAndServer, [uriformat]::SafeUnescaped)
                        $expressScanResults[$uri] = $_
                    }
                }
    
                $recordsByHardware = $recorderGroup.Group | Group-Object Address
                foreach ($hardwareGroup in $recordsByHardware) {
                    $recordsProcessed += $hardwareGroup.Count
                    $progress.PercentComplete = $recordsProcessed / $records.Count * 100
                    Write-Progress @progress

                    # If hardware already exists, update DriverNumber for related CSV records
                    if ($existingHardware.ContainsKey($hardwareGroup.Name)) {
                        $currentDriver = $existingHardware[$hardwareGroup.Name] | Get-VmsHardwareDriver
                        if ($currentDriver.Number -ne $hardwareGroup.Group[0].DriverNumber) {
                            $hardwareGroup.Group | ForEach-Object {
                                $_.DriverNumber = $currentDriver.Number
                            }
                        }
                    }
                    $driver = ($availableDrivers | Where-Object Number -EQ $hardwareGroup.Group[0].DriverNumber).Path
                    if ($null -eq $driver) {
                        # Discover driver via hardware scans - first check express scan results, then do targetted scan.
                        if ($expressScanResults[$hardwareGroup.Name]) {
                            $driver = $expressScanResults[$hardwareGroup.Name].HardwareDriverPath
                            $hardwareGroup.Group[0].UserName = $expressScanResults[$hardwareGroup.Name].UserName
                            $hardwareGroup.Group[0].Password = $expressScanResults[$hardwareGroup.Name].Password
                            $hardwareGroup.Group | ForEach-Object {
                                $_.DriverNumber = ($availableDrivers | Where-Object Path -eq $driver).Number
                            }
                            Write-Verbose "Adding $($hardwareGroup.Name) to $($recorder.Name) using DriverNumber $(($availableDrivers | Where-Object Path -EQ $driver).Number) discovered during express scan."
                        } else {
                            $progress.CurrentOperation = "Trying to determine the correct driver for $($hardwareGroup.Name)"
                            Write-Progress @progress
    
                            # Hardware not found in express scan. Perform targetted scan on hardware address
                            $scanSplat = @{
                                Address               = $hardwareGroup.Name
                                UseDefaultCredentials = $true
                                Credential            = [collections.generic.list[pscredential]]::new()
                            }
                            if (-not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].DriverGroup)) {
                                $scanSplat.DriverGroup = $hardwareGroup.Group[0].DriverGroup -split ';' | Where-Object {
                                    -not [string]::IsNullOrWhiteSpace($_)
                                } | ForEach-Object { $_.Trim() }
                            }
    
                            # Build credential set for hardware scan using credentials from row if available along with
                            # credentials provided using the Credential parameter.
                            if (![string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].UserName) -and ![string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].Password)) {
                                $scanSplat.Credential.Add([pscredential]::new($hardwareGroup.Group[0].UserName, ($hardwareGroup.Group[0].Password | ConvertTo-SecureString -AsPlainText -Force)))
                            }
                            $Credential | ForEach-Object {
                                if ($null -ne $_) {
                                    $scanSplat.Credential.Add($_)
                                }
                            }
                            
                            # Perform targetted hardware scan. Multiple scans may be performed depending on the number of credentials provided
                            # so return the first validated scan.
                            $hardwareScan = $recorder | Start-VmsHardwareScan @scanSplat -Verbose:$false | Where-Object HardwareScanValidated | Select-Object -First 1
                            if ($hardwareScan.HardwareScanValidated) {
                                $driver = $hardwareScan.HardwareDriverPath
                                $hardwareGroup.Group[0].UserName = $hardwareScan.UserName
                                $hardwareGroup.Group[0].Password = $hardwareScan.Password
                            } else {
                                Write-Error "Hardware scan was unsuccessful for $($hardwareGroup.Name) on RecordingServer $($recorder.Name). Check the provided credentials, and driver, and try again." -Category InvalidResult -ErrorId 'AddHardwareFailed' -TargetObject $hardwareGroup.Group
                                $hardwareGroup.Group | ForEach-Object {
                                    $_.Result += "Failed to detect the correct driver for the hardware based on the provided credential(s), and DriverGroup. Note that a small number of drivers do not support hardware scanning and the exact driver is required."
                                }
                                continue
                            }
                        }
                    }
                    if ($null -eq $driver) {
                        $hardwareGroup.Group | ForEach-Object {
                            $_.Result += 'DriverNumber not found on RecordingServer.'
                        }
                        Write-Error -Message "No hardware driver found for device at $($hardwareGroup.Name) with DriverNumber $($hardwareGroup.Group[0].DriverNumber) on RecordingServer $($recorder.Name)." -TargetObject $hardwareGroup.Group
                        continue
                    }
                    $recordsByDeviceType = $hardwareGroup.Group | Group-Object DeviceType
                    $skipHardware = $false
                    foreach ($deviceTypeGroup in $recordsByDeviceType) {
                        $invalidDeviceTypeGroup = $deviceTypeGroup.Group | Group-Object Channel | Where-Object Count -GT 1
                        if ($invalidDeviceTypeGroup) {
                            $skipHardware = $true
                            $hardwareGroup.Group | ForEach-Object {
                                $_.Result += 'One or more devices with this Address have the same DeviceType and Channel.'
                            }
                            Write-Error -Message "Multiple $($deviceTypeGroup.Name) records found for $($hardwareGroup.Name) with the same channel number. Please add, or correct the Channel field in your CSV file." -TargetObject $invalidDeviceTypeGroup.Group
                        }
                    }
                    if ($skipHardware) {
                        continue
                    }
    
                    try {
                        $hwSplat = @{
                            Enabled  = $true
                            PassThru = $true
                        }
    
                        if (-not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].HardwareName)) {
                            $hwSplat.Name = $hardwareGroup.Group[0].HardwareName
                        }
                        if (-not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].UserName)) {
                            $hwSplat.UserName = $hardwareGroup.Group[0].UserName
                        }
                        if (-not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].Password)) {
                            $hwSplat.Password = $hardwareGroup.Group[0].Password | ConvertTo-SecureString -AsPlainText -Force
                        }
    
                        if ($existingHardware.ContainsKey($hardwareGroup.Name)) {
                            if (-not $UpdateExisting) {
                                Write-Verbose "Skipping row(s) $($hardwareGroup.Group.Row -join ', ') because the hardware is already added and the UpdateExisting parameter was omitted."
                                $hardwareGroup.Group | ForEach-Object {
                                    $_.Result += 'Skipped because the hardware already exists and -UpdateExisting was not used.'
                                    $_.Path = ($existingHardware[$hardwareGroup.Name] | Get-VmsDevice -Type $_.DeviceType -Channel $_.Channel).Path
                                }
                                continue
                            }
                            Write-Verbose "Updating existing device(s) for hardware at $($hardwareGroup.Name) defined in row(s) $($hardwareGroup.Group.Row -join ', ')"
                            $hardwareGroup.Group | ForEach-Object { $_.Result += 'Updating existing hardware.' }
                            $hardware = $existingHardware[$hardwareGroup.Name]
                        } else {
                            $skipHardware = $true
                            $credentials = [collections.generic.list[pscredential[]]]::new()
                            if (-not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].UserName) -and -not [string]::IsNullOrWhiteSpace($hardwareGroup.Group[0].Password)) {
                                $credentials.Add([pscredential]::new($hardwareGroup.Group[0].UserName, ($hardwareGroup.Group[0].Password | ConvertTo-SecureString -AsPlainText -Force)))
                            }
                            foreach ($c in $Credential) {
                                $credentials.Add($c)
                            }
                            if ($credentials.Count -eq 0) {
                                $hardwareGroup.Group | ForEach-Object {
                                    $_.Result += "Hardware not added - no credentials provided."
                                }
                                Write-Warning "Skipping $($hardware.Name) as no credentials have been provided."
                            }
                            for ($credIndex = 0; $credIndex -lt $credentials.Count; $credIndex++) {
                                $cred = $credentials[$credIndex]
                                try {
                                    $progress.CurrentOperation = "Adding $($hardwareGroup.Name)"
                                    Write-Progress @progress
                                    $task = $recorder.AddHardware($hardwareGroup.Name, $driver, $cred.UserName, $cred.Password, $null) | Wait-VmsTask -Cleanup
                                    if (($task.Properties | Where-Object Key -EQ 'State').Value -eq 'Error') {
                                        if ($credIndex -ge $credentials.Count - 1) {
                                            $hardwareGroup.Group | ForEach-Object { $_.Result += 'Failed to add hardware.' }
                                            Write-Error -Message "Failed to add $($hardwareGroup.Name) in row(s) $($hardwareGroup.Group.Row -join ', ') to RecordingServer $($recorder.Name): $(($task.Properties | Where-Object Key -EQ 'ErrorText').Value)" -Category InvalidResult -ErrorId 'AddHardwareFailure' -TargetObject $hardwareGroup.Group
                                            break
                                        } else {
                                            Write-Warning "Failed to add $($hardwareGroup.Name) in row(s) $($hardwareGroup.Group.Row -join ', ') to RecordingServer $($recorder.Name). Retrying with another credential..."
                                        }
                                        continue
                                    } else {
                                        
                                        $skipHardware = $false
                                        break
                                    }
                                } catch {
                                    throw
                                }
                            }
                            if ($skipHardware) {
                                $hardwareGroup.Group | ForEach-Object {
                                    $_.Result += 'Hardware successfully added.'
                                }
                                continue
                            } else {
                                $hardwareGroup.Group | ForEach-Object {
                                    $_.Result += 'Hardware successfully added.'
                                }
                            }
    
                            $hardware = [VideoOS.Platform.ConfigurationItems.Hardware]::new($recorder.ServerId, ($task.Properties | Where-Object Key -EQ 'Path').Value)
                            'UserName', 'Password' | ForEach-Object {
                                if ($hwSplat.ContainsKey($_)) { $hwSplat.Remove($_) }
                            }
                        }
                        $hardware = $hardware | Set-VmsHardware @hwSplat
                        
                        $progress.CurrentOperation = "Updating settings for $($hardwareGroup.Count) devices on $($hardwareGroup.Name)"
                        Write-Progress @progress
                        foreach ($deviceRecord in $hardwareGroup.Group) {
                            $splat = @{
                                Enabled = $deviceRecord.Enabled
                            }
                            if (-not [string]::IsNullOrWhiteSpace($deviceRecord.Name)) {
                                $splat.Name = $deviceRecord.Name
                            }
                            if (-not [string]::IsNullOrWhiteSpace($deviceRecord.Coordinates)) {
                                $splat.Coordinates = $deviceRecord.Coordinates
                            }
                            
                            $device = $hardware | Get-VmsDevice -Type $deviceRecord.DeviceType -Channel $deviceRecord.Channel -EnableFilter All
                            if ($null -eq $device) {
                                Write-Error "$($deviceRecord.DeviceType) channel $($deviceRecord.Channel) not found on hardware with address $($hardwareGroup.Name) on RecordingServer $($recorder.Name) defined in row $($deviceRecord.Row)."
                                $deviceRecord.Result += 'Channel does not exist on hardware.'
                                continue
                            }
                            $deviceRecord.Path = $device.Path
                            $device | Set-VmsDevice @splat
                            $deviceRecord.DeviceGroups -split ';' | ForEach-Object {
                                if ([string]::IsNullOrWhiteSpace($_)) { return }
                                $deviceGroup = New-VmsDeviceGroup -Type $deviceRecord.DeviceType -Path $_.Trim()
                                $deviceGroup | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                            }
    
                            if ($device.RecordingStorage -and -not [string]::IsNullOrWhiteSpace($deviceRecord.StorageName)) {
                                if ($existingStorage.ContainsKey($deviceRecord.StorageName)) {
                                    if ($device.RecordingStorage -ne $existingStorage[$deviceRecord.StorageName].Path) {
                                        $tries = 0
                                        $maxTries = 5
                                        $delay = [timespan]::FromSeconds(10)
                                        do {
                                            try {
                                                $device | Set-VmsDeviceStorage -Destination $deviceRecord.StorageName -ErrorAction Stop
                                                break
                                            } catch {
                                                $tries += 1
                                                if ($tries -ge $maxTries) {
                                                    $deviceRecord.Result += 'Failed to assign device to the specified storage.'
                                                    Write-Error -Message "Failed to assign $($deviceRecord.DeviceType) `"$($deviceRecord.Name)`" with address $($deviceRecord.Address) to storage `"$($deviceRecord.StorageName)`". $($_.Exception.Message)" -Exception $_.Exception -Category InvalidResult -ErrorId 'StorageAssignmentFailed' -TargetObject $deviceRecord
                                                } else {
                                                    Write-Warning "Failed to assign $($deviceRecord.DeviceType) `"$($deviceRecord.Name)`" with address $($deviceRecord.Address) to storage `"$($deviceRecord.StorageName)`". Attempt $tries of $maxTries. Retrying in $($delay.Seconds) seconds. $($_.Exception.Message)"
                                                    Start-Sleep -Seconds $delay.Seconds
                                                }
                                            }
                                        } while ($tries -lt $maxTries)
                                    }
                                } else {
                                    $storageGroup.Group | ForEach-Object {
                                        $_.Result += 'StorageName not found.'
                                    }
                                    Write-Warning "Cannot update the storage configuration for $($deviceRecord.Name) at $($hardware.Address) because StorageName $($deviceRecord.StorageName) does not exist on RecordingServer $($recorder.Name)."
                                }
                            }
                        }
                    } catch {
                        throw $_
                    }
                }
                $recorder.HardwareFolder.ClearChildrenCache()
            }
        } finally {
            $progress.Completed = $true
            Write-Progress @progress
        }
        $records
    }
}

$script:TruthyFalsey = [regex]::new('^\s*(true|yes|yep|affirmative|1|false|no|nope|negative|0)\s*$', [RegexOptions]::IgnoreCase)
$script:Truthy = [regex]::new('^\s*(true|yes|yep|affirmative|1)\s*$', [RegexOptions]::IgnoreCase)

function Show-FileDialog {
    [CmdletBinding(DefaultParameterSetName = 'OpenFile')]
    param (
        [Parameter(ParameterSetName = 'OpenFile')]
        [switch]
        $OpenFile,

        [Parameter(Mandatory, ParameterSetName = 'SaveFile')]
        [switch]
        $SaveFile
    )

    process {
        $params = @{
            Title            = 'ImportVmsHardwareExcel'
            Filter           = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
            DefaultExt       = '.xlsx'
            RestoreDirectory = $true
            AddExtension     = $true
        }
        switch ($PSCmdlet.ParameterSetName) {
            'OpenFile' {
                $dialog = [OpenFileDialog]$params
            }
            'SaveFile' {
                $params.FileName = 'Hardware_{0}.xlsx' -f (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')
                $dialog = [SaveFileDialog]$params
            }
            Default {
                throw "ParameterSetName '$_' not implemented."
            }
        }

        try {
            $form = [form]@{
                TopMost = $true
            }
            if ($dialog.ShowDialog($form) -eq 'OK') {
                $dialog.FileName
            } else {
                throw "$($PSCmdlet.ParameterSetName) aborted."
            }
        } finally {
            if ($dialog) {
                $dialog.Dispose()
            }
            if ($form) {
                $form.Dispose()
            }
        }
    }
}

function Resolve-Path2 {
    <#
    .SYNOPSIS
    Resolves paths like the PowerShell-native `Resolve-Path` cmdlet, even for
    paths that don't exist yet.

    .NOTES
    Inspired by a [blog post](http://devhawk.net/blog/2010/1/22/fixing-powershells-busted-resolve-path-cmdlet)
    by DevHawk, aka Harry Pierson, linked to by joshuapoehls on [stackoverflow.com](https://stackoverflow.com/a/12605755/3736007).
    #>
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Path')]
        [SupportsWildcards()]
        [string[]]
        $Path,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'LiteralPath')]
        [string[]]
        $LiteralPath,

        [Parameter()]
        [switch]
        $Relative,

        [Parameter()]
        [switch]
        $NoValidation,

        [Parameter(ParameterSetName = 'Path')]
        [switch]
        $ExpandEnvironmentVariables
    )

    process {
        foreach ($unresolvedPath in $MyInvocation.BoundParameters[$PSCmdlet.ParameterSetName]) {
            if ($ExpandEnvironmentVariables) {
                $unresolvedPath = [environment]::ExpandEnvironmentVariables($unresolvedPath)
            }
            $params = @{
                $($PSCmdlet.ParameterSetName) = $unresolvedPath
                ErrorAction                   = 'SilentlyContinue'
                ErrorVariable                 = 'resolvePathError'
            }
            $resolvedPath = Resolve-Path @params
            if ($null -eq $resolvedPath) {
                if ($NoValidation) {
                    $resolvedPath = $resolvePathError[0].TargetObject
                } elseif ($resolvePathError) {
                    Write-Error -ErrorRecord $resolvePathError[0]
                    Remove-Variable -Name resolvePathError
                    continue
                }
            }

            foreach ($pathInfo in $resolvedPath) {
                if ($Relative) {
                    $separator = [io.path]::DirectorySeparatorChar
                    $currentPathUri = [uri]::new($pwd.Path, [urikind]::Absolute)
                    $resolvedPathUri = [uri]::new(($pathInfo.Path -replace "([^$([regex]::Escape($separator))])`$", "`$1$([regex]::Escape($separator))"), [UriKind]::Absolute)
                    $relativePath = $currentPathUri.MakeRelativeUri($resolvedPathUri).ToString() -replace '/', [io.path]::DirectorySeparatorChar
                    if ($relativePath -notmatch "^\.+\$([io.path]::DirectorySeparatorChar)") {
                        $relativePath = '.{0}{1}' -f [io.path]::DirectorySeparatorChar, $relativePath
                    }
                    $relativePath
                } else {
                    $pathInfo
                }
            }
        }
    }
}

function Export-DeviceEventConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [MilestonePSTools.ValidateVmsItemType('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device
    )

    process {
        # Using Get-ConfigurationItem just so that we have the display names since
        # they aren't available on the strongly typed HardwareDeviceEventChildItems.
        $eventDisplayNames = @{}
        (Get-ConfigurationItem -Path "HardwareDeviceEvent[$($Device.Id)]").Children | ForEach-Object {
            $id = ($_.Properties | Where-Object Key -eq 'Id').Value
            $displayName = ($_.Properties | Where-Object Key -eq 'EventIndex').DisplayName
            $eventDisplayNames[$id] = $displayName
        }
        foreach ($deviceEvent in $Device | Get-VmsDeviceEvent) {
            [pscustomobject]@{
                Event      = $deviceEvent.DisplayName
                Used       = $deviceEvent.EventUsed
                Enabled    = $deviceEvent.Enabled
                EventIndex = $deviceEvent.EventIndex
                IndexName  = $eventDisplayNames[$deviceEvent.Id]
            }
        }
    }
}

function Get-DevicePropertyList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device
    )

    begin {
        $excludedProperties = 'Icon', 'ItemCategory', 'Methods', 'ServerId', 'CreatedDate', 'DisplayName', 'ParentItemPath', 'StreamDefinitions', 'StreamUsages', 'Guid'
        $orderPriority = 'Name', 'ShortName', 'HostName', 'WebServerUri', 'Address', 'UserName', 'Password', 'Enabled', 'Channel', 'GisPoint', 'ActiveWebServerUri', 'PublicAccessEnabled', 'PublicWebserverHostName', 'PublicWebserverPort'
        $rearOrderPriority = 'LastModified', 'Id'

        $pathNameMap = @{}
        $childToParentMap = @{}
        $recordingStorage = @{}
        foreach ($rec in Get-VmsRecordingServer) {
            foreach ($storage in $rec | Get-VmsStorage) {
                $recordingStorage[$storage.Path] = $storage
                $pathNameMap[$storage.Path] = $storage.Name
                $pathNameMap[$rec.Path] = $rec.Name
                $childToParentMap[$storage.Path] = $rec.Path
            }
        }

        # Use translations to take an existing device property/value, and modify the column name and value in some way.
        # For example, the GisPoint property has a name unfamiliar to most users, and the "POINT(X Y)" value is even more unfamiliar.
        # Also useful for translating a config API path like "Storage[guid]" to the name of that storage.
        $translations = @{
            'GisPoint'         = {
                @{
                    Name  = 'Coordinates'
                    Value = $_.GisPoint | ConvertFrom-GisPoint
                }
            }
            'RecordingStorage' = {
                @{
                    Name  = 'Storage'
                    Value = $recordingStorage[$_.RecordingStorage].Name
                }
            }
        }

        # Properties to be added. Keys represent the name of a property after which these new properties will be added. Each scriptblock can return one or more Name/Value pairs
        $additionalProperties = @{
            'UserName'                      = {
                $hwPassword = ''
                try {
                    $hwPassword = $_ | Get-VmsHardwarePassword -ErrorAction Stop
                } catch {
                    Write-Warning "Failed to retrieve hardware password. $($_.Exception.Message)"
                }
                [pscustomobject]@{
                    Name  = 'Password'
                    Value = $hwPassword
                }
            }

            'RecordOnRelatedDevices'        = {
                $motion = $_.MotionDetectionFolder.MotionDetections[0]
                [pscustomobject]@{ Name = 'MotionEnabled'; Value = $motion.Enabled }
                [pscustomobject]@{ Name = 'MotionManualSensitivityEnabled'; Value = $motion.ManualSensitivityEnabled }
                [pscustomobject]@{ Name = 'MotionManualSensitivity'; Value = $motion.ManualSensitivity }
                [pscustomobject]@{ Name = 'MotionThreshold'; Value = $motion.Threshold }
                [pscustomobject]@{ Name = 'MotionKeyframesOnly'; Value = $motion.KeyframesOnly }
                [pscustomobject]@{ Name = 'MotionProcessTime'; Value = $motion.ProcessTime }
                [pscustomobject]@{ Name = 'MotionDetectionMethod'; Value = $motion.DetectionMethod }
                [pscustomobject]@{ Name = 'MotionGenerateMotionMetadata'; Value = $motion.GenerateMotionMetadata }
                [pscustomobject]@{ Name = 'MotionUseExcludeRegions'; Value = $motion.UseExcludeRegions }
                [pscustomobject]@{ Name = 'MotionGridSize'; Value = $motion.GridSize }
                [pscustomobject]@{ Name = 'MotionExcludeRegions'; Value = $motion.ExcludeRegions }
                [pscustomobject]@{ Name = 'MotionHardwareAccelerationMode'; Value = $motion.HardwareAccelerationMode }
            }

            'ManualRecordingTimeoutMinutes' = {
                $ptzTimeout = $_.DeviceDriverSettingsFolder.DeviceDriverSettings[0].PTZSessionTimeoutChildItem
                [pscustomobject]@{ Name = 'ManualPTZTimeout'; Value = $ptzTimeout.ManualPTZTimeout }
                [pscustomobject]@{ Name = 'PausePatrollingTimeout'; Value = $ptzTimeout.PausePatrollingTimeout }
                [pscustomobject]@{ Name = 'ReservedPTZTimeout'; Value = $ptzTimeout.ReservedPTZTimeout }
            }

            'RecordingFramerate'            = {
                $privacyMask = $_.PrivacyProtectionFolder.PrivacyProtections[0]
                [pscustomobject]@{ Name = 'PrivacyMaskEnabled'; Value = $privacyMask.Enabled }
                [pscustomobject]@{ Name = 'PrivacyMaskXml'; Value = $privacyMask.PrivacyMaskXml }
            }

            'Channel'                       = {
                $hwId = [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.ParentItemPath).Id
                $hw = [VideoOS.Platform.Configuration]::Instance.GetItem($hwId , [VideoOS.Platform.Kind]::Hardware)
                # If the hardware is disabled, the above command returns $null so we need to fallback to a reliable, but slower, method.
                if ([string]::IsNullOrEmpty($hw)) {
                    $hw = Get-VmsHardware -Id $hwId
                    $recId = [regex]::Matches($hw.ParentItemPath, '(?<=\[)[^]]+(?=\])').Value

                    [pscustomobject]@{
                        Name  = 'Address'
                        Value = $hw.Address
                    }
                    [pscustomobject]@{
                        Name  = 'Hardware'
                        Value = $hw.Name
                    }
                    [pscustomobject]@{
                        Name  = 'RecordingServer'
                        Value = $pathNameMap["RecordingServer[$($recId)]"]
                    }
                } else {
                    [pscustomobject]@{
                        Name  = 'Address'
                        Value = $hw.Properties.Address
                    }
                    [pscustomobject]@{
                        Name  = 'Hardware'
                        Value = $hw.Name
                    }
                    [pscustomobject]@{
                        Name  = 'RecordingServer'
                        Value = $pathNameMap["RecordingServer[$($hw.FQID.ServerId.Id)]"]
                    }
                }
            }

            'EdgeStoragePlaybackEnabled'    = {
                $clientSettings = $_.ClientSettingsFolder.ClientSettings[0]
                if ($clientSettings.Shortcut -eq 0 -or [string]::IsNullOrEmpty($clientSettings.Shortcut)) {
                    [pscustomobject]@{ Name = 'Shortcut'; Value = $null }
                } else {
                    [pscustomobject]@{ Name = 'Shortcut'; Value = $clientSettings.Shortcut }
                }
                [pscustomobject]@{ Name = 'MulticastEnabled'; Value = $clientSettings.MulticastEnabled }
            }

            # Add driver and recording server info after model column for hardware objects
            'Model'                         = {
                if ($hwSettings = ($_ | Get-HardwareSetting -ErrorAction SilentlyContinue)) {
                    [pscustomobject]@{
                        Name  = 'MACAddress'
                        Value = $hwSettings.MacAddress
                    }
                    [pscustomobject]@{
                        Name  = 'SerialNumber'
                        Value = $hwSettings.SerialNumber
                    }
                    [pscustomobject]@{
                        Name  = 'FirmwareVersion'
                        Value = $hwSettings.FirmwareVersion
                    }
                }
                if ($driver = ($_ | Get-VmsHardwareDriver -ErrorAction SilentlyContinue)) {
                    [pscustomobject]@{
                        Name  = 'DriverNumber'
                        Value = $driver.Number
                    }
                    [pscustomobject]@{
                        Name  = 'DriverGroup'
                        Value = $driver.GroupName
                    }
                    [pscustomobject]@{
                        Name  = 'DriverDriverType'
                        Value = $driver.DriverType
                    }
                    [pscustomobject]@{
                        Name  = 'DriverVersion'
                        Value = $driver.DriverVersion
                    }
                    [pscustomobject]@{
                        Name  = 'DriverRevision'
                        Value = $driver.DriverRevision
                    }
                }
                [pscustomobject]@{
                    Name  = 'RecordingServer'
                    Value = $pathNameMap[$_.ParentItemPath]
                }
            }
        }
    }

    process {
        $properties = ($Device | Get-Member -MemberType Property | Where-Object { $_.Name -notlike '*Folder' -and $_.Name -notlike '*Path' -and $_.Name -notin $excludedProperties }).Name

        $obj = [ordered]@{}
        foreach ($property in $orderPriority) {
            if ($null -ne $Device.$property) {
                if ($translations.ContainsKey($property)) {
                    $translations[$property].Invoke($Device) | ForEach-Object {
                        $obj.Add($_.Name, $_.Value)
                    }
                } else {
                    $obj.Add($property, $Device.$property)
                }
                if ($additionalProperties.ContainsKey($property)) {
                    $additionalProperties[$property].Invoke($Device) | ForEach-Object {
                        $obj.Add($_.Name, $_.Value)
                    }
                }
            }
        }
        foreach ($property in $properties | Where-Object { $_ -notin $orderPriority -and $_ -notin $rearOrderPriority }) {
            if ($translations.ContainsKey($property)) {
                $translations[$property].Invoke($Device) | ForEach-Object {
                    $obj.Add($_.Name, $_.Value)
                }
            } else {
                $obj.Add($property, $Device.$property)
            }
            if ($additionalProperties.ContainsKey($property)) {
                $additionalProperties[$property].Invoke($Device) | ForEach-Object {
                    $obj.Add($_.Name, $_.Value)
                }
            }
        }
        foreach ($property in $rearOrderPriority) {
            if ($null -ne $Device.$property) {
                if ($translations.ContainsKey($property)) {
                    $translations[$property].Invoke($Device) | ForEach-Object {
                        $obj.Add($_.Name, $_.Value)
                    }
                } else {
                    $obj.Add($property, $Device.$property)
                }
                if ($additionalProperties.ContainsKey($property)) {
                    $additionalProperties[$property].Invoke($Device) | ForEach-Object {
                        $obj.Add($_.Name, $_.Value)
                    }
                }
            }
        }
        [pscustomobject]$obj
    }
}

function Get-GeneralSettingList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [MilestonePSTools.ValidateVmsItemType('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent', 'Output', 'Metadata')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device
    )

    process {
        $itemType = Split-VmsConfigItemPath -Path $Device.Path -ItemType
        $parentItemId = Split-VmsConfigItemPath -Path $Device.ParentItemPath -Id
        $parentItemType = Split-VmsConfigItemPath -Path $Device.ParentItemPath -ItemType

        $commonProperties = [ordered]@{}
        switch ($parentItemType) {
            'Hardware' {
                $hwItem = [videoos.platform.configuration]::Instance.GetItem($parentItemId, [videoos.platform.kind]::Hardware)
                if ($hwItem) {
                    $recorderItem = [videoos.platform.configuration]::Instance.GetItem($hwItem.FQID.ServerId.Id, [videoos.platform.kind]::Server)
                } else {
                    # If the hardware is disabled, the $hwItem will be $null so we need to fallback to a reliable, but slower, method.
                    $hwItem = Get-VmsHardware -Id $parentItemId
                    $recId = Split-VmsConfigItemPath -Path $hwItem.ParentItemPath -Id
                    $recorderItem = [videoos.platform.configuration]::Instance.GetItem($recId, [videoos.platform.kind]::Server)
                }
                $commonProperties['RecordingServer'] = $recorderItem.Name
                $commonProperties['Hardware'] = $hwItem.Name
            }

            'RecordingServer' {
                $recorderItem = [videoos.platform.configuration]::Instance.GetItem($parentItemId, [videoos.platform.kind]::Server)
                $commonProperties['RecordingServer'] = $recorderItem.Name
            }

            Default {}
        }
        $commonProperties[$itemType] = $Device.Name
        if ($Device.Channel) {
            $commonProperties['Channel'] = $Device.Channel
        }

        $typePrefix = if ($itemType -eq 'Hardware') { 'Hardware' } else { 'Device' }
        Get-ConfigurationItem -Path "$($typePrefix)DriverSettings[$($Device.Id)]" | Select-Object -ExpandProperty Children | Where-Object ItemType -EQ "$($typePrefix)DriverSettings" | Select-Object -ExpandProperty Properties | ForEach-Object {
            $property = $_
            $displayValue = ($property.ValueTypeInfos | Where-Object Value -EQ $property.Value).Name
            $key = $property.Key
            if ($key -match '^([^/]+/)(?<key>[^/]+)(/[^/]+)?$') {
                $key = $Matches.key
            }
            $row = [ordered]@{}
            $commonProperties.Keys | ForEach-Object { $row[$_] = $commonProperties[$_] }
            $row.Setting = $key
            $row.Value = $property.Value
            $row.DisplayValue = if ($property.ValueType -eq 'Enum' -and $displayValue -ne $property.Value) { $displayValue } else { $null }
            $row.ReadOnly = !$property.IsSettable
            [pscustomobject]$row
        }
    }
}

function Import-DeviceEventConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [MilestonePSTools.ValidateVmsItemType('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device,

        [Parameter(Mandatory)]
        [pscustomobject[]]
        $Settings
    )

    process {
        foreach ($eventRow in $Settings) {
            if ($deviceEvent = $Device | Get-VmsDeviceEvent -Name $eventRow.EventName) {
                $setEventArgs = @{
                    Used    = $script:Truthy.IsMatch($eventRow.Used)
                    Enabled = $script:Truthy.IsMatch($eventRow.Enabled)
                    Index   = $eventRow.EventIndex
                    Verbose = $VerbosePreference
                }
                if ($deviceEvent.EventUsed -ne $setEventArgs.Used) {
                    $deviceEvent | Set-VmsDeviceEvent @setEventArgs
                }
            } else {
                Write-Warning "Device '$($Device.Name)' does not have a device event setting with the key '$($eventRow.EventName)'."
            }
        }
    }
}

function Import-DevicePropertyList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device,

        [Parameter(Mandatory)]
        [pscustomobject]
        $Settings
    )

    begin {
        $ignoredColumns = 'RecordingServer', 'Hardware', 'Address', 'LastModified', 'Id', 'MotionDetectionMethod', 'MotionGenerateMotionMetadata', 'MotionGridSize', 'MotionExcludeRegions', 'MotionHardwareAccelerationMode', 'MotionKeyframesOnly', 'MotionManualSensitivity', 'MotionManualSensitivityEnabled', 'MotionProcessTime', 'MotionThreshold', 'MotionUseExcludeRegions', 'PrivacyMaskXml', 'PausePatrollingTimeout', 'ReservedPTZTimeout', 'MulticastEnabled', 'Guid'
        $recordingStorage = @{}
        Get-VmsRecordingServer -Name $Settings.RecordingServer | Get-VmsStorage | ForEach-Object {
            $recordingStorage[$_.Name] = $_
        }

        $translations = @{
            'Coordinates' = {
                param($item, $settings)
                try {
                    @{
                        Name  = 'GisPoint'
                        Value = if ($settings.Coordinates -eq 'Unknown' -or [string]::IsNullOrWhiteSpace($settings.Coordinates)) { 'POINT EMPTY' } else { ConvertTo-GisPoint -Coordinates $settings.Coordinates -ErrorAction Stop }
                    }
                } catch {
                    Write-Warning "Failed to convert value '$($settings.Coordinates)' to a GisPoint value compatible with Milestone."
                }
            }
            'Storage'     = {
                param($item, $settings)
                if ($recordingStorage.ContainsKey($settings.Storage)) {
                    @{
                        Name  = 'RecordingStorage'
                        Value = $recordingStorage[$settings.Storage].Path
                    }
                } else {
                    Write-Warning "Storage configuration '$($settings.Storage)' not found on recording server $($settings.RecordingServer)"
                }
            }
        }

        $customHandlers = @{
            'Enabled'            = {
                param($item, $settings)
                $enabled = $false
                if (-not [string]::IsNullOrWhiteSpace($settings.Enabled) -and [bool]::TryParse($settings.Enabled, [ref]$enabled) -and $item.EnableProperty.Enabled -ne $enabled) {
                    Write-Verbose "Changing 'Enabled' to $enabled on $($item.DisplayName)"
                    $item.EnableProperty.Enabled = $enabled
                    return $true
                }
                return $false
            }

            'RecordingStorage'   = {
                param($item, $settings)
                try {
                    $storagePath = $recordingStorage[$settings.Storage].Path
                    if ($null -eq $storagePath) {
                        throw "Storage configuration named '$($settings.Storage)' not found."
                    }
                    if ($storagePath -eq ($item.Properties | Where-Object Key -EQ 'RecordingStorage').Value) {
                        return $true
                    }
                    $invokeInfo = $item | Invoke-Method -MethodId 'ChangeDeviceRecordingStorage'
                    foreach ($p in $invokeInfo.Properties) {
                        switch ($p.Key) {
                            'ItemSelection' { $p.Value = $storagePath }
                            'moveData' { $p.Value = $false }
                        }
                    }
                    $invokeResult = $invokeInfo | Invoke-Method -MethodId 'ChangeDeviceRecordingStorage'
                    $taskPath = ($invokeResult.Properties | Where-Object Key -EQ 'Path').Value
                    if ($taskPath) {
                        $null = Wait-VmsTask -Path $taskPath -Cleanup
                    }
                    return $true
                } catch {
                    Write-Warning $_.Exception.Message
                }
                return $false
            }

            'MotionEnabled'      = {
                param($item, $settings)
                $motion = Get-ConfigurationItem -Path "MotionDetection[$(($item.Properties | Where-Object Key -EQ Id).Value)]"
                $dirty = $false
                foreach ($column in $settings | Get-Member -MemberType NoteProperty -Name Motion* | Select-Object -ExpandProperty Name) {
                    if ([string]::IsNullOrWhiteSpace($settings.$column)) {
                        continue
                    }
                    $key = $column -replace '^Motion', ''
                    if ($key -eq 'Enabled') {
                        $newValue = $script:Truthy.IsMatch($settings.$column)
                        if ($motion.EnableProperty.Enabled -ne $newValue) {
                            $motion.EnableProperty.Enabled = $newValue
                            $dirty = $true
                        }
                    } else {
                        $property = $motion.Properties | Where-Object Key -EQ $key
                        if ($property.Value -ne $settings.$column) {
                            $property.Value = $settings.$column
                            $dirty = $true
                        }
                    }
                }
                if ($dirty) {
                    $result = $motion | Set-ConfigurationItem
                    if (-not $result.ValidatedOk) {
                        foreach ($errorResult in $result.ErrorResults) {
                            Write-Warning "Failed to update motion detection settings for $($item.DisplayName). $($errorResult.ErrorText)."
                        }
                    }
                }
            }

            'ManualPTZTimeout'   = {
                param($item, $settings)
                $deviceDriverSettings = Get-ConfigurationItem -Path "DeviceDriverSettings[$(($item.Properties | Where-Object Key -EQ Id).Value)]"
                $dirty = $false
                foreach ($column in $settings | Get-Member -MemberType NoteProperty -Name ManualPTZTimeout, PausePatrollingTimeout, ReservedPTZTimeout | Select-Object -ExpandProperty Name) {
                    if ([string]::IsNullOrWhiteSpace($settings.$column)) {
                        continue
                    }
                    $key = $column
                    $property = ($deviceDriverSettings.Children | Where-Object { $_.ItemType -eq 'PTZSessionTimeout' }).Properties | Where-Object Key -EQ $key
                    if ($property) {
                        if ($property.Value -ne $settings.$column) {
                            $property.Value = $settings.$column
                            $dirty = $true
                        }
                    } else {
                        Write-Warning "No PTZSessionTimeout property found in $($item.DisplayName) DeviceDriverSettings named $key"
                    }
                }
                if ($dirty) {
                    $result = $deviceDriverSettings | Set-ConfigurationItem
                    if (-not $result.ValidatedOk) {
                        foreach ($errorResult in $result.ErrorResults) {
                            Write-Warning "Failed to update PTZ session timeout settings for $($item.DisplayName). $($errorResult.ErrorText)."
                        }
                    }
                }
            }

            'PrivacyMaskEnabled' = {
                param($item, $settings)
                $privacyMask = Get-ConfigurationItem -Path "PrivacyProtection[$(($item.Properties | Where-Object Key -EQ Id).Value)]"
                $dirty = $false
                foreach ($column in $settings | Get-Member -MemberType NoteProperty -Name PrivacyMask* | Select-Object -ExpandProperty Name) {
                    if ([string]::IsNullOrWhiteSpace($settings.$column)) {
                        continue
                    }

                    $key = $column
                    if ($key -eq 'PrivacyMaskEnabled') {
                        $newValue = 'True' -eq $settings.$column
                        if ($privacyMask.EnableProperty.Enabled -ne $newValue) {
                            $privacyMask.EnableProperty.Enabled = $newValue
                            $dirty = $true
                        }
                    } elseif ($key -eq 'PrivacyMaskXml') {
                        $property = $privacyMask.Properties | Where-Object Key -EQ $key
                        if ($property.Value -ne $settings.$column) {
                            $property.Value = $settings.$column
                            $dirty = $true
                        }
                    }
                }
                if ($dirty) {
                    $result = $privacyMask | Set-ConfigurationItem
                    if (-not $result.ValidatedOk) {
                        foreach ($errorResult in $result.ErrorResults) {
                            Write-Warning "Failed to update privacy mask settings for $($item.DisplayName). $($errorResult.ErrorText)."
                        }
                    }
                }
            }

            'Shortcut'           = {
                param($item, $settings)
                $clientSettings = Get-ConfigurationItem -Path "ClientSettings[$(($item.Properties | Where-Object Key -EQ Id).Value)]"
                $dirty = $false
                foreach ($column in $settings | Get-Member -MemberType NoteProperty -Name Shortcut, MulticastEnabled | Select-Object -ExpandProperty Name) {
                    if ([string]::IsNullOrWhiteSpace($settings.$column)) {
                        continue
                    }

                    $key = $column
                    if ($key -eq 'MulticastEnabled') {
                        $newValue = 'True' -eq $settings.$column
                        $property = $clientSettings.Properties | Where-Object Key -EQ $key
                        if ($null -eq $property) {
                            Write-Verbose "Property '$column' not found in ClientSettings for $($item.DisplayName). It may not be available on this VMS version."
                            continue
                        }
                        if ($property.Value -ne $newValue) {
                            $property.Value = $newValue
                            $dirty = $true
                        }
                    } elseif ($key -eq 'Shortcut') {
                        $property = $clientSettings.Properties | Where-Object Key -EQ $key
                        if ($null -eq $property) {
                            Write-Verbose "Property '$column' not found in ClientSettings for $($item.DisplayName). It may not be available on this VMS version."
                            continue
                        }
                        if ($property.Value -ne $settings.$column -and $settings.$column -ge 1) {
                            $property.Value = $settings.$column
                            $dirty = $true
                        }
                    }
                }
                if ($dirty) {
                    $result = $clientSettings | Set-ConfigurationItem
                    if (-not $result.ValidatedOk) {
                        foreach ($errorResult in $result.ErrorResults) {
                            Write-Warning "Failed to update privacy mask settings for $($item.DisplayName). $($errorResult.ErrorText)."
                        }
                    }
                }
            }
        }
    }

    process {
        $dirty = $false
        $properties = @{}
        $item = $Device | Get-ConfigurationItem
        $item.Properties | ForEach-Object { $properties[$_.Key] = $_ }

        foreach ($columnName in $Settings | Get-Member -MemberType NoteProperty | Where-Object Name -NotIn $ignoredColumns | Select-Object -ExpandProperty Name) {
            $newValue = $Settings.$columnName
            if ($translations.ContainsKey($columnName)) {
                $columnName, $newValue = $translations[$columnName].Invoke($item, $Settings) | ForEach-Object {
                    Write-Verbose "Translating column name '$($columnName)' to '$($_.Name)', and value '$($newValue)' to '$($_.Value)'"
                    @($_.Name, $_.Value)
                }
                if ($null -eq $columnName -or $null -eq $newValue) {
                    Write-Verbose 'Failed to translate column/value. No change will be made for this property.'
                    continue
                }
            }

            if ($customHandlers.ContainsKey($columnName)) {
                Write-Verbose "Invoking custom handler for column $columnName on device $($Device.Name)"
                if ($customHandlers[$columnName].Invoke($item, $Settings)) {
                    $dirty = $true
                }
            } else {
                $property = $properties[$columnName]
                if ($property) {
                    if ($property.Value -ne $newValue) {
                        Write-Verbose "Setting $columnName to $newValue on $($Device.Name)"
                        $property.Value = $newValue
                        $dirty = $true
                    } else {
                        Write-Verbose "Setting $columnName already has value $newValue on $($Device.Name)"
                    }
                } else {
                    Write-Warning "Property '$($columnName)' not found on $($Device.Name)"
                }
            }
        }

        # Update the name for the in-memory copy of $Device so that the verbose logging doesn't mention the old name anymore.
        $Device.Name = ($item.Properties | Where-Object Key -EQ 'Name').Value

        if ($dirty) {
            Write-Verbose "Saving changes to $($Device.Name)"
            $result = $item | Set-ConfigurationItem
            foreach ($entry in $result.ErrorResults) {
                Write-Error -Message "Validation error: $($entry.ErrorText) on '$($Device.Name)'."
            }
        } else {
            Write-Verbose "No changes made to $($Device.Name)"
        }
    }
}

function Import-GeneralSettingList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device,

        [Parameter(Mandatory)]
        [pscustomobject[]]
        $Settings
    )

    begin {
        $validDeviceTypes = @('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent', 'Output', 'Metadata')
    }

    process {
        $devicePath = [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($Device.Path)
        if ($devicePath.ItemType -notin $validDeviceTypes) {
            Write-Error 'Invalid device type for this cmdlet.'
            return
        }

        $itemType = if ($devicePath.ItemType -eq 'Hardware') { 'Hardware' } else { 'Device' }
        Write-Verbose "$($devicePath.ItemType)GeneralSettings: Checking general settings for '$($Device.Name)'"
        $item = Get-ConfigurationItem -Path "$($itemType)DriverSettings[$($Device.Id)]"
        $general = $item.Children | Where-Object ItemType -EQ "$($itemType)DriverSettings"
        $dirty = $false
        foreach ($setting in $Settings) {
            $property = $general.Properties | Where-Object Key -Match "^([^/]+/)?(?<key>$([regex]::Escape($setting.Setting)))(/[^/]+)?$" | Select-Object -First 1
            $key = $setting.Setting
            
            if ($null -eq $property) {
                Write-Warning "$($devicePath.ItemType)GeneralSettings: Device '$($Device.Name)' does not have a general setting with the key '$($setting.Setting)'."
                continue
            } elseif (!$property.IsSettable) {
                continue
            }
            $incomingValue = $setting.Value
            if ($property.ValueType -eq 'Enum' -and $incomingValue -in $property.ValueTypeInfos.Value) {
                # Handle incorrect case for incoming settings by doing case-insensitive check against
                # available enum values.
                $incomingValue = $property.ValueTypeInfos.Value | Where-Object { $_ -eq $incomingValue }
            }
            if ($property.Value -cne $incomingValue) {
                Write-Verbose "$($devicePath.ItemType)GeneralSettings: Changing $($property.DisplayName) ($key) to '$($incomingValue)'"
                $property.Value = $incomingValue
                $dirty = $true
            } else {
                Write-Verbose "$($devicePath.ItemType)GeneralSettings: Keeping $($property.DisplayName) ($key) value '$($property.Value)'"
            }
        }

        if (-not $dirty) {
            Write-Verbose "$($devicePath.ItemType)GeneralSettings: No changes to general settings were required for '$($Device.Name)'"
            return
        }

        Write-Verbose "$($devicePath.ItemType)GeneralSettings: Saving changes to general settings for '$($Device.Name)'"
        $result = $item | Set-ConfigurationItem
        foreach ($entry in $result.ErrorResults) {
            Write-Error -Message "$($devicePath.ItemType)GeneralSettings: Validation error: $($entry.ErrorText) on '$($Device.Name)'."
        }
    }
}


function Export-VmsHardwareExcel {
    <#
    .SYNOPSIS
    Exports hardware configuration in Microsoft Excel XLSX format.

    .DESCRIPTION
    The `Export-VmsHardwareExcel` cmdlet accepts one or more Hardware objects
    from `Get-VmsHardware` and exports detailed configuration to an Excel XLSX
    document.

    The document will contain multiple worksheets, depending on which device
    types are specified in the `IncludedDevices` parameter. Each area of the
    hardware configuration is represented in it's own worksheet which makes it
    possible to represent many different types of objects and settings in the
    same document while keeping it human-readable and easy to modify.

    .PARAMETER Hardware
    Specifies one or more Hardware objects returned by `Get-VmsHardware`. If no
    hardware is provided, then all hardware found in the VMS matching the
    desired `EnableState` will be exported.

    .PARAMETER Path
    The absolute, or relative path, including filename, where the .XLSX file
    should be saved. If no path is provided, a save-file dialog will be shown.

    .PARAMETER IncludedDevices
    Defaults to "Cameras". Specifies the types of child devices to include in the export. It can be
    very time consuming to export configuration for thousands of devices, and
    if you only need camera and metadata settings, you can specify this and
    avoid retrieving detailed configuration on microphones, speakers, inputs,
    and outputs.

    .PARAMETER EnableFilter
    Defaults to "Enabled". Filters the exported hardware and devices to only
    those matching the specified EnableFilter.

    .PARAMETER Force
    Overwrite an existing file if the file specified in `Path` already exists.

    .EXAMPLE
    Export-VmsHardwareExcel -Path ~\Documents\hardware.xlsx -Verbose

    Exports configuration for all enabled hardware, and cameras to the current
    user's Documents directory.

    .EXAMPLE
    Export-VmsHardwareExcel -Path ~\Documents\hardware.xlsx -IncludedDevices Cameras, Microphones -Verbose

    Exports configuration for all enabled hardware, cameras, and microphones to
    the current user's Documents directory.

    .EXAMPLE
    $hardware = Get-VmsRecordingServer -Name Recorder1 | Get-VmsHardware
    Export-VmsHardwareExcel -Hardware $hardware -Path ~\Desktop\hardware.xlsx -Verbose

    Exports configuration for all enabled hardware, and cameras on the
    recording server named "Recorder1" tp the current user's Desktop.

    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Hardware[]]
        $Hardware,

        [Parameter()]
        [string]
        $Path,

        [Parameter()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output')]
        [string[]]
        $IncludedDevices = @('Camera'),

        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'Enabled',

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $Path = Show-FileDialog -SaveFile
        }
        if (Test-Path $Path) {
            throw ([io.ioexception]::new("File $Path already exists."))
        } else {
            $directoryPath = Split-Path -Path $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) -Parent
            $null = New-Item -Path $directoryPath -ItemType Directory -Force
        }
        $excelPackage = Open-ExcelPackage -Path $Path -Create
        $worksheets = @(
            'Hardware',
            'HardwareGeneralSettings',
            'HardwarePtzSettings',
            'HardwareEvents',
            'Cameras',
            'CameraGeneralSettings',
            'CameraStreams',
            'CameraStreamSettings',
            'CameraPtzPresets',
            'CameraPtzPatrols',
            'CameraPtzPatrolPresets',
            'CameraRelatedDevices',
            'CameraEvents',
            'CameraGroups',
            'Microphones',
            'MicrophoneGeneralSettings',
            'MicrophoneStreamSettings',
            'MicrophoneEvents',
            'MicrophoneGroups',
            'Speakers',
            'SpeakerGeneralSettings',
            'SpeakerEvents',
            'SpeakerGroups',
            'Metadata',
            'MetadataGeneralSettings',
            'MetadataGroups',
            'Inputs',
            'InputGeneralSettings',
            'InputEvents',
            'InputGroups'
            'Outputs',
            'OutputGeneralSettings',
            'OutputGroups'
        )
        $null = $worksheets | ForEach-Object { $excelPackage.Workbook.Worksheets.Add($_) }
        Clear-VmsCache
    }

    process {
        $progress = @{
            Activity         = 'Exporting hardware configuration to {0}' -f $Path
            Id               = 11
            PercentComplete  = 0
            CurrentOperation = 'Preparing'
        }
        Write-Progress @progress

        if ($IncludedDevices) {
            $IncludedDevices = $IncludedDevices | Group-Object | Select-Object -ExpandProperty Name
        }


        $progress.CurrentOperation = 'Retrieving list of recording servers'
        Write-Progress @progress
        Write-Verbose 'Retrieving recording server list'
        $recorderMap = @{}
        Get-VmsRecordingServer | ForEach-Object {
            $recorderMap[$_.Path] = $_
        }

        if ($null -eq $Hardware) {
            $progress.CurrentOperation = 'Retrieving list of hardware to be exported'
            Write-Progress @progress
            $Hardware = Get-VmsHardware
        }

        Write-Verbose 'Loading device groups'
        $deviceGroups = @{}
        $IncludedDevices | ForEach-Object {
            $type = $_ -replace 's$', ''
            foreach ($group in Get-VmsDeviceGroup -Type $type -Recurse) {
                $members = $group | Get-VmsDeviceGroupMember -EnableFilter $EnableFilter
                if ($members.Count -eq 0) { continue }
                
                $groupPath = $group | Resolve-VmsDeviceGroupPath -NoTypePrefix
                foreach ($member in $members) {
                    if (-not $deviceGroups.ContainsKey($member.Id)) {
                        $deviceGroups[$member.Id] = [list[string]]::new()
                    }
                    $deviceGroups[$member.Id].Add($groupPath)
                }
            }
        }

        $excelParams = @{
            ExcelPackage       = $excelPackage
            TableStyle         = 'Medium9'
            AutoSize           = $true
            Append             = $true
            NoNumberConversion = 'Value', 'DisplayValue', 'MotionExcludeRegions', 'MACAddress', 'SerialNumber', 'FirmwareVersion', 'Password'
            PassThru           = $true
        }

        $totalHardwareCount = $Hardware.Count
        $processedHardwareCount = 0
        $Hardware | ForEach-Object {
            $hw = $_
            $progress.PercentComplete = [math]::Round(($processedHardwareCount++) / $totalHardwareCount * 100)
            $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
            Write-Progress @progress
            if (($EnableFilter -eq 'Enabled' -and -not $hw.Enabled) -or ($EnableFilter -eq 'Disabled' -and $hw.Enabled)) {
                Write-Verbose "Skipping hardware $($hw.Name) due to the EnableFilter value of $EnableFilter"
                return
            }
            Write-Verbose "Retrieving hardware properties for $($hw.Name)"
            $null = $hw | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Hardware -TableName HardwareList }

            Write-Verbose "Retrieving general setting properties for $($hw.Name)"
            $null = $hw | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName HardwareGeneralSettings -TableName HardwareGeneralSettingsList }

            $channel = 0
            $hw.HardwarePtzSettingsFolder.HardwarePtzSettings.HardwarePtzDeviceSettingChildItems | Where-Object { $null -ne $_ } | ForEach-Object {
                $hwPtzSettings = @(

                    @{
                        Name       = 'RecordingServer'
                        Expression = { $recorderMap[$hw.ParentItemPath].Name }
                    },
                    @{
                        Name       = 'Hardware'
                        Expression = { $hw.Name }
                    },
                    @{
                        Name       = 'Camera'
                        Expression = { $_.DisplayName }
                    },
                    @{
                        Name       = 'Channel'
                        Expression = { $channel }
                    },
                    @{
                        Name       = 'PTZEnabled'
                        Expression = { $_.Properties.GetValue('PTZEnabled') }
                    },
                    @{
                        Name       = 'PTZDeviceID'
                        Expression = { $_.Properties.GetValue('PTZDeviceID') }
                    },
                    @{
                        Name       = 'PTZCOMPort'
                        Expression = { $_.Properties.GetValue('PTZCOMPort') }
                    },
                    @{
                        Name       = 'PTZProtocol'
                        Expression = { $_.Properties.GetValue('PTZProtocol') }
                    }
                )
                $null = $_ | Select-Object $hwPtzSettings | Export-Excel @excelParams -WorksheetName HardwarePtzSettings -TableName HardwarePtzSettingsList
                $channel += 1
            }

            Write-Verbose "Retrieving event properties for $($hw.Name)"
            $obj = [ordered]@{
                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                Hardware        = $hw.Name
            }
            $null = $hw | Export-DeviceEventConfig | ForEach-Object {
                $eventInfo = $_
                $obj.EventName = $eventInfo.Event
                $obj.Used = $eventInfo.Used
                $obj.Enabled = $eventInfo.Enabled
                $obj.EventIndex = $eventInfo.EventIndex
                $obj.IndexName = $eventInfo.IndexName
                [pscustomobject]$obj | Export-Excel @excelParams -WorksheetName HardwareEvents -TableName HardwareEventsList
            }

            if ('Camera' -in $IncludedDevices) {
                $hw | Get-VmsCamera -EnableFilter $EnableFilter | ForEach-Object {
                    Write-Verbose "Retrieving camera properties for $($_.Name)"
                    $cam = $_
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $null = $cam | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Cameras -TableName CamerasList }

                    Write-Verbose "Retrieving general setting properties for $($cam.Name)"
                    $null = $cam | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName CameraGeneralSettings -TableName CameraGeneralSettingsList }

                    Write-Verbose "Retrieving stream properties for $($cam.Name)"
                    $recordingTrack = @{
                        '16ce3aa1-5f93-458a-abe5-5c95d9ed1372' = 'Primary'
                        '84fff8b9-8cd1-46b2-a451-c4a87d4cbbb0' = 'Secondary'
                        ''                                     = 'None'
                    }
                    $supportsAdaptivePlayback = [version](Get-VmsManagementServer).Version -ge '23.2'
                    $cam | Get-VmsCameraStream -Enabled -RawValues | ForEach-Object {
                        $stream = $_
                        $obj = [pscustomobject]@{
                            RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                            Hardware        = $hw.Name
                            Camera          = $cam.Name
                            Channel         = $cam.Channel

                            Name            = $stream.Name
                            DisplayName     = $stream.DisplayName
                            LiveMode        = $stream.LiveMode
                            LiveDefault     = $stream.LiveDefault
                            PlaybackDefault = $stream.PlaybackDefault
                            RecordingTrack  = if ($supportsAdaptivePlayback) { $recordingTrack["$($stream.RecordingTrack)"] } elseif ($stream.Recorded) { 'Primary' } else { 'None' }
                            UseEdge         = $stream.UseEdge
                        }
                        $null = $obj | Export-Excel @excelParams -WorksheetName CameraStreams -TableName CameraStreamsList

                        $null = $stream.Settings.Keys | ForEach-Object {
                            $key = $_
                            $displayValue = ($stream.ValueTypeInfo[$key] | Where-Object { $_.Value -eq $property.Value -and $_.Name -notlike '*Value' }).Name
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Camera          = $cam.Name
                                Channel         = $cam.Channel
                                Stream          = $stream.Name

                                Setting         = $key
                                Value           = $stream.Settings[$key]
                                DisplayValue    = if ($stream.Settings[$key] -ne $displayValue) { $displayValue } else { $null }
                            } | Export-Excel @excelParams -WorksheetName CameraStreamSettings -TableName CameraStreamSettingsList
                        }
                    }

                    Write-Verbose "Retrieving system PTZ presets for $($cam.Name)"
                    $cam.PtzPresetFolder.PtzPresets | Where-Object { $null -ne $_ } | ForEach-Object {
                        $ptzPreset = $_
                        if ($ptzPreset.DevicePreset) {
                            Write-Verbose "Camera $($cam.Name) has preset positions defined on camera and not in the VMS."
                            return
                        }
                        $obj = [pscustomobject]@{
                            RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                            Hardware        = $hw.Name
                            Camera          = $cam.Name
                            Channel         = $cam.Channel

                            DefaultPreset   = $ptzPreset.DefaultPreset
                            Pan             = $ptzPreset.Pan 
                            Tilt            = $ptzPreset.Tilt
                            Zoom            = $ptzPreset.Zoom
                            Name            = $ptzPreset.Name
                            Description     = $ptzPreset.Description
                        }
                        $null = $obj | Export-Excel @excelParams -WorksheetName CameraPtzPresets -TableName CameraPtzPresetsList
                    }

                    Write-Verbose "Retrieving system PTZ patrols for $($cam.Name)"
                    $ptzPresets = $cam.PtzPresetFolder.PtzPresets
                    $cam.PatrollingProfileFolder.PatrollingProfiles | Where-Object { $null -ne $_ } | ForEach-Object {
                        $ptzPatrol = $_
                        if ($cam.PtzPresetFolder.PtzPresets[0].DevicePreset) {
                            Write-Verbose "Camera $($cam.Name) has preset positions defined on camera so skipping PTZ patrolling profiles."
                            return
                        }
                        $obj = [pscustomobject]@{
                            RecordingServer      = $recorderMap[$hw.ParentItemPath].Name
                            Hardware             = $hw.Name
                            Camera               = $cam.Name
                            Channel              = $cam.Channel

                            Name                 = $ptzPatrol.Name
                            Description          = $ptzPatrol.Description
                            CustomizeTransitions = $ptzPatrol.CustomizeTransitions
                            InitSpeed            = $ptzPatrol.InitSpeed
                            InitTransitionTime   = $ptzPatrol.InitTransitionTime
                            EndPresetId          = $ptzPatrol.EndPresetId
                            EndPresetName        = ($ptzPresets | Where-Object Id -EQ $ptzPatrol.EndPresetId).Name
                            EndSpeed             = $ptzPatrol.EndSpeed
                            EndTransitionTime    = $ptzPatrol.EndTransitionTime

                        }
                        $null = $obj | Export-Excel @excelParams -WorksheetName CameraPtzPatrols -TableName CameraPtzPatrolsList

                        Write-Verbose "Retrieving PTZ patrols presets for $($ptzPatrol.Name) on $($cam.Name)"
                        $patrolChildren = (Get-ConfigurationItem -Path "PatrollingProfile[$($ptzPatrol.Id)]").Children

                        for ($i = 0; $i -lt $patrolChildren.Count; $i++) {
                            $patrolChild = $patrolChildren | Where-Object { $_.Path -eq "PatrollingEntry[$($i)]" }
                            $presetId = ($patrolChild.Properties | Where-Object Key -EQ PresetId).Value

                            $obj = [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Camera          = $cam.Name
                                Channel         = $cam.Channel
                                Patrol          = $ptzPatrol.Name

                                Order           = ($patrolChild.Properties | Where-Object Key -EQ Order).Value
                                WaitTime        = ($patrolChild.Properties | Where-Object Key -EQ WaitTime).Value
                                Speed           = ($patrolChild.Properties | Where-Object Key -EQ Speed).Value
                                TransitionTime  = ($patrolChild.Properties | Where-Object Key -EQ TransitionTime).Value
                                PresetName      = ($ptzPresets | Where-Object { $_.Id -eq $presetId }).Name
                            }
                            $null = $obj | Export-Excel @excelParams -WorksheetName CameraPtzPatrolPresets -TableName CameraPtzPatrolPresetsList
                        }
                    }

                    Write-Verbose "Retrieving related devices, shortcut number, and multicast setting for $($cam.Name)"
                    $clientSettings = $cam.ClientSettingsFolder.ClientSettings[0]
                    if (-not [string]::IsNullOrEmpty($clientSettings.Related)) {
                        $relatedDevices = [list[pscustomobject]]::new()
                        $clientSettings.Related.Split(',') | ForEach-Object {
                            $deviceType = $_.Split('[') | Select-Object -First 1
                            $deviceCI = Get-ConfigurationItem -Path $_
                            $deviceProperties = $deviceCI.Properties
                            $hardwarePath = $deviceCI.ParentPath.Split('/') | Select-Object -First 1
                            $hardwareCI = Get-ConfigurationItem -Path $hardwarePath
                            $hardwareProperties = $hardwareCI.Properties
                            $recordingServerPath = $hardwareCI.ParentPath.Split('/') | Select-Object -First 1
                            $recordingServerCI = Get-ConfigurationItem -Path $recordingServerPath
                            $recordingServerProperties = $recordingServerCI.Properties

                            $row = [PSCustomObject]@{
                                RelatedDeviceType              = $deviceType
                                RelatedRecordingServerName     = ($recordingServerProperties | Where-Object Key -EQ Name).Value
                                RelatedRecordingServerHostName = ($recordingServerProperties | Where-Object Key -EQ HostName).Value
                                RelatedHardwareName            = ($hardwareProperties | Where-Object Key -EQ Name).Value
                                RelatedHardwareAddress         = ($hardwareProperties | Where-Object Key -EQ Address).Value
                                RelatedDeviceName              = ($deviceProperties | Where-Object Key -EQ Name).Value
                                RelatedDeviceChannel           = ($deviceProperties | Where-Object Key -EQ Channel).Value
                            }
                            $relatedDevices.Add($row)
                        }
                    } else {
                        $relatedDevices = $null
                    }

                    foreach ($relatedDevice in $relatedDevices) {
                        $obj = [pscustomobject]@{
                            RecordingServer                = $recorderMap[$hw.ParentItemPath].Name
                            Hardware                       = $hw.Name
                            Camera                         = $cam.Name
                            Channel                        = $cam.Channel

                            RelatedDeviceType              = $relatedDevice.RelatedDeviceType
                            RelatedRecordingServerName     = $relatedDevice.RelatedRecordingServerName
                            RelatedRecordingServerHostName = $relatedDevice.RelatedRecordingServerHostName
                            RelatedHardwareName            = $relatedDevice.RelatedHardwareName
                            RelatedHardwareAddress         = $relatedDevice.RelatedHardwareAddress
                            RelatedDeviceName              = $relatedDevice.RelatedDeviceName
                            RelatedDeviceChannel           = $relatedDevice.RelatedDeviceChannel
                            Shortcut                       = $clientSettings.Shortcut
                            MulticastEnabled               = $clientSettings.MulticastEnabled
                        }
                        $null = $obj | Export-Excel @excelParams -WorksheetName CameraRelatedDevices -TableName CameraRelatedDevicesList
                    }

                    Write-Verbose "Retrieving event properties for $($cam.Name)"
                    $obj = [ordered]@{
                        RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                        Hardware        = $hw.Name
                        Camera          = $cam.Name
                    }
                    $null = $cam | Export-DeviceEventConfig | ForEach-Object {
                        $eventInfo = $_
                        $obj.EventName = $eventInfo.Event
                        $obj.Used = $eventInfo.Used
                        $obj.Enabled = $eventInfo.Enabled
                        $obj.EventIndex = $eventInfo.EventIndex
                        $obj.IndexName = $eventInfo.IndexName
                        [pscustomobject]$obj | Export-Excel @excelParams -WorksheetName CameraEvents -TableName CameraEventsList
                    }

                    Write-Verbose "Retrieving device groups for $($cam.Name)"
                    if ($deviceGroups.ContainsKey($cam.Id)) {
                        $null = $deviceGroups[$cam.Id] | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Camera          = $cam.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName CameraGroups -TableName CameraGroupsList
                    }
                }
            }

            if ('Microphone' -in $IncludedDevices) {
                $deviceType = 'Microphone'
                $deviceTypePlural = "Microphones"

                $hw | Get-VmsMicrophone -EnableFilter $EnableFilter | ForEach-Object {
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $device = $_
                    if (($EnableFilter -eq 'Enabled' -and -not $device.Enabled) -or ($EnableFilter -eq 'Disabled' -and $device.Enabled)) {
                        Write-Verbose "Skipping $deviceType $($device.Name) due to the EnableFilter value of $EnableFilter"
                        return
                    }

                    Write-Verbose "Retrieving $deviceType properties for $($device.Name)"
                    $null = $device | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName $deviceTypePlural -TableName "$($deviceTypePlural)List" }

                    Write-Verbose "Retrieving general setting properties for $($device.Name)"
                    $null = $device | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName "$($deviceType)GeneralSettings" -TableName "$($deviceType)GeneralSettingsList" }

                    Write-Verbose "Retrieving stream properties for $($device.Name)"
                    $deviceDriverSettings | Select-Object -ExpandProperty Children | Where-Object ItemType -EQ Stream | Select-Object -ExpandProperty Properties | Where-Object IsSettable | ForEach-Object {
                        if ($null -eq $_) {
                            return
                        }
                        $property = $_
                        $key = $property.Key
                        $displayValue = ($property.ValueTypeInfos | Where-Object Value -EQ $property.Value).Name
                        if ($key -match '^[^/]+/(?<key>.*?)/[^/]+$') {
                            # If the value of $property.Key is in the format 'device:0:1/KeyName/usually-a-guid', we just want the KeyName value in the middle
                            $key = $Matches.key
                        }
                        $obj = [pscustomobject]@{
                            RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                            Hardware        = $hw.Name
                            Microphone      = $device.Name
                            Channel         = $device.Channel

                            Setting         = $key
                            Value           = $property.Value
                            DisplayValue    = if ($property.ValueType -eq 'Enum' -and $displayValue -ne $property.Value) { $displayValue } else { $null }
                        }
                        $null = $obj | Export-Excel @excelParams -WorksheetName MicrophoneStreamSettings -TableName MicrophoneStreamSettingsList
                    }

                    Write-Verbose "Retrieving event properties for $($device.Name)"
                    $obj = [ordered]@{
                        RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                        Hardware        = $hw.Name
                        Microphone      = $device.Name
                    }
                    $null = $device | Export-DeviceEventConfig | ForEach-Object {
                        $eventInfo = $_
                        $obj.EventName = $eventInfo.Event
                        $obj.Used = $eventInfo.Used
                        $obj.Enabled = $eventInfo.Enabled
                        $obj.EventIndex = $eventInfo.EventIndex
                        $obj.IndexName = $eventInfo.IndexName
                        [pscustomobject]$obj | Export-Excel @excelParams -WorksheetName MicrophoneEvents -TableName MicrophoneEventsList
                    }

                    Write-Verbose "Retrieving device groups for $($device.Name)"
                    if ($deviceGroups.ContainsKey($device.Id)) {
                        $null = $deviceGroups[$device.Id] | Where-Object { $null -ne $_ } | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Microphone      = $device.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName MicrophoneGroups -TableName MicrophoneGroupsList
                    }
                }
            }

            if ('Speaker' -in $IncludedDevices) {
                $hw | Get-VmsSpeaker -EnableFilter $EnableFilter | ForEach-Object {
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $device = $_
                    if (($EnableFilter -eq 'Enabled' -and -not $device.Enabled) -or ($EnableFilter -eq 'Disabled' -and $device.Enabled)) {
                        Write-Verbose "Skipping speaker $($device.Name) due to the EnableFilter value of $EnableFilter"
                        return
                    }
                    Write-Verbose "Retrieving speaker properties for $($device.Name)"
                    $null = $device | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Speakers -TableName SpeakersList }

                    Write-Verbose "Retrieving general setting properties for $($device.Name)"
                    $null = $device | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName SpeakerGeneralSettings -TableName SpeakerGeneralSettingsList }

                    Write-Verbose "Retrieving event properties for $($device.Name)"
                    $obj = [ordered]@{
                        RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                        Hardware        = $hw.Name
                        Speaker         = $device.Name
                    }
                    $null = $device | Export-DeviceEventConfig | ForEach-Object {
                        $eventInfo = $_
                        $obj.EventName = $eventInfo.Event
                        $obj.Used = $eventInfo.Used
                        $obj.Enabled = $eventInfo.Enabled
                        $obj.EventIndex = $eventInfo.EventIndex
                        $obj.IndexName = $eventInfo.IndexName
                        [pscustomobject]$obj | Export-Excel @excelParams -WorksheetName SpeakerEvents -TableName SpeakerEventsList
                    }

                    Write-Verbose "Retrieving device groups for $($device.Name)"
                    if ($deviceGroups.ContainsKey($device.Id)) {
                        $null = $deviceGroups[$device.Id] | Where-Object { $null -ne $_ } | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Speaker         = $device.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName SpeakerGroups -TableName SpeakerGroupsList
                    }
                }
            }

            if ('Metadata' -in $IncludedDevices) {
                $hw | Get-VmsMetadata -EnableFilter $EnableFilter | ForEach-Object {
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $device = $_
                    if (($EnableFilter -eq 'Enabled' -and -not $device.Enabled) -or ($EnableFilter -eq 'Disabled' -and $device.Enabled)) {
                        Write-Verbose "Skipping metadata $($device.Name) due to the EnableFilter value of $EnableFilter"
                        return
                    }
                    Write-Verbose "Retrieving metadata properties for $($device.Name)"
                    $null = $device | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Metadata -TableName MetadataList }

                    Write-Verbose "Retrieving metadata general settings for $($device.Name)"
                    $null = $device | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName MetadataGeneralSettings -TableName MetadataGeneralSettingsList }

                    Write-Verbose "Retrieving device groups for $($device.Name)"
                    if ($deviceGroups.ContainsKey($device.Id)) {
                        $null = $deviceGroups[$device.Id] | Where-Object { $null -ne $_ } | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Metadata        = $device.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName MetadataGroups -TableName MetadataGroupsList
                    }
                }
            }

            if ('Input' -in $IncludedDevices) {
                $hw | Get-VmsInput -EnableFilter $EnableFilter | ForEach-Object {
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $device = $_
                    if (($EnableFilter -eq 'Enabled' -and -not $device.Enabled) -or ($EnableFilter -eq 'Disabled' -and $device.Enabled)) {
                        Write-Verbose "Skipping input $($device.Name) due to the EnableFilter value of $EnableFilter"
                        return
                    }
                    Write-Verbose "Retrieving input properties for $($device.Name)"
                    $null = $device | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Inputs -TableName InputsList }

                    Write-Verbose "Retrieving input general settings for $($device.Name)"
                    $null = $device | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName InputGeneralSettings -TableName InputGeneralSettingsList }

                    Write-Verbose "Retrieving event properties for $($device.Name)"
                    $obj = [ordered]@{
                        RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                        Hardware        = $hw.Name
                        Input           = $device.Name
                    }
                    $null = $device | Export-DeviceEventConfig | ForEach-Object {
                        $eventInfo = $_
                        $obj.EventName = $eventInfo.Event
                        $obj.Used = $eventInfo.Used
                        $obj.Enabled = $eventInfo.Enabled
                        $obj.EventIndex = $eventInfo.EventIndex
                        $obj.IndexName = $eventInfo.IndexName
                        [pscustomobject]$obj | Export-Excel @excelParams -WorksheetName InputEvents -TableName InputEventsList
                    }

                    Write-Verbose "Retrieving device groups for $($device.Name)"
                    if ($deviceGroups.ContainsKey($device.Id)) {
                        $null = $deviceGroups[$device.Id] | Where-Object { $null -ne $_ } | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Input           = $device.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName InputGroups -TableName InputGroupsList
                    }
                }
            }

            if ('Output' -in $IncludedDevices) {
                $hw | Get-VmsOutput -EnableFilter $EnableFilter | ForEach-Object {
                    $progress.CurrentOperation = '{0} "{1}"' -f [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($_.Path).ItemType, $_.Name
                    Write-Progress @progress
                    $device = $_
                    if (($EnableFilter -eq 'Enabled' -and -not $device.Enabled) -or ($EnableFilter -eq 'Disabled' -and $device.Enabled)) {
                        Write-Verbose "Skipping output $($device.Name) due to the EnableFilter value of $EnableFilter"
                        return
                    }
                    Write-Verbose "Retrieving output properties for $($device.Name)"
                    $null = $device | Get-DevicePropertyList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName Outputs -TableName OutputsList }

                    Write-Verbose "Retrieving output general settings for $($device.Name)"
                    $null = $device | Get-GeneralSettingList | ForEach-Object { $_ | Export-Excel @excelParams -WorksheetName OutputGeneralSettings -TableName OutputGeneralSettingsList }

                    Write-Verbose "Retrieving device groups for $($device.Name)"
                    if ($deviceGroups.ContainsKey($device.Id)) {
                        $null = $deviceGroups[$device.Id] | Where-Object { $null -ne $_ } | ForEach-Object {
                            [pscustomobject]@{
                                RecordingServer = $recorderMap[$hw.ParentItemPath].Name
                                Hardware        = $hw.Name
                                Output          = $device.Name
                                Group           = $_
                            }
                        } | Export-Excel @excelParams -WorksheetName OutputGroups -TableName OutputGroupsList
                    }
                }
            }
        }
        $progress.PercentComplete = 100
        $progress.Completed = $true
        Write-Progress @progress
    }

    end {
        $excelPackage.Workbook.Worksheets.Name | ForEach-Object {
            if ($null -eq $excelPackage.Workbook.Worksheets[$_].GetValue(1, 1)) {
                $excelPackage.Workbook.Worksheets.Delete($_)
            }
        }
        $excelPackage | Close-ExcelPackage
    }
}

function Import-VmsHardwareExcel {
    <#
    .SYNOPSIS
    Imports hardware configuration from an Excel .XLSX document and adds and
    optionally updates hardware based.

    .DESCRIPTION
    The `Import-VmsHardwareExcel` cmdlet accepts a path to an existing Excel
    .XLSX document, and imports the hardware configuration. The cmdlet can add
    new devices and update the settings of existing devices if the values in
    the Excel document differ from the live values.

    Depending on the content of the Excel document, the settings imported can
    include hardware, general settings, cameras, microphones, speakers, inputs,
    outputs, metadata, and the corresponding general settings, settings for
    streams, recording, events, motion, and more.

    The format of the Excel document, and the valid values for various settings
    is challenging to document. The best way to perform a successful import is
    to add and configure a representative sample of devices, and then use
    `Export-VmsHardwareExcel` to generate a configuration export. You can then
    use the export as a reference to build a document to import.

    .PARAMETER Path
    Specifies a path to an existing Excel document in .XLSX format. While the
    `ImportExcel` module supports reading from password protected files, this
    has not been extended to this cmdlet. If no path is provided, an open-file
    dialog will be shown.

    .PARAMETER UpdateExisting
    If hardware defined in the Excel document is already added, it will not be
    modified by default. If you wish to update the settings for existing
    hardware during an import, this switch can be used.

    .EXAMPLE
    Import-VmsHardwareExcel -Path ~\Desktop\hardware.xlsx -Verbose

    Imports the hardware.xlsx file on the current user's desktop. If any cameras
    in the Excel document are already added, they will be ignored and their
    settings will not be modified if they have drifted from the configuration
    defined in the document.

    .EXAMPLE
    Import-VmsHardwareExcel -Path ~\Desktop\hardware.xlsx -UpdateExisting -Verbose

    Imports the hardware.xlsx file on the current user's desktop. If any cameras
    in the Excel document are already added, they will be updated to reflect the
    configuration defined in the document.

    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $Path,

        [Parameter()]
        [VideoOS.Platform.ConfigurationItems.RecordingServer]
        $RecordingServer,

        [Parameter()]
        [pscredential[]]
        $Credential,

        [Parameter()]
        [switch]
        $UpdateExisting
    )

    begin {
        if ($null -eq (Get-VmsManagementServer -ErrorAction 'SilentlyContinue')) {
            Connect-Vms -ShowDialog -AcceptEula -ErrorAction Stop
        }
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $Path = Show-FileDialog -OpenFile
        }
        try {
            $excelPackage = Open-ExcelPackage -Path $Path
            $worksheets = $excelPackage.Workbook.Worksheets.Name
            $data = @{
                Hardware                  = [list[pscustomobject]]::new()
                HardwareGeneralSettings   = [list[pscustomobject]]::new()
                HardwarePtzSettings       = [list[pscustomobject]]::new()
                HardwareEvents            = [list[pscustomobject]]::new()
                Cameras                   = [list[pscustomobject]]::new()
                CameraGeneralSettings     = [list[pscustomobject]]::new()
                CameraStreams             = [list[pscustomobject]]::new()
                CameraStreamSettings      = [list[pscustomobject]]::new()
                CameraPtzPresets          = [list[pscustomobject]]::new()
                CameraPtzPatrols          = [list[pscustomobject]]::new()
                CameraPtzPatrolPresets    = [list[pscustomobject]]::new()
                CameraRelatedDevices      = [list[pscustomobject]]::new()
                CameraEvents              = [list[pscustomobject]]::new()
                CameraGroups              = [list[pscustomobject]]::new()
                Microphones               = [list[pscustomobject]]::new()
                MicrophoneGeneralSettings = [list[pscustomobject]]::new()
                MicrophoneStreamSettings  = [list[pscustomobject]]::new()
                MicrophoneEvents          = [list[pscustomobject]]::new()
                MicrophoneGroups          = [list[pscustomobject]]::new()
                Speakers                  = [list[pscustomobject]]::new()
                SpeakerGeneralSettings    = [list[pscustomobject]]::new()
                SpeakerEvents             = [list[pscustomobject]]::new()
                SpeakerGroups             = [list[pscustomobject]]::new()
                Metadata                  = [list[pscustomobject]]::new()
                MetadataGeneralSettings   = [list[pscustomobject]]::new()
                MetadataGroups            = [list[pscustomobject]]::new()
                Inputs                    = [list[pscustomobject]]::new()
                InputGeneralSettings      = [list[pscustomobject]]::new()
                InputEvents               = [list[pscustomobject]]::new()
                InputGroups               = [list[pscustomobject]]::new()
                Outputs                   = [list[pscustomobject]]::new()
                OutputGeneralSettings     = [list[pscustomobject]]::new()
                OutputGroups              = [list[pscustomobject]]::new()
            }
            foreach ($key in $data.Keys) {
                if ($key -in $worksheets) {
                    if ($excelPackage.Workbook.Worksheets[$key].GetValue(1, 1)) {
                        Import-Excel -ExcelPackage $excelPackage -WorksheetName $key | ForEach-Object {
                            if ($null -ne $_.RecordingServer -and $PSBoundParameters.ContainsKey('RecordingServer')) {
                                $_.RecordingServer = $RecordingServer.Name
                            }
                            $data[$key].Add($_)
                        }
                    } else {
                        Write-Verbose "Ignoring worksheet '$key' because the value at 1,1 is null."
                    }
                }
            }
        } finally {
            if ($excelPackage) {
                $excelPackage | Close-ExcelPackage -NoSave
            }
        }
    }

    process {
        if ($data.Hardware.Count -eq 0) {
            Write-Error 'No hardware entries found in the Hardware worksheet.'
            return
        }

        $totalRows = $data.Hardware.Count
        $processedRows = 0
        $progressParams = @{
            Activity         = 'Importing hardware configuration from {0}' -f $Path
            Id               = 42
            PercentComplete  = 0
            CurrentOperation = 'Preparing'
        }
        Write-Progress @progressParams

        $recorders = @{}
        $existingHardware = @{}
        foreach ($rec in Get-VmsRecordingServer) {
            $recorders[$rec.Name] = $rec
            $existingHardware[$rec.Name] = @{}
            foreach ($hw in $rec | Get-VmsHardware) {
                if ($uri = $hw.Address -as [uri]) {
                    $hostAndPort = $uri.GetComponents([UriComponents]::HostAndPort, [uriformat]::Unescaped)
                    $existingHardware[$rec.Name][$hostAndPort] = $hw
                }
            }
        }

        foreach ($row in $data.Hardware | Sort-Object RecordingServer) {
            $progressParams.PercentComplete = [math]::Round(($processedRows++) / $totalRows * 100)
            $progressParams.CurrentOperation = '{0} ({1})' -f $row.Name, $row.Address
            Write-Progress @progressParams
            try {
                $recorder = if ($row.RecordingServer) { $recorders[$row.RecordingServer] } else { $null }
                if ($null -eq $recorder) {
                    Write-Warning "Recording server '$($row.RecordingServer)' not found. Skipping hardware '$($row.Name)' ($($row.Address))."
                    continue
                }


                $params = @{
                    HardwareAddress = $row.Address -as [uri]
                    Credential      = [collections.generic.list[pscredential]]::new()
                    DriverNumber    = $row.DriverNumber -as [int]
                    RecordingServer = $recorder
                    ErrorAction     = 'Stop'
                }

                if ($row.UserName -and $row.Password) {
                    $params.Credential.Add([pscredential]::new($row.UserName, ($row.Password | ConvertTo-SecureString -AsPlainText -Force)))
                }
                foreach ($pscredential in $Credential){
                    $params.Credential.Add($pscredential)
                }

                if (-not $params.HardwareAddress -or -not $params.HardwareAddress.IsAbsoluteUri) {
                    Write-Warning "Hardware '$($row.Name)' must have a valid address in the Address column. The value '$($row.Address)' is not a valid absolute URI. Example: http://192.168.1.101"
                    continue
                }

                $hostAndPort = $params.HardwareAddress.GetComponents([UriComponents]::HostAndPort, [uriformat]::Unescaped)
                if (($hardware = $existingHardware[$row.RecordingServer][$hostAndPort])) {
                    if (-not $UpdateExisting) {
                        Write-Verbose "Skipping the hardware at $($params.HardwareAddress) because it is already added to $($recorder.Name). To Update existing hardware/devices, use the 'UpdateExisting' switch."
                        continue
                    }
                } else {
                    if (-not $params.DriverNumber) {
                        $scanParams = @{
                            RecordingServer = $recorder
                            Address         = $params.HardwareAddress
                        }
                        if ($row.DriverGroup) {
                            $scanParams.DriverFamily = $row.DriverGroup
                        }
                        if ($params.Credential) {
                            $scanParams.Credential = $params.Credential
                        } else {
                            $scanParams.UseDefaultCredentials
                        }
                        Write-Verbose "Scanning hardware at $($row.Address) for driver discovery"
                        $scans = Start-VmsHardwareScan @scanParams
                        $scan = if ($null -eq ($scans | Where-Object HardwareScanValidated)) {
                            $scans | Select-Object -Last 1
                        } else {
                            $scans | Where-Object HardwareScanValidated | Select-Object -First 1
                        }
                        if ($scan.HardwareScanValidated) {
                            $params.Remove('DriverNumber')
                            $params.HardwareDriverPath = $scan.HardwareDriverPath
                            $params.Credential = [pscredential]::new($scan.UserName, ($scan.Password | ConvertTo-SecureString -AsPlainText -Force))
                        } else {
                            Write-Error -Message "Hardware scan failed for '$($row.Name)' ($($params.HardwareAddress)). ErrorText: '$($scan.ErrorText)'" -TargetObject $scan
                            continue
                        }
                    }
                    $credentials = $params.Credential
                    foreach ($hwCredential in $credentials) {
                        try {
                            $params.Credential = $hwCredential
                            $hardware = Add-VmsHardware @params
                        } catch {
                            Write-Error -ErrorRecord $_
                        }
                    }
                }
                $setHwParams = @{
                    Name        = if ($row.Name) { $row.Name } else { $hardware.Name }
                    Enabled     = $script:Truthy.IsMatch($row.Enabled)
                    Description = $row.Description
                    Verbose     = $VerbosePreference
                }
                $hardware | Set-VmsHardware @setHwParams

                $settings = $data.HardwareGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($settings) {
                    Import-GeneralSettingList -Device $hardware -Settings $settings
                }

                $settings = $data.HardwarePtzSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name } | Sort-Object Channel
                if ($settings) {
                    if ($hardware.HardwarePtzSettingsFolder.HardwarePtzSettings.Count) {
                        try {
                            $ptzSettings = $hardware.HardwarePtzSettingsFolder.HardwarePtzSettings
                            $channel = 0
                            foreach ($ptzChannel in $ptzSettings.HardwarePtzDeviceSettingChildItems) {
                                if ($settings.Count -lt ($channel + 1)) {
                                    Write-Warning "No HardwarePtzSettings available for channel $channel"
                                    continue
                                }
                                'PTZEnabled', 'PTZDeviceID', 'PTZCOMPort', 'PTZProtocol' | ForEach-Object {
                                    if ([string]::IsNullOrWhiteSpace($settings[$channel])) {
                                        Write-Warning "The supplied value for HardwarePtzSetting '$_' for $($hardware.Name) channel $channel is null or empty"
                                        return
                                    }
                                    if ($ptzChannel.Properties.GetValue($_) -cne $settings[$channel].$_) {
                                        $ptzChannel.Properties.SetValue($_, $settings[$channel].$_)
                                    }
                                }
                                
                                $channel += 1
                            }
                            $ptzSettings.Save()
                        } catch {
                            throw
                        }
                    } else {
                        Write-Warning "Unable to import HardwarePtzSettings for '$($hardware.Name)'. It may not be supported on the current VMS version or on this device."
                    }
                }

                $settings = $data.HardwareEvents | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($settings) {
                    Import-DeviceEventConfig -Device $hardware -Settings $settings
                }


                $cameraRows = $data.Cameras | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($cameraRows) {
                    Write-Verbose "Updating camera properties for $($hardware.Name)"
                    $hardware | Get-VmsCamera -EnableFilter All | Where-Object Channel -In $cameraRows.Channel | ForEach-Object {
                        $camera = $_

                        Import-DevicePropertyList -Device $_ -Settings ($cameraRows | Where-Object Channel -EQ $_.Channel | Select-Object -First 1)

                        $generalSettings = $data.CameraGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $_ -Settings $generalSettings
                        }

                        $eventSettings = $data.CameraEvents | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name }
                        if ($eventSettings) {
                            Import-DeviceEventConfig -Device $camera -Settings $eventSettings
                        }

                        $data.CameraStreams | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name } | Sort-Object Channel | ForEach-Object {
                            $streamRow = $_
                            $stream = $camera | Get-VmsCameraStream -Name $streamRow.Name -ErrorAction SilentlyContinue
                            if ($stream) {
                                $streamParams = @{}
                                if (-not [string]::IsNullOrWhiteSpace($streamRow.DisplayName)) {
                                    $streamParams.DisplayName = $streamRow.DisplayName
                                }
                                if ($streamRow.LiveMode -in 'Always', 'Never', 'WhenNeeded') {
                                    $streamParams.LiveMode = $streamRow.LiveMode
                                }
                                if ($script:TruthyFalsey.IsMatch($streamRow.LiveDefault)) {
                                    $streamParams.LiveDefault = $script:Truthy.IsMatch($streamRow.LiveDefault)
                                }
                                if ($script:TruthyFalsey.IsMatch($streamRow.PlaybackDefault)) {
                                    if (Test-VmsLicensedFeature -Name MultistreamRecording) {
                                        $streamParams.PlaybackDefault = $script:Truthy.IsMatch($streamRow.PlaybackDefault)
                                    } else {
                                        Write-Verbose "PlaybackDefault cannot be set because your VMS version does not include the MultistreamRecording feature."
                                    }
                                }
                                if ($streamRow.RecordingTrack -in 'Primary', 'Secondary', 'None') {
                                    $streamParams.RecordingTrack = $streamRow.RecordingTrack
                                }
                                if ($script:TruthyFalsey.IsMatch($streamRow.UseEdge)) {
                                    $streamParams.UseEdge = $script:Truthy.IsMatch($streamRow.UseEdge)
                                }
                                if ($streamParams.Count -gt 0) {
                                    $streamParams.Verbose = $VerbosePreference
                                    $stream | Set-VmsCameraStream @streamParams
                                }
                            } else {
                                Write-Warning "No stream found on $($camera.Name) with the name '$($streamRow.Name)'"
                            }
                        }

                        $data.CameraStreamSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name -and $_.Setting -and $_.Value } | Group-Object Stream | ForEach-Object {
                            $streamName = $_.Name
                            $streamSettings = @{}
                            $_.Group | ForEach-Object { $streamSettings[$_.Setting] = $_.Value }
                            $stream = $camera | Get-VmsCameraStream -Name $streamName -ErrorAction Ignore
                            if ($stream) {
                                $stream | Set-VmsCameraStream -Settings $streamSettings -Verbose:($VerbosePreference -eq 'Continue' )
                            } else {
                                Write-Warning "No stream found on $($camera.Name) with the name '$($streamRow.Name)'"
                            }
                        }

                        $data.CameraPtzPresets | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name } | ForEach-Object {
                            $ptzPresetRow = $_
                            if ($ptzPresetRow.Name -notin $camera.PtzPresetFolder.PtzPresets.Name) {
                                $newPtzPreset = $camera.PtzPresetFolder.AddPtzPreset($ptzPresetRow.Name, $ptzPresetRow.Description, $ptzPresetRow.Pan, $ptzPresetRow.Tilt, $ptzPresetRow.Zoom)
                                if ($ptzPresetRow.DefaultPreset -eq $true) {
                                    $null = $camera.PtzPresetFolder.DefaultPtzPreset($newPtzPreset.Path)
                                }
                            }
                        }

                        $data.CameraPtzPatrols | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name } | ForEach-Object {
                            $ptzPatrolRow = $_
                            if ($ptzPatrolRow.Name -notin $camera.PatrollingProfileFolder.PatrollingProfiles.Name) {
                                $endPresetId = ($camera.PtzPresetFolder.PtzPresets | Where-Object { $_.Name -eq $ptzPatrolRow.EndPresetName }).Id
                                $newPtzPatrol = $camera.PatrollingProfileFolder.AddPatrollingProfile($ptzPatrolRow.Name, $ptzPatrolRow.Description, $ptzPatrolRow.CustomizeTransitions, $ptzPatrolRow.InitSpeed, $ptzPatrolRow.InitTransitionTime, $endPresetId, $ptzPatrolRow.EndSpeed, $ptzPatrolRow.EndTransitionTime)
                                $newPtzPatrol = $camera.PatrollingProfileFolder.PatrollingProfiles | Where-Object { $_.Path -eq $newPtzPatrol.Path }

                                $index = 0
                                $data.CameraPtzPatrolPresets | Where-Object { $_.Patrol -eq $newPtzPatrol.Name } | ForEach-Object {
                                    $ptzPatrolPresetRow = $_
                                    $patrolPresetId = ($camera.PtzPresetFolder.PtzPresets | Where-Object { $_.Name -eq $ptzPatrolPresetRow.PresetName }).Id
                                    $null = $newPtzPatrol.AddPatrollingEntry($ptzPatrolPresetRow.Order, $patrolPresetId, $ptzPatrolPresetRow.WaitTime)

                                    $patrol = Get-ConfigurationItem -Path "PatrollingProfile[$($newPtzPatrol.Id)]"
                                    if ($newPtzPatrol.CustomizeTransitions) {
                                        ($patrol.Children[$index].Properties | Where-Object { $_.Key -eq 'Speed' }).Value = $ptzPatrolPresetRow.Speed
                                        ($patrol.Children[$index].Properties | Where-Object { $_.Key -eq 'TransitionTime' }).Value = $ptzPatrolPresetRow.TransitionTime
                                    }
                                    ### TODO: Refactor this section so Set-ConfigurationItem only needs to be called after the entire Patrol object has been updated
                                    $null = Set-ConfigurationItem -ConfigurationItem $patrol
                                    $index++
                                }
                            }
                        }

                        $data.CameraGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($camera.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Camera -Path $_.Group | Add-VmsDeviceGroupMember -Device $camera -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No cameras to configure for $($hardware.Name)"
                }

                $rows = $data.Microphones | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($rows) {
                    Write-Verbose "Updating microphone properties for $($hardware.Name)"
                    $hardware | Get-VmsMicrophone -EnableFilter All | Where-Object Channel -In $rows.Channel | ForEach-Object {
                        $device = $_

                        Import-DevicePropertyList -Device $device -Settings ($rows | Where-Object Channel -EQ $device.Channel | Select-Object -First 1)

                        $generalSettings = $data.MicrophoneGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Microphone -eq $device.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $device -Settings $generalSettings
                        }

                        $eventSettings = $data.MicrophoneEvents | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Microphone -eq $device.Name }
                        if ($eventSettings) {
                            Import-DeviceEventConfig -Device $device -Settings $eventSettings
                        }

                        $data.MicrophoneGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Microphone -eq $device.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($device.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Microphone -Path $_.Group | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No microphones to configure for $($hardware.Name)"
                }


                $rows = $data.Speakers | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($rows) {
                    Write-Verbose "Updating speaker properties for $($hardware.Name)"
                    $hardware | Get-VmsSpeaker -EnableFilter All | Where-Object Channel -In $rows.Channel | ForEach-Object {
                        $device = $_

                        Import-DevicePropertyList -Device $device -Settings ($rows | Where-Object Channel -EQ $device.Channel | Select-Object -First 1)

                        $generalSettings = $data.SpeakerGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Speaker -eq $device.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $device -Settings $generalSettings
                        }

                        $data.SpeakerGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Speaker -eq $device.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($device.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Speaker -Path $_.Group | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No speakers to configure for $($hardware.Name)"
                }


                $rows = $data.Metadata | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($rows) {
                    Write-Verbose "Updating metadata properties for $($hardware.Name)"
                    $hardware | Get-VmsMetadata -EnableFilter All | Where-Object Channel -In $rows.Channel | ForEach-Object {
                        $device = $_

                        Import-DevicePropertyList -Device $device -Settings ($rows | Where-Object Channel -EQ $device.Channel | Select-Object -First 1)

                        $generalSettings = $data.MetadataGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Metadata -eq $device.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $device -Settings $generalSettings
                        }

                        $data.MetadataGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Metadata -eq $device.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($device.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Metadata -Path $_.Group | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No metadata to configure for $($hardware.Name)"
                }


                $rows = $data.Inputs | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($rows) {
                    Write-Verbose "Updating IO input properties for $($hardware.Name)"
                    $hardware | Get-VmsInput -EnableFilter All | Where-Object Channel -In $rows.Channel | ForEach-Object {
                        $device = $_

                        Import-DevicePropertyList -Device $device -Settings ($rows | Where-Object Channel -EQ $device.Channel | Select-Object -First 1)

                        $generalSettings = $data.InputGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.InputEvent -eq $device.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $device -Settings $generalSettings
                        }

                        $data.InputGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Input -eq $device.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($device.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Input -Path $_.Group | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No inputs to configure for $($hardware.Name)"
                }

                $rows = $data.Outputs | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
                if ($rows) {
                    Write-Verbose "Updating IO output properties for $($hardware.Name)"
                    $hardware | Get-VmsOutput -EnableFilter All | Where-Object Channel -In $rows.Channel | ForEach-Object {
                        $device = $_

                        Import-DevicePropertyList -Device $device -Settings ($rows | Where-Object Channel -EQ $device.Channel | Select-Object -First 1)

                        $generalSettings = $data.OutputGeneralSettings | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Output -eq $device.Name }
                        if ($generalSettings) {
                            Import-GeneralSettingList -Device $device -Settings $generalSettings
                        }

                        $data.OutputGroups | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Output -eq $device.Name } | ForEach-Object {
                            # Device may already be added to the destination device group. If so, SilentlyContinue will hide the ArgumentMIPException error.
                            Write-Verbose "Adding $($device.Name) to device group $($_.Group)"
                            New-VmsDeviceGroup -Type Output -Path $_.Group | Add-VmsDeviceGroupMember -Device $device -ErrorAction SilentlyContinue
                        }
                    }
                } else {
                    Write-Verbose "No outputs to configure for $($hardware.Name)"
                }
            } catch {
                Write-Error -ErrorRecord $_
            }
        }
        $progressParams.PercentComplete = 100
        $progressParams.Completed = $true
        Write-Progress @progressParams

        Clear-VmsCache

        $totalRows = $data.Hardware.Count
        $processedRows = 0
        $progressParams = @{
            Activity        = 'Configuring related devices'
            Id              = 43
            PercentComplete = 0
        }
        Write-Progress @progressParams

        foreach ($row in $data.Hardware | Sort-Object RecordingServer) {
            $progressParams.PercentComplete = [math]::Round(($processedRows++) / $totalRows * 100)
            $progressParams.CurrentOperation = '{0} ({1})' -f $row.Name, $row.Address
            Write-Progress @progressParams

            $recorder = if ($row.RecordingServer) { $recorders[$row.RecordingServer] } else { $null }
            if ($null -eq $recorder) {
                continue
            }

            $params = @{
                Name            = $row.Name
                HardwareAddress = $row.Address -as [uri]
                RecordingServer = $recorder
                ErrorAction     = 'Stop'
            }

            if ([string]::IsNullOrWhiteSpace($params.Name)) {
                $params.Remove('Name')
            }

            $hostAndPort = $params.HardwareAddress.GetComponents([UriComponents]::HostAndPort, [uriformat]::Unescaped)
            if (($hardware = $existingHardware[$row.RecordingServer][$hostAndPort])) {
                if (-not $UpdateExisting) {
                    Write-Verbose "Skipping the hardware at $($params.HardwareAddress) because it is already added to $($recorder.Name). To Update existing hardware/devices, use the 'UpdateExisting' switch."
                    continue
                }
            }

            $hardware = Get-VmsHardware -Name $row.Name
            $cameraRows = $data.Cameras | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name }
            if ($cameraRows) {
                foreach ($camera in $hardware | Get-VmsCamera -EnableFilter All) {
                    $relatedDevicesString = $null
                    $data.CameraRelatedDevices | Where-Object { $_.RecordingServer -eq $recorder.Name -and $_.Hardware -eq $hardware.Name -and $_.Camera -eq $camera.Name -and $_.Channel -eq $camera.Channel } | ForEach-Object {
                        $relatedDevicesRow = $_
                        if ([string]::IsNullOrEmpty($relatedDevicesString)) {
                            $relatedRec = $recorders[$relatedDevicesRow.RelatedRecordingServerName]
                            $relatedHW = Get-VmsHardware -RecordingServer $relatedRec | Where-Object Address -EQ $relatedDevicesRow.RelatedHardwareAddress
                            switch ($relatedDevicesRow.RelatedDeviceType) {
                                Metadata { $relatedDeviceItem = Get-VmsMetadata -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                                Microphone { $relatedDeviceItem = Get-VmsMicrophone -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                                Speaker { $relatedDeviceItem = Get-VmsSpeaker -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                            }
                            [string]$relatedDevicesString = $relatedDeviceItem.Path
                        } else {
                            $relatedRec = $recorders[$relatedDevicesRow.RelatedRecordingServerName]
                            $relatedHW = Get-VmsHardware -RecordingServer $relatedRec | Where-Object Address -EQ $relatedDevicesRow.RelatedHardwareAddress
                            switch ($relatedDevicesRow.RelatedDeviceType) {
                                Metadata { $relatedDeviceItem = Get-VmsMetadata -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                                Microphone { $relatedDeviceItem = Get-VmsMicrophone -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                                Speaker { $relatedDeviceItem = Get-VmsSpeaker -EnableFilter All -Hardware $relatedHW -Channel $relatedDevicesRow.Channel }
                            }
                            [string]$relatedDevicesString += ",$($relatedDeviceItem.Path)"
                        }
                    }
                    $clientSettingsItem = $camera.ClientSettingsFolder.ClientSettings[0]
                    $clientSettingsItem.Related = $relatedDevicesString
                    $clientSettingsItem.Save()
                }
            }
        }
        $progressParams.PercentComplete = 100
        $progressParams.Completed = $true
        Write-Progress @progressParams
    }
}

function New-CameraViewItemDefinition {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VmsCameraViewItemProperties]
        $Properties
    )

    process {
        $template = @"
<viewitem id="{0}" displayname="Camera ViewItem" shortcut="{1}" type="VideoOS.RemoteClient.Application.Data.ContentTypes.CameraContentType.CameraViewItem, VideoOS.RemoteClient.Application" smartClientId="{2}">
    <iteminfo cameraid="{3}" lastknowncameradisplayname="{4}" livestreamid="{5}" imagequality="{6}" framerate="{7}" maintainimageaspectratio="{8}" usedefaultdisplaysettings="{9}" showtitlebar="{10}" keepimagequalitywhenmaximized="{11}" updateonmotiononly="{12}" soundonmotion="{13}" soundonevent="{14}" smartsearchgridwidth="{15}" smartsearchgridheight="{16}" smartsearchgridmask="{17}" pointandclickmode="{18}" usingproperties="True" />
    <properties>
        <property name="cameraid" value="{3}" />
        <property name="livestreamid" value="{5}" />
        <property name="framerate" value="{7}" />
        <property name="imagequality" value="{6}" />
        <property name="lastknowncameradisplayname" value="{4}" />
    </properties>
</viewitem>
"@
        $soundOnMotion = if ($Properties.SoundOnMotion) { 1 } else { 0 }
        $soundOnEvent  = if ($Properties.SoundOnEvent)  { 1 } else { 0 }
        $values = @(
            $Properties.Id,
            $Properties.Shortcut,
            $Properties.SmartClientId,
            $Properties.CameraId,
            $Properties.CameraName,
            $Properties.LiveStreamId,
            $Properties.ImageQuality,
            $Properties.Framerate,
            $Properties.MaintainImageAspectRatio,
            $Properties.UseDefaultDisplaySettings,
            $Properties.ShowTitleBar,
            $Properties.KeepImageQualityWhenMaximized,
            $Properties.UpdateOnMotionOnly,
            $soundOnMotion,
            $soundOnEvent,
            $Properties.SmartSearchGridWidth,
            $Properties.SmartSearchGridHeight,
            $Properties.SmartSearchGridMask,
            $Properties.PointAndClickMode
        )
        Write-Output ($template -f $values)
    }
}


function New-VmsViewItemProperties {
    [CmdletBinding()]
    [OutputType([VmsCameraViewItemProperties])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Private function.')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('CameraId')]
        [guid]
        $Id,

        [Parameter()]
        [guid]
        $SmartClientId
    )

    process {
        $properties = [VmsCameraViewItemProperties]::new()
        $properties.CameraName = $Name
        $properties.CameraId = $Id
        if ($MyInvocation.BoundParameters.ContainsKey('SmartClientId')) {
            $properties.SmartClientId = $SmartClientId
        }
        Write-Output $properties
    }
}



function New-VmsViewLayout {
    [CmdletBinding(DefaultParameterSetName = 'Simple')]
    [OutputType([string])]
    param (
        [Parameter(ParameterSetName = 'Simple')]
        [ValidateRange(0, 100)]
        [int]
        $ViewItemCount = 1,

        [Parameter(ParameterSetName = 'Custom')]
        [ValidateRange(1, 100)]
        [int]
        $Columns,

        [Parameter(ParameterSetName = 'Custom')]
        [ValidateRange(1, 100)]
        [int]
        $Rows
    )

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'Simple' {
                $size = 1
                if ($ViewItemCount -gt 0) {
                    $sqrt = [math]::Sqrt($ViewItemCount)
                    $size = [math]::Floor($sqrt)
                    if ($sqrt % 1) {
                        $size++
                    }
                }
                $Columns = $Rows = $size
                $width = $height = [math]::Floor(1000 / $size)
            }

            'Custom' {
                $width = [math]::Floor(1000 / $Columns)
                $height = [math]::Floor(1000 / $Rows)
            }
        }

        $template = '<ViewItem><Position><X>{0}</X><Y>{1}</Y></Position><Size><Width>{2}</Width><Height>{3}</Height></Size></ViewItem>'
        $xmlBuilder = [text.stringbuilder]::new()
        $null = $xmlBuilder.Append("<ViewItems>")
        for ($posY = 0; $posY -lt $Rows; $posY++) {
            for ($posX = 0; $posX -lt $Columns; $posX++) {
                $x = $width  * $posX
                $y = $height * $posY
                $null = $xmlBuilder.Append(($template -f $x, $y, $width, $height))
            }
        }
        $null = $xmlBuilder.Append("</ViewItems>")
        Write-Output $xmlBuilder.ToString()
    }
}


function NewVmsAppDataPath {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    process {
        $appDataRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'MilestonePSTools\'
        (New-Item -Path $appDataRoot -ItemType Directory -Force).FullName
    }
}

function OwnerInfoPropertyCompleter {
    param (
        $commandName,
        $parameterName,
        $wordToComplete,
        $commandAst,
        $fakeBoundParameters
    )

    $ownerPath = 'BasicOwnerInformation[{0}]' -f (Get-VmsManagementServer).Id
    $ownerInfo = Get-ConfigurationItem -Path $ownerPath
    $invokeInfo = $ownerInfo | Invoke-Method -MethodId AddBasicOwnerInfo
    $tagTypeInfo = $invokeInfo.Properties | Where-Object Key -eq 'TagType'
    $tagTypeInfo.ValueTypeInfos.Value | ForEach-Object { $_ }
}


function Set-CertKeyPermission {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        # Specifies the certificate store path to locate the certificate specified in Thumbprint. Example: Cert:\LocalMachine\My
        [Parameter()]
        [string]
        $CertificateStore = 'Cert:\LocalMachine\My',

        # Specifies the thumbprint of the certificate to which private key access should be updated.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Thumbprint,

        # Specifies the Windows username for the identity to which permissions should be granted.
        [Parameter(Mandatory)]
        [string]
        $UserName,

        # Specifies the level of access to grant to the private key.
        [Parameter()]
        [ValidateSet('Read', 'FullControl')]
        [string]
        $Permission = 'Read',

        # Specifies the access type for the Access Control List rule.
        [Parameter()]
        [ValidateSet('Allow', 'Deny')]
        [string]
        $PermissionType = 'Allow'
    )

    process {
        <#
            There is a LOT of error checking in this function as it seems that certificates are not
            always consistently storing their private keys in predictable places. I've found private
            keys for RSA certs in ProgramData\Microsoft\Crypto\Keys instead of
            ProgramData\Microsoft\Crypto\RSA\MachineKeys, I've seen the UniqueName property contain
            a value representing the file name of the certificate private key file somewhere in the
            ProgramData\Microsoft\Crypto folder, and I've seen the UniqueName property contain a
            full file path to the private key file. I've also found that some RSA certs require you
            to use the RSA extension method to retrieve the private key, even though it seems like
            you should expect to find it in the PrivateKey property when retrieving the certificate
            from Get-ChildItem Cert:\LocalMachine\My.
        #>

        $certificate = Get-ChildItem -Path $CertificateStore | Where-Object Thumbprint -eq $Thumbprint
        Write-Verbose "Processing certificate for $($certificate.Subject) with thumbprint $($certificate.Thumbprint)"
        if ($null -eq $certificate) {
            Write-Error "Certificate not found in certificate store '$CertificateStore' matching thumbprint '$Thumbprint'"
            return
        }
        if (-not $certificate.HasPrivateKey) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a private key attached."
            return
        }
        $privateKey = $null
        switch ($certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName) {
            'RSA' {
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
            }

            'ECC' {
                $privateKey = [System.Security.Cryptography.X509Certificates.ECDsaCertificateExtensions]::GetECDsaPrivateKey($certificate)
            }

            'DSA' {
                Write-Error "Use of DSA-based certificates is not recommended, and not supported by this command. See https://docs.microsoft.com/en-us/dotnet/api/system.security.cryptography.dsa?view=net-5.0"
                return
            }

            Default { Write-Error "`$certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName was '$($certificate.PublicKey.EncodedKeyValue.Oid.FriendlyName)'. Expected RSA, DSA or ECC."; return }
        }
        if ($null -eq $privateKey) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a private key attached."
            return
        }
        if ([string]::IsNullOrWhiteSpace($privateKey.Key.UniqueName)) {
            Write-Error "Certificate with friendly name '$($certificate.FriendlyName)' issued to subject '$($certificate.Subject)' does not have a value for the private key's UniqueName property so we cannot find the file on the filesystem associated with the private key."
            return
        }

        if (Test-Path -LiteralPath $privateKey.Key.UniqueName) {
            $privateKeyFile = Get-Item -Path $privateKey.Key.UniqueName
        }
        else {
            $privateKeyFile = Get-ChildItem -Path (Join-Path -Path ([system.environment]::GetFolderPath([system.environment+specialfolder]::CommonApplicationData)) -ChildPath ([io.path]::combine('Microsoft', 'Crypto'))) -Filter $privateKey.Key.UniqueName -Recurse -ErrorAction Ignore
            if ($null -eq $privateKeyFile) {
                Write-Error "No private key file found matching UniqueName '$($privateKey.Key.UniqueName)'"
                return
            }
            if ($privateKeyFile.Count -gt 1) {
                Write-Error "Found more than one private key file matching UniqueName '$($privateKey.Key.UniqueName)'"
                return
            }
        }

        $privateKeyPath = $privateKeyFile.FullName
        if (-not (Test-Path -Path $privateKeyPath)) {
            Write-Error "Expected to find private key file at '$privateKeyPath' but the file does not exist. You may need to re-install the certificate in the certificate store"
            return
        }

        $acl = Get-Acl -Path $privateKeyPath
        $rule = [Security.AccessControl.FileSystemAccessRule]::new($UserName, $Permission, $PermissionType)
        $acl.AddAccessRule($rule)
        if ($PSCmdlet.ShouldProcess($privateKeyPath, "Add FileSystemAccessRule")) {
            $acl | Set-Acl -Path $privateKeyPath
        }
    }
}


function Show-DeprecationWarning {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [System.Management.Automation.InvocationInfo]
        $InvocationInfo
    )

    process {
        $oldName = $InvocationInfo.InvocationName
        if ($script:Deprecations.ContainsKey($oldName)) {
            $newName = $script:deprecations[$oldName]
            Write-Warning "The '$oldName' cmdlet is deprecated. To minimize the risk of being impacted by a breaking change in the future, please use '$newName' instead."
            $script:Deprecations.Remove($oldName)
        }
    }
}


function ValidateHardwareCsvRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]
        $Rows
    )

    process {
        $ErrorActionPreference = 'Stop'
        $defaultValues = @{
            DeviceType      = 'Camera'
            Name            = $null
            Address         = $null
            Channel         = 0
            UserName        = $null
            Password        = $null
            RecordingServer = $null
            DriverNumber    = 0
            DriverGroup     = $null
            Enabled         = $true
            StorageName     = $null
            HardwareName    = $null
            Coordinates     = $null
            DeviceGroups    = '/Imported from CSV'
            Path            = $null
            Result          = [string]::Empty
        }

        $supportedValues = @{
            DeviceType = @('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output')
        }

        for ($i = 0; $i -lt $Rows.Count; $i++) {
            $row = $Rows[$i]

            $record = [pscustomobject]@{
                Row             = $i + 1
                DeviceType      = $defaultValues['DeviceType']
                Name            = $defaultValues['Name']
                Address         = $defaultValues['Address']
                Channel         = $defaultValues['Channel']
                UserName        = $defaultValues['UserName']
                Password        = $defaultValues['Password']
                RecordingServer = $defaultValues['RecordingServer']
                DriverNumber    = $defaultValues['DriverNumber']
                DriverGroup    = $defaultValues['DriverGroup']
                Enabled         = $defaultValues['Enabled']
                StorageName     = $defaultValues['StorageName']
                HardwareName    = $defaultValues['HardwareName']
                Coordinates     = $defaultValues['Coordinates']
                DeviceGroups    = $defaultValues['DeviceGroups']
                Path            = $defaultValues['Path']
                Result          = $defaultValues['Result']
            }
            
            $headersProvided = ($row | Get-Member -MemberType NoteProperty).Name
            foreach ($property in $headersProvided) {
                if (-not $defaultValues.ContainsKey($property)) {
                    Write-Warning "Ignoring unsupported header `"$property`""
                    continue
                }
                if ($property -in @('Path', 'Result')) {
                    continue
                }
                $record.$property = $row.$property
            }
            $recorders = @{}
            $storages = @{}
            $driversByRecorder = @{}
            foreach ($recorder in Get-VmsRecordingServer) {
                $recorders[$recorder.Name] = $recorder
                foreach ($storage in $recorder | Get-VmsStorage) {
                    $storages["$($recorder.Name).$($storage.Name)"] = $null
                }
                foreach ($driver in $recorder | Get-VmsHardwareDriver) {
                    $driversByRecorder["$($recorder.Name).$($driver.Number)"] = $null
                }
            }
            foreach ($property in ($record | Get-Member -MemberType NoteProperty).Name) {
                switch ($property) {
                    'DeviceType' {
                        if ($record.DeviceType -notin $supportedValues.DeviceType) {
                            Write-Error -Message "Invalid DeviceType value `"$($row.DeviceType)`" in row $($i + 1). Supported values are $($supportedValues.DeviceType -join ', ')." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                    }
                    'Name' {}
                    'Address' {
                        try {
                            $record.Address = ([uribuilder]$record.Address).Uri.GetComponents([uricomponents]::SchemeAndServer, [uriformat]::SafeUnescaped)
                            if ($record.Address -notmatch '^https?') {
                                throw [argumentexception]::new("Invalid address scheme. Supported schemes are http and https.")
                            }
                        } catch {
                            $errorParams = @{
                                Message      = 'Invalid Address value "{0}" in row {0}.' -f $row.Address, ($i + 1)
                                Category     = 'InvalidData'
                                ErrorId      = 'InvalidValue'
                                TargetObject = $row
                            }
                            if ($null -ne $_.Exception) {
                                $errorParams.Exception = $_.Exception
                            }
                            Write-Error @errorParams
                        }
                    }
                    'Channel' {
                        $channelNumber = 0
                        if (-not [int]::TryParse($record.Channel, [ref]$channelNumber)) {
                            Write-Error -Message "Invalid Channel value `"$($row.Channel)`" in row $($i + 1)." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                        $record.Channel = $channelNumber
                    }
                    'UserName' {}
                    'Password' {}
                    'RecordingServer' {
                        if (-not [string]::IsNullOrWhiteSpace($record.RecordingServer) -and -not $recorders.ContainsKey($record.RecordingServer)) {
                            Write-Error -Message "Invalid RecordingServer value `"$($row.Channel)`" in row $($i + 1)." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                    }
                    'DriverNumber' {
                        $driverNumber = 0
                        if ([string]::IsNullOrWhiteSpace($record.DriverNumber)) {
                            $record.DriverNumber = 0
                        }
                        if (-not [int]::TryParse($record.DriverNumber, [ref]$driverNumber)) {
                            Write-Error -Message "Invalid DriverNumber value `"$($row.DriverNumber)`" in row $($i + 1)." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                        if (-not [string]::IsNullOrWhiteSpace($record.RecordingServer) -and -not [string]::IsNullOrWhiteSpace($row.DriverNumber) -and -not $driversByRecorder.ContainsKey("$($record.RecordingServer).$($record.DriverNumber)")) {
                            Write-Error -Message "DriverNumber `"$($row.DriverNumber)`" in row $($i + 1) not found on RecordingServer `"$($record.RecordingServer)`". You may need to install a newer device pack version or custom device driver." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                        $record.DriverNumber = $driverNumber
                    }
                    'DriverGroup' {}
                    'Enabled' {
                        $enabled = $true
                        if (-not [bool]::TryParse($record.Enabled, [ref]$enabled)) {
                            Write-Error -Message "Invalid Enabled value `"$($row.Enabled)`" in row $($i + 1)." -Category InvalidData -ErrorId "InvalidValue" -TargetObject $row
                        }
                        $record.Enabled = $enabled
                    }
                    'StorageName' {}
                    'HardwareName' {}
                    'Coordinates' {
                        if (-not [string]::IsNullOrWhiteSpace($record.Coordinates)) {
                            try {
                                $null = ConvertTo-GisPoint -Coordinates $record.Coordinates
                            } catch {
                                $errorParams = @{
                                    Message      = 'Invalid Coordinates value "{0}" in row {0}.' -f $row.Coordinates, ($i + 1)
                                    Category     = 'InvalidData'
                                    ErrorId      = 'InvalidValue'
                                    TargetObject = $row
                                }
                                if ($null -ne $_.Exception) {
                                    $errorParams.Exception = $_.Exception
                                }
                                Write-Error @errorParams
                            }
                        }
                    }
                    'DeviceGroups' {}
                    'Row' {}
                    'Path' {}
                    'Result' {}
                    Default {
                        Write-Verbose "Ignoring header `"$_`""
                    }
                }
            }
            
            $record
        }
    }
}

function ValidateSiteInfoTagName {
    $ownerPath = 'BasicOwnerInformation[{0}]' -f (Get-VmsManagementServer).Id
    $ownerInfo = Get-ConfigurationItem -Path $ownerPath
    $invokeInfo = $ownerInfo | Invoke-Method -MethodId AddBasicOwnerInfo
    $tagTypeInfo = $invokeInfo.Properties | Where-Object Key -eq 'TagType'
    if ($_ -cin $tagTypeInfo.ValueTypeInfos.Value) {
        $true
    } else {
        throw "$_ is not a valid BasicOwnerInformation property key."
    }
}


function Add-VmsDeviceGroupMember {
    [CmdletBinding()]
    [Alias('Add-DeviceGroupMember')]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'SpeakerGroup', 'MetadataGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Group,

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'ByObject')]
        [ValidateVmsItemType('Camera', 'Microphone', 'Speaker', 'Metadata', 'InputEvent', 'Output')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Device,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'ById')]
        [Alias('Id')]
        [guid[]]
        $DeviceId
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $dirty = $false
        $groupItemType = ($Group | Split-VmsConfigItemPath -ItemType) -replace 'Group$', ''
        try {
            if ($Device) {
                $DeviceId = $Device.Id
            }
            foreach ($id in $DeviceId) {
                try {
                    $path = '{0}[{1}]' -f $groupItemType, $id
                    $null = $Group."$($groupItemType)Folder".AddDeviceGroupMember($path)
                    $dirty = $true
                } catch [VideoOS.Platform.ArgumentMIPException] {
                    Write-Error -Message "Failed to add device group member: $_.Exception.Message" -Exception $_.Exception
                }
            }
        }
        finally {
            if ($dirty) {
                $Group."$($groupItemType)GroupFolder".ClearChildrenCache()
                (Get-VmsManagementServer)."$($groupItemType)GroupFolder".ClearChildrenCache()
            }
        }
    }
}


function Add-VmsHardware {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.Hardware])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ParameterSetName = 'FromHardwareScan', Mandatory, ValueFromPipeline)]
        [VmsHardwareScanResult[]]
        $HardwareScan,

        [Parameter(ParameterSetName = 'Manual', Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer]
        $RecordingServer,

        [Parameter(ParameterSetName = 'Manual', Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Address')]
        [uri]
        $HardwareAddress,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(ParameterSetName = 'Manual')]
        [int]
        $DriverNumber,

        [Parameter(ParameterSetName = 'Manual', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]
        $HardwareDriverPath,

        [Parameter(ParameterSetName = 'Manual', Mandatory)]
        [pscredential]
        $Credential,

        [Parameter()]
        [switch]
        $SkipConfig,

        # Specifies that the hardware should be added, even if it already exists on another recording server.
        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $recorders = @{}
        $tasks = New-Object System.Collections.Generic.List[VideoOS.Platform.ConfigurationItems.ServerTask]
        switch ($PSCmdlet.ParameterSetName) {
            'Manual' {
                if ([string]::IsNullOrWhiteSpace($HardwareDriverPath)) {
                    if ($MyInvocation.BoundParameters.ContainsKey('DriverNumber')) {
                        $hardwareDriver = $RecordingServer.HardwareDriverFolder.HardwareDrivers | Where-Object Number -eq $DriverNumber
                        if ($null -ne $hardwareDriver) {
                            Write-Verbose "Mapped DriverNumber $DriverNumber to $($hardwareDriver.Name)"
                            $HardwareDriverPath = $hardwareDriver.Path
                        } else {
                            Write-Error "Failed to find hardware driver matching driver number $DriverNumber on Recording Server '$($RecordingServer.Name)'"
                            return
                        }
                    } else {
                        Write-Error "Add-VmsHardware cannot continue without either the HardwareDriverPath or the user-friendly driver number found in the supported hardware list."
                        return
                    }
                }
                $serverTask = $RecordingServer.AddHardware($HardwareAddress, $HardwareDriverPath, $Credential.UserName, $Credential.Password)
                $tasks.Add($serverTask)
                $recorders[$RecordingServer.Path] = $RecordingServer
            }
            'FromHardwareScan' {
                if ($HardwareScan.HardwareScanValidated -contains $false) {
                    Write-Warning "One or more scanned hardware could not be validated. These entries will be skipped."
                }
                if ($HardwareScan.MacAddressExistsLocal -contains $true) {
                    Write-Warning "One or more scanned hardware already exist on the target recording server. These entries will be skipped."
                }
                if ($HardwareScan.MacAddressExistsGlobal -contains $true -and -not $Force) {
                    Write-Warning "One or more scanned hardware already exist on another recording server. These entries will be skipped since the Force switch was not used."
                }
                foreach ($scan in $HardwareScan | Where-Object { $_.HardwareScanValidated -and -not $_.MacAddressExistsLocal }) {
                    if ($scan.MacAddressExistsGlobal -and -not $Force) {
                        continue
                    }
                    Write-Verbose "Adding $($scan.HardwareAddress) to $($scan.RecordingServer.Name) using driver identified by $($scan.HardwareDriverPath)"
                    $serverTask = $scan.RecordingServer.AddHardware($scan.HardwareAddress, $scan.HardwareDriverPath, $scan.UserName, $scan.Password)
                    $tasks.Add($serverTask)
                }
            }
        }
        if ($tasks.Count -eq 0) {
            return
        }
        Write-Verbose "Awaiting $($tasks.Count) AddHardware requests"
        Write-Verbose "Tasks: $([string]::Join(', ', $tasks.Path))"
        Wait-VmsTask -Path $tasks.Path -Title "Adding hardware to recording server(s) on site $((Get-VmsSite).Name)" -Cleanup | Foreach-Object {
            $vmsTask = [VmsTaskResult]$_
            if ($vmsTask.State -eq [VmsTaskState]::Success) {
                $hardwareId = $vmsTask | Split-VmsConfigItemPath -Id
                $newHardware = Get-VmsHardware -Id $hardwareId
                if ($null -eq $recorders[$newHardware.ParentItemPath]) {
                    Get-VmsRecordingServer | Where-Object Path -eq $newHardware.ParentItemPath | Foreach-Object {
                        $recorders[$_.Path] = $_
                    }
                }

                if (-not $SkipConfig) {
                    Set-NewHardwareConfig -Hardware $newHardware -Name $Name
                }
                if ($null -ne $newHardware) {
                    $newHardware
                }
            } else {
                Write-Error "Add-VmsHardware failed with error code $($vmsTask.ErrorCode). $($vmsTask.ErrorText)"
            }
        }

        $recorders.Values | Foreach-Object {
            $_.HardwareFolder.ClearChildrenCache()
        }
    }
}

function Set-NewHardwareConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [VideoOS.Platform.ConfigurationItems.Hardware]
        $Hardware,

        [Parameter()]
        [string]
        $Name
    )

    process {
        $systemInfo = [videoos.platform.configuration]::Instance.FindSystemInfo((Get-VmsSite).FQID.ServerId, $true)
        $version = $systemInfo.Properties.ProductVersion -as [version]
        $itemTypes = @('Camera')
        if (-not [string]::IsNullOrWhiteSpace($Name)) {
            $itemTypes += 'Microphone', 'Speaker', 'Metadata', 'InputEvent', 'Output'
        }
        if ($version -ge '20.2') {
            $Hardware.FillChildren($itemTypes)
        }

        $Hardware.Enabled = $true
        if (-not [string]::IsNullOrWhiteSpace($Name)) {
            $Hardware.Name = $Name
        }
        $Hardware.Save()

        foreach ($itemType in $itemTypes) {
            foreach ($item in $Hardware."$($itemType)Folder"."$($itemType)s") {
                if (-not [string]::IsNullOrWhiteSpace($Name)) {
                    $newName = '{0} - {1} {2}' -f $Name, $itemType.Replace('Event', ''), ($item.Channel + 1)
                    $item.Name = $newName
                }
                if ($itemType -eq 'Camera' -and $item.Channel -eq 0) {
                    $item.Enabled = $true
                }
                $item.Save()
            }
        }
    }
}

function Add-VmsLoginProviderClaim {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter(Mandatory)]
        [string[]]
        $Name,

        [Parameter()]
        [string[]]
        $DisplayName,

        [Parameter()]
        [switch]
        $CaseSensitive
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($DisplayName.Count -gt 0 -and $DisplayName.Count -ne $Name.Count) {
            Write-Error "Number of claim names does not match the number of display names. When providing display names for claims, the number of DisplayName values must match the number of Name values."
            return
        }
        try {
            for ($index = 0; $index -lt $Name.Count; $index++) {
                $claimName = $Name[$index]
                $claimDisplayName = $Name[$index]
                if ($DisplayName.Count -gt 0) {
                    $claimDisplayName = $DisplayName[$index]
                }
                if ($PSCmdlet.ShouldProcess("Login provider '$($LoginProvider.Name)'", "Add claim '$claimName'")) {
                    $null = $LoginProvider.RegisteredClaimFolder.AddRegisteredClaim($claimName, $claimDisplayName, $CaseSensitive)
                }
            }
        } catch {
            Write-Error -Message $_.Exception.Message -TargetObject $LoginProvider
        }
    }
}

function Clear-VmsSiteInfo {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    param (
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $ownerInfoFolder = (Get-VmsManagementServer).BasicOwnerInformationFolder
        $ownerInfoFolder.ClearChildrenCache()
        $ownerInfo = $ownerInfoFolder.BasicOwnerInformations[0]
        foreach ($key in $ownerInfo.Properties.KeysFullName) {
            if ($key -match '^\[(?<id>[a-fA-F0-9\-]{36})\]/(?<tagtype>[\w\.]+)$') {
                if ($PSCmdlet.ShouldProcess((Get-VmsSite).Name, "Remove $($Matches.tagtype) entry with value '$($ownerInfo.Properties.GetValue($key))' in site information")) {
                    $invokeResult = $ownerInfo.RemoveBasicOwnerInfo($Matches.id)
                    if ($invokeResult.State -ne 'Success') {
                        Write-Error "An error occurred while removing a site information property: $($invokeResult.ErrorText)"
                    }
                }
            } else {
                Write-Warning "Site information property key format unrecognized: $key"
            }
        }
    }
}


function Clear-VmsView {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.View])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 1)]
        [VideoOS.Platform.ConfigurationItems.View[]]
        $View,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($v in $View) {
            if ($PSCmdlet.ShouldProcess($v.DisplayName, "Reset to empty ViewItem layout")) {
                foreach ($viewItem in $v.ViewItemChildItems) {
                    $id = New-Guid
                    $viewItem.ViewItemDefinitionXml = '<viewitem id="{0}" displayname="Empty ViewItem" shortcut="" type="VideoOS.RemoteClient.Application.Data.Configuration.EmptyViewItem, VideoOS.RemoteClient.Application"><properties /></viewitem>' -f $id.ToString()
                }
                $v.Save()
            }
            if ($PassThru) {
                Write-Output $View
            }
        }
    }
}


function ConvertFrom-ConfigurationItem {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [OutputType([VideoOS.Platform.ConfigurationItems.IConfigurationItem])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly', '', Justification='ManagementServer inherits from IConfigurationItem')]
    param(
        # Specifies the Milestone Configuration API 'Path' value of the configuration item. For example, 'Hardware[a6756a0e-886a-4050-a5a5-81317743c32a]' where the guid is the ID of an existing Hardware item.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Path,

        # Specifies the Milestone 'ItemType' value such as 'Camera', 'Hardware', or 'InputEvent'
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $ItemType
    )

    begin {
        Assert-VmsRequirementsMet
        $assembly = [System.Reflection.Assembly]::GetAssembly([VideoOS.Platform.ConfigurationItems.Hardware])
        $serverId = (Get-VmsSite -ErrorAction Stop).FQID.ServerId
    }

    process {
        if ($Path -eq '/') {
            [VideoOS.Platform.ConfigurationItems.ManagementServer]::new($serverId)
        } else {
            $instance = $assembly.CreateInstance("VideoOS.Platform.ConfigurationItems.$ItemType", $false, [System.Reflection.BindingFlags]::Default, $null, (@($serverId, $Path)), $null, $null)
            Write-Output $instance
        }
    }
}


function Copy-VmsView {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.View])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.View[]]
        $View,

        [Parameter(Mandatory)]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $DestinationViewGroup,

        [Parameter()]
        [switch]
        $Force,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($v in $View) {
            $newName = $v.Name
            if ($DestinationViewGroup.ViewFolder.Views.Name -contains $newName) {
                if ($Force) {
                    $existingView = $DestinationViewGroup.ViewFolder.Views | Where-Object Name -eq $v.Name
                    $existingView | Remove-VmsView -Confirm:$false
                } else {
                    while ($newName -in $DestinationViewGroup.ViewFolder.Views.Name) {
                        $newName = '{0} - Copy' -f $newName
                    }
                }
            }
            $params = @{
                Name = $newName
                LayoutDefinitionXml = $v.LayoutViewItems
                ViewItemDefinitionXml = $v.ViewItemChildItems.ViewItemDefinitionXml
            }
            $newView = $DestinationViewGroup | New-VmsView @params
            Write-Output $newView
        }
    }
}


function Copy-VmsViewGroup {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ViewGroup])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.ViewGroup[]]
        $ViewGroup,

        [Parameter()]
        [ValidateNotNull()]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $DestinationViewGroup,

        [Parameter()]
        [switch]
        $Force,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($vg in $ViewGroup) {
            $source = $vg | Get-ConfigurationItem -Recurse | ConvertTo-Json -Depth 100 -Compress | ConvertFrom-Json
            $destFolder = (Get-VmsManagementServer).ViewGroupFolder
            if ($MyInvocation.BoundParameters.ContainsKey('DestinationViewGroup')) {
                $destFolder = $DestinationViewGroup.ViewGroupFolder
            }
            $destFolder.ClearChildrenCache()
            $nameProp = $source.Properties | Where-Object Key -eq 'Name'
            if ($nameProp.Value -in $destFolder.ViewGroups.DisplayName -and $Force) {
                $existingGroup = $destFolder.ViewGroups | Where-Object DisplayName -eq $nameProp.Value
                if ($existingGroup.Path -ne $source.Path) {
                    Remove-VmsViewGroup -ViewGroup $existingGroup -Recurse
                }
            }
            while ($nameProp.Value -in $destFolder.ViewGroups.DisplayName) {
                $nameProp.Value = '{0} - Copy' -f $nameProp.Value
            }
            $params = @{
                Source = $source
            }
            if ($MyInvocation.BoundParameters.ContainsKey('DestinationViewGroup')) {
                $params.ParentViewGroup = $DestinationViewGroup
            }
            $newViewGroup = Copy-ViewGroupFromJson @params
            if ($PassThru) {
                Write-Output $newViewGroup
            }
        }
    }
}


function Export-VmsHardware {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'LiteralPath')]
        [ValidateNotNull()]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware,

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Path')]
        [string]
        $Path,

        [Parameter(Mandatory, ParameterSetName = 'LiteralPath')]
        [string]
        $LiteralPath,

        [Parameter()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input', 'Output')]
        [string[]]
        $DeviceType = @('Camera'),

        [Parameter()]
        [ValidateSet('All', 'Enabled', 'Disabled')]
        [string]
        $EnableFilter = 'Enabled',

        [Parameter()]
        [char]
        $Delimiter = ','
    )

    begin {
        Assert-VmsRequirementsMet
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $LiteralPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }
        $records = [collections.generic.list[VideoOS.Platform.ConfigurationItems.Hardware]]::new()
    }
    
    process {
        if ($Hardware.Count -eq 0) {
            $Hardware = Get-VmsHardware
        }
        
        foreach ($hw in $Hardware) {
            $records.Add($hw)
        }
    }
    
    end {
        if ($LiteralPath -match '\.csv$') {
            $splat = @{
                Hardware     = $records
                EnableFilter = $EnableFilter
                DeviceType   = $DeviceType
            }
            ExportHardwareCsv @splat | Export-Csv -LiteralPath $LiteralPath -Delimiter $Delimiter -NoTypeInformation
        } elseif ($LiteralPath -match '\.xlsx$') {
            if ($null -eq (Get-Module ImportExcel)) {
                if (Get-module ImportExcel -ListAvailable) {
                    Import-Module ImportExcel
                } else {
                    Import-Module "$PSScriptRoot\modules\ImportExcel\7.8.9\ImportExcel.psd1"
                }
            }
            $splat = @{
                Path            = $LiteralPath
                Hardware        = $records
                EnableFilter    = $EnableFilter
                IncludedDevices = $DeviceType
            }
            Export-VmsHardwareExcel @splat
        } else {
            Write-Error -Message 'Invalid file extension. Please specify a file path with either a .CSV or .XLSX extension.' -ErrorId 'InvalidExtension' -Category InvalidArgument
        }
    }
}


function Export-VmsLicenseRequest {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([System.IO.FileInfo])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $Force,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if ((Test-Path $filePath) -and -not $Force) {
                Write-Error "File '$Path' already exists. To overwrite an existing file, specify the -Force switch."
                return
            }
            $ms = Get-VmsManagementServer
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].RequestLicense()
            if ($result.State -ne 'Success') {
                Write-Error "Failed to create license request. $($result.ErrorText.Trim('.'))."
                return
            }

            $content = [Convert]::FromBase64String($result.GetProperty('License'))
            [io.file]::WriteAllBytes($filePath, $content)

            if ($PassThru) {
                Get-Item -Path $filePath
            }
        }
        catch {
            Write-Error $_
        }
    }
}


function Export-VmsViewGroup {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion(21.1)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup]
        $ViewGroup,

        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        [environment]::CurrentDirectory = Get-Location
        $Path = [io.path]::GetFullPath($Path)
        $fileInfo = [io.fileinfo]::new($Path)

        if (-not $fileInfo.Directory.Exists) {
            if ($Force) {
                $null = New-Item -Path $fileInfo.Directory.FullName -ItemType Directory -Force
            } else {
                throw [io.DirectoryNotfoundexception]::new("Directory does not exist: $($fileInfo.Directory.FullName). Create the directory manually, or use the -Force switch.")
            }
        }

        if ($fileInfo.Exists -and -not $Force) {
            throw [invalidoperationexception]::new("File already exists. Use -Force to overwrite the existing file.")
        }
        $item = $ViewGroup | Get-ConfigurationItem -Recurse
        $json = $item | ConvertTo-Json -Depth 100 -Compress
        [io.file]::WriteAllText($Path, $json)
    }
}


function Find-ConfigurationItem {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        # Specifies all, or part of the display name of the configuration item to search for. For example, if you want to find a camera named "North West Parking" and you specify the value 'Parking', you will get results for any camera where 'Parking' appears in the name somewhere. The search is not case sensitive.
        [Parameter()]
        [string]
        $Name,

        # Specifies the type(s) of items to include in the results. The default is to include only 'Camera' items.
        [Parameter()]
        [string[]]
        $ItemType = 'Camera',

        # Specifies whether all matching items should be included, or whether only enabled, or disabled items should be included in the results. The default is to include all items regardless of state.
        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'All',

        # An optional hashtable of additional property keys and values to filter results. Properties must be string types, and the results will be included if the property key exists, and the value contains the provided string.
        [Parameter()]
        [hashtable]
        $Properties = @{}
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $svc = Get-IConfigurationService -ErrorAction Stop
        $itemFilter = [VideoOS.ConfigurationApi.ClientService.ItemFilter]::new()
        $itemFilter.EnableFilter = [VideoOS.ConfigurationApi.ClientService.EnableFilter]::$EnableFilter

        $propertyFilters = New-Object System.Collections.Generic.List[VideoOS.ConfigurationApi.ClientService.PropertyFilter]
        if (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -ne '*') {
            $Properties.Name = $Name
        }
        foreach ($key in $Properties.Keys) {
            $propertyFilters.Add([VideoOS.ConfigurationApi.ClientService.PropertyFilter]::new(
                    $key,
                    [VideoOS.ConfigurationApi.ClientService.Operator]::Contains,
                    $Properties.$key
                ))
        }
        $itemFilter.PropertyFilters = $propertyFilters

        foreach ($type in $ItemType) {
            $itemFilter.ItemType = $type
            $svc.QueryItems($itemFilter, [int]::MaxValue) | Foreach-Object {
                Write-Output $_
            }
        }
    }
}

$ItemTypeArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    ([VideoOS.ConfigurationAPI.ItemTypes] | Get-Member -Static -MemberType Property).Name | Where-Object {
        $_ -like "$wordToComplete*"
    } | Foreach-Object {
        "'$_'"
    }
}
Register-ArgumentCompleter -CommandName Find-ConfigurationItem -ParameterName ItemType -ScriptBlock $ItemTypeArgCompleter
Register-ArgumentCompleter -CommandName ConvertFrom-ConfigurationItem -ParameterName ItemType -ScriptBlock $ItemTypeArgCompleter


function Find-XProtectDevice {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    param(
        # Specifies the ItemType such as Camera, Microphone, or InputEvent. Default is 'Camera'.
        [Parameter()]
        [ValidateSet('Hardware', 'Camera', 'Microphone', 'Speaker', 'InputEvent', 'Output', 'Metadata')]
        [string[]]
        $ItemType = 'Camera',

        # Specifies name, or part of the name of the device(s) to find.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        # Specifies all or part of the IP or hostname of the hardware device to search for.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Address,

        # Specifies all or part of the MAC address of the hardware device to search for. Note: Searching by MAC is significantly slower than searching by IP.
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $MacAddress,

        # Specifies whether all devices should be returned, or only enabled or disabled devices. Default is to return all matching devices.
        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'All',

        # Specifies an optional hash table of key/value pairs matching properties on the items you're searching for.
        [Parameter()]
        [hashtable]
        $Properties = @{},

        [Parameter(ParameterSetName = 'ShowDialog')]
        [switch]
        $ShowDialog
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($ShowDialog) {
            Find-XProtectDeviceDialog
            return
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Address')) {
            $ItemType = 'Hardware'
            $Properties.Address = $Address
        }

        if ($MyInvocation.BoundParameters.ContainsKey('MacAddress')) {
            $ItemType = 'Hardware'
            $MacAddress = $MacAddress.Replace(':', '').Replace('-', '')
        }
        # When many results are returned, this hashtable helps avoid unnecessary configuration api queries by caching parent items and indexing by their Path property
        $pathToItemMap = @{}

        Find-ConfigurationItem -ItemType $ItemType -EnableFilter $EnableFilter -Name $Name -Properties $Properties | Foreach-Object {
            $item = $_
            if (![string]::IsNullOrWhiteSpace($MacAddress)) {
                $hwid = ($item.Properties | Where-Object Key -eq 'Id').Value
                $mac = ((Get-ConfigurationItem -Path "HardwareDriverSettings[$hwid]").Children[0].Properties | Where-Object Key -like '*/MacAddress/*' | Select-Object -ExpandProperty Value).Replace(':', '').Replace('-', '')
                if ($mac -notlike "*$MacAddress*") {
                    return
                }
            }
            $deviceInfo = [ordered]@{}
            while ($true) {
                $deviceInfo.($item.ItemType) = $item.DisplayName
                if ($item.ItemType -eq 'RecordingServer') {
                    break
                }
                $parentItemPath = $item.ParentPath -split '/' | Select-Object -First 1

                # Set $item to the cached copy of that parent item if available. If not, retrieve it using configuration api and cache it.
                if ($pathToItemMap.ContainsKey($parentItemPath)) {
                    $item = $pathToItemMap.$parentItemPath
                } else {
                    $item = Get-ConfigurationItem -Path $parentItemPath
                    $pathToItemMap.$parentItemPath = $item
                }
            }
            [pscustomobject]$deviceInfo
        }
    }
}


function Get-ManagementServerConfig {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param()

    begin {
        Assert-VmsRequirementsMet
        $configXml = Join-Path ([system.environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData)) 'milestone\xprotect management server\serverconfig.xml'
        if (-not (Test-Path $configXml)) {
            throw [io.filenotfoundexception]::new('Management Server configuration file not found', $configXml)
        }
    }

    process {
        $xml = [xml](Get-Content -Path $configXml)
        
        $versionNode = $xml.SelectSingleNode('/server/version')
        $clientRegistrationIdNode = $xml.SelectSingleNode('/server/ClientRegistrationId')
        $webApiPortNode = $xml.SelectSingleNode('/server/WebApiConfig/Port')
        $authServerAddressNode = $xml.SelectSingleNode('/server/WebApiConfig/AuthorizationServerUri')


        $serviceProperties = 'Name', 'PathName', 'StartName', 'ProcessId', 'StartMode', 'State', 'Status'
        $serviceInfo = Get-CimInstance -ClassName 'Win32_Service' -Property $serviceProperties -Filter "name = 'Milestone XProtect Management Server'"

        $config = @{
            Version = if ($null -ne $versionNode) { [version]::Parse($versionNode.InnerText) } else { [version]::new(0, 0) }
            ClientRegistrationId = if ($null -ne $clientRegistrationIdNode) { [guid]$clientRegistrationIdNode.InnerText } else { [guid]::Empty }
            WebApiPort = if ($null -ne $webApiPortNode) { [int]$webApiPortNode.InnerText } else { 0 }
            AuthServerAddress = if ($null -ne $authServerAddressNode) { [uri]$authServerAddressNode.InnerText } else { $null }
            ServerCertHash = $null
            InstallationPath = $serviceInfo.PathName.Trim('"')
            ServiceInfo = $serviceInfo
        }

        $netshResult = Get-ProcessOutput -FilePath 'netsh.exe' -ArgumentList "http show sslcert ipport=0.0.0.0:$($config.WebApiPort)"
        if ($netshResult.StandardOutput -match 'Certificate Hash\s+:\s+(\w+)\s+') {
            $config.ServerCertHash = $Matches.1
        }

        Write-Output ([pscustomobject]$config)
    }
}

function Get-PlaybackInfo {
    [CmdletBinding(DefaultParameterSetName = 'FromPath')]
    [RequiresVmsConnection()]
    param (
        # Accepts a Milestone Configuration Item path string like Camera[A64740CF-5511-4957-9356-2922A25FF752]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'FromPath')]
        [ValidateScript( {
                if ($_ -notmatch '^(?<ItemType>\w+)\[(?<Id>[a-fA-F0-9\-]{36})\]$') {
                    throw "$_ does not a valid Milestone Configuration API Item path"
                }
                if ($Matches.ItemType -notin @('Camera', 'Microphone', 'Speaker', 'Metadata')) {
                    throw "$_ represents an item of type '$($Matches.ItemType)'. Only camera, microphone, speaker, or metadata item types are allowed."
                }
                return $true
            })]
        [string[]]
        $Path,

        # Accepts a Camera, Microphone, Speaker, or Metadata object
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromDevice')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Device,

        [Parameter()]
        [ValidateSet('MotionSequence', 'RecordingSequence', 'TimelineMotionDetected', 'TimelineRecording')]
        [string]
        $SequenceType = 'RecordingSequence',

        [Parameter()]
        [switch]
        $Parallel,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [VideoOS.Platform.ConfigurationItems.Camera]
        $Camera,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [guid]
        $CameraId,

        [Parameter(ParameterSetName = 'DeprecatedParameterSet')]
        [switch]
        $UseLocalTime
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'DeprecatedParameterSet') {
            Write-Warning 'The Camera, CameraId, and UseLocalTime parameters are deprecated. See "Get-Help Get-PlaybackInfo -Full" for more information.'
            if ($null -ne $Camera) {
                $Path = $Camera.Path
            }
            else{
                $Path = "Camera[$CameraId]"
            }
        }
        if ($PSCmdlet.ParameterSetName -eq 'FromDevice') {
            $Path = $Device.Path
        }
        if ($Path.Count -le 60 -and $Parallel) {
            Write-Warning "Ignoring the Parallel switch since there are only $($Path.Count) devices to query."
            $Parallel = $false
        }

        if ($Parallel) {
            $jobRunner = [LocalJobRunner]::new()
        }


        $script = {
            param([string]$Path, [string]$SequenceType)
            if ($Path -notmatch '^(?<ItemType>\w+)\[(?<Id>[a-fA-F0-9\-]{36})\]$') {
                Write-Error "Path '$Path' is not a valid Milestone Configuration API item path."
                return
            }
            try {
                $site = Get-VmsSite
                $epoch = [datetime]::SpecifyKind([datetimeoffset]::FromUnixTimeSeconds(0).DateTime, [datetimekind]::utc)
                $item = [videoos.platform.Configuration]::Instance.GetItem($site.FQID.ServerId, $Matches.Id, [VideoOS.Platform.Kind]::($Matches.ItemType))
                if ($null -eq $item) {
                    Write-Error "Camera not available. It may be disabled, or it may not belong to a camera group."
                    return
                }
                $sds = [VideoOS.Platform.Data.SequenceDataSource]::new($item)
                $sequenceTypeGuid = [VideoOS.Platform.Data.DataType+SequenceTypeGuids]::$SequenceType
                $first = $sds.GetData($epoch, [timespan]::zero, 0, ([datetime]::utcnow - $epoch), 1, $sequenceTypeGuid) | Select-Object -First 1
                $last = $sds.GetData([datetime]::utcnow, ([datetime]::utcnow - $epoch), 1, [timespan]::zero, 0, $sequenceTypeGuid) | Select-Object -First 1
                if ($first.EventSequence -and $last.EventSequence) {
                    [PSCustomObject]@{
                        Begin = $first.EventSequence.StartDateTime
                        End   = $last.EventSequence.EndDateTime
                        Retention = $last.EventSequence.EndDateTime - $first.EventSequence.StartDateTime
                        Path = $Path
                    }
                }
                else {
                    Write-Warning "No sequences of type '$SequenceType' found for $(($Matches.ItemType).ToLower()) $($item.Name) ($($item.FQID.ObjectId))"
                }
            } finally {
                if ($sds) {
                    $sds.Close()
                }
            }
        }

        try {
            foreach ($p in $Path) {
                if ($Parallel) {
                    $null = $jobRunner.AddJob($script, @{Path = $p; SequenceType = $SequenceType})
                }
                else {
                    $script.Invoke($p, $SequenceType) | Foreach-Object {
                        if ($UseLocalTime) {
                            $_.Begin = $_.Begin.ToLocalTime()
                            $_.End = $_.End.ToLocalTime()
                        }
                        $_
                    }
                }
            }

            if ($Parallel) {
                while ($jobRunner.HasPendingJobs()) {
                    $jobRunner.ReceiveJobs() | Foreach-Object {
                        if ($_.Output) {
                            if ($UseLocalTime) {
                                $_.Output.Begin = $_.Output.Begin.ToLocalTime()
                                $_.Output.End = $_.Output.End.ToLocalTime()
                            }
                            Write-Output $_.Output
                        }
                        if ($_.Errors) {
                            $_.Errors | Foreach-Object {
                                Write-Error $_
                            }
                        }
                    }
                    Start-Sleep -Milliseconds 200
                }
            }
        }
        finally {
            if ($jobRunner) {
                $jobRunner.Dispose()
            }
        }
    }
}


function Get-RecorderConfig {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param()

    begin {
        Assert-VmsRequirementsMet
        $configXml = Join-Path ([system.environment]::GetFolderPath([System.Environment+SpecialFolder]::CommonApplicationData)) 'milestone\xprotect recording server\recorderconfig.xml'
        if (-not (Test-Path $configXml)) {
            throw [io.filenotfoundexception]::new('Recording Server configuration file not found', $configXml)
        }
    }

    process {
        $xml = [xml](Get-Content -Path $configXml)
        
        $versionNode = $xml.SelectSingleNode('/recorderconfig/version')
        $recorderIdNode = $xml.SelectSingleNode('/recorderconfig/recorder/id')
        $clientRegistrationIdNode = $xml.SelectSingleNode('/recorderconfig/recorder/ClientRegistrationId')
        $webServerPortNode = $xml.SelectSingleNode('/recorderconfig/webserver/port')        
        $alertServerPortNode = $xml.SelectSingleNode('/recorderconfig/driverservices/alert/port')
        $serverAddressNode = $xml.SelectSingleNode('/recorderconfig/server/address')        
        $serverPortNode = $xml.SelectSingleNode('/recorderconfig/server/webapiport')        
        $localServerPortNode = $xml.SelectSingleNode('/recorderconfig/webapi/port')
        $authServerAddressNode = $xml.SelectSingleNode('/recorderconfig/server/authorizationserveraddress')

        $serviceProperties = 'Name', 'PathName', 'StartName', 'ProcessId', 'StartMode', 'State', 'Status'
        $serviceInfo = Get-CimInstance -ClassName 'Win32_Service' -Property $serviceProperties -Filter "name = 'Milestone XProtect Recording Server'"

        $config = @{
            Version = if ($null -ne $versionNode) { [version]::Parse($versionNode.InnerText) } else { [version]::new(0, 0) }
            RecorderId = if ($null -ne $recorderIdNode) { [guid]$recorderIdNode.InnerText } else { [guid]::Empty }
            ClientRegistrationId = if ($null -ne $clientRegistrationIdNode) { [guid]$clientRegistrationIdNode.InnerText } else { [guid]::Empty }
            WebServerPort = if ($null -ne $webServerPortNode) { [int]$webServerPortNode.InnerText } else { 0 }
            AlertServerPort = if ($null -ne $alertServerPortNode) { [int]$alertServerPortNode.InnerText } else { 0 }
            ServerAddress = $serverAddressNode.InnerText
            ServerPort = if ($null -ne $serverPortNode) { [int]$serverPortNode.InnerText } else { 0 }
            LocalServerPort = if ($null -ne $localServerPortNode) { [int]$localServerPortNode.InnerText } else { 0 }
            AuthServerAddress = if ($null -ne $authServerAddressNode) { [uri]$authServerAddressNode.InnerText } else { $null }
            ServerCertHash = $null
            InstallationPath = $serviceInfo.PathName.Trim('"')
            DevicePackPath = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\WOW6432Node\VideoOS\DeviceDrivers -Name InstallPath
            ServiceInfo = $serviceInfo
        }

        $netshResult = Get-ProcessOutput -FilePath 'netsh.exe' -ArgumentList "http show sslcert ipport=0.0.0.0:$($config.LocalServerPort)"
        if ($netshResult.StandardOutput -match 'Certificate Hash\s+:\s+(\w+)\s+') {
            $config.ServerCertHash = $Matches.1
        }

        Write-Output ([pscustomobject]$config)
    }
}

function Get-VmsBasicUser {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.BasicUser])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[BasicUser]])]
        [string]
        $Name,

        [Parameter()]
        [ValidateSet('Enabled', 'LockedOutByAdmin', 'LockedOutBySystem')]
        [string]
        $Status,

        [Parameter()]
        [switch]
        $External
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $matchFound = $false
        foreach ($user in (Get-VmsManagementServer).BasicUserFolder.BasicUsers){
            if ($MyInvocation.BoundParameters.ContainsKey('Status') -and $user.Status -ne $Status) {
                continue
            }

            if ($MyInvocation.BoundParameters.ContainsKey('External') -and $user.IsExternal -ne $External) {
                continue
            }

            if ($MyInvocation.BoundParameters.ContainsKey('Name') -and $user.Name -ne $Name) {
                continue
            }
            $matchFound = $true
            $user
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Name') -and -not $matchFound) {
            Write-Error "No basic user found matching the name '$Name'"
        }
    }
}



function Get-VmsBasicUserClaim {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClaimChildItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [VideoOS.Platform.ConfigurationItems.BasicUser[]]
        $InputObject
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($user in $InputObject) {
            $user.ClaimFolder.ClaimChildItems | ForEach-Object {
                $_
            }
        }
    }
}


function Get-VmsCameraStream {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [OutputType([VmsCameraStreamConfig])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        [Parameter(ParameterSetName = 'ByName')]
        [string]
        $Name,

        [Parameter(Mandatory, ParameterSetName = 'Enabled')]
        [switch]
        $Enabled,

        [Parameter(Mandatory, ParameterSetName = 'LiveDefault')]
        [switch]
        $LiveDefault,

        [Parameter(ParameterSetName = 'PlaybackDefault')]
        [switch]
        $PlaybackDefault,

        [Parameter(Mandatory, ParameterSetName = 'Recorded')]
        [switch]
        $Recorded,

        [Parameter(ParameterSetName = 'RecordingTrack')]
        [ValidateSet('Primary', 'Secondary', 'None')]
        [string]
        $RecordingTrack,

        [Parameter()]
        [switch]
        $RawValues
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($cam in $Camera) {
            $streamUsages = ($cam.StreamFolder.Streams | Select-Object -First 1).StreamUsageChildItems
            if ($null -eq $streamUsages) {
                $message = 'Camera "{0}" does not support simultaneous use of multiple streams. The following properties should be ignored for streams on this camera: DisplayName, Enabled, LiveMode, LiveDefault, Recorded.' -f $cam.Name
                Write-Warning $message
            }
            $deviceDriverSettings = $cam.DeviceDriverSettingsFolder.DeviceDriverSettings
            if ($null -eq $deviceDriverSettings -or $deviceDriverSettings.Count -eq 0 -or $deviceDriverSettings[0].StreamChildItems.Count -eq 0) {
                # Added this due to a situation where a camera/driver is in a weird state where maybe a replace hardware
                # is needed to bring it online and until then there are no stream settings listed in the settings tab
                # for the camera. This block allows us to return _something_ even though there are no stream settings available.
                $message = 'Camera "{0}" has no device driver settings available.' -f $cam.Name
                Write-Warning $message
                foreach ($streamUsage in $streamUsages) {
                    if ($LiveDefault -and -not $streamUsage.LiveDefault) {
                        continue
                    }
                    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Recorded') -and $Recorded -ne $streamUsage.Record) {
                        continue
                    }
                    [VmsCameraStreamConfig]@{
                        Name              = $streamUsage.Name
                        DisplayName       = $streamUsage.Name
                        Enabled           = $true
                        LiveDefault       = $streamUsage.LiveDefault
                        LiveMode          = $streamUsage.LiveMode
                        Recorded          = $streamUsage.Record
                        Settings          = @{}
                        ValueTypeInfo     = @{}
                        Camera            = $cam
                        StreamReferenceId = $streamUsage.StreamReferenceId
                    }
                }

                continue
            }

            foreach ($stream in $deviceDriverSettings[0].StreamChildItems) {
                $streamUsage = if ($streamUsages) { $streamUsages | Where-Object { $_.StreamReferenceId -eq $_.StreamReferenceIdValues[$stream.DisplayName] } }

                if ($LiveDefault -and -not $streamUsage.LiveDefault) {
                    continue
                }
                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Recorded') -and $Recorded -ne $streamUsage.Record) {
                    continue
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecordingTrack')) {
                    if ($RecordingTrack -eq 'Primary' -and -not $streamUsage.Record) {
                        continue
                    } elseif ($RecordingTrack -eq 'Secondary' -and $streamUsage.RecordTo -ne '84fff8b9-8cd1-46b2-a451-c4a87d4cbbb0') {
                        continue
                    } elseif ($RecordingTrack -eq 'None' -and ($streamUsage.Record -or -not [string]::IsNullOrEmpty($streamUsage.RecordTo))) {
                        continue
                    }
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('PlaybackDefault') -and (($streamUsage.RecordToValues.Count -eq 0 -and $streamUsage.Record -ne $PlaybackDefault) -or ($streamUsage.RecordToValues.Count -gt 0 -and $streamUsage.DefaultPlayback -ne $PlaybackDefault))) {
                    continue
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Enabled') -and $streamUsages -and $Enabled -eq ($null -eq $streamUsage)) {
                    continue
                }

                if ($MyInvocation.BoundParameters.ContainsKey('Name') -and $stream.DisplayName -notlike $Name) {
                    continue
                }

                $streamConfig = [VmsCameraStreamConfig]@{
                    Name         = $stream.DisplayName
                    Camera       = $cam
                    UseRawValues = $RawValues
                }
                $streamConfig.Update()
                $streamConfig
            }
        }
    }
}


function Get-VmsConnectionString {
    [CmdletBinding()]
    [Alias('Get-ConnectionString')]
    [OutputType([string])]
    [RequiresVmsConnection($false)]
    param (
        [Parameter(Position = 0)]
        [string]
        $Component = 'ManagementServer'
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (Get-Item -Path HKLM:\SOFTWARE\VideoOS\Server\ConnectionString -ErrorAction Ignore) {
            Get-ItemPropertyValue -Path HKLM:\SOFTWARE\VideoOS\Server\ConnectionString -Name $Component
        } else {
            if ($Component -ne 'ManagementServer') {
                Write-Warning "Specifying a component name is only allowed on a management server running version 2022 R3 (22.3) or greater."
            }
            Get-ItemPropertyValue -Path HKLM:\SOFTWARE\VideoOS\Server\Common -Name 'Connectionstring'
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsConnectionString -ParameterName Component -ScriptBlock {
    $values = Get-Item HKLM:\SOFTWARE\videoos\Server\ConnectionString\ -ErrorAction Ignore | Select-Object -ExpandProperty Property
    if ($values) {
        Complete-SimpleArgument $args $values
    }
}


function Get-VmsDeviceGroup {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [Alias('Get-DeviceGroup')]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, ParameterSetName = 'ByName')]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'SpeakerGroup', 'MetadataGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $ParentGroup,

        [Parameter(Position = 0, ParameterSetName = 'ByName')]
        [string]
        $Name = '*',

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ByPath')]
        [string[]]
        $Path,

        [Parameter(Position = 2, ParameterSetName = 'ByName')]
        [Parameter(Position = 2, ParameterSetName = 'ByPath')]
        [Alias('DeviceCategory')]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Input', 'Output', 'Metadata')]
        [string]
        $Type = 'Camera',

        [Parameter(ParameterSetName = 'ByName')]
        [Parameter(ParameterSetName = 'ByPath')]
        [switch]
        $Recurse
    )

    begin {
        Assert-VmsRequirementsMet
        $adjustedType = $Type
        if ($adjustedType -eq 'Input') {
            # Inputs on cameras have an object type called "InputEvent"
            # but we don't want the user to have to remember that.
            $adjustedType = 'InputEvent'
        }
    }

    process {
        $rootGroup = Get-VmsManagementServer
        if ($ParentGroup) {
            $rootGroup = $ParentGroup
        }

        $matchFound = $false
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                $subGroups = $rootGroup."$($adjustedType)GroupFolder"."$($adjustedType)Groups"
                $subGroups | Where-Object Name -like $Name | Foreach-Object {
                    if ($null -eq $_) { return }
                    $matchFound = $true
                    $_
                    if ($Recurse) {
                        $_ | Get-VmsDeviceGroup -Type $Type -Recurse
                    }
                }
            }

            'ByPath' {
                foreach ($groupPath in $Path) {
                    $pathPrefixPattern = '^/(?<type>(Camera|Microphone|Speaker|Metadata|Input|Output))(Event)?GroupFolder'
                    if ($groupPath -match $pathPrefixPattern) {
                        $pathPrefix = $groupPath -replace '^/(Camera|Microphone|Speaker|Metadata|Input|Output)(Event)?GroupFolder.*', '$1'
                        if ($pathPrefix -ne $Type) {
                            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Type')) {
                                throw "The device group prefix '$pathPrefix' does not match the specified device group type '$Type'. Either remove the prefix from the path, or do not specify a value for the Type parameter."
                            } else {
                                Write-Verbose "Device type '$pathPrefix' determined from the provided path."
                                $Type = $pathPrefix
                            }
                        }
                    }
                    $params = @{
                        Type        = $Type
                        ErrorAction = 'SilentlyContinue'
                    }
                    $pathInterrupted = $false
                    $groupPath = $groupPath -replace '^/(Camera|Microphone|Speaker|Metadata|InputEvent|Output)GroupFolder', ''
                    $pathParts = $groupPath | Split-VmsDeviceGroupPath
                    foreach ($name in $pathParts) {
                        $params.Name = $name
                        $group = Get-VmsDeviceGroup @params
                        if ($null -eq $group) {
                            $pathInterrupted = $true
                            break
                        }
                        $params.ParentGroup = $group
                    }
                    if ($pathParts -and -not $pathInterrupted) {
                        $matchFound = $true
                        $params.ParentGroup
                        if ($Recurse) {
                            $params.ParentGroup | Get-VmsDeviceGroup -Type $Type -Recurse
                        }
                    }
                    if ($null -eq $pathParts -and $Recurse) {
                        Get-VmsDeviceGroup -Type $Type -Recurse
                    }
                }
            }
        }

        if (-not $matchFound -and -not [management.automation.wildcardpattern]::ContainsWildcardCharacters($Name)) {
            Write-Error "No $Type group found with the name '$Name'"
        }
    }
}

function Get-VmsDeviceGroupMember {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline)]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'MetadataGroup', 'SpeakerGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Group,

        [Parameter()]
        [VideoOS.ConfigurationApi.ClientService.EnableFilter]
        $EnableFilter = [VideoOS.ConfigurationApi.ClientService.EnableFilter]::Enabled
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $deviceType = ($Group | Split-VmsConfigItemPath -ItemType) -replace 'Group$', ''
        $Group."$($deviceType)Folder"."$($deviceType)s" | ForEach-Object {
            if ($_.Enabled -and $EnableFilter -eq 'Disabled') {
                return
            }
            if (-not $_.Enabled -and $EnableFilter -eq 'Enabled') {
                return
            }
            $_
        }
    }
}


function Get-VmsDeviceStatus {
    [CmdletBinding()]
    [OutputType([VmsStreamDeviceStatus])]
    [RequiresVmsConnection()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', IgnoreCase = $false)]
        [string[]]
        $DeviceType = 'Camera'
    )

    begin {
        Assert-VmsRequirementsMet
        $scriptBlock = {
            param([guid]$RecorderId, [VideoOS.Platform.Item[]]$Devices, [type]$VmsStreamDeviceStatusClass)
            $recorderItem = [VideoOS.Platform.Configuration]::Instance.GetItem($RecorderId, [VideoOS.Platform.Kind]::Server)
            $svc = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($recorderItem.FQID.ServerId.Uri)
            $status = @{}
            $currentStatus = $svc.GetCurrentDeviceStatus((Get-VmsToken), $Devices.FQID.ObjectId)
            foreach ($kind in 'Camera', 'Microphone', 'Speaker', 'Metadata') {
                foreach ($entry in $currentStatus."$($kind)DeviceStatusArray") {
                    $status[$entry.DeviceId] = $entry
                }
            }
            foreach ($item in $Devices) {
                $obj = $VmsStreamDeviceStatusClass::new($status[$item.FQID.ObjectId])
                $obj.DeviceName = $item.Name
                $obj.DeviceType = [VideoOS.Platform.Kind]::DefaultTypeToNameTable[$item.FQID.Kind]
                $obj.RecorderName = $recorderItem.Name
                $obj.RecorderId = $RecorderItem.FQID.ObjectId
                Write-Output $obj
            }
        }
    }

    process {
        <# TODO: Once a decision is made on how to handle the PoshRSJob
           dependency, uncomment the bits below and remove the line right
           after the opening foreach curly brace as it's already handled
           in the else block.
        #>
        $recorderCameraMap = Get-DevicesByRecorder -Id $RecordingServerId -DeviceType $DeviceType
        # $jobs = [system.collections.generic.list[RSJob]]::new()
        foreach ($recorderId in $recorderCameraMap.Keys) {
            $scriptBlock.Invoke($recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus]))
            # if ($Parallel -and $RecordingServerId.Count -gt 1) {
            #     $job = Start-RSJob -ScriptBlock $scriptBlock -ArgumentList $recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus])
            #     $jobs.Add($job)
            # } else {
            #     $scriptBlock.Invoke($recorderId, $recorderCameraMap.$recorderId, ([VmsStreamDeviceStatus]))
            # }
        }
        # if ($jobs.Count -gt 0) {
        #     $jobs | Wait-RSJob -ShowProgress:($ProgressPreference -eq 'Continue') | Receive-RSJob
        #     $jobs | Remove-RSJob
        # }
    }
}


function Get-VmsHardwareDriver {
    [CmdletBinding(DefaultParameterSetName = 'Hardware')]
    [OutputType([VideoOS.Platform.ConfigurationItems.HardwareDriver])]
    [Alias('Get-HardwareDriver')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'RecordingServer')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer[]]
        $RecordingServer,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Hardware')]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware
    )

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
        $driversByRecorder = @{}
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'RecordingServer' {
                foreach ($rec in $RecordingServer) {
                    foreach ($driver in $rec.HardwareDriverFolder.HardwareDrivers | Sort-Object DriverType) {
                        $driver
                    }
                }
            }
            'Hardware' {
                foreach ($hw in $Hardware) {
                    if (-not $driversByRecorder.ContainsKey($hw.ParentItemPath)) {
                        $driversByRecorder[$hw.ParentItemPath] = @{}
                        $rec = [VideoOS.Platform.ConfigurationItems.RecordingServer]::new($hw.ServerId, $hw.ParentItemPath)
                        $rec.HardwareDriverFolder.HardwareDrivers | ForEach-Object {
                            $driversByRecorder[$hw.ParentItemPath][$_.Path] = $_
                        }
                    }
                    $driver = $driversByRecorder[$hw.ParentItemPath][$hw.HardwareDriverPath]
                    if ($null -eq $driver) {
                        Write-Error "HardwareDriver '$($hw.HardwareDriverPath)' for hardware '$($hw.Name)' not found on the parent recording server."
                        continue
                    }
                    $driver
                }
            }
            Default {
                throw "Support for ParameterSetName '$_' not implemented."
            }
        }
    }

    end {
        $driversByRecorder.Clear()
    }
}

function Get-VmsHardwarePassword {
    [CmdletBinding()]
    [OutputType([string])]
    [Alias('Get-HardwarePassword')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware]
        $Hardware
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $serverTask = $Hardware.ReadPasswordHardware()
            if ($serverTask.State -ne [VideoOS.Platform.ConfigurationItems.StateEnum]::Success) {
                Write-Error -Message "ReadPasswordHardware error: $(t.ErrorText)" -TargetObject $Hardware
                return
            }
            $serverTask.GetProperty('Password')
        } catch {
            Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $Hardware
        }
    }
}


function Get-VmsLoginProvider {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LoginProvider])]
    param (
        [Parameter(Position = 0)]
        [ArgumentCompleter([MilestonePSTools.Utility.MipItemNameCompleter[VideoOS.Platform.ConfigurationItems.LoginProvider]])]
        [string]
        $Name
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {

        if ($MyInvocation.BoundParameters.ContainsKey('Name')) {
            $loginProviders = (Get-VmsManagementServer).LoginProviderFolder.LoginProviders | Where-Object Name -EQ $Name
        } else {
            $loginProviders = (Get-VmsManagementServer).LoginProviderFolder.LoginProviders | ForEach-Object { $_ }
        }
        if ($loginProviders) {
            $loginProviders
        } elseif ($MyInvocation.BoundParameters.ContainsKey('Name')) {
            Write-Error 'No matching login provider found.'
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsLoginProvider -ParameterName Name -ScriptBlock {
    $values = (Get-VmsLoginProvider).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsLoginProviderClaim {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.RegisteredClaim])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter()]
        [string]
        $Name
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $LoginProvider.RegisteredClaimFolder.RegisteredClaims | Foreach-Object {
            if ($MyInvocation.BoundParameters.ContainsKey('Name') -and $_.Name -ne $Name) {
                return
            }
            $_
        }
    }
}

function Get-VmsRecordingServer {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [Alias('Get-RecordingServer')]
    [OutputType([VideoOS.Platform.ConfigurationItems.RecordingServer])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ByName')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [string]
        $Name = '*',

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ById')]
        [guid]
        $Id,

        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'ByHostname')]
        [Alias('ComputerName')]
        [string]
        $HostName = '*'
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                $matchFound = $false
                foreach ($rec in (Get-VmsManagementServer).RecordingServerFolder.RecordingServers | Where-Object Name -like $Name) {
                    $matchFound = $true
                    $rec
                }
                if (-not $matchFound -and -not [system.management.automation.wildcardpattern]::ContainsWildcardCharacters($Name)) {
                    Write-Error "No item found with name matching '$Name'"
                }
            }
            'ById' {
                try {
                    [VideoOS.Platform.ConfigurationItems.RecordingServer]::new((Get-VmsManagementServer).ServerId, "RecordingServer[$Id]")
                }
                catch [VideoOS.Platform.PathNotFoundMIPException] {
                    Write-Error -Message "No item found with id matching '$Id'" -Exception $_.Exception
                }
            }
            'ByHostname' {
                $matchFound = $false
                foreach ($rec in (Get-VmsManagementServer).RecordingServerFolder.RecordingServers | Where-Object HostName -like $HostName) {
                    $matchFound = $true
                    $rec
                }
                if (-not $matchFound -and -not [system.management.automation.wildcardpattern]::ContainsWildcardCharacters($HostName)) {
                    Write-Error "No item found with hostname matching '$HostName'"
                }
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsRecordingServer -ParameterName HostName -ScriptBlock {
    $values = (Get-VmsRecordingServer).HostName | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

Register-ArgumentCompleter -CommandName Get-VmsRecordingServer -ParameterName Id -ScriptBlock {
    $values = (Get-VmsRecordingServer | Sort-Object Name).Id
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsSiteInfo {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({ ValidateSiteInfoTagName @args })]
        [SupportsWildcards()]
        [string]
        $Property = '*'
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $ownerPath = 'BasicOwnerInformation[{0}]' -f (Get-VmsManagementServer).Id
        $ownerInfo = Get-ConfigurationItem -Path $ownerPath
        $resultFound = $false
        foreach ($p in $ownerInfo.Properties) {
            if ($p.Key -match '^\[(?<id>[a-fA-F0-9\-]{36})\]/(?<tagtype>[\w\.]+)$') {
                if ($Matches.tagtype -like $Property) {
                    $resultFound = $true
                    [pscustomobject]@{
                        DisplayName  = $p.DisplayName
                        Property   = $Matches.tagtype
                        Value = $p.Value
                    }
                }
            } else {
                Write-Warning "Site information property key format unrecognized: $($p.Key)"
            }
        }
        if (-not $resultFound -and -not [system.management.automation.wildcardpattern]::ContainsWildcardCharacters($Property)) {
            Write-Error "Site information property with key '$Property' not found."
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsSiteInfo -ParameterName Property -ScriptBlock { OwnerInfoPropertyCompleter @args }


function Get-VmsStorageRetention {
    [CmdletBinding()]
    [OutputType([timespan])]
    [RequiresVmsConnection()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [StorageNameTransformAttribute()]
        [VideoOS.Platform.ConfigurationItems.Storage[]]
        $Storage
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($Storage.Count -lt 1) {
            $Storage = Get-VmsStorage
        }
        foreach ($s in $Storage) {
            $retention = [int]$s.RetainMinutes
            foreach ($archive in $s.ArchiveStorageFolder.ArchiveStorages) {
                if ($archive.RetainMinutes -gt $retention) {
                    $retention = $archive.RetainMinutes
                }
            }
            [timespan]::FromMinutes($retention)
        }
    }
}


Register-ArgumentCompleter -CommandName Get-VmsStorageRetention -ParameterName Storage -ScriptBlock {
    $values = (Get-VmsRecordingServer | Get-VmsStorage).Name | Select-Object -Unique | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsToken {
    [CmdletBinding(DefaultParameterSetName = 'CurrentSite')]
    [OutputType([string])]
    [Alias('Get-Token')]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ServerId')]
        [VideoOS.Platform.ServerId]
        $ServerId,

        [Parameter(ValueFromPipeline, ParameterSetName = 'Site')]
        [VideoOS.Platform.Item]
        $Site
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'CurrentSite' {
                    [VideoOS.Platform.Login.LoginSettingsCache]::GetLoginSettings((Get-VmsSite).FQID).Token
                }

                'ServerId' {
                    [VideoOS.Platform.Login.LoginSettingsCache]::GetLoginSettings($ServerId).Token
                }

                'Site' {
                    [VideoOS.Platform.Login.LoginSettingsCache]::GetLoginSettings($Site.FQID).Token
                }

                Default {
                    throw "ParameterSet '$_' not implemented."
                }
            }
        } catch {
            Write-Error -ErrorRecord $_
        }
    }
}


function Get-VmsView {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.View])]
    param (
        [Parameter(ValueFromPipeline, ParameterSetName = 'Default')]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup[]]
        $ViewGroup,

        [Parameter(ParameterSetName = 'Default', Position = 1)]
        [ArgumentCompleter([MilestonePSTools.Utility.MipItemNameCompleter[VideoOS.Platform.ConfigurationItems.View]])]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]
        $Name = '*',

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ById', Position = 2)]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'Default' {
                if ($null -eq $ViewGroup) {
                    $ViewGroup = Get-VmsViewGroup -Recurse
                }
                $count = 0
                foreach ($vg in $ViewGroup) {
                    foreach ($view in $vg.ViewFolder.Views) {
                        if ($view.Path -in $vg.ViewGroupFolder.ViewGroups.ViewFolder.Views.Path) {
                            # TODO: Remove this someday when bug 479533 is no longer an issue.
                            Write-Verbose "Ignoring duplicate view caused by configuration api issue resolved in later VMS versions."
                            continue
                        }
                        foreach ($n in $Name) {
                            if ($view.DisplayName -like $n) {
                                Write-Output $view
                                $count++
                            }
                        }
                    }
                }

                if ($count -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
                    Write-Error "View ""$Name"" not found."
                }
            }

            'ById' {
                $path = 'View[{0}]' -f $Id.ToString().ToUpper()
                Write-Output ([VideoOS.Platform.ConfigurationItems.View]::new((Get-VmsSite).FQID.ServerId, $path))
            }
        }
    }
}

function ViewArgumentCompleter{
    param ( $commandName,
            $parameterName,
            $wordToComplete,
            $commandAst,
            $fakeBoundParameters )

    if ($fakeBoundParameters.ContainsKey('ViewGroup')) {
        $folder = $fakeBoundParameters.ViewGroup.ViewFolder
        $possibleValues = $folder.Views.Name
        $wordToComplete = $wordToComplete.Trim("'").Trim('"')
        if (-not [string]::IsNullOrWhiteSpace($wordToComplete)) {
            $possibleValues = $possibleValues | Where-Object { $_ -like "$wordToComplete*" }
        }
        $possibleValues | Foreach-Object {
            if ($_ -like '* *') {
                "'$_'"
            } else {
                $_
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsView -ParameterName Name -ScriptBlock (Get-Command ViewArgumentCompleter).ScriptBlock


function Get-VmsViewGroup {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ViewGroup])]
    param (
        [Parameter(ValueFromPipeline, ParameterSetName = 'Default')]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $Parent,

        [Parameter(ParameterSetName = 'Default', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [string[]]
        $Name = '*',

        [Parameter(ParameterSetName = 'Default')]
        [switch]
        $Recurse,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ById', Position = 2)]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                $vg = [VideoOS.Platform.ConfigurationItems.ViewGroup]::new((Get-VmsSite).FQID.ServerId, "ViewGroup[$Id]")
                Write-Output $vg
            } catch [System.Management.Automation.MethodInvocationException] {
                if ($_.FullyQualifiedErrorId -eq 'PathNotFoundMIPException') {
                    Write-Error "No ViewGroup found with ID matching $Id"
                    return
                }
            }
        } else {
            if ($null -ne $Parent) {
                $vgFolder = $Parent.ViewGroupFolder
            } else {
                $vgFolder = (Get-VmsManagementServer).ViewGroupFolder
            }

            $count = 0
            foreach ($vg in $vgFolder.ViewGroups) {
                foreach ($n in $Name) {
                    if ($vg.DisplayName -notlike $n) {
                        continue
                    }
                    $count++
                    if (-not $Recurse -or ($Recurse -and $Name -eq '*')) {
                        Write-Output $vg
                    }
                    if ($Recurse) {
                        $vg | Get-VmsViewGroup -Recurse
                    }
                    continue
                }
            }

            if ($count -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
                Write-Error "ViewGroup ""$Name"" not found."
            }
        }
    }
}

function Get-VmsViewGroupAcl {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VmsViewGroupAcl])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup]
        $ViewGroup,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'FromRole')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromRoleId')]
        [VideoOS.Platform.ConfigurationItems.Role]
        $RoleId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromRoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [string]
        $RoleName
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'FromRole' { }
            'FromRoleId' { $Role = Get-VmsRole -Id $RoleId -ErrorAction Stop }
            'FromRoleName' { $Role = Get-VmsRole -Name $RoleName -ErrorAction Stop }
            Default { throw "Unexpected ParameterSetName ""$($PSCmdlet.ParameterSetName)""" }
        }
        if ($Role.Count -eq 0) {
            $Role = Get-VmsRole -RoleType UserDefined
        }
        foreach ($r in $Role) {
            $invokeInfo = $ViewGroup.ChangeSecurityPermissions($r.Path)
            if ($null -eq $invokeInfo) {
                Write-Error "Permissions can not be read or modified on view group ""$($ViewGroup.DisplayName)""."
                continue
            }
            $acl = [VmsViewGroupAcl]@{
                Role               = $r
                Path               = $ViewGroup.Path
                SecurityAttributes = @{}
            }
            foreach ($key in $invokeInfo.GetPropertyKeys()) {
                if ($key -eq 'UserPath') { continue }
                $acl.SecurityAttributes[$key] = $invokeInfo.GetProperty($key)
            }
            Write-Output $acl
        }
    }
}


function Import-VmsHardware {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [string[]]
        $Path,
        
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'LiteralPath')]
        [string[]]
        $LiteralPath,

        [Parameter()]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer]
        $RecordingServer,

        [Parameter()]
        [pscredential[]]
        $Credential,

        [Parameter()]
        [switch]
        $UpdateExisting,

        [Parameter()]
        [char]
        $Delimiter = ','
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $validExtensions = '.csv', '.xlsx'
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $LiteralPath = $Path | ForEach-Object {
                $ExecutionContext.SessionState.Path.GetResolvedPSPathFromPSPath($_)
            } | ForEach-Object {
                $fileInfo = [io.fileinfo](Split-Path $_.Path -Leaf)
                if ($fileInfo.Extension -notin $validExtensions) {
                    throw "Invalid file extension $($fileInfo.Extension). Valid extensions include $($validExtensions -join ', ')"
                }
                $_.Path
            }
        }

        
        foreach ($filePath in $LiteralPath) {
            $splat = @{
                Path           = $filePath
                UpdateExisting = $UpdateExisting
            }
            if ($Credential.Count -gt 0) {
                $splat.Credential = $Credential
            }
            if ($null -ne $RecordingServer) {
                $splat.RecordingServer = $RecordingServer
            }
            $fileInfo = [io.fileinfo](Split-Path $filePath -Leaf)
            switch ($fileInfo.Extension) {
                '.csv' {
                    $splat.Delimiter = $Delimiter
                    ImportHardwareCsv @splat
                }

                '.xlsx' {
                    if ($null -eq (Get-Module ImportExcel)) {
                        if (Get-module ImportExcel -ListAvailable) {
                            Import-Module ImportExcel
                        } else {
                            Import-Module "$PSScriptRoot\modules\ImportExcel\7.8.9\ImportExcel.psd1"
                        }
                    }
                    Import-VmsHardwareExcel @splat
                }

                default {
                    throw "Support for file extension $_ not implemented."
                }
            }
        }
    }
}

function Import-VmsLicense {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (-not (Test-Path $filePath)) {
                throw [System.IO.FileNotFoundException]::new('Import-VmsLicense could not find the file.', $filePath)
            }
            $bytes = [IO.File]::ReadAllBytes($filePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $ms = Get-VmsManagementServer
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].UpdateLicense($b64)
            if ($result.State -eq 'Success') {
                $ms.LicenseInformationFolder.ClearChildrenCache()
                Write-Output $ms.LicenseInformationFolder.LicenseInformations[0]
            }
            else {
                Write-Error "Failed to import updated license file. $($result.ErrorText.Trim('.'))."
            }
        }
        catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}


function Import-VmsViewGroup {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ViewGroup])]
    param(
        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $NewName,

        [Parameter()]
        [ValidateNotNull()]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $ParentViewGroup
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        [environment]::CurrentDirectory = Get-Location
        $Path = [io.path]::GetFullPath($Path)

        $source = [io.file]::ReadAllText($Path) | ConvertFrom-Json -ErrorAction Stop
        if ($source.ItemType -ne 'ViewGroup') {
            throw "Invalid file specified in Path parameter. File must be in JSON format and the root object must have an ItemType value of ViewGroup."
        }
        if ($MyInvocation.BoundParameters.ContainsKey('NewName')) {
            ($source.Properties | Where-Object Key -eq 'Name').Value = $NewName
        }
        $params = @{
            Source = $source
        }
        if ($MyInvocation.BoundParameters.ContainsKey('ParentViewGroup')) {
            $params.ParentViewGroup = $ParentViewGroup
        }
        Copy-ViewGroupFromJson @params
    }
}


function Join-VmsDeviceGroupPath {
    [CmdletBinding()]
    [OutputType([string])]
    [RequiresVmsConnection($false)]
    param (
        # Specifies a device group path in unix directory form with forward-slashes as separators.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [string[]]
        $PathParts
    )

    begin {
        Assert-VmsRequirementsMet
        $sb = [text.stringbuilder]::new()
    }

    process {

        foreach ($part in $PathParts) {
            $part | Foreach-Object {
                $null = $sb.Append('/{0}' -f ($_ -replace '(?<!`)/', '`/'))
            }
        }
    }

    end {
        $sb.ToString()
    }
}


function New-VmsBasicUser {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.BasicUser])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [SecureStringTransformAttribute()]
        [securestring]
        $Password,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Description,

        [Parameter(ValueFromPipelineByPropertyName)]
        [BoolTransformAttribute()]
        [bool]
        $CanChangePassword = $true,

        [Parameter(ValueFromPipelineByPropertyName)]
        [BoolTransformAttribute()]
        [bool]
        $ForcePasswordChange,

        [Parameter(ValueFromPipelineByPropertyName)]
        [ValidateSet('Enabled', 'LockedOutByAdmin')]
        [string]
        $Status = 'Enabled'
    )

    begin {
        Assert-VmsRequirementsMet
        $ms = Get-VmsManagementServer
    }

    process {
        try {
            $result = $ms.BasicUserFolder.AddBasicUser($Name, $Description, $CanChangePassword, $ForcePasswordChange, $Password, $Status)
            [VideoOS.Platform.ConfigurationItems.BasicUser]::new($ms.ServerId, $result.Path)
        } catch {
            Write-Error -ErrorRecord $_
        }
    }
}


function New-VmsDeviceGroup {
    [CmdletBinding()]
    [Alias('Add-DeviceGroup')]
    [OutputType([VideoOS.Platform.ConfigurationItems.IConfigurationItem])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, Position = 0, ParameterSetName = 'ByName')]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'SpeakerGroup', 'MetadataGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $ParentGroup,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ByName')]
        [string[]]
        $Name,

        [Parameter(Mandatory, Position = 2, ParameterSetName = 'ByPath')]
        [string[]]
        $Path,

        [Parameter(Position = 3, ParameterSetName = 'ByName')]
        [Parameter(Position = 3, ParameterSetName = 'ByPath')]
        [string]
        $Description,

        [Parameter(Position = 4, ParameterSetName = 'ByName')]
        [Parameter(Position = 4, ParameterSetName = 'ByPath')]
        [Alias('DeviceCategory')]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Input', 'Output', 'Metadata')]
        [string]
        $Type = 'Camera'
    )

    begin {
        Assert-VmsRequirementsMet
        $adjustedType = $Type
        if ($adjustedType -eq 'Input') {
            # Inputs on cameras have an object type called "InputEvent"
            # but we don't want the user to have to remember that.
            $adjustedType = 'InputEvent'
        }
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                $getGroupParams = @{
                    Type = $Type
                }
                $rootGroup = Get-VmsManagementServer
                if ($ParentGroup) {
                    $getGroupParams.ParentGroup = $ParentGroup
                    $rootGroup = $ParentGroup
                }
                foreach ($n in $Name) {
                    try {
                        $getGroupParams.Name = $n
                        $group = Get-VmsDeviceGroup @getGroupParams -ErrorAction SilentlyContinue
                        if ($null -eq $group) {
                            $serverTask = $rootGroup."$($adjustedType)GroupFolder".AddDeviceGroup($n, $Description)
                            $rootGroup."$($adjustedType)GroupFolder".ClearChildrenCache()
                            New-Object -TypeName "VideoOS.Platform.ConfigurationItems.$($adjustedType)Group" -ArgumentList $rootGroup.ServerId, $serverTask.Path
                        } else {
                            $group
                        }
                    } catch {
                        Write-Error -ErrorRecord $_
                    }
                }
            }
            'ByPath' {
                $params = @{
                    Type = $Type
                }
                foreach ($p in $Path) {
                    try {
                        $skip = 0
                        $pathPrefixPattern = '^/(?<type>(Camera|Microphone|Speaker|Metadata|Input|Output))(Event)?GroupFolder'
                        if ($p -match $pathPrefixPattern) {
                            $pathPrefix = $p -replace '^/(Camera|Microphone|Speaker|Metadata|Input|Output)(Event)?GroupFolder.*', '$1'
                            if ($pathPrefix -ne $params.Type) {
                                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Type')) {
                                    throw "The device group prefix '$pathPrefix' does not match the specified device group type '$Type'. Either remove the prefix from the path, or do not specify a value for the Type parameter."
                                } else {
                                    Write-Verbose "Device type '$pathPrefix' determined from the provided path."
                                    $params.Type = $pathPrefix
                                }
                            }
                            $skip = 1
                        }
                        $p | Split-VmsDeviceGroupPath | Select-Object -Skip $skip | ForEach-Object {
                            $params.Remove('Name')
                            $group = Get-VmsDeviceGroup @params -Name ($_ -replace '([\*\?\[\]])', '`$1') -ErrorAction SilentlyContinue
                            $params.Name = $_
                            if ($null -eq $group) {
                                $group = New-VmsDeviceGroup @params -ErrorAction Stop
                            }
                            $params.ParentGroup = $group
                        }
                        if (-not [string]::IsNullOrWhiteSpace($Description)) {
                            $group.Description = $Description
                            $group.Save()
                        }
                        $group
                    } catch {
                        Write-Error -ErrorRecord $_
                    }
                }
            }
            Default {
                throw "Parameter set '$_' not implemented."
            }
        }
    }
}

function New-VmsLoginProvider {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.LoginProvider])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory)]
        [string]
        $Name,

        [Parameter(Mandatory)]
        [string]
        $ClientId,

        [Parameter(Mandatory)]
        [SecureStringTransformAttribute()]
        [securestring]
        $ClientSecret,

        [Parameter()]
        [string]
        $CallbackPath = '/signin-oidc',

        [Parameter(Mandatory)]
        [uri]
        $Authority,

        [Parameter()]
        [string]
        $UserNameClaim,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Scopes = @(),

        [Parameter()]
        [bool]
        $PromptForLogin = $true,

        [Parameter()]
        [bool]
        $Enabled = $true
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $credential = [pscredential]::new($ClientId, $ClientSecret)
            $folder = (Get-VmsManagementServer).LoginProviderFolder
            $serverTask = $folder.AddLoginProvider([guid]::Empty, $Name, $ClientId, $credential.GetNetworkCredential().Password, $CallbackPath, $Authority, $UserNameClaim, $Scopes, $PromptForLogin, $Enabled)
            $loginProvider = Get-VmsLoginProvider | Where-Object Path -eq $serverTask.Path
            if ($null -ne $loginProvider) {
                $loginProvider
            }
        } catch {
            Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $serverTask
        }
    }
}


function New-VmsView {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup]
        $ViewGroup,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Position = 2)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Cameras,

        [Parameter(ParameterSetName = 'Default')]
        [string]
        $StreamName,

        [Parameter(ParameterSetName = 'Custom')]
        [ValidateRange(1, 100)]
        [int]
        $Columns,

        [Parameter(ParameterSetName = 'Custom')]
        [ValidateRange(1, 100)]
        [int]
        $Rows,

        [Parameter(ParameterSetName = 'Advanced')]
        [string]
        $LayoutDefinitionXml,

        [Parameter(ParameterSetName = 'Advanced')]
        [string[]]
        $ViewItemDefinitionXml
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            if ($null -eq $ViewGroup.ViewFolder) {
                throw "Top-level view groups cannot contain views. Views may only be added to child view groups."
            }
            switch ($PSCmdlet.ParameterSetName) {
                'Default' { $LayoutDefinitionXml = New-VmsViewLayout -ViewItemCount $Cameras.Count }
                'Custom'  { $LayoutDefinitionXml = New-VmsViewLayout -Columns $Columns -Rows $Rows }
            }

            $invokeInfo = $ViewGroup.ViewFolder.AddView($LayoutDefinitionXml)
            if ($invokeInfo.State -ne 'Success') {
                throw $invokeInfo.ErrorText
            }
            $invokeInfo.SetProperty('Name', $Name)
            $invokeResult = $invokeInfo.ExecuteDefault()
            if ($invokeResult.State -ne 'Success') {
                throw $invokeResult.ErrorText
            }
            $ViewGroup.ViewFolder.ClearChildrenCache()
            $view = $ViewGroup.ViewFolder.Views | Where-Object Path -eq $invokeResult.Path
            $dirty = $false

            if ($PSCmdlet.ParameterSetName -ne 'Advanced') {
                $smartClientId = GetSmartClientId -View $view
                $i = 0
                if ($Cameras.Count -gt $view.ViewItemChildItems.Count) {
                    Write-Warning "The view is not large enough for the number of cameras selected. Only the first $($view.ViewItemChildItems.Count) of $($Cameras.Count) cameras will be included."
                }
                foreach ($cam in $Cameras) {
                    $streamId = [guid]::Empty
                    if (-not [string]::IsNullOrWhiteSpace($StreamName)) {
                        $stream = $cam | Get-VmsCameraStream | Where-Object DisplayName -eq $StreamName | Select-Object -First 1

                        if ($null -ne $stream) {
                            $streamId = $stream.StreamReferenceId
                        } else {
                            Write-Warning "Stream named ""$StreamName"" not found on $($cam.Name). Default live stream will be used instead."
                        }
                    }
                    $properties = $cam | New-VmsViewItemProperties -SmartClientId $smartClientId
                    $properties.LiveStreamId = $streamId
                    $viewItemDefinition = $properties | New-CameraViewItemDefinition
                    $view.ViewItemChildItems[$i++].SetProperty('ViewItemDefinitionXml', $viewItemDefinition)
                    $dirty = $true
                    if ($i -ge $view.ViewItemChildItems.Count) {
                        break
                    }
                }
            } else {
                for ($i = 0; $i -lt $ViewItemDefinitionXml.Count; $i++) {
                    $view.ViewItemChildItems[$i].SetProperty('ViewItemDefinitionXml', $ViewItemDefinitionXml[$i])
                    $dirty = $true
                }
            }

            if ($dirty) {
                $view.Save()
            }
            Write-Output $view
        } catch {
            Write-Error $_
        }
    }
}

function GetSmartClientId ($View) {
    $id = New-Guid
    if ($view.ViewItemChildItems[0].GetProperty('ViewItemDefinitionXml') -match 'smartClientId="(?<id>.{36})"') {
        $id = $Matches.id
    }
    Write-Output $id
}


function New-VmsViewGroup {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ViewGroup])]
    param (
        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.ViewGroup]
        $Parent,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $vgFolder = (Get-VmsManagementServer).ViewGroupFolder
        if ($null -ne $Parent) {
            $vgFolder = $Parent.ViewGroupFolder
        }
        if ($Force) {
            $vg = $vgFolder.ViewGroups | Where-Object DisplayName -eq $Name
            if ($null -ne $vg) {
                Write-Output $vg
                return
            }
        }
        try {
            $result = $vgFolder.AddViewGroup($Name, $Description)
            if ($result.State -eq 'Success') {
                $vgFolder.ClearChildrenCache()
                Get-VmsViewGroup -Name $Name -Parent $Parent
            } else {
                Write-Error $result.ErrorText
            }
        } catch {
            if ($Force -and $_.Exception.Message -like '*Group name already exist*') {
                Get-VmsViewGroup -Name $Name
            } else {
                Write-Error $_
            }
        }
    }
}


function Remove-VmsBasicUser {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[BasicUser]])]
        [MipItemTransformation([BasicUser])]
        [BasicUser[]]
        $InputObject
    )

    begin {
        Assert-VmsRequirementsMet
        $folder = (Get-VmsManagementServer).BasicUserFolder
    }

    process {
        foreach ($user in $InputObject) {
            $target = "Basic user $($InputObject.Name)"
            if ($user.IsExternal) {
                $target += " <External IDP>"
            }
            if ($PSCmdlet.ShouldProcess($target, "Remove")) {
                try {
                    $null = $folder.RemoveBasicUser($user.Path)
                } catch {
                    Write-Error -Message $_.Exception.Message -TargetObject $user
                }
            }
        }
    }
}


function Remove-VmsDeviceGroup {
    [CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess)]
    [Alias('Remove-DeviceGroup')]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline)]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'MetadataGroup', 'SpeakerGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Group,

        [Parameter()]
        [switch]
        $Recurse
    )

    begin {
        Assert-VmsRequirementsMet
        $cacheToClear = @{}
    }

    process {
        foreach ($g in $Group) {
            $itemType = $g | Split-VmsConfigItemPath -ItemType
            $target = "$itemType '$($g.Name)'"
            $action = "Delete"
            if ($Recurse) {
                $target += " and all group members"
            }
            if ($PSCmdlet.ShouldProcess($target, $action)) {
                try {
                    $parentFolder = Get-ConfigurationItem -Path $g.ParentPath
                    $invokeInfo = $parentFolder | Invoke-Method -MethodId RemoveDeviceGroup
                    if ($Recurse -and ($prop = $invokeInfo.Properties | Where-Object Key -eq 'RemoveMembers')) {
                        $prop.Value = $Recurse.ToString()
                    } elseif ($Recurse) {
                        # Versions around 2019 and older apparently didn't have a "RemoveMembers" option for recursively deleting device groups.
                        $members = $g | Get-VmsDeviceGroupMember -EnableFilter All
                        if ($members.Count -gt 0) {
                            $g | Remove-VmsDeviceGroupMember -Device $members -Confirm:$false
                        }
                        $g | Get-VmsDeviceGroup | Remove-VmsDeviceGroup -Recurse -Confirm:$false
                    }

                    ($invokeInfo.Properties | Where-Object Key -eq 'ItemSelection').Value = $g.Path
                    $null = $invokeInfo | Invoke-Method -MethodId RemoveDeviceGroup -ErrorAction Stop
                    $cacheToClear[$itemType] = $null
                } catch {
                    Write-Error -ErrorRecord $_
                }
            }
        }

    }

    end {
        $cacheToClear.Keys | Foreach-Object {
            Write-Verbose "Clearing $_ cache"
            (Get-VmsManagementServer)."$($_)Folder".ClearChildrenCache()
        }
    }
}


function Remove-VmsDeviceGroupMember {
    [CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess)]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'MetadataGroup', 'SpeakerGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Group,

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'ByObject')]
        [ValidateVmsItemType('Camera', 'Microphone', 'Metadata', 'Speaker', 'InputEvent', 'Output')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Device,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ById')]
        [guid[]]
        $DeviceId
    )
    
    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        $groupItemType = ($Group | Split-VmsConfigItemPath -ItemType) -replace 'Group$', ''
        $dirty = $false
        if ($Device) {
            $DeviceId = [guid[]]$Device.Id
            $map = @{}; $Device | Foreach-Object { $map[[guid]$_.Id] = $_ }
        }
        if ($PSCmdlet.ShouldProcess("$groupItemType group '$($Group.Name)'", "Remove $($DeviceId.Count) device group member(s)")) {
            foreach ($id in $DeviceId) {
                try {
                    $path = '{0}[{1}]' -f $groupItemType, $id
                    $null = $Group."$($groupItemType)Folder".RemoveDeviceGroupMember($path)
                    $dirty = $true
                } catch [VideoOS.Platform.ArgumentMIPException] {
                    Write-Error -Message "Failed to remove device group member: $_.Exception.Message" -Exception $_.Exception
                }
            }
        }
    }

    end {
        if ($dirty) {
            $Group."$($groupItemType)GroupFolder".ClearChildrenCache()
            (Get-VmsManagementServer)."$($groupItemType)GroupFolder".ClearChildrenCache()
        }
    }
}


function Remove-VmsHardware {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
    [Alias('Remove-Hardware')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware
    )

    begin {
        Assert-VmsRequirementsMet
        $recorders = @{}
        (Get-VmsManagementServer).RecordingServerFolder.RecordingServers | Foreach-Object {
            $recorders[$_.Path] = $_
        }
        $foldersNeedingCacheReset = @{}
    }

    process {
        try {
            $action = 'Permanently delete hardware and all associated video, audio and metadata from the VMS'
            foreach ($hw in $Hardware) {
                try {
                    $target = "$($hw.Name) with ID $($hw.Id)"
                    if ($PSCmdlet.ShouldProcess($target, $action)) {
                        $folder = $recorders[$hw.ParentItemPath].HardwareFolder
                        $result = $folder.DeleteHardware($hw.Path) | Wait-VmsTask -Title "Removing hardware $($hw.Name)" -Cleanup
                        $properties = @{}
                        $result.Properties | Foreach-Object { $properties[$_.Key] = $_.Value}
                        if ($properties.State -eq 'Success') {
                            $foldersNeedingCacheReset[$folder.Path] = $folder
                        } else {
                            Write-Error "An error occurred while deleting the hardware. $($properties.ErrorText.Trim('.'))."
                        }
                    }
                }
                catch [VideoOS.Platform.PathNotFoundMIPException] {
                    Write-Error "The hardware named $($hw.Name) with ID $($hw.Id) was not found."
                }
            }
        }
        catch [VideoOS.Platform.PathNotFoundMIPException] {
            Write-Error "One or more recording servers for the provided hardware values do not exist."
        }
    }

    end {
        $foldersNeedingCacheReset.Values | Foreach-Object {
            $_.ClearChildrenCache()
        }
    }
}


function Remove-VmsLoginProvider {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ShouldProcess("Login provider '$($LoginProvider.Name)'", 'Remove')) {
            if ($Force) {
                # Disable the login provider to ensure no external users login
                # and generate a new external basic user between the time the users
                # are removed and the provider is deleted.
                $LoginProvider | Set-VmsLoginProvider -Enabled $false -ErrorAction Stop -Verbose:($VerbosePreference -eq 'Continue')

                # The basic user folder may be cached already, and there may be
                # new external users on the VMS that are not present in the cache.
                # By clearing the cache we ensure that the next step removes all
                # external users.
                (Get-VmsManagementServer).BasicUserFolder.ClearChildrenCache()

                # Remove all basic users with claims associated with this login provider
                Get-VmsBasicUser -External | Where-Object {
                    $_.ClaimFolder.ClaimChildItems.ClaimProvider -contains $LoginProvider.Id
                } | Remove-VmsBasicUser -ErrorAction Stop -Verbose:($VerbosePreference -eq 'Continue')

                # Remove all claims associated with this login provider from all roles
                foreach ($role in Get-VmsRole) {
                    $claims = $role | Get-VmsRoleClaim | Where-Object ClaimProvider -EQ $LoginProvider.Id
                    if ($claims.Count -gt 0) {
                        $role | Remove-VmsRoleClaim -ClaimName $claims.ClaimName -ErrorAction Stop -Verbose:($VerbosePreference -eq 'Continue')
                    }
                }

                # Remove all claims registered on this login provider
                $LoginProvider | Remove-VmsLoginProviderClaim -All -ErrorAction Stop
            }
            $null = (Get-VmsManagementServer).LoginProviderFolder.RemoveLoginProvider($LoginProvider.Path)
        }
    }
}

function Remove-VmsLoginProviderClaim {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'All')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'Name')]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter(Mandatory, ParameterSetName = 'All')]
        [switch]
        $All,

        [Parameter(Mandatory, ParameterSetName = 'Name')]
        [string]
        $ClaimName,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($Force) {
            Get-VmsRole | Foreach-Object {
                $currentRole = $_
                $claims = $currentRole | Get-VmsRoleClaim -LoginProvider $LoginProvider | Where-Object {
                    $All -or $_.ClaimName -eq $ClaimName
                }
                if ($claims.Count -eq 0) {
                    return
                }
                $currentRole | Remove-VmsRoleClaim -ClaimName $claims.ClaimName
            }
        }
        $folder = $LoginProvider.RegisteredClaimFolder
        $LoginProvider | Get-VmsLoginProviderClaim | Foreach-Object {
            if (-not [string]::IsNullOrWhiteSpace($ClaimName) -and $_.Name -notlike $ClaimName) {
                return
            }
            if ($PSCmdlet.ShouldProcess("Registered claim '$($_.DisplayName)'", "Remove")) {
                $null = $folder.RemoveRegisteredClaim($_.Path)
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsLoginProviderClaim -ParameterName ClaimName -ScriptBlock {
    $values = (Get-VmsLoginProvider | Get-VmsLoginProviderClaim).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Remove-VmsView {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.View])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.View[]]
        $View
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($v in $View) {
            if ($PSCmdlet.ShouldProcess($($v.Name), "Remove view")) {
                $viewFolder = [VideoOS.Platform.ConfigurationItems.ViewFolder]::new($v.ServerId, $v.ParentPath)
                $result = $viewFolder.RemoveView($v.Path)
                if ($result.State -ne 'Success') {
                    Write-Error $result.ErrorText
                }
            }
        }
    }
}


function Remove-VmsViewGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup[]]
        $ViewGroup,

        [Parameter()]
        [switch]
        $Recurse
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($vg in $ViewGroup) {
            if ($PSCmdlet.ShouldProcess($vg.DisplayName, "Remove ViewGroup")) {
                try {
                    $viewGroupFolder = [VideoOS.Platform.ConfigurationItems.ViewGroupFolder]::new($vg.ServerId, $vg.ParentPath)
                    $result = $viewGroupFolder.RemoveViewGroup($Recurse, $vg.Path)
                    if ($result.State -eq 'Success') {
                        $viewGroupFolder.ClearChildrenCache()
                    } else {
                        Write-Error $result.ErrorText
                    }
                } catch {
                    Write-Error $_
                }
            }
        }
    }
}


function Resolve-VmsDeviceGroupPath {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('DeviceGroup')]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'SpeakerGroup', 'MetadataGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Group,

        [Parameter()]
        [switch]
        $NoTypePrefix
    )

    begin {
        Assert-VmsRequirementsMet
        $ctor = $null
        $sb = [text.stringbuilder]::new()
    }

    process {
        if ($null -eq $ctor -or $ctor.ReflectedType -ne $Group.GetType()) {
            $ctor = $Group.GetType().GetConstructor(@([videoos.platform.serverid], [string]))
        }
        try {
            $current = $Group
            $null = $sb.Clear().Insert(0, "/$($current.Name -replace '(?<!`)/', '`/')")
            while ($current.ParentItemPath -ne '/') {
                $current = $ctor.Invoke(@($current.ServerId, $current.ParentItemPath))
                $null = $sb.Insert(0, "/$($current.Name -replace '(?<!`)/', '`/')")
            }
            if (-not $NoTypePrefix) {
                $null = $sb.Insert(0, $current.ParentPath)
            }
            $sb.ToString()
        } catch {
            throw
        }
    }
}

function Set-VmsBasicUser {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.BasicUser])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[BasicUser]])]
        [MipItemTransformation([BasicUser])]
        [BasicUser]
        $BasicUser,

        [Parameter()]
        [SecureStringTransformAttribute()]
        [securestring]
        $Password,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [BoolTransformAttribute()]
        [bool]
        $CanChangePassword,

        [Parameter()]
        [BoolTransformAttribute()]
        [bool]
        $ForcePasswordChange,

        [Parameter()]
        [ValidateSet('Enabled', 'LockedOutByAdmin')]
        [string]
        $Status,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            if ($PSCmdlet.ShouldProcess("Basic user '$($BasicUser.Name)'", "Update")) {
                $dirty = $false
                $dirtyPassword = $false
                $initialName = $BasicUser.Name
                foreach ($key in @(($MyInvocation.BoundParameters.GetEnumerator() | Where-Object Key -in $BasicUser.GetPropertyKeys()).Key) + @('Password')) {
                    $newValue = (Get-Variable -Name $key).Value
                    if ($MyInvocation.BoundParameters.ContainsKey('Password') -and $key -eq 'Password') {
                        if ($BasicUser.IsExternal -or -not $BasicUser.CanChangePassword) {
                            Write-Error "Password can not be changed for '$initialName'. IsExternal = $($BasicUser.IsExternal), CanChangePassword = $($BasicUser.CanChangePassword)" -TargetObject $BasicUser
                        } else {
                            Write-Verbose "Updating $key on '$initialName'"
                            $null = $BasicUser.ChangePasswordBasicUser($Password)
                            $dirtyPassword = $true
                        }
                    } elseif ($BasicUser.$key -cne $newValue) {
                        Write-Verbose "Updating $key on '$initialName'"
                        $BasicUser.$key = $newValue
                        $dirty = $true
                    }
                }
                if ($dirty) {
                    $BasicUser.Save()
                } elseif (-not $dirtyPassword) {
                    Write-Verbose "No changes were made to '$initialName'."
                }
            }

            if ($PassThru) {
                $BasicUser
            }
        } catch {
            Write-Error -Message $_.Exception.Message -TargetObject $BasicUser
        }
    }
}


function Set-VmsCameraStream {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ParameterSetName = 'RemoveStream')]
        [switch]
        $Disabled,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'AddOrUpdateStream')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'RemoveStream')]
        [VmsCameraStreamConfig[]]
        $Stream,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [string]
        $DisplayName,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [ValidateSet('Always', 'Never', 'WhenNeeded')]
        [string]
        $LiveMode,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [switch]
        $LiveDefault,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [switch]
        $Recorded,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [ValidateSet('Primary', 'Secondary', 'None')]
        [string]
        $RecordingTrack,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [ValidateVmsVersion('23.2')]
        [ValidateVmsFeature('MultistreamRecording')]
        [switch]
        $PlaybackDefault,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [switch]
        $UseEdge,

        [Parameter(ParameterSetName = 'AddOrUpdateStream')]
        [hashtable]
        $Settings
    )

    begin {
        Assert-VmsRequirementsMet

        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Recorded') -and $Recorded) {
            Write-Warning "The 'Recorded' switch parameter is deprecated with MilestonePSTools version 2023 R2 and later due to the added support for adaptive playback. For compatibility reasons, the '-Recorded' switch has the same meaning as '-RecordingTrack Primary -PlaybackDefault' unless one or both of these parameters were also specified."
            if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecordingTrack')) {
                Write-Verbose "Setting RecordingTrack parameter to 'Primary'"
                $PSCmdlet.MyInvocation.BoundParameters['RecordingTrack'] = $RecordingTrack = 'Primary'
            }
            if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('PlaybackDefault')) {
                Write-Verbose "Setting PlaybackDefault parameter to `$true"
                $PSCmdlet.MyInvocation.BoundParameters['PlaybackDefault'] = $PlaybackDefault = [switch]::new($true)
            }
            $null = $PSCmdlet.MyInvocation.BoundParameters.Remove('Recorded')
            Remove-Variable -Name 'Recorded'
        }
        $updatedItems = [system.collections.generic.list[pscustomobject]]::new()
        $itemCache = @{}
    }

    process {
        foreach ($s in $Stream) {
            $target = "$($s.Name) on $($s.Camera.Name)"
            $deviceDriverSettings = $s.Camera.DeviceDriverSettingsFolder.DeviceDriverSettings[0]
            if ($itemCache.ContainsKey($deviceDriverSettings.Path)) {
                $deviceDriverSettings = $itemCache[$deviceDriverSettings.Path]
            } else {
                $itemCache[$deviceDriverSettings.Path] = $deviceDriverSettings
            }
            $streamUsages = $s.Camera.StreamFolder.Streams | Select-Object -First 1
            if ($null -ne $streamUsages -and $itemCache.ContainsKey($streamUsages.Path)) {
                $streamUsages = $itemCache[$streamUsages.Path]
            } elseif ($null -ne $streamUsages) {
                $itemCache[$streamUsages.Path] = $streamUsages
            }

            $streamRefToName = @{}
            if ($streamUsages.StreamUsageChildItems.Count -gt 0) {
                $streamNameToRef = $streamUsages.StreamUsageChildItems[0].StreamReferenceIdValues
                foreach ($key in $streamNameToRef.Keys) {
                    $streamRefToName[$streamNameToRef.$key] = $key
                }
                $streamUsageChildItem = $streamUsages.StreamUsageChildItems | Where-Object StreamReferenceId -eq $streamNameToRef[$s.Name]
            }

            if ($PSCmdlet.ParameterSetName -eq 'RemoveStream' -and $null -ne $streamUsageChildItem -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling stream '$($s.Name)'")) {
                if ($streamUsages.StreamUsageChildItems.Count -eq 1) {
                    Write-Error "Stream $($s.Name) cannot be removed because it is the only enabled stream."
                } else {
                    $result = $streamUsages.RemoveStream($streamUsageChildItem.StreamReferenceId)
                    if ($result.State -eq 'Success') {
                        $s.Update()
                        $streamUsages = $s.Camera.StreamFolder.Streams[0]
                        $itemCache[$streamUsages.Path] = $streamUsages
                    } else {
                        Write-Error $result.ErrorText
                    }
                }
            } elseif ($PSCmdlet.ParameterSetName -eq 'AddOrUpdateStream') {
                $dirtyStreamUsages = $false
                $parametersRequiringStreamUsage = @('DisplayName', 'LiveDefault', 'LiveMode', 'PlaybackDefault', 'Recorded', 'RecordingTrack', 'UseEdge')
                if ($null -eq $streamUsageChildItem -and ($PSCmdlet.MyInvocation.BoundParameters.Keys | Where-Object { $_ -in $parametersRequiringStreamUsage } ) -and $PSCmdlet.ShouldProcess($s.Camera.Name, 'Adding a new stream usage')) {
                    try {
                        $result = $streamUsages.AddStream()
                        if ($result.State -ne 'Success') {
                            throw $result.ErrorText
                        }
                        $s.Update()
                        $streamUsages = $s.Camera.StreamFolder.Streams[0]
                        $itemCache[$streamUsages.Path] = $streamUsages
                        $streamUsageChildItem = $streamUsages.StreamUsageChildItems | Where-Object StreamReferenceId -eq $result.GetProperty('StreamReferenceId')
                        $streamUsageChildItem.StreamReferenceId = $streamNameToRef[$s.Name]
                        $streamUsageChildItem.Name = $s.Name
                        $dirtyStreamUsages = $true
                    } catch {
                        Write-Error $_
                    }
                }

                if ($RecordingTrack -eq 'Secondary' -and $streamUsageChildItem.RecordToValues.Count -eq 0) {
                    Write-Error "Adaptive playback is not available. RecordingTrack parameter must be Primary or None."
                    continue
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('DisplayName') -and $DisplayName -ne $streamUsageChildItem.Name) {
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Setting DisplayName on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.Name = $DisplayName
                    }
                    $dirtyStreamUsages = $true
                }

                $recordingTrackId = @{
                    Primary   = '16ce3aa1-5f93-458a-abe5-5c95d9ed1372'
                    Secondary = '84fff8b9-8cd1-46b2-a451-c4a87d4cbbb0'
                    None      = ''
                }
                $compatibilityRecord = if ($RecordingTrack -eq 'Primary') { $true } else { $false }
                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RecordingTrack') -and (($streamUsageChildItem.RecordToValues.Count -gt 0 -and $recordingTrackId[$RecordingTrack] -ne $streamUsageChildItem.RecordTo) -or ($streamUsageChildItem.RecordToValues.Count -eq 0 -and $compatibilityRecord -ne $streamUsageChildItem.Record))) {
                    if ($streamUsageChildItem.RecordToValues.Count -gt 0) {
                        # 2023 R2 or later
                        $primaryStreamUsage = $streamUsages.StreamUsageChildItems | Where-Object RecordTo -eq $recordingTrackId.Primary
                        $secondaryStreamUsage = $streamUsages.StreamUsageChildItems | Where-Object RecordTo -eq $recordingTrackId.Secondary
                        switch ($RecordingTrack) {
                            'Primary' {
                                if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Record $($streamUsageChildItem.Name) to the primary recording track")) {
                                    $streamUsageChildItem.RecordTo = $recordingTrackId.Primary

                                    Write-Verbose "Disabling recording on current primary stream '$($primaryStreamUsage.Name)'."
                                    $primaryStreamUsage.RecordTo = $recordingTrackId.None

                                    if ($primaryStreamUsage.LiveMode -eq 'Never') {
                                        Write-Verbose "Changing LiveMode from Never to WhenNeeded on $($primaryStreamUsage.Name)"
                                        $primaryStreamUsage.LiveMode = 'WhenNeeded'
                                    }

                                    if ($streamUsageChildItem.LiveMode -eq 'Never') {
                                        Write-Verbose "Changing LiveMode from Never to WhenNeeded on $($streamUsageChildItem.Name)"
                                        $streamUsageChildItem.LiveMode = 'WhenNeeded'
                                    }

                                    $dirtyStreamUsages = $true
                                }
                            }
                            'Secondary' {
                                if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Record $($streamUsageChildItem.Name) to the secondary recording track")) {
                                    $streamUsageChildItem.RecordTo = $recordingTrackId.Secondary
                                    if ($streamUsageChildItem.LiveMode -eq 'Never') {
                                        Write-Verbose "Changing LiveMode from Never to WhenNeeded on $($streamUsageChildItem.Name)"
                                        $streamUsageChildItem.LiveMode = 'WhenNeeded'
                                    }

                                    if ($secondaryStreamUsage) {
                                        Write-Verbose "Disabling recording on current secondary stream '$($secondaryStreamUsage.Name)'."
                                        $secondaryStreamUsage.RecordTo = $recordingTrackId.None

                                        if ($secondaryStreamUsage.LiveMode -eq 'Never') {
                                            Write-Verbose "Changing LiveMode from Never to WhenNeeded on $($secondaryStreamUsage.Name)"
                                            $secondaryStreamUsage.LiveMode = 'WhenNeeded'
                                        }
                                    }

                                    $dirtyStreamUsages = $true
                                }
                            }
                            'None' {
                                if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Disable recording of stream $($streamUsageChildItem.Name)")) {
                                    $streamUsageChildItem.RecordTo = $recordingTrackId.None
                                    if ($streamUsageChildItem.LiveMode -eq 'Never') {
                                        Write-Verbose "Changing LiveMode from Never to WhenNeeded on $($streamUsageChildItem.Name)"
                                        $streamUsageChildItem.LiveMode = 'WhenNeeded'
                                    }

                                    $streamUsages.StreamUsageChildItems | Where-Object {
                                        $_.StreamReferenceId -ne $streamUsageChildItem.StreamReferenceId -and -not [string]::IsNullOrWhiteSpace($_.RecordTo)
                                    } | Select-Object -First 1 | ForEach-Object {
                                        Write-Verbose "Setting the default playback stream to $($_.Name)"
                                        $_.DefaultPlayback = $true
                                    }

                                    $dirtyStreamUsages = $true
                                }
                            }
                        }
                    } else {
                        # 2023 R1 or earlier
                        $recordedStream = $streamUsages.StreamUsageChildItems | Where-Object Record
                        if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling recording on $($recordedStream.Name)")) {
                            $recordedStream.Record = $false
                            if ($recordedStream.LiveMode -eq 'Never' -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Changing LiveMode from Never to WhenNeeded on $($recordedStream.Name)")) {
                                # This avoids a validation exception error.
                                $recordedStream.LiveMode = 'WhenNeeded'
                            }
                        }

                        if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Enabling recording on $($streamUsageChildItem.Name)")) {
                            $streamUsageChildItem.Record = $true
                            $dirtyStreamUsages = $true
                        }
                    }
                }

                if ($PlaybackDefault -and $PlaybackDefault -ne $streamUsageChildItem.DefaultPlayback) {
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Set the default playback stream to $($streamUsageChildItem.Name)")) {
                        $streamUsages.StreamUsageChildItems | ForEach-Object {
                            $_.DefaultPlayback = $false
                        }
                        $streamUsageChildItem.DefaultPlayback = $PlaybackDefault
                        $dirtyStreamUsages = $true
                    }
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('UseEdge') -and $UseEdge -ne $streamUsageChildItem.UseEdge) {
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Enable use of edge storage on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.UseEdge = $UseEdge
                        $dirtyStreamUsages = $true
                    }
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('LiveDefault') -and $LiveDefault -and $LiveDefault -ne $streamUsageChildItem.LiveDefault) {
                    $liveStream = $streamUsages.StreamUsageChildItems | Where-Object LiveDefault
                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Disabling LiveDefault on $($liveStream.Name)")) {
                        $liveStream.LiveDefault = $false
                    }

                    if ($PSCmdlet.ShouldProcess($s.Camera.Name, "Enabling LiveDefault on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.LiveDefault = $true
                        $dirtyStreamUsages = $true
                    }
                }

                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('LiveMode') -and $LiveMode -ne $streamUsageChildItem.LiveMode -and -not [string]::IsNullOrWhiteSpace($LiveMode)) {
                    if ($LiveMode -eq 'Never' -and (-not $streamUsageChildItem.Record -or $streamUsageChildItem.LiveDefault)) {
                        Write-Warning 'The LiveMode property can only be set to "Never" the recorded stream, and only when that stream is not used as the LiveDefault stream.'
                    } elseif ($PSCmdlet.ShouldProcess($s.Camera.Name, "Setting LiveMode on $($streamUsageChildItem.Name)")) {
                        $streamUsageChildItem.LiveMode = $LiveMode
                        $dirtyStreamUsages = $true
                    }
                }

                if ($dirtyStreamUsages -and $PSCmdlet.ShouldProcess($s.Camera.Name, "Saving StreamUsages")) {
                    $updatedItems.Add(
                        [pscustomobject]@{
                            Item         = $streamUsages
                            Parent       = $s.Camera
                            StreamConfig = $s
                        }
                    )
                }

                $streamChildItem = $deviceDriverSettings.StreamChildItems.Where( { $_.DisplayName -eq $s.Name })
                if ($Settings.Keys.Count -gt 0) {
                    $dirty = $false
                    foreach ($key in $Settings.Keys) {
                        if ($key -notin $s.Settings.Keys) {
                            Write-Warning "A setting with the key '$key' was not found for stream $($streamChildItem.DisplayName) on $($s.Camera.Name)."
                            continue
                        }

                        $currentValue = $streamChildItem.Properties.GetValue($key)
                        if ($currentValue -eq $Settings.$key) {
                            continue
                        }

                        if ($PSCmdlet.ShouldProcess($target, "Changing $key from $currentValue to $($Settings.$key)")) {
                            $streamChildItem.Properties.SetValue($key, $Settings.$key)
                            $dirty = $true
                        }
                    }
                    if ($dirty -and $PSCmdlet.ShouldProcess($target, "Save changes")) {
                        $updatedItems.Add(
                            [pscustomobject]@{
                                Item         = $deviceDriverSettings
                                Parent       = $s.Camera
                                StreamConfig = $s
                            }
                        )
                    }
                }
            }
        }
    }

    end {
        $updatedStreamConfigs = [system.collections.generic.list[object]]::new()
        foreach ($update in $updatedItems) {
            try {
                $item = $itemCache[$update.Item.Path]
                if ($null -ne $item) {
                    $item.Save()
                }
                if ($update.StreamConfig -notin $updatedStreamConfigs) {
                    $update.StreamConfig.Update()
                    $updatedStreamConfigs.Add($update.StreamConfig)
                }
            } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                $update.Parent.ClearChildrenCache()
                $_ | HandleValidateResultException -TargetObject $item
            } finally {
                if ($null -ne $item) {
                    $itemCache.Remove($item.Path)
                    $item = $null
                }
            }
        }
    }
}


function Set-VmsConnectionString {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection($false)]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]
        $Component,

        [Parameter(Mandatory, Position = 1)]
        [string]
        $ConnectionString,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq (Get-Item -Path HKLM:\SOFTWARE\VideoOS\Server\ConnectionString -ErrorAction Ignore)) {
            Write-Error "Could not find the registry key 'HKLM:\SOFTWARE\VideoOS\Server\ConnectionString'. This key was introduced in 2022 R3, and this cmdlet is only compatible with VMS versions 2022 R3 and later."
            return
        }

        $currentValue = Get-VmsConnectionString -Component $Component -ErrorAction SilentlyContinue
        if ($null -eq $currentValue) {
            if ($Force) {
                if ($PSCmdlet.ShouldProcess((hostname), "Create new connection string value for $Component")) {
                    $null = New-ItemProperty -Path HKLM:\SOFTWARE\VideoOS\Server\ConnectionString -Name $Component -Value $ConnectionString
                }
            } else {
                Write-Error "A connection string for $Component does not exist. Retry with the -Force switch to create one anyway."
            }
        } else {
            if ($PSCmdlet.ShouldProcess((hostname), "Change connection string value of $Component")) {
                Set-ItemProperty -Path HKLM:\SOFTWARE\VideoOS\Server\ConnectionString -Name $Component -Value $ConnectionString
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Set-VmsConnectionString -ParameterName Component -ScriptBlock {
    $values = Get-Item HKLM:\SOFTWARE\videoos\Server\ConnectionString\ -ErrorAction Ignore | Select-Object -ExpandProperty Property
    if ($values) {
        Complete-SimpleArgument $args $values
    }
}


function Set-VmsDeviceGroup {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [OutputType([VideoOS.Platform.ConfigurationItems.IConfigurationItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateVmsItemType('CameraGroup', 'MicrophoneGroup', 'MetadataGroup', 'SpeakerGroup', 'InputEventGroup', 'OutputGroup')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Group,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $groupType = $Group | Split-VmsConfigItemPath -ItemType
        $dirty = $false
        $keys = $MyInvocation.BoundParameters.Keys | Where-Object { $_ -in @('Name', 'Description') }
        if ($PSCmdlet.ShouldProcess("$groupType '$($Group.Name)", "Update $([string]::Join(', ', $keys))")) {
            foreach ($key in $keys) {
                if ($Group.$key -cne $MyInvocation.BoundParameters[$key]) {
                    $Group.$key = $MyInvocation.BoundParameters[$key]
                    $dirty = $true
                }
            }
            if ($dirty) {
                Write-Verbose "Saving changes to $groupType '$($Group.Name)'"
                $Group.Save()
            } else {
                Write-Verbose "No changes made to $groupType '$($Group.Name)'"
            }
        }
        if ($PassThru) {
            $Group
        }
    }
}


function Set-VmsHardware {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.Hardware])]
    [Alias('Set-HardwarePassword')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware,

        [Parameter()]
        [bool]
        $Enabled,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [uri]
        $Address,

        [Parameter()]
        [string]
        $UserName,

        [Parameter()]
        [Alias('NewPassword')]
        [ValidateVmsVersion('11.3')]
        [SecureStringTransformAttribute()]
        [ValidateScript({
            if ($_.Length -gt 64) {
                throw "The maximum password length is 64 characters. See Get-Help Set-VmsHardware -Online for more information."
            }
            $true
        })]
        [securestring]
        $Password,

        [Parameter()]
        [ValidateVmsVersion('23.2')]
        [switch]
        $UpdateRemoteHardware,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        if ($UpdateRemoteHardware -and -not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Password')) {
            Write-Warning "The UpdateRemoteHardware parameter is ignored because no value was provided for the Password parameter."
        }
    }

    process {
        foreach ($hw in $Hardware) {
            if ($MyInvocation.BoundParameters.ContainsKey('WhatIf') -and $WhatIf -eq $true) {
                # Operate on a separate hardware record to avoid modifying local properties when using WhatIf.
                $hw = Get-VmsHardware -Id $hw.Id
            }
            $initialName = $hw.Name
            $initialAddress = $hw.Address
            $dirty = $false

            foreach ($key in $MyInvocation.BoundParameters.Keys) {
                switch ($key) {
                    'Enabled' {
                        if ($Enabled -ne $hw.Enabled) {
                            Write-Verbose "Changing value of '$key' from $($hw.Enabled) to $Enabled on $initialName."
                            $hw.Enabled = $Enabled
                            $dirty = $true
                        }
                    }

                    'Name' {
                        if ($Name -cne $hw.Name) {
                            Write-Verbose "Changing value of '$key' from $($hw.Name) to $Name."
                            $hw.Name = $Name
                            $dirty = $true
                        }
                    }

                    'Address' {
                        if ($Address -ne [uri]$hw.Address) {
                            Write-Verbose "Changing value of '$key' from $($hw.Address) to $Address on $initialName."
                            $hw.Address = $Address
                            $dirty = $true
                        }
                    }

                    'UserName' {
                        if ($UserName -cne $hw.UserName) {
                            Write-Verbose "Changing value of '$key' from $($hw.UserName) to $UserName on $initialName."
                            $hw.UserName = $UserName
                            $dirty = $true
                        }
                    }

                    'Password' {
                        $action = "Change password in the VMS"
                        if ($UpdateRemoteHardware) {
                            $action += ' and on remote hardware device'
                        }
                        if ($PSCmdlet.ShouldProcess("$initialName", $action)) {
                            try {
                                $invokeResult = $hw.ChangePasswordHardware($Password, $UpdateRemoteHardware.ToBool())
                                if ($invokeResult.Path -match '^Task') {
                                    $invokeResult = $invokeResult | Wait-VmsTask -Title "Updating hardware password for $initialName"
                                }
                                if (($invokeResult.Properties | Where-Object Key -eq 'State').Value -eq 'Error') {
                                    Write-Error -Message "ChangePasswordHardware error: $(($invokeResult.Properties | Where-Object Key -eq 'ErrorText').Value)" -TargetObject $hw
                                }
                            } catch {
                                Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $hw
                            }
                        }
                    }

                    'Description' {
                        if ($Description -cne $hw.Description) {
                            Write-Verbose "Changing value of '$key' on $initialName."
                            $hw.Description = $Description
                            $dirty = $true
                        }
                    }
                }
            }

            $target = "Hardware '$initialName' ($initialAddress)"
            if ($dirty) {
                if ($PSCmdlet.ShouldProcess($target, "Save changes")) {
                    try {
                        $hw.Save()
                    } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                        $errorResults = $_.Exception.InnerException.ValidateResult.ErrorResults
                        if ($null -eq $errorResults -or $errorResults.Count -eq 0) {
                            throw
                        }
                        foreach ($result in $errorResults) {
                            Write-Error -Message "Validation error on property '$($result.ErrorProperty)': $($result.ErrorText)"
                        }
                    } catch {
                        Write-Error -ErrorRecord $_ -Exception $_.Exception -TargetObject $hw
                    }
                }
            }

            if ($PassThru) {
                $hw
            }
        }
    }
}


function Set-VmsHardwareDriver {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Hardware])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware,

        [Parameter()]
        [uri]
        $Address,

        [Parameter()]
        [pscredential]
        $Credential,

        [Parameter()]
        [HardwareDriverTransformAttribute()]
        [VideoOS.Platform.ConfigurationItems.HardwareDriver]
        $Driver,

        [Parameter()]
        [string]
        $CustomDriverData,

        [Parameter()]
        [switch]
        $AllowDeletingDisabledDevices,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        $tasks = [system.collections.generic.list[VideoOS.ConfigurationApi.ClientService.ConfigurationItem]]::new()
        $taskInfo = @{}
        $recorderPathByHwPath = @{}
    }

    process {
        $hwParams = @{
            AllowDeletingDisabledDevices = $AllowDeletingDisabledDevices.ToString()
        }

        if ($MyInvocation.BoundParameters.ContainsKey('Address')) {
            if ($Address.Scheme -notin 'https', 'http') {
                Write-Error "Address must be in the format http://address or https://address"
                return
            }
            $hwParams.Address   = $Address.Host
            $hwParams.Port      = if ($Address.Scheme -eq 'http') { $Address.Port } else { 80 }
            $hwParams.UseHttps  = if ($Address.Scheme -eq 'https') { 'True' } else { 'False' }
            $hwParams.HttpsPort = if ($Address.Scheme -eq 'https') { $Address.Port } else { 443 }
        }

        if ($MyInvocation.BoundParameters.ContainsKey('Credential')) {
            $hwParams.UserName = $Credential.UserName
            $hwParams.Password = $Credential.GetNetworkCredential().Password
        } else {
            $hwParams.UserName = $Hardware.UserName
            $hwParams.Password = $Hardware | Get-VmsHardwarePassword
        }

        if ($MyInvocation.BoundParameters.ContainsKey('Driver')) {
            $hwParams.Driver = $Driver.Number.ToString()
        }

        if ($MyInvocation.BoundParameters.ContainsKey('CustomDriverData')) {
            $hwParams.CustomDriverData = $CustomDriverData
        }

        foreach ($hw in $Hardware) {
            if ($PSCmdlet.ShouldProcess("$($hw.Name) ($($hw.Address))", "Replace hardware")) {
                $recorderPathByHwPath[$hw.Path] = [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]::new($hw.ParentItemPath)
                $method = 'ReplaceHardware'
                $item = $hw | Get-ConfigurationItem
                if ($method -notin $item.MethodIds) {
                    throw "The $method MethodId is not present. This method was introduced in XProtect VMS version 2023 R1."
                }
                $invokeInfo = $item | Invoke-Method -MethodId $method

                foreach ($key in $hwParams.Keys) {
                    if ($prop = $invokeInfo.Properties | Where-Object Key -eq $key) {
                        $prop.Value = $hwParams[$key]
                    }
                }

                Write-Verbose "ReplaceHardware task properties`r`n$($invokeInfo.Properties | Select-Object Key, @{Name = 'Value'; Expression = {if ($_.Key -eq 'Password') {'*' * 8} else {$_.Value}}} | Out-String)"
                $invokeResult = $invokeInfo | Invoke-Method ReplaceHardware
                $taskPath = ($invokeResult.Properties | Where-Object Key -eq 'Path').Value
                $tasks.Add((Get-ConfigurationItem -Path $taskPath))
                $taskInfo[$taskPath] = @{
                    HardwareName = $hw.Name
                    HardwarePath = [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]::new($hw.Path)
                    RecorderPath = $recorderPathByHwPath[$hw.Path]
                    Task         = $null
                }
            }
        }
    }

    end {
        $recorders = @{}
        $replacedHardwarePaths = [system.collections.generic.list[string]]::new()
        foreach ($task in $tasks) {
            $task = $task | Wait-VmsTask -Cleanup
            if (($task.Properties | Where-Object Key -eq 'State').Value -ne 'Success') {
                $info = $taskInfo[$task.Path]
                $info.Task = $task
                $message = "Unknown error during ReplaceHardware for $($info.HardwareName) ($info.HardwarePath.Id)."
                $taskError = ($task.Properties | Where-Object Key -eq 'ErrorText').Value
                if (-not [string]::IsNullOrWhiteSpace($taskError)) {
                    $message = $taskError
                }
                Write-Error -Message $message -TargetObject ([ReplaceHardwareTaskInfo]$info)
            } else {
                $hwPath = ($task.Properties | Where-Object Key -eq 'HardwareId').Value
                $recPath = $recorderPathByHwPath[$hwPath]
                if (-not $recorders.ContainsKey($recPath.Id)) {
                    $recorders[$recPath.Id] = Get-VmsRecordingServer -Id $recPath.Id
                }
                $replacedHardwarePaths.Add($hwPath)
            }
        }
        foreach ($rec in $recorders.Values) {
            $rec.HardwareFolder.ClearChildrenCache()
        }
        if ($PassThru) {
            foreach ($path in $replacedHardwarePaths) {
                $itemPath = [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]::new($path)
                Get-VmsHardware -HardwareId $itemPath.Id
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Set-VmsHardwareDriver -ParameterName Driver -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    $values = Get-VmsRecordingServer | Select-Object -First 1 | Get-VmsHardwareDriver |
        Where-Object Name -like "$wordToComplete*" |
        Sort-Object Name |
        Select-Object -ExpandProperty Name -Unique
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Set-VmsLicense {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param (
        [Parameter(Mandatory)]
        [string]
        $Path
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $filePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            if (-not (Test-Path $filePath)) {
                throw [System.IO.FileNotFoundException]::new('Set-VmsLicense could not find the file.', $filePath)
            }
            $bytes = [IO.File]::ReadAllBytes($filePath)
            $b64 = [Convert]::ToBase64String($bytes)
            $result = $ms.LicenseInformationFolder.LicenseInformations[0].ChangeLicense($b64)
            if ($result.State -eq 'Success') {
                $oldSlc = $ms.LicenseInformationFolder.LicenseInformations[0].Slc
                $ms.ClearChildrenCache()
                $newSlc = $ms.LicenseInformationFolder.LicenseInformations[0].Slc
                if ($oldSlc -eq $newSlc) {
                    Write-Verbose "The software license code in the license file passed to Set-VmsLicense is the same as the existing software license code."
                }
                else {
                    Write-Verbose "Set-VmsLicense changed the software license code from $oldSlc to $newSlc."
                }
                Write-Output $ms.LicenseInformationFolder.LicenseInformations[0]
            }
            else {
                Write-Error "Call to ChangeLicense failed. $($result.ErrorText.Trim('.'))."
            }
        }
        catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}


function Set-VmsLoginProvider {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.LoginProvider])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $ClientId,

        [Parameter()]
        [SecureStringTransformAttribute()]
        [securestring]
        $ClientSecret,

        [Parameter()]
        [string]
        $CallbackPath,

        [Parameter()]
        [uri]
        $Authority,

        [Parameter()]
        [string]
        $UserNameClaim,

        [Parameter()]
        [string[]]
        $Scopes,

        [Parameter()]
        [bool]
        $PromptForLogin,

        [Parameter()]
        [bool]
        $Enabled,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        try {
            if ($PSCmdlet.ShouldProcess("Login provider '$($LoginProvider.Name)'", "Update")) {
                $dirty = $false
                $initialName = $LoginProvider.Name
                $keys = @()
                $MyInvocation.BoundParameters.GetEnumerator() | Where-Object Key -in $LoginProvider.GetPropertyKeys() | Foreach-Object {
                    $keys += $_.Key
                }
                if ($MyInvocation.BoundParameters.ContainsKey('Enabled')) {
                    $keys += 'Enabled'
                }
                if ($MyInvocation.BoundParameters.ContainsKey('UserNameClaim')) {
                    $keys += 'UserNameClaim'
                }
                foreach ($key in $keys) {
                    if ($key -eq 'Scopes') {
                        $differences = (($Scopes | Foreach-Object { $_ -in $LoginProvider.Scopes}) -eq $false).Count + (($LoginProvider.Scopes | Foreach-Object { $_ -in $Scopes}) -eq $false).Count
                        if ($differences -gt 0) {
                            Write-Verbose "Updating $key on login provider '$initialName'"
                            $LoginProvider.Scopes.Clear()
                            $Scopes | Foreach-Object {
                                $LoginProvider.Scopes.Add($_)
                            }
                            $dirty = $true
                        }
                    } elseif ($key -eq 'ClientSecret') {
                        Write-Verbose "Updating $key on login provider '$initialName'"
                        $cred = [pscredential]::new('a', $ClientSecret)
                        $LoginProvider.ClientSecret = $cred.GetNetworkCredential().Password
                        $dirty = $true
                    } elseif ($key -eq 'Enabled' -and $LoginProvider.Enabled -ne $Enabled) {
                        Write-Verbose "Setting Enabled to $Enabled on login provider '$initialName'"
                        $LoginProvider.Enabled = $Enabled
                        $dirty = $true
                    } elseif ($key -eq 'UserNameClaim') {
                        Write-Verbose "Setting UserNameClaimType to $UserNameClaim on login provider '$initialName'"
                        $LoginProvider.UserNameClaimType = $UserNameClaim
                        $dirty = $true
                    } elseif ($LoginProvider.$key -cne (Get-Variable -Name $key).Value) {
                        Write-Verbose "Updating $key on login provider '$initialName'"
                        $LoginProvider.$key = (Get-Variable -Name $key).Value
                        $dirty = $true
                    }
                }
                if ($dirty) {
                    $LoginProvider.Save()
                } else {
                    Write-Verbose "No changes were made to login provider '$initialName'."
                }
            }

            if ($PassThru) {
                $LoginProvider
            }
        } catch {
            Write-Error -Message $_.Exception.Message -TargetObject $LoginProvider
        }
    }
}

function Set-VmsLoginProviderClaim {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.RegisteredClaim]
        $Claim,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $DisplayName,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $item = $Claim | Get-ConfigurationItem
        $nameProperty = $item.Properties | Where-Object Key -eq 'Name'
        $dirty = $false
        if ($MyInvocation.BoundParameters.ContainsKey('Name') -and $Name -cne $nameProperty.Value) {
            if ($nameProperty.Value -ceq $item.DisplayName) {
                $item.DisplayName = $Name
            }
            $nameProperty.Value = $Name
            $dirty = $true
        }
        if ($MyInvocation.BoundParameters.ContainsKey('DisplayName') -and $DisplayName -cne $item.DisplayName) {
            $item.DisplayName = $DisplayName
            $dirty = $true
        }
        if ($dirty -and $PSCmdlet.ShouldProcess("Registered claim '$($Claim.Name)'", "Update")) {
            $result = $item | Set-ConfigurationItem
        }
        if ($PassThru -and $result.ValidatedOk) {
            $loginProvider = (Get-VmsLoginProvider | Where-Object Path -eq $Claim.ParentItemPath)
            $loginProvider.ClearChildrenCache()
            $loginProvider | Get-VmsLoginProviderClaim -Name $nameProperty.Value
        }
    }
}


function Set-VmsRecordingServer {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('Recorder')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer[]]
        $RecordingServer,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Description,

        [Parameter(ValueFromPipelineByPropertyName)]
        [BooleanTransformAttribute()]
        [bool]
        $PublicAccessEnabled,

        [Parameter()]
        [ValidateRange(0, 65535)]
        [int]
        $PublicWebserverPort,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $PublicWebserverHostName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [BooleanTransformAttribute()]
        [bool]
        $ShutdownOnStorageFailure,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $MulticastServerAddress,

        [Parameter()]
        [ValidateVmsFeature('RecordingServerFailover')]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $PrimaryFailoverGroup,

        [Parameter()]
        [ValidateVmsFeature('RecordingServerFailover')]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $SecondaryFailoverGroup,

        [Parameter()]
        [ValidateVmsFeature('RecordingServerFailover')]
        [ArgumentCompleter([MipItemNameCompleter[FailoverRecorder]])]
        [MipItemTransformation([FailoverRecorder])]
        [FailoverRecorder]
        $HotStandbyFailoverRecorder,

        [Parameter()]
        [ValidateVmsFeature('RecordingServerFailover')]
        [switch]
        $DisableFailover,

        [Parameter()]
        [ValidateVmsFeature('RecordingServerFailover')]
        [ValidateRange(0, 65535)]
        [int]
        $FailoverPort,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        $updateFailoverSettings = $false
        'PrimaryFailoverGroup', 'SecondaryFailoverGroup', 'HotStandbyFailoverRecorder', 'DisableFailover', 'FailoverPort' | Foreach-Object {
            if ($MyInvocation.BoundParameters.ContainsKey($_)) {
                $updateFailoverSettings = $true
            }
        }
    }

    process {
        if ($HotStandbyFailoverRecorder -and ($PrimaryFailoverGroup -or $SecondaryFailoverGroup)) {
            throw "Invalid combination of failover parameters. When specifying a hot standby failover recorder, you may not also assign a primary or secondary failover group."
        }
        if ($PrimaryFailoverGroup -and ($PrimaryFailoverGroup.Path -eq $SecondaryFailoverGroup.Path)) {
            throw "The same failover group cannot be used for both the primary, and secondary failover groups."
        }

        foreach ($rec in $RecordingServer) {
            try {
                foreach ($property in $rec | Get-Member -MemberType Property | Where-Object Definition -like '*set;*' | Select-Object -ExpandProperty Name) {
                    $parameterName = $property
                    if (-not $PSBoundParameters.ContainsKey($parameterName)) {
                        continue
                    }
                    $newValue = $PSBoundParameters[$parameterName]
                    if ($newValue -ceq $rec.$property) {
                        continue
                    }
                    if ($PSCmdlet.ShouldProcess($rec.Name, "Set $property to $newValue")) {
                        $rec.$property = $newValue
                        $dirty = $true
                    }
                }

                if ($updateFailoverSettings) {

                    $dirtyFailover = $false
                    $failoverSettings = $rec.RecordingServerFailoverFolder.recordingServerFailovers[0]

                    if ($MyInvocation.BoundParameters.ContainsKey('PrimaryFailoverGroup') -and $PrimaryFailoverGroup.Path -ne $failoverSettings.PrimaryFailoverGroup) {
                        $targetName, $targetPath = $PrimaryFailoverGroup.Name, $PrimaryFailoverGroup.Path
                        if ($null -eq $targetName) {
                            $targetName, $targetPath = 'Not used', $failoverSettings.PrimaryFailoverGroupValues['Not used']
                        }

                        if ($PSCmdlet.ShouldProcess($rec.Name, "Set PrimaryFailoverGroup to $targetName")) {
                            $failoverSettings.PrimaryFailoverGroup = $targetPath
                            $failoverSettings.HotStandby = $failoverSettings.HotStandbyValues['Not used']
                            if ($targetPath -eq $failoverSettings.PrimaryFailoverGroupValues['Not used']) {
                                $failoverSettings.SecondaryFailoverGroup = $failoverSettings.SecondaryFailoverGroupValues['Not used']
                            }
                            $dirtyFailover = $true
                        }
                    }

                    if ($MyInvocation.BoundParameters.ContainsKey('SecondaryFailoverGroup') -and $SecondaryFailoverGroup.Path -ne $failoverSettings.SecondaryFailoverGroup) {
                        $targetName, $targetPath = $SecondaryFailoverGroup.Name, $SecondaryFailoverGroup.Path
                        if ($null -eq $targetName) {
                            $targetName, $targetPath = 'Not used', $failoverSettings.SecondaryFailoverGroupValues['Not used']
                        }

                        if ($failoverSettings.PrimaryFailoverGroup -eq 'FailoverGroup[00000000-0000-0000-0000-000000000000]') {
                            Write-Error -Message "You must specify a primary failover group to set the secondary failover group."
                        } elseif ($targetPath -eq $failoverSettings.PrimaryFailoverGroup) {
                            Write-Error -Message "The PrimaryFailoverGroup and SecondaryFailoverGroup must not be the same."
                        } elseif ($PSCmdlet.ShouldProcess($rec.Name, "Set SecondaryFailoverGroup to $targetName")) {
                            $failoverSettings.SecondaryFailoverGroup = $targetPath
                            $failoverSettings.HotStandby = $failoverSettings.HotStandbyValues['Not used']
                            $dirtyFailover = $true
                        }
                    }

                    if ($MyInvocation.BoundParameters.ContainsKey('HotStandbyFailoverRecorder') -and $HotStandbyFailoverRecorder.Path -ne $failoverSettings.HotStandby) {
                        $targetName, $targetPath = $HotStandbyFailoverRecorder.Name, $HotStandbyFailoverRecorder.Path
                        if ($null -eq $targetName) {
                            $targetName, $targetPath = 'Not used', $failoverSettings.HotStandbyValues['Not used']
                        }

                        if ($PSCmdlet.ShouldProcess($rec.Name, "Set hot standby server to $targetName")) {
                            $failoverSettings.PrimaryFailoverGroup = $failoverSettings.PrimaryFailoverGroupValues['Not used']
                            $failoverSettings.SecondaryFailoverGroup = $failoverSettings.SecondaryFailoverGroupValues['Not used']

                            if (-not [string]::IsNullOrWhiteSpace($failoverSettings.HotStandby)) {
                                # Fix for bug #593838. If bug is fixed, consider adding a version check and skip this extra call to Save()
                                $failoverSettings.HotStandby = $failoverSettings.HotStandbyValues['Not used']
                                $failoverSettings.Save()
                            }
                            $failoverSettings.HotStandby = $targetPath
                            $dirtyFailover = $true
                        }
                    }

                    if ($DisableFailover) {
                        if ($PSCmdlet.ShouldProcess($rec.Name, "Disable failover recording")) {
                            $failoverSettings.PrimaryFailoverGroup = $failoverSettings.PrimaryFailoverGroupValues['Not used']
                            $failoverSettings.SecondaryFailoverGroup = $failoverSettings.SecondaryFailoverGroupValues['Not used']
                            $failoverSettings.HotStandby = $failoverSettings.HotStandbyValues['Not used']
                            $dirtyFailover = $true
                        }
                    }

                    if ($MyInvocation.BoundParameters.ContainsKey('FailoverPort') -and $FailoverPort -ne $failoverSettings.FailoverPort) {
                        if ($PSCmdlet.ShouldProcess($rec.Name, "Set failover port to $FailoverPort")) {
                            $failoverSettings.FailoverPort = $FailoverPort
                            $dirtyFailover = $true
                        }
                    }

                    if ($dirtyFailover) {
                        $failoverSettings.Save()
                    }
                }

                if ($dirty) {
                    $rec.Save()
                }

                if ($PassThru) {
                    $rec
                }
            } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                $rec.RecordingServerFailoverFolder.ClearChildrenCache()
                $_ | HandleValidateResultException -TargetObject $rec -ItemName $rec.Name
            }
        }
    }
}

function Set-VmsSiteInfo {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({ ValidateSiteInfoTagName @args })]
        [string]
        $Property,

        [Parameter(Mandatory, Position = 1, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateLength(1, 256)]
        [string]
        $Value,

        [Parameter()]
        [switch]
        $Append
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $ownerPath = 'BasicOwnerInformation[{0}]' -f (Get-VmsManagementServer).Id
        $ownerInfo = Get-ConfigurationItem -Path $ownerPath

        $existingProperties = $ownerInfo.Properties.Key | Foreach-Object { $_ -split '/' | Select-Object -Last 1 }
        if ($Property -in $existingProperties -and -not $Append) {
            # Update existing entry instead of adding a new one
            if ($PSCmdlet.ShouldProcess((Get-VmsSite).Name, "Change $Property entry value to '$Value' in site information")) {
                $p = $ownerInfo.Properties | Where-Object { $_.Key.EndsWith($Property) }
                if ($p.Count -gt 1) {
                    Write-Warning "Site information has multiple values for $Property. Only the first value can be updated with this command."
                    $p = $p[0]
                }
                $p.Value = $Value
                $invokeResult = $ownerInfo | Set-ConfigurationItem
                if (($invokeResult.Properties | Where-Object Key -eq 'State').Value -ne 'Success') {
                    Write-Error "Failed to update Site Information: $($invokeResult.Properties | Where-Object Key -eq 'ErrorText')"
                }
            }
        } elseif ($PSCmdlet.ShouldProcess((Get-VmsSite).Name, "Add $Property entry with value '$Value' to site information")) {
            # Add new, or additional entry for the given property value
            $invokeInfo = $ownerInfo | Invoke-Method -MethodId AddBasicOwnerInfo
            foreach ($p in $invokeInfo.Properties) {
                switch ($p.Key) {
                    'TagType' { $p.Value = $Property }
                    'TagValue' { $p.Value = $Value }
                }
            }
            $invokeResult = $invokeInfo | Invoke-Method -MethodId AddBasicOwnerInfo
            if (($invokeResult.Properties | Where-Object Key -eq 'State').Value -ne 'Success') {
                Write-Error "Failed to update Site Information: $($invokeResult.Properties | Where-Object Key -eq 'ErrorText')"
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Set-VmsSiteInfo -ParameterName Property -ScriptBlock { OwnerInfoPropertyCompleter @args }


function Set-VmsView {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.View])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.View]
        $View,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [Nullable[int]]
        $Shortcut,

        [Parameter()]
        [string[]]
        $ViewItemDefinition,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $dirty = $false
        foreach ($key in 'Name', 'Description', 'Shortcut') {
            if ($MyInvocation.BoundParameters.ContainsKey($key)) {
                $value = $MyInvocation.BoundParameters[$key]
                if ($View.$key -ceq $value) { continue }
                if ($PSCmdlet.ShouldProcess($View.DisplayName, "Changing $key from $($View.$key) to $value")) {
                    $View.$key = $value
                    $dirty = $true
                }
            }
        }

        if ($MyInvocation.BoundParameters.ContainsKey('ViewItemDefinition')) {
            for ($i = 0; $i -lt $ViewItemDefinition.Count; $i++) {
                $definition = $ViewItemDefinition[$i]
                if ($PSCmdlet.ShouldProcess($View.DisplayName, "Update ViewItem $($i + 1)")) {
                    $View.ViewItemChildItems[$i].ViewItemDefinitionXml = $definition
                    $dirty = $true
                }
            }
        }

        if ($dirty -and $PSCmdlet.ShouldProcess($View.DisplayName, 'Saving changes')) {
            $View.Save()
        }

        if ($PassThru) {
            Write-Output $View
        }
    }
}


function Set-VmsViewGroup {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ViewGroup])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ViewGroup]])]
        [MipItemTransformation([ViewGroup])]
        [ViewGroup]
        $ViewGroup,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($key in 'Name', 'Description') {
            if ($MyInvocation.BoundParameters.ContainsKey($key)) {
                $value = $MyInvocation.BoundParameters[$key]
                if ($ViewGroup.$key -ceq $value) { continue }
                if ($PSCmdlet.ShouldProcess($ViewGroup.DisplayName, "Changing $key from $($ViewGroup.$key) to $value")) {
                    $ViewGroup.$key = $value
                    $ViewGroup.Save()
                }
            }
        }
        if ($PassThru) {
            Write-Output $ViewGroup
        }
    }
}

function Set-VmsViewGroupAcl {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VmsViewGroupAcl[]]
        $ViewGroupAcl
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($acl in $ViewGroupAcl) {
            $path = [VideoOS.Platform.Proxy.ConfigApi.ConfigurationItemPath]::new($acl.Path)
            $viewGroup = Get-VmsViewGroup -Id $path.Id
            $target = "View group ""$($viewGroup.DisplayName)"""
            if ($PSCmdlet.ShouldProcess($target, "Updating security permissions for role $($acl.Role.Name)")) {
                $invokeInfo = $viewGroup.ChangeSecurityPermissions($acl.Role.Path)
                $dirty = $false
                foreach ($key in $acl.SecurityAttributes.Keys) {
                    $newValue = $acl.SecurityAttributes[$key]
                    $currentValue = $invokeInfo.GetProperty($key)
                    if ($newValue -cne $currentValue -and $PSCmdlet.ShouldProcess($target, "Changing $key from $currentValue to $newValue")) {
                        $invokeInfo.SetProperty($key, $newValue)
                        $dirty = $true
                    }

                }
                if ($dirty -and $PSCmdlet.ShouldProcess($target, "Saving security permission changes for role $($acl.Role.Name)")) {
                    $invokeResult = $invokeInfo.ExecuteDefault()
                    if ($invokeResult.State -ne 'Success') {
                        Write-Error $invokeResult.ErrorText
                    }
                }
            }
        }
    }
}


function Split-VmsDeviceGroupPath {
    [CmdletBinding()]
    [OutputType([string[]])]
    [RequiresVmsConnection($false)]
    param (
        # Specifies a device group path in unix directory form with forward-slashes as separators.
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline, Position = 0)]
        [string]
        $Path
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        <#
        Path separator   = /
        Escape character = `
        Steps:
            1. Remove unescaped leading and trailing path separator characters
            2. Split path string on unescaped path separators
            3. In each path part, replace the `/ character sequence with /
        #>
        $Path.TrimStart('/') -replace '(?<!`)/$', '' -split '(?<!`)/' | Foreach-Object { $_ -replace '`/', '/' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }
}


function Start-VmsHardwareScan {
    [CmdletBinding()]
    [OutputType([VmsHardwareScanResult])]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer[]]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'Express')]
        [switch]
        $Express,

        [Parameter(ParameterSetName = 'Manual')]
        [uri[]]
        $Address = @(),

        [Parameter(ParameterSetName = 'Manual')]
        [ipaddress]
        $Start,

        [Parameter(ParameterSetName = 'Manual')]
        [ipaddress]
        $End,

        [Parameter(ParameterSetName = 'Manual')]
        [string]
        $Cidr,

        [Parameter(ParameterSetName = 'Manual')]
        [int]
        $HttpPort = 80,

        [Parameter(ParameterSetName = 'Manual')]
        [int[]]
        $DriverNumber = @(),

        [Parameter(ParameterSetName = 'Manual')]
        [string[]]
        $DriverFamily,

        [Parameter()]
        [pscredential[]]
        $Credential,

        [Parameter()]
        [switch]
        $UseDefaultCredentials,

        [Parameter()]
        [switch]
        $UseHttps,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $tasks = @()
        $recorderPathMap = @{}
        $progressParams = @{
            Activity        = 'Initiating VMS hardware scan'
            PercentComplete = 0
        }
        Write-Progress @progressParams
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Express' {
                    foreach ($credentialSet in $Credential | BuildGroupsOfGivenSize -GroupSize 3 -EmptyItem $null) {
                        try {
                            $credentials = $credentialSet | ForEach-Object {
                                [pscustomobject]@{
                                    UserName = $_.UserName
                                    Password = if ($null -eq $_) { $null } else { $_.GetNetworkCredential().Password }
                                }
                            }
                            foreach ($recorder in $RecordingServer) {
                                $recorderPathMap.($recorder.Path) = $recorder
                                $tasks += $recorder.HardwareScanExpress(
                                    $credentials[0].UserName, $credentials[0].Password,
                                    $credentials[1].UserName, $credentials[1].Password,
                                    $credentials[2].UserName, $credentials[2].Password,
                                    ($null -eq $Credential -or $UseDefaultCredentials), $UseHttps)
                            }
                        } catch {
                            throw
                        }
                    }
                }

                'Manual' {
                    $rangeParameters = ($MyInvocation.BoundParameters.Keys | Where-Object { $_ -in @('Start', 'End') }).Count
                    if ($rangeParameters -eq 1) {
                        Write-Error 'When using the Start or End parameters, you must provide both Start and End parameter values'
                        return
                    }

                    $Address = $Address | ForEach-Object {
                        if ($_.IsAbsoluteUri) {
                            $_
                        } else {
                            [uri]"http://$($_.OriginalString)"
                        }
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('UseHttps') -or $MyInvocation.BoundParameters.ContainsKey('HttpPort')) {
                        $Address = $Address | Foreach-Object {
                            $a = [uribuilder]$_
                            if ($MyInvocation.BoundParameters.ContainsKey('UseHttps')) {
                                $a.Scheme = if ($UseHttps) { 'https' } else { 'http' }
                            }
                            if ($MyInvocation.BoundParameters.ContainsKey('HttpPort')) {
                                $a.Port = $HttpPort
                            }
                            $a.Uri
                        }
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('Start')) {
                        $Address += Expand-IPRange -Start $Start -End $End | ConvertTo-Uri -UseHttps:$UseHttps -HttpPort $HttpPort
                    }
                    if ($MyInvocation.BoundParameters.ContainsKey('Cidr')) {
                        $Address += Expand-IPRange -Cidr $Cidr | Select-Object -Skip 1 | Select-Object -SkipLast 1 | ConvertTo-Uri -UseHttps:$UseHttps -HttpPort $HttpPort
                    }

                    foreach ($entry in $Address) {
                        try {
                            foreach ($cred in $Credential | BuildGroupsOfGivenSize -GroupSize 1 -EmptyItem $null) {
                                $user = $cred[0].UserName
                                $pass = $cred[0].Password
                                foreach ($recorder in $RecordingServer) {
                                    if ($MyInvocation.BoundParameters.ContainsKey('DriverFamily')) {
                                        $DriverNumber += $recorder | Get-VmsHardwareDriver | Where-Object { $_.GroupName -in $DriverFamily -and $_.Number -notin $DriverNumber } | Select-Object -ExpandProperty Number
                                    }
                                    if ($DriverNumber.Count -eq 0) {
                                        Write-Warning "Start-VmsHardwareScan is about to scan $($Address.Count) addresses from $($recorder.Name) without specifying one or more hardware device drivers. This can take a very long time."
                                    }
                                    $driverNumbers = $DriverNumber -join ';'
                                    Write-Verbose "Adding HardwareScan task for $($entry) using driver numbers $driverNumbers"
                                    $recorderPathMap.($recorder.Path) = $recorder
                                    $tasks += $RecordingServer.HardwareScan($entry.ToString(), $driverNumbers, $user, $pass, ($null -eq $Credential -or $UseDefaultCredentials))
                                }
                            }
                        } catch {
                            throw
                        }
                    }
                }
            }
        } finally {
            $progressParams.Completed = $true
            $progressParams.PercentComplete = 100
            Write-Progress @progressParams
        }

        if ($PassThru) {
            Write-Output $tasks
        } else {
            Wait-VmsTask -Path $tasks.Path -Title "Running $(($PSCmdlet.ParameterSetName).ToLower()) hardware scan" -Cleanup | Foreach-Object {
                $state = $_.Properties | Where-Object Key -eq 'State'
                if ($state.Value -eq 'Error') {
                    $errorText = $_.Properties | Where-Object Key -eq 'ErrorText'
                    Write-Error $errorText.Value
                } else {
                    $results = if ($_.Children.Count -gt 0) { [VmsHardwareScanResult[]]$_.Children } else {
                        [VmsHardwareScanResult]$_
                    }
                    foreach ($result in $results) {
                        $result.RecordingServer = $recorderPathMap.($_.ParentPath)
                        # TODO: Remove this entire if block when bug 487881 is fixed and hotfixes for supported versions are available.
                        if ($result.MacAddressExistsLocal) {
                            if ($result.MacAddress -notin ($result.RecordingServer | Get-VmsHardware | Get-HardwareSetting).MacAddress) {
                                Write-Verbose "MacAddress $($result.MacAddress) incorrectly reported as already existing on recorder. Changing MacAddressExistsLocal to false."
                                $result.MacAddressExistsLocal = $false
                            }
                        }
                        Write-Output $result
                    }
                }
            }
        }
    }
}

function Wait-VmsTask {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidateVmsItemType('Task')]
        [string[]]
        $Path,

        [Parameter()]
        [string]
        $Title,

        [Parameter()]
        [switch]
        $Cleanup
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $tasks = New-Object 'System.Collections.Generic.Queue[VideoOS.ConfigurationApi.ClientService.ConfigurationItem]'
        $Path | Foreach-Object {
            $item = $null
            $errorCount = 0
            while ($null -eq $item) {
                try {
                    $item = Get-ConfigurationItem -Path $_
                }
                catch {
                    $errorCount++
                    if ($errorCount -ge 5) {
                        throw
                    }
                    else {
                        Write-Verbose 'Wait-VmsTask received an error when communicating with Configuration API. The communication channel will be re-established and the connection will be attempted up to 5 times.'
                        Start-Sleep -Seconds 2
                        Get-VmsSite | Select-VmsSite
                    }
                }
            }

            if ($item.ItemType -ne 'Task') {
                Write-Error "Configuration Item with path '$($item.Path)' is incompatible with Wait-VmsTask. Expected an ItemType of 'Task' and received a '$($item.ItemType)'."
            }
            else {
                $tasks.Enqueue($item)
            }
        }
        $completedStates = 'Error', 'Success', 'Completed'
        $totalTasks = $tasks.Count
        $progressParams = @{
            Activity = if ([string]::IsNullOrWhiteSpace($Title)) { 'Waiting for VMS Task(s) to complete' } else { $Title }
            PercentComplete = 0
            Status = 'Processing'
        }
        try {
            Write-Progress @progressParams
            $stopwatch = [diagnostics.stopwatch]::StartNew()
            while ($tasks.Count -gt 0) {
                Start-Sleep -Milliseconds 500
                $taskInfo = $tasks.Dequeue()
                $completedTaskCount = $totalTasks - ($tasks.Count + 1)
                $tasksRemaining = $totalTasks - $completedTaskCount
                $percentComplete = [int]($taskInfo.Properties | Where-Object Key -eq 'Progress' | Select-Object -ExpandProperty Value)

                if ($completedTaskCount -gt 0) {
                    $timePerTask = $stopwatch.ElapsedMilliseconds / $completedTaskCount
                    $remainingTime = [timespan]::FromMilliseconds($tasksRemaining * $timePerTask)
                    $progressParams.SecondsRemaining = [int]$remainingTime.TotalSeconds
                }
                elseif ($percentComplete -gt 0){
                    $pointsRemaining = 100 - $percentComplete
                    $timePerPoint = $stopwatch.ElapsedMilliseconds / $percentComplete
                    $remainingTime = [timespan]::FromMilliseconds($pointsRemaining * $timePerPoint)
                    $progressParams.SecondsRemaining = [int]$remainingTime.TotalSeconds
                }

                if ($tasks.Count -eq 0) {
                    $progressParams.Status = "$($taskInfo.Path) - $($taskInfo.DisplayName)."
                    $progressParams.PercentComplete = $percentComplete
                    Write-Progress @progressParams
                }
                else {
                    $progressParams.Status = "Completed $completedTaskCount of $totalTasks tasks. Remaining tasks: $tasksRemaining"
                    $progressParams.PercentComplete = [int]($completedTaskCount / $totalTasks * 100)
                    Write-Progress @progressParams
                }
                $errorCount = 0
                while ($null -eq $taskInfo) {
                    try {
                        $taskInfo = $taskInfo | Get-ConfigurationItem
                        break
                    }
                    catch {
                        $errorCount++
                        if ($errorCount -ge 5) {
                            throw
                        }
                        else {
                            Write-Verbose 'Wait-VmsTask received an error when communicating with Configuration API. The communication channel will be re-established and the connection will be attempted up to 5 times.'
                            Start-Sleep -Seconds 2
                            Get-VmsSite | Select-VmsSite
                        }
                    }
                }
                $taskInfo = $taskInfo | Get-ConfigurationItem
                if (($taskInfo | Get-ConfigurationItemProperty -Key State) -notin $completedStates) {
                    $tasks.Enqueue($taskInfo)
                    continue
                }
                Write-Output $taskInfo
                if ($Cleanup -and $taskInfo.MethodIds -contains 'TaskCleanup') {
                    $null = $taskInfo | Invoke-Method -MethodId 'TaskCleanup'
                }
            }
        }
        finally {
            $progressParams.Completed = $true
            Write-Progress @progressParams
        }
    }
}


function Get-VmsAlarmDefinition {
    [CmdletBinding()]
    [Alias("Get-AlarmDefinition")]
    [OutputType([VideoOS.Platform.ConfigurationItems.AlarmDefinition])]
    [RequiresVmsConnection()]
    param (
        [Parameter()]
        [SupportsWildcards()]
        [ArgumentCompleter([MipItemNameCompleter[AlarmDefinition]])]
        [string]
        $Name
    )

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        if ($PSBoundParameters.ContainsKey('Name')) {
            # Since the QueryItems feature doesn't support wildcards or regex, it isn't very good for searching as the
            # PowerShell user is used to using "Server*" and advanced users are used to using regex like '^Server'.
            # Because of this, this cmdlet is just going to use the -Like operator against all alarm definitions.
            (Get-VmsManagementServer).AlarmDefinitionFolder.AlarmDefinitions | Where-Object Name -like $Name
        } else {
            (Get-VmsManagementServer).AlarmDefinitionFolder.AlarmDefinitions
        }
    }
}

function New-VmsAlarmDefinition {
    [CmdletBinding()]
    [MilestonePSTools.RequiresVmsConnection()]
    [OutputType([VideoOS.Platform.ConfigurationItems.AlarmDefinition])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        [Parameter(Mandatory)]
        [string]
        $EventTypeGroup,

        [Parameter(Mandatory)]
        [string]
        $EventType,

        # Specifies one or more devices in the form of a Configuration Item
        # Path like "Camera[e6d71e26-4c27-447d-b719-7db14fef8cd7]".
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Path')]
        [string[]]
        $Source,

        [Parameter()]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $RelatedCameras,

        [Parameter()]
        [string]
        $TimeProfile,

        # UDEs and Inputs
        [Parameter()]
        [string[]]
        $EnabledBy,
        
        # UDEs and Inputs
        [Parameter()]
        [string[]]
        $DisabledBy,

        [Parameter()]
        [string]
        $Instructions,

        [Parameter()]
        [string]
        $Priority,

        [Parameter()]
        [string]
        $Category,

        [Parameter()]
        [switch]
        $AssignableToAdmins,

        [Parameter()]
        [timespan]
        $Timeout = [timespan]::FromMinutes(1),

        [Parameter()]
        [string[]]
        $TimeoutAction,

        [Parameter(ParameterSetName="SmartMap")]
        [switch]
        $SmartMap,

        [Parameter(ParameterSetName="RelatedMap")]
        [string]
        $RelatedMap,

        [Parameter()]
        [string]
        $Owner,

        # UDE's or outputs?
        [Parameter()]
        [string[]]
        $EventsToTrigger
    )

    begin {
        Assert-VmsRequirementsMet
        $sources = [system.collections.generic.list[string]]::new()
    }
    
    process {
        foreach ($path in $Source) {
            $sources.Add($path)
        }
    }

    end {
        $def = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition()
        $def.Name = $Name
        $def.Description = $Instructions
        $def.AssignableToAdmins = $AssignableToAdmins.ToBool()
        $def.TriggerEventlist = $EventsToTrigger -join ','
        $def.Owner = $Owner

        $eventTypeGroupId = [guid]::Empty
        if (![guid]::TryParse($EventTypeGroup, [ref]$eventTypeGroupId)) {
            $groupName = $def.EventTypeGroupValues.Keys | Where-Object { $_ -eq $EventTypeGroup }
            if ($null -eq $groupName) {
                Write-Error "EventTypeGroup '$EventTypeGroup' is not a valid EventTypeGroup name, or GUID."
                return
            }
            $eventTypeGroupid = $def.EventTypeGroupValues[$groupName]
        }
        $def.EventTypeGroup = $eventTypeGroupId
        
        $eventTypeId = [guid]::Empty
        if (![guid]::TryParse($EventType, [ref]$eventTypeId)) {
            $null = $def.ValidateItem()
            $eventName = $def.EventTypeValues.Keys | Where-Object { $_ -eq $EventType }
            if ($null -eq $eventName) {
                Write-Error "EventType '$EventType' is not a valid event name, or GUID. For a list of system events, try running (Get-VmsManagementServer).EventTypeGroupFolder.EventTypeGroups.EventTypeFolder.EventTypes | Select Name, DisplayName, Id"
                return
            }
            $eventTypeId = $def.EventTypeValues[$eventName]
        }
        $def.EventType = $eventTypeId

        $boundParameters = $PSCmdlet.MyInvocation.BoundParameters
        if (($boundParameters.ContainsKey('EnabledBy') -or $boundParameters.ContainsKey('DisabledBy') -and $boundParameters.ContainsKey('TimeProfile'))) {
            Write-Error 'Rules for when an alarm definition is enabled may either be based on a time profile, or a specified enable/disable event, but not both.'
            return
        }

        # TODO: Use switch on parametersetname to determine enablerule
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('TimeProfile')) {
            $timeProfiles = @{
                'Always' = 'TimeProfile[00000000-0000-0000-0000-000000000000]'
            }
            (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | ForEach-Object {
                if ($null -eq $_) { return }
                $timeProfiles[$_.Name] = $_.Path
            }

            if (!$timeProfiles.ContainsKey($TimeProfile)) {
                Write-Error "No TimeProfile found matching '$TimeProfile'"
                return
            }
            $def.TimeProfile = $timeProfiles[$TimeProfile]
            $def.EnableRule = 1
        }

        if ($PSCmdlet.ParameterSetName -eq 'EventTriggered') {
            $def.EnableEventList = $EnabledBy -join ','
            $def.DisableEventList = $DisabledBy -join ','
            $def.EnableRule = 2
        }

        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Priority')) {
            if (!$def.PriorityValues.ContainsKey($Priority)) {
                Write-Error "No alarm priority found with the name '$Priority'. Check your Alarm Data Settings in Management Client."
                return
            }
            $def.Priority = $def.PriorityValues[$Priority]
        }
        
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Category')) {
            if (!$def.CategoryValues.ContainsKey($Category)) {
                Write-Error "No alarm category found with the name '$Category'. Check your Alarm Data Settings in Management Client."
                return
            }
            $def.Category = $def.CategoryValues[$Category]
        }

        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RelatedMap')) {
            if (!$def.RelatedMapValues.ContainsKey($RelatedMap)) {
                Write-Error "No related map found with the name '$RelatedMap'. Check the map name and try again."
                return
            }
            $def.MapType = 1
            $def.RelatedMap = $def.RelatedMapValues[$RelatedMap]
        }

        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('SmartMap')) {
            $def.MapType = 2
            $def.RelatedMap = ''
        }

        $def.ManagementTimeoutTime = $Timeout.ToString()
        $def.ManagementTimeoutEventList = $TimeoutAction -join ','

        $sourceHelpers = @{
            'AllCameras'     = '/CameraFolder'
            'AllMicrophones' = '/MicrophoneFolder'
            'AllSpeakers'    = '/SpeakerFolder'
            'AllInputs'      = '/InputEventFolder'
            'AllOutputs'     = '/OutputFolder'
            'AllEvents'      = '/UserDefinedEventFolder'
            'AllServers'     = '/'
        }
        $def.SourceList = ($sources | ForEach-Object {
            if ($sourceHelpers.ContainsKey($_)) { $sourceHelpers[$_] } else { $_ }
        }) -join ','

        if ($RelatedCameras.Count -gt 0) {
            $def.RelatedCameraList = ($RelatedCameras | ForEach-Object Path) -join ','
        }
        
        try {
            $taskResult = $def.ExecuteDefault()
            if ($taskResult.State -ne 'Success') {
                Write-Error "New-VmsAlarmDefinition failed1: $($taskResult.ErrorText)" -TargetObject $def
                return
            }
        } catch {
            Write-Error "New-VmsAlarmDefinition failed2: $($_.Exception.Message)" -TargetObject $def
            return
        }
        
        
        (Get-VmsManagementServer).AlarmDefinitionFolder.ClearChildrenCache()
        (Get-VmsManagementServer).AlarmDefinitionFolder.AlarmDefinitions | Where-Object Path -EQ $taskResult.Path
    }
}


$eventTypeGroupArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition().EventTypeGroupValues.Keys | Sort-Object
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$eventTypeArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $group = $fakeBoundParameters['EventTypeGroup']
    if ([string]::IsNullOrWhiteSpace($group)) {
        "'Tab completion unavailable until EventTypeGroup is set'"
        return
    }
    $info = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition()
    $groupNames = @{}
    $info.EventTypeGroupValues.Keys | ForEach-Object { $groupNames[$_] = $info.EventTypeGroupValues[$_] }
    if ($groupNames.ContainsKey($group)) {
        $group = $info.EventTypeGroupValues.Keys | Where-Object { $_ -eq $group }
        $info.EventTypeGroup = $info.EventTypeGroupValues[$group]
    } elseif ($groupNames.Values -contains $group) {
        $info.EventTypeGroup = $group
    } else {
        "'Invalid EventTypeGroup `"$group`"'"
        return
    }

    $null = $info.ValidateItem()
    $values = $info.EventTypeValues.Keys | Sort-Object
    if ($null -eq $values) {
        "'No events available for EventTypeGroup $group'"
    }
    
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$timeProfileArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = [system.collections.generic.list[string]]::new()
    $values.Add('Always')
    (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | ForEach-Object {
        $values.Add($_.Name)
    } | Sort-Object
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$priorityArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition().PriorityValues.Keys | Where-Object {
        ![string]::IsNullOrEmpty($_)
    } | Sort-Object
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$categoryArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition().CategoryValues.Keys | Where-Object {
        ![string]::IsNullOrEmpty($_)
    } | Sort-Object
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$relatedMapArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = (Get-VmsManagementServer).AlarmDefinitionFolder.AddAlarmDefinition().RelatedMapValues.Keys | Where-Object {
        ![string]::IsNullOrEmpty($_)
    } | Sort-Object
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

$sourceArgCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    $values = 'AllCameras', 'AllMicrophones', 'AllSpeakers', 'AllMetadatas', 'AllInputs', 'AllOutputs', 'AllServers', 'AllEvents'
    $values | Where-Object {
        $_ -like "$($wordToComplete.Trim('"', "'"))*"
    } | ForEach-Object {
        if ($_ -match '.*\s+.*') {
            "'$_'"
        } else {
            $_
        }
    }
}

Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName EventTypeGroup -ScriptBlock $eventTypeGroupArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName EventType -ScriptBlock $eventTypeArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName TimeProfile -ScriptBlock $timeProfileArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName Priority -ScriptBlock $priorityArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName Category -ScriptBlock $categoryArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName RelatedMap -ScriptBlock $relatedMapArgCompleter
Register-ArgumentCompleter -CommandName New-VmsAlarmDefinition -ParameterName Source -ScriptBlock $sourceArgCompleter

function Remove-VmsAlarmDefinition {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[AlarmDefinition]])]
        [MipItemTransformation([AlarmDefinition])]
        [AlarmDefinition[]]
        $AlarmDefinition
    )

    begin {
        Assert-VmsRequirementsMet
        $folder = (Get-VmsManagementServer).AlarmDefinitionFolder
    }

    process {
        foreach ($definition in $AlarmDefinition) {
            try {
                if ($PSCmdlet.ShouldProcess($definition.Name, 'Remove Alarm Definition')) {
                    $result = $folder.RemoveAlarmDefinition($definition.Path)
                    if ($result.State -ne 'Success') {
                        Write-Error "An error was returned while removing the alarm definition. $($result.ErrorText)" -TargetObject $definition
                    }
                }
            } catch [VideoOS.Platform.PathNotFoundMIPException] {
                Write-Error -Message "Alarm definition '$($definition.Name)' with Id '$($definition.Id)' does not exist." -Exception $_.Exception -TargetObject $definition
            }
        }
    }
}

function Set-VmsAlarmDefinition {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ArgumentCompleter([MipItemNameCompleter[AlarmDefinition]])]
        [MipItemTransformation([AlarmDefinition])]
        [AlarmDefinition[]]
        $AlarmDefinition,

        [Parameter()]
        [string]
        $Name,

        # Specifies one or more devices in the form of a Configuration Item
        # Path like "Camera[e6d71e26-4c27-447d-b719-7db14fef8cd7]".
        [Parameter()]
        [Alias('Path')]
        [string[]]
        $Source,

        # Specifies the related cameras in the form of a comma-separated list of Configuration Item paths.
        [Parameter()]
        [string[]]
        $RelatedCameras,

        [Parameter()]
        [string]
        $TimeProfile,

        # UDEs and Inputs
        [Parameter()]
        [string[]]
        $EnabledBy,
        
        # UDEs and Inputs
        [Parameter()]
        [string[]]
        $DisabledBy,

        [Parameter()]
        [string]
        $Instructions,

        [Parameter()]
        [string]
        $Priority,

        [Parameter()]
        [string]
        $Category,

        [Parameter()]
        [switch]
        $AssignableToAdmins,

        [Parameter()]
        [timespan]
        $Timeout = [timespan]::FromMinutes(1),

        [Parameter()]
        [string[]]
        $TimeoutAction,

        [Parameter(ParameterSetName="SmartMap")]
        [switch]
        $SmartMap,

        [Parameter(ParameterSetName="RelatedMap")]
        [string]
        $RelatedMap,

        [Parameter()]
        [string]
        $Owner,

        # UDE's or outputs?
        [Parameter()]
        [string[]]
        $EventsToTrigger,

        [Parameter()]
        [switch]
        $PassThru
    )
    
    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($def in $AlarmDefinition) {
            $alarmName = $def.Name
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Name')) {
                $def.Name = $Name
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Instructions')) {
                $def.Description = $Instructions
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('AssignableToAdmins')) {
                $def.AssignableToAdmins = $AssignableToAdmins.ToBool()
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EventsToTrigger')) {
                $def.TriggerEventlist = $EventsToTrigger -join ','
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Owner')) {
                $def.Owner = $Owner
            }
            
            # TODO: Use switch on parametersetname to determine enablerule
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('TimeProfile')) {
                $timeProfiles = @{
                    'Always' = 'TimeProfile[00000000-0000-0000-0000-000000000000]'
                }
                (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | ForEach-Object {
                    if ($null -eq $_) { return }
                    $timeProfiles[$_.Name] = $_.Path
                }

                if (!$timeProfiles.ContainsKey($TimeProfile)) {
                    Write-Error "No TimeProfile found matching '$TimeProfile'"
                    return
                }
                $def.TimeProfile = $timeProfiles[$TimeProfile]
                $def.EnableRule = 1
            }

            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('EventTriggered')) {
                $def.EnableEventList = $EnabledBy -join ','
                $def.DisableEventList = $DisabledBy -join ','
                $def.EnableRule = 2
            }

            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Priority')) {
                if (!$def.PriorityValues.ContainsKey($Priority)) {
                    Write-Error "No alarm priority found with the name '$Priority'. Check your Alarm Data Settings in Management Client."
                    return
                }
                $def.Priority = $def.PriorityValues[$Priority]
            }
    
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Category')) {
                if (!$def.CategoryValues.ContainsKey($Category)) {
                    Write-Error "No alarm category found with the name '$Category'. Check your Alarm Data Settings in Management Client."
                    return
                }
                $def.Category = $def.CategoryValues[$Category]
            }

            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('RelatedMap')) {
                if (!$def.RelatedMapValues.ContainsKey($RelatedMap)) {
                    Write-Error "No related map found with the name '$RelatedMap'. Check the map name and try again."
                    return
                }
                $def.MapType = 1
                $def.RelatedMap = $def.RelatedMapValues[$RelatedMap]
            }

            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('SmartMap')) {
                $def.MapType = 2
                $def.RelatedMap = ''
            }

            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Timeout')) {
                $def.ManagementTimeoutTime = $Timeout.ToString()
            }
            
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('TimeoutAction')) {
                $def.ManagementTimeoutEventList = $TimeoutAction -join ','
            }
            
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Source')) {
                $sourceHelpers = @{
                    'AllCameras'     = '/CameraFolder'
                    'AllMicrophones' = '/MicrophoneFolder'
                    'AllSpeakers'    = '/SpeakerFolder'
                    'AllInputs'      = '/InputEventFolder'
                    'AllOutputs'     = '/OutputFolder'
                    'AllEvents'      = '/UserDefinedEventFolder'
                    'AllServers'     = '/'
                }
    
                $def.SourceList = ($Source | ForEach-Object {
                        if ($sourceHelpers.ContainsKey($_)) { $sourceHelpers[$_] } else { $_ }
                    }) -join ','
            }
            

            if ($RelatedCameras.Count -gt 0) {
                $def.RelatedCameraList = $RelatedCameras -join ','
            }
    
            try {
                if ($PSCmdlet.ShouldProcess($alarmName, 'Set Alarm Definition')) {
                    $null = $def.Save()
                    if ($PassThru) {
                        $def
                    }
                }
            } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                HandleValidateResultException -ErrorRecord $_ -TargetObject $def -ItemName $def.Name
            } catch {
                Write-Error "Set-VmsAlarmDefinition failed: $($_.Exception.Message)" -TargetObject $def
                return
            }
        }
    }

    end {
        (Get-VmsManagementServer).AlarmDefinitionFolder.ClearChildrenCache()
    }
}

Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName EventTypeGroup -ScriptBlock $eventTypeGroupArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName EventType -ScriptBlock $eventTypeArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName TimeProfile -ScriptBlock $timeProfileArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName Priority -ScriptBlock $priorityArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName Category -ScriptBlock $categoryArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName RelatedMap -ScriptBlock $relatedMapArgCompleter
Register-ArgumentCompleter -CommandName Set-VmsAlarmDefinition -ParameterName Source -ScriptBlock $sourceArgCompleter

function Copy-VmsClientProfile {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClientProfile])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter(Mandatory, Position = 0)]
        [string]
        $NewName
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $newProfile = New-VmsClientProfile -Name $NewName -Description $ClientProfile.Description -ErrorAction Stop
        if ($ClientProfile.IsDefaultProfile) {
            # New client profiles are by default an exact copy of the default profile. No need to copy attributes to the new profile.
            $newProfile
            return
        }

        foreach ($attributes in $ClientProfile | Get-VmsClientProfileAttributes) {
            $newProfile | Set-VmsClientProfileAttributes -Attributes $attributes -Verbose:($VerbosePreference -eq 'Continue')
        }
    }
}


function Export-VmsClientProfile {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param (
        [Parameter(ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile[]]
        $ClientProfile,

        [Parameter(Mandatory, Position = 0)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $ValueTypeInfo
    )

    begin {
        Assert-VmsRequirementsMet

        $resolvedPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue -ErrorVariable rpError).Path
        if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
            $resolvedPath = $rpError.TargetObject
        }
        $Path = $resolvedPath
        $fileInfo = [io.fileinfo]$Path
        if (-not $fileInfo.Directory.Exists) {
            throw ([io.directorynotfoundexception]::new("Directory not found: $($fileInfo.Directory.FullName)"))
        }
        if (($fi = [io.fileinfo]$Path).Extension -ne '.json') {
            Write-Verbose "A .json file extension will be added to the file '$($fi.Name)'"
            $Path += ".json"
        }
        $results = [system.collections.generic.list[pscustomobject]]::new()
    }

    process {
        if ($ClientProfile.Count -eq 0) {
            $ClientProfile = Get-VmsClientProfile
        }
        foreach ($p in $ClientProfile) {
            $results.Add([pscustomobject]@{
                Name        = $p.Name
                Description = $p.Description
                Attributes  = $p | Get-VmsClientProfileAttributes -ValueTypeInfo:$ValueTypeInfo
            })
        }
    }

    end {
        $json = ConvertTo-Json -InputObject $results -Depth 10 -Compress
        [io.file]::WriteAllText($Path, $json, [text.encoding]::UTF8)
    }
}


function Get-VmsClientProfile {
    [CmdletBinding(DefaultParameterSetName = 'Name')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClientProfile])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param (
        [Parameter(ParameterSetName = 'Name', ValueFromPipelineByPropertyName, Position = 0)]
        [ArgumentCompleter([MilestonePSTools.Utility.MipItemNameCompleter[VideoOS.Platform.ConfigurationItems.ClientProfile]])]
        [SupportsWildcards()]
        [string]
        $Name,

        [Parameter(Mandatory, ParameterSetName = 'Id', ValueFromPipelineByPropertyName)]
        [guid]
        $Id,

        [Parameter(Mandatory, ParameterSetName = 'DefaultProfile')]
        [switch]
        $DefaultProfile
    )

    begin {
        Assert-VmsRequirementsMet
        $folder = (Get-VmsManagementServer -ErrorAction Stop).ClientProfileFolder
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'Id' {
                [VideoOS.Platform.ConfigurationItems.ClientProfile]::new($folder.ServerId, "ClientProfile[$Id]")
            }

            'Name' {
                $matchingProfiles = $folder.ClientProfiles | Where-Object {
                    [string]::IsNullOrWhiteSpace($Name) -or $_.Name -like $Name
                }
                if ($matchingProfiles) {
                    $matchingProfiles
                } elseif (-not [system.management.automation.wildcardpattern]::ContainsWildcardCharacters($Name)) {
                    Write-Error -Message "ClientProfile '$Name' not found."
                }
            }

            'DefaultProfile' {
                Get-VmsClientProfile | Where-Object IsDefaultProfile -eq $DefaultProfile
            }

            default {
                throw "ParameterSetName '$_' not implemented."
            }
        }
    }
}


function Get-VmsClientProfileAttributes {
    [CmdletBinding()]
    [OutputType([System.Collections.IDictionary])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter(Position = 0)]
        [string[]]
        $Namespace,

        [Parameter()]
        [switch]
        $ValueTypeInfo
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $namespaces = ($ClientProfile | Get-Member -MemberType Property -Name 'ClientProfile*ChildItems').Name -replace 'ClientProfile(.+)ChildItems', '$1'
        if ($Namespace.Count -eq 0) {
            $Namespace = $namespaces
        }

        foreach ($ns in $Namespace) {
            if ($ns -notin $namespaces) {
                Write-Error "Property 'ClientProfile$($ns)ChildItems' does not exist on client profile '$($ClientProfile.DisplayName)'"
                continue
            }
            $settings = $ClientProfile."ClientProfile$($ns)ChildItems"
            $attributes = [ordered]@{
                Namespace = $ns
            }
            if ($settings.Count -eq 0) {
                Write-Verbose "Ignoring empty client profile namespace '$ns'."
                continue
            }
            foreach ($key in $settings.GetPropertyKeys() | Where-Object { $_ -notmatch '(?<!Locked)Locked$' } | Sort-Object) {
                $attributes[$key] = [pscustomobject]@{
                    Value         = $settings.GetProperty($key)
                    ValueTypeInfo = if ($ValueTypeInfo) { $settings.GetValueTypeInfoList($key) } else { $null }
                    Locked        = $settings."$($key)Locked"
                }
            }
            $attributes
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsClientProfileAttributes -ParameterName Namespace -ScriptBlock {
    $values = (Get-VmsClientProfile -DefaultProfile | Get-VmsClientProfileAttributes).Namespace | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Import-VmsClientProfile {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClientProfile])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param (
        [Parameter(Mandatory, Position = 0)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
        $Path = (Resolve-Path -Path $Path -ErrorAction Stop).Path
        (Get-VmsManagementServer -ErrorAction Stop).ClientProfileFolder.ClearChildrenCache()
        $existingProfiles = @{}
        Get-VmsClientProfile | Foreach-Object {
            $existingProfiles[$_.Name] = $_
        }
        $showVerbose = $VerbosePreference -eq 'Continue'
    }

    process {
        $definitions = [io.file]::ReadAllText($Path, [text.encoding]::UTF8) | ConvertFrom-Json
        foreach ($def in $definitions) {
            try {
                if ($existingProfiles.ContainsKey($def.Name)) {
                    if ($Force) {
                        $current = $existingProfiles[$def.Name]
                        $current | Set-VmsClientProfile -Description $def.Description -ErrorAction Stop -Verbose:$showVerbose
                    } else {
                        Write-Error "ClientProfile '$($def.Name)' already exists. To overwrite existing profiles, try including the -Force switch."
                        continue
                    }
                } else {
                    $current = New-VmsClientProfile -Name $def.Name -Description $def.Description -ErrorAction Stop
                    $existingProfiles[$current.Name] = $current
                }
                foreach ($psObj in $def.Attributes) {
                    $attributes = @{}
                    foreach ($memberName in ($psObj | Get-Member -MemberType NoteProperty).Name) {
                        $attributes[$memberName] = $psObj.$memberName
                    }
                    $current | Set-VmsClientProfileAttributes -Attributes $attributes -Verbose:$showVerbose
                }
                $current
            } catch {
                Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $def
            }
        }
    }
}


function New-VmsClientProfile {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClientProfile])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [string]
        $Description
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $serverTask = (Get-VmsManagementServer -ErrorAction Stop).ClientProfileFolder.AddClientProfile($Name, $Description)
            if ($serverTask.State -ne 'Success') {
                Write-Error -Message "Error creating new client profile: $($serverTask.ErrorText)" -TargetObject $serverTask
                return
            }
            Get-VmsClientProfile -Id ($serverTask.Path -replace 'ClientProfile\[(.+)\]', '$1')
        } catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}


function Remove-VmsClientProfile {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile[]]
        $ClientProfile
    )

    begin {
        Assert-VmsRequirementsMet
        $folder = (Get-VmsManagementServer -ErrorAction Stop).ClientProfileFolder
    }

    process {
        foreach ($p in $ClientProfile) {
            try {
                if ($PSCmdlet.ShouldProcess("ClientProfile $($p.Name)", "Remove")) {
                    $serverTask = $folder.RemoveClientProfile($p.Path)
                    if ($serverTask.State -ne 'Success') {
                        Write-Error -Message "Error creating new client profile: $($serverTask.ErrorText)" -TargetObject $serverTask
                        return
                    }
                }
            } catch {
                Write-Error -Message $_.Message -Exception $_.Exception -TargetObject $p
            }
        }
    }
}


function Set-VmsClientProfile {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClientProfile])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $Priority,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        if ($MyInvocation.BoundParameters.ContainsKey('Priority')) {
            (Get-VmsManagementServer -ErrorAction Stop).ClientProfileFolder.ClearChildrenCache()
            $clientProfiles = Get-VmsClientProfile
        }
    }

    process {
        try {
            $dirty = $false
            if (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -cne $ClientProfile.Name) {
                $dirty = $true
            } else {
                $Name = $ClientProfile.Name
            }
            if ($MyInvocation.BoundParameters.ContainsKey('Description') -and $Description -cne $ClientProfile.Description) {
                $dirty = $true
            } else {
                $Description = $ClientProfile.Description
            }

            $priorityDifference = 0
            if ($MyInvocation.BoundParameters.ContainsKey('Priority')) {
                $currentPriority = 1..($clientProfiles.Count) | Where-Object { $ClientProfile.Path -eq $clientProfiles[$_ - 1].Path }
                $priorityDifference = $Priority - $currentPriority
                if ($priorityDifference) {
                    $dirty = $true
                }
            }

            if ($dirty -and $PSCmdlet.ShouldProcess("ClientProfile '$($ClientProfile.Name)'", "Update")) {
                if ($MyInvocation.BoundParameters.ContainsKey('Name') -or $MyInvocation.BoundParameters.ContainsKey('Description')) {
                    $ClientProfile.Name = $Name
                    $ClientProfile.Description = $Description
                    $ClientProfile.Save()
                }

                if ($priorityDifference -lt 0) {
                    do {
                        $null = $ClientProfile.ClientProfileUpPriority()
                    } while ((++$priorityDifference))
                } elseif ($priorityDifference -gt 0) {
                    $priorityDifference = [math]::Min($priorityDifference, $clientProfiles.Count)
                    do {
                        $null = $ClientProfile.ClientProfileDownPriority()
                    } while ((--$priorityDifference))
                }
            }

            if ($PassThru) {
                $ClientProfile
            }
        } catch {
            Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $ClientProfile
        }
    }

    end {
        (Get-VmsManagementServer).ClientProfileFolder.ClearChildrenCache()
    }
}


function Set-VmsClientProfileAttributes {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('SmartClientProfiles')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter(Position = 0)]
        [System.Collections.IDictionary]
        $Attributes,

        [Parameter()]
        [string]
        $Namespace
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $namespaces = ($ClientProfile | Get-Member -MemberType Property -Name 'ClientProfile*ChildItems').Name -replace 'ClientProfile(.+)ChildItems', '$1'
        if (-not $MyInvocation.BoundParameters.ContainsKey('Namespace')) {
            $Namespace = $Attributes.Namespace
        }
        if ([string]::IsNullOrWhiteSpace($Namespace)) {
            Write-Error "Client profile attribute namespace required. Either supply the namespace using the Namespace parameter, or include a Namespace key in the Attributes dictionary with the appropriate namespace name as a string value."
            return
        } elseif ($Namespace -notin $namespaces) {
            Write-Error "Client profile namespace '$Namespace' not found. Namespaces include $($namespaces -join ', ')."
            return
        }

        $settings = $ClientProfile."ClientProfile$($Namespace)ChildItems"
        $availableKeys = $settings.GetPropertyKeys()
        $dirty = $false
        foreach ($key in $Attributes.Keys | Where-Object { $_ -ne 'Namespace'}) {
            if ($key -notin $availableKeys) {
                Write-Warning "Client profile attribute with key '$key' not found in client profile namespace '$Namespace'."
                continue
            }

            if ($Attributes[$key].Value) {
                $newValue = $Attributes[$key].Value.ToString()
            } else {
                $newValue = $Attributes[$key].ToString()
            }

            if ($settings.GetProperty($key) -cne $newValue -and $PSCmdlet.ShouldProcess("$($ClientProfile.Name)/$Namespace/$key", "Change value from '$($settings.GetProperty($key))' to '$newValue'")) {
                $settings.SetProperty($key, $newValue)
                $dirty = $true
            }

            $locked = $null
            if ("$($key)Locked" -in $availableKeys) {
                $locked = $settings.GetProperty("$($key)Locked")
            }
            if ($null -ne $locked -and $null -ne $Attributes[$key].Locked -and $locked -ne $Attributes[$key].Locked.ToString() -and $PSCmdlet.ShouldProcess("$($ClientProfile.Name)/$Namespace/$($key)Locked", "Change value from '$locked' to '$($Attributes[$key].Locked.ToString())'")) {
                $settings.SetProperty("$($key)Locked", $Attributes[$key].Locked.ToString())
                $dirty = $true
            }
        }
        if ($dirty) {
            $ClientProfile.Save()
        }
    }
}


function Connect-Vms {
    [CmdletBinding(DefaultParameterSetName = 'ConnectionProfile')]
    [OutputType([VideoOS.Platform.ConfigurationItems.ManagementServer])]
    [RequiresVmsConnection($false)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='The noun is an acronym.')]
    param (
        [Parameter(ParameterSetName = 'ConnectionProfile', ValueFromPipelineByPropertyName, Position = 0)]
        [Parameter(ParameterSetName = 'ServerAddress')]
        [Parameter(ParameterSetName = 'ShowDialog')]
        [string]
        $Name = 'default',

        [Parameter(ParameterSetName = 'ShowDialog', ValueFromPipelineByPropertyName)]
        [RequiresInteractiveSession()]
        [switch]
        $ShowDialog,

        [Parameter(ParameterSetName = 'ServerAddress', Mandatory, ValueFromPipelineByPropertyName)]
        [uri]
        $ServerAddress,

        [Parameter(ParameterSetName = 'ServerAddress', ValueFromPipelineByPropertyName)]
        [pscredential]
        $Credential,

        [Parameter(ParameterSetName = 'ServerAddress', ValueFromPipelineByPropertyName)]
        [switch]
        $BasicUser,

        [Parameter(ParameterSetName = 'ServerAddress', ValueFromPipelineByPropertyName)]
        [switch]
        $SecureOnly,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $IncludeChildSites,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $AcceptEula,

        [Parameter(ParameterSetName = 'ConnectionProfile')]
        [switch]
        $NoProfile
    )

    begin {
        Assert-VmsRequirementsMet
    }
        
    process {
        Disconnect-Vms
        
        switch ($PSCmdlet.ParameterSetName) {
            'ConnectionProfile' {
                $vmsProfile = GetVmsConnectionProfile -Name $Name
                if ($vmsProfile) {
                    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('IncludeChildSites')) {
                        $vmsProfile['IncludeChildSites'] = $IncludeChildSites
                    }
                    Connect-ManagementServer @vmsProfile -Force -ErrorAction Stop
                } else {
                    Connect-ManagementServer -ShowDialog -AcceptEula:$AcceptEula -IncludeChildSites:$IncludeChildSites -Force -ErrorAction Stop
                }
            }

            'ServerAddress' {
                $connectArgs = @{
                    ServerAddress     = $ServerAddress
                    SecureOnly        = $SecureOnly
                    IncludeChildSites = $IncludeChildSites
                    AcceptEula        = $AcceptEula
                }
                if ($Credential) {
                    $connectArgs.Credential = $Credential
                    $connectArgs.BasicUser = $BasicUser
                }
                Connect-ManagementServer @connectArgs -ErrorAction Stop
            }

            'ShowDialog' {
                if ($ShowDialog) {
                    $connectArgs = @{
                        ShowDialog        = $ShowDialog
                        IncludeChildSites = $IncludeChildSites
                        AcceptEula        = $AcceptEula
                    }
                    Connect-ManagementServer @connectArgs -ErrorAction Stop
                }
            }

            Default {
                throw "ParameterSetName '$_' not implemented."
            }
        }

        if (Test-VmsConnection) {
            if (-not $NoProfile -and ($PSCmdlet.ParameterSetName -eq 'ConnectionProfile' -or $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Name'))) {
                Save-VmsConnectionProfile -Name $Name -Force
            }
            
            Get-VmsManagementServer
        }
    }
}

Register-ArgumentCompleter -CommandName Connect-Vms, Get-VmsConnectionProfile, Save-VmsConnectionProfile, Remove-VmsConnectionProfile -ParameterName Name -ScriptBlock {
    $options = (GetVmsConnectionProfile -All).Keys | Sort-Object
    if ([string]::IsNullOrWhiteSpace($args[2])) {
        $wordToComplete = '*'
    } else {
        $wordToComplete = $args[2].Trim('''').Trim('"')
    }

    $options | ForEach-Object {
        if ($_ -like "$wordToComplete*") {
            if ($_ -match '\s') {
                "'$_'"
            } else {
                $_
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Connect-Vms -ParameterName ServerAddress -ScriptBlock {
    $options = (GetVmsConnectionProfile -All).Values | ForEach-Object { $_.ServerAddress.ToString() } | Sort-Object
    if ([string]::IsNullOrWhiteSpace($args[2])) {
        $wordToComplete = '*'
    } else {
        $wordToComplete = $args[2].Trim('''').Trim('"')
    }

    $options | ForEach-Object {
        if ($_ -like "$wordToComplete*") {
            if ($_ -match '\s') {
                "'$_'"
            } else {
                $_
            }
        }
    }
}

function Disconnect-Vms {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='The noun is an acronym.')]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ([milestonepstools.connection.milestoneconnection]::Instance) {
            Disconnect-ManagementServer
        }
    }
}

function Get-VmsConnectionProfile {
    [CmdletBinding(DefaultParameterSetName = 'Name')]
    [Alias('Get-Vms')]
    [OutputType([pscustomobject])]
    [RequiresVmsConnection($false)]
    param(
        [Parameter(ParameterSetName = 'Name', ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name = 'default',

        [Parameter(ParameterSetName = 'All')]
        [switch]
        $All
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $vmsProfiles = GetVmsConnectionProfile -All
        foreach ($profileName in $vmsProfiles.Keys | Sort-Object) {
            if ($All -or $profileName -eq $Name) {
                [pscustomobject]@{
                    Name              = $profileName
                    ServerAddress     = $vmsProfiles[$profileName].ServerAddress
                    Credential        = $vmsProfiles[$profileName].Credential
                    BasicUser         = $vmsProfiles[$profileName].BasicUser
                    SecureOnly        = $vmsProfiles[$profileName].SecureOnly
                    IncludeChildSites = $vmsProfiles[$profileName].SecureOnly
                    AcceptEula        = $vmsProfiles[$profileName].AcceptEula
                }
            }
        }
    }
}

function Remove-VmsConnectionProfile {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param (
        [Parameter(ParameterSetName = 'Name', Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [string[]]
        $Name,

        [Parameter(ParameterSetName = 'All')]
        [switch]
        $All
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        $vmsProfiles = GetVmsConnectionProfile -All
        if ($All) {
            $vmsProfiles.Clear()
        } else {
            $Name | ForEach-Object {
                $vmsProfiles.Remove($_)
            }
        }

        $vmsProfiles | Export-Clixml -Path (GetVmsConnectionProfilePath) -Force
    }
}

function Save-VmsConnectionProfile {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param(
        [Parameter(Position = 0)]
        [string]
        $Name = 'default',

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        $vmsProfiles = GetVmsConnectionProfile -All
        if ($vmsProfiles.ContainsKey($Name) -and -not $Force) {
            Write-Error "Connection profile '$Name' already exists. To overwrite it, use the -Force parameter."
            return
        }
        
        $vmsProfiles[$Name] = ExportVmsLoginSettings -ErrorAction Stop
        $vmsProfiles | Export-Clixml -Path (GetVmsConnectionProfilePath) -Force
    }
}

function Test-VmsConnection {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        $null -ne [MilestonePSTools.Connection.MilestoneConnection]::Instance
    }
}

function Get-VmsCameraMotion {
    [OutputType([VideoOS.Platform.ConfigurationItems.MotionDetection])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera
    )

    begin {
        Assert-VmsRequirementsMet -ErrorAction Stop
    }

    process {
        foreach ($currentDevice in $Camera) {
            $currentDevice.MotionDetectionFolder.MotionDetections[0] | Add-Member -MemberType NoteProperty -Name Camera -Value $currentDevice -PassThru
        }
    }
}

function Get-VmsDeviceEvent {
    [CmdletBinding()]
    [MilestonePSTools.RequiresVmsConnection()]
    [MilestonePSTools.RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.HardwareDeviceEventChildItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateVmsItemType('Hardware', 'Camera', 'Microphone', 'Speaker', 'Metadata', 'InputEvent', 'Output')]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem]
        $Device,

        [Parameter()]
        [SupportsWildcards()]
        [string]
        $Name = '*',

        [Parameter()]
        [bool]
        $Used,

        [Parameter()]
        [bool]
        $Enabled
    )

    begin {
        Assert-VmsRequirementsMet
    }
   
    process {
        $Device.HardwareDeviceEventFolder.ClearChildrenCache()
        $hardwareDeviceEvent = $Device.HardwareDeviceEventFolder.HardwareDeviceEvents | Select-Object -First 1
        $wildcardPattern = [system.management.automation.wildcardpattern]::new($Name, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        foreach ($childItem in $hardwareDeviceEvent.HardwareDeviceEventChildItems | Sort-Object DisplayName) {
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Name')) {
                if (-not $wildcardPattern.IsMatch($childItem.DisplayName)) {
                    continue
                }
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Used') -and $childItem.EventUsed -ne $Used) {
                continue
            }
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Enabled') -and $childItem.Enabled -ne $Enabled) {
                continue
            }
            
            # Used in Set-VmsDeviceEvent for more useful log messages and so that it's easy to know which event is associated with which device
            $childItem | Add-Member -MemberType NoteProperty -Name Device -Value $Device
            # Used in Set-VmsDeviceEvent because the .Save() method is on the parent HardwareDeviceEvent, not the HardwareDeviceEventChildItem.
            $childItem | Add-Member -MemberType NoteProperty -Name HardwareDeviceEvent -Value $hardwareDeviceEvent
            # Used in Set-VmsDeviceEvent to know whether to refresh our HardwareDeviceEvent before calling .Save().
            $hwPath = if ($Device.ParentItemPath -match '^Hardware') { $Device.ParentItemPath } else { $Device.Path }
            $childItem | Add-Member -MemberType NoteProperty -Name HardwarePath -Value $hwPath
            
            $childItem
        }
    }
}


function Set-VmsCameraMotion {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.Camera])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [VideoOS.Platform.ConfigurationItems.Camera[]]
        $Camera,

        [Parameter()]
        [ValidateSet('Normal', 'Optimized', 'Fast')]
        [string]
        $DetectionMethod,

        [Parameter()]
        [bool]
        $Enabled,

        [Parameter()]
        [string]
        $ExcludeRegions,

        [Parameter()]
        [bool]
        $GenerateMotionMetadata,

        [Parameter()]
        [ValidateSet('Grid8X8', 'Grid16X16', 'Grid32X32', 'Grid64X64')]
        [string]
        $GridSize,

        [Parameter()]
        [ValidateSet('Automatic', 'Off')]
        [RequiresVmsFeature('HardwareAcceleratedVMD')]
        [string]
        $HardwareAccelerationMode,

        [Parameter()]
        [bool]
        $KeyframesOnly,

        [Parameter()]
        [ValidateRange(0, 300)]
        [int]
        $ManualSensitivity,

        [Parameter()]
        [bool]
        $ManualSensitivityEnabled,

        [Parameter()]
        [ValidateSet('Ms100', 'Ms250', 'Ms500', 'Ms750', 'Ms1000')]
        [string]
        $ProcessTime,

        [Parameter()]
        [ValidateRange(0, 10000)]
        [int]
        $Threshold,

        [Parameter()]
        [bool]
        $UseExcludeRegions,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet -ErrorAction Stop
        $members = @{}
    }
    
    process {
        foreach ($currentDevice in $Camera) {
            $dirty = $false
            try {
                $motion = $currentDevice.MotionDetectionFolder.MotionDetections[0]
                if ($members.Count -eq 0) {
                    # Cache settable property names as keys in hashtable
                    $motion | Get-Member -MemberType Property | Where-Object Definition -match 'set;' | ForEach-Object {
                        $members[$_.Name] = $null
                    }
                }
                foreach ($parameter in $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator()) {
                    $key, $newValue = $parameter.Key, $parameter.Value
                    if (!$members.ContainsKey($key)) {
                        continue
                    } elseif ($motion.$key -eq $newValue) {
                        Write-Verbose "Motion detection setting '$key' is already '$newValue' on $currentDevice"
                        continue
                    }
                    Write-Verbose "Changing motion detection setting '$key' to '$newValue' on $currentDevice"
                    $motion.$key = $newValue
                    $dirty = $true
                }
                if ($PSCmdlet.ShouldProcess($currentDevice, "Update motion detection settings")) {
                    if ($dirty) {
                        $motion.Save()
                    }
                    if ($PassThru) {
                        $currentDevice
                    }
                }
            } catch {
                Write-Error -TargetObject $currentDevice -Exception $_.Exception -Message $_.Exception.Message -Category $_.CategoryInfo.Category
            }
        }
    }
}


function Set-VmsDeviceEvent {
    [CmdletBinding(SupportsShouldProcess)]
    [MilestonePSTools.RequiresVmsConnection()]
    [MilestonePSTools.RequiresVmsVersion('21.1')]
    [OutputType([VideoOS.Platform.ConfigurationItems.HardwareDeviceEventChildItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateScript({
                if ($null -eq ($_ | Get-Member -MemberType NoteProperty -Name HardwareDeviceEvent)) {
                    throw 'DeviceEvent must be returned by Get-VmsDeviceEvent or it does not have a NoteProperty member named HardwareDeviceEvent.'
                }
                $true
            })]
        [VideoOS.Platform.ConfigurationItems.HardwareDeviceEventChildItem]
        $DeviceEvent,

        [Parameter()]
        [bool]
        $Used,

        [Parameter()]
        [bool]
        $Enabled,

        [Parameter()]
        [string]
        $Index,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        $modified = @{}
    }
   
    process {
        $changes = @{}
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Used') -and $DeviceEvent.EventUsed -ne $Used) {
            $changes['EventUsed'] = $Used            
        }
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Enabled') -and $DeviceEvent.Enabled -ne $Enabled) {
            $changes['Enabled'] = $Enabled         
        }
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Index') -and $DeviceEvent.EventIndex -ne $Index) {
            $changes['EventIndex'] = $Index
        }

        # Management Client sets EventUsed and Enabled to the same value when you add or remove them in the UI.
        if ($changes.ContainsKey('EventUsed') -and $changes['EventUsed'] -ne $DeviceEvent.Enabled) {
            $changes['Enabled'] = $changes['EventUsed']
        }

        if ($changes.Count -gt 0 -and $PSCmdlet.ShouldProcess($DeviceEvent.Device.Name, "Update '$($DeviceEvent.DisplayName)' device event settings")) {
            <#
             # BUG #627670 - This method does not work because you can only call Save() on the most recently queried HardwareDeviceEvent.
             # The LastModified datetime for the Hardware associated with the most recently queried HardwareDeviceEvent must match the
             # LastModified timestamp of the hardware associated with the HardwareDeviceEvent.Save() method.
             # This method will be ~30% faster if we can change the server-side behavior.
            
            foreach ($kvp in $changes.GetEnumerator()) {
                $DeviceEvent.($kvp.Key) = $kvp.Value
            }
            $modified[$DeviceEvent.Path] = $DeviceEvent
            
            #>
            
            
            # Alternate method to work around issue described in BUG #627670
            if (-not $modified.ContainsKey($DeviceEvent.Path)) {
                $modified[$DeviceEvent.Path] = [pscustomobject]@{
                    Device              = $DeviceEvent.Device
                    HardwareDeviceEvent = $DeviceEvent.HardwareDeviceEvent
                    Changes             = @{}
                }
            }
            $modified[$DeviceEvent.Path].Changes[$DeviceEvent.Id] = $changes
        } elseif ($PassThru) {
            $DeviceEvent
        }
    }

    end {
        <#
             # BUG #627670 - This method does not work because you can only call Save() on the most recently queried HardwareDeviceEvent.
             # The LastModified datetime for the Hardware associated with the most recently queried HardwareDeviceEvent must match the
             # LastModified timestamp of the hardware associated with the HardwareDeviceEvent.Save() method.
             # This method will be ~30% faster if we can change the server-side behavior.

             foreach ($item in $modified.Values) {
                try {
                    Write-Verbose "Saving device event changes on $($item.Device.Name)."
                    $item.HardwareDeviceEvent.Save()
                    if ($PassThru) {
                        $item
                    }
                } catch {
                    throw
                }
            }
        #>

        # Alternate method to work around issue described in BUG #627670
        foreach ($record in $modified.Values) {
            $record.Device.HardwareDeviceEventFolder.ClearChildrenCache()
            $hardwareDeviceEvent = [VideoOS.Platform.ConfigurationItems.HardwareDeviceEvent]::new($record.HardwareDeviceEvent.ServerId, $record.HardwareDeviceEvent.Path)
            $modifiedChildItems = $hardwareDeviceEvent.HardwareDeviceEventChildItems | Where-Object { $record.Changes.ContainsKey($_.Id) }
            foreach ($eventId in $record.Changes.Keys) {
                if (($childItem = $modifiedChildItems | Where-Object Id -eq $eventId)) {
                    foreach ($change in $record.Changes[$eventId].GetEnumerator()) {
                        Write-Verbose "Setting $($change.Key) = $($change.Value) for event '$($childItem.DisplayName)' on $($record.Device.Name)."
                        $childItem.($change.Key) = $change.Value
                    }
                } else {
                    throw "HardwareDeviceEventChildItem with ID $eventId not found on $($record.Device.Name)."
                }
            }
            Write-Verbose "Saving changes to HardwareDeviceEvents on $($record.Device.Name)"
            $hardwareDeviceEvent.Save()
            if ($PassThru) {
                $record.Device.HardwareDeviceEventFolder.ClearChildrenCache()
                $record.Device | Get-VmsDeviceEvent | Where-Object Id -in $modifiedChildItems.Id
            }
        }
    }
}


function Add-VmsFailoverRecorder {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $FailoverGroup,

        [Parameter(Mandatory, Position = 0)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverRecorder]])]
        [MipItemTransformation([FailoverRecorder])]
        [FailoverRecorder[]]
        $FailoverRecorder,

        [Parameter()]
        [int]
        $Position = 0
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($failover in $FailoverRecorder) {
            if ($PSCmdlet.ShouldProcess("FailoverGroup $($FailoverGroup.Name)", "Add $($failover.Name)")) {
                try {
                    $serverTask = (Get-VmsManagementServer).FailoverGroupFolder.MoveFailoverGroup($failover.Path, $FailoverGroup.Path, $Position)
                    while ($serverTask.Progress -lt 100) {
                        Start-Sleep -Milliseconds 100
                        $serverTask.UpdateState()
                    }
                    if ($serverTask.State -ne 'Success') {
                        Write-Error -Message "MoveFailoverGroup returned with ErrorCode $($serverTask.ErrorCode). $($serverTask.ErrorText)" -TargetObject $serverTask
                        return
                    }
                } catch {
                    throw
                }
            }
        }
    }
}

function Get-VmsFailoverGroup {
    [CmdletBinding(DefaultParameterSetName = 'Name')]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverGroup])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Id')]
        [guid]
        $Id,

        [Parameter(Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'Name')]
        [string]
        $Name = '*'
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'Id' {
                try {
                    $serverId = (Get-VmsManagementServer).ServerId
                    $path = 'FailoverGroup[{0}]' -f $Id
                    [VideoOS.Platform.ConfigurationItems.FailoverGroup]::new($serverId, $path)
                } catch {
                    throw
                }
            }
            'Name' {
                foreach ($group in (Get-VmsManagementServer).FailoverGroupFolder.FailoverGroups | Where-Object Name -like $Name) {
                    $group
                }
            }
            Default {
                throw "ParameterSetName '$_' not implemented."
            }
        }
    }
}


Register-ArgumentCompleter -CommandName Get-VmsFailoverGroup -ParameterName Name -ScriptBlock {
    $values = (Get-VmsFailoverGroup).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsFailoverRecorder {
    [CmdletBinding(DefaultParameterSetName = 'FailoverGroup')]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverRecorder])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'FailoverGroup')]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $FailoverGroup,

        [Parameter(ParameterSetName = 'FailoverGroup')]
        [switch]
        $Recurse,

        [Parameter(Mandatory, ParameterSetName = 'HotStandby')]
        [switch]
        $HotStandby,

        [Parameter(Mandatory, ParameterSetName = 'Unassigned')]
        [switch]
        $Unassigned,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Id')]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
        if ($HotStandby -or $Unassigned) {
            $failovers = (Get-VmsManagementServer).FailoverGroupFolder.FailoverRecorders
            $hotFailovers = Get-VmsRecordingServer | Foreach-Object {
                $_.RecordingServerFailoverFolder.RecordingServerFailovers[0].HotStandby
            } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'FailoverGroup' {
                if ($FailoverGroup) {
                    $FailoverGroup.FailoverRecorderFolder.FailoverRecorders
                } else {
                    (Get-VmsManagementServer).FailoverGroupFolder.FailoverRecorders
                    if ($Recurse) {
                        Get-VmsFailoverGroup | Get-VmsFailoverRecorder
                    }
                }
            }
            'HotStandby' {
                if ($failovers.Count -eq 0) {
                    return
                }
                $failovers | Where-Object Path -in $hotFailovers
            }
            'Unassigned' {
                if ($failovers.Count -eq 0) {
                    return
                }
                $failovers | Where-Object Path -notin $hotFailovers
            }
            'Id' {
                try {
                    $serverId = (Get-VmsManagementServer).ServerId
                    $path = 'FailoverRecorder[{0}]' -f $Id
                    [VideoOS.Platform.ConfigurationItems.FailoverRecorder]::new($serverId, $path)
                } catch {
                    throw
                }
            }
            Default {
                throw "ParameterSetName '$_' not implemented."
            }
        }
    }
}

function New-VmsFailoverGroup {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverGroup])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Description
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (-not $PSCmdlet.ShouldProcess("FailoverGroup $Name", "Create")) {
            return
        }
        try {
            $serverTask = (Get-VmsManagementServer).FailoverGroupFolder.AddFailoverGroup($Name, $Description)
            while ($serverTask.Progress -lt 100) {
                Start-Sleep -Milliseconds 100
                $serverTask.UpdateState()
            }
            if ($serverTask.State -ne 'Success') {
                Write-Error -Message "AddFailoverGroup returned with ErrorCode $($serverTask.ErrorCode). $($serverTask.ErrorText)" -TargetObject $serverTask
                return
            }
            $id = $serverTask.Path -replace 'FailoverGroup\[(.+)\]', '$1'
            Get-VmsFailoverGroup -Id $id
        } catch {
            throw
        }
    }
}


function Remove-VmsFailoverGroup {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param(
        [Parameter(Mandatory, ValueFromPipeline, Position = 0)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $FailoverGroup,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ShouldProcess($FailoverGroup.Name, "Remove FailoverGroup")) {
            if ($FailoverGroup.FailoverRecorderFolder.FailoverRecorders.Count -gt 0) {
                if (-not $Force) {
                    throw "Cannot delete FailoverGroup with members. Try again with -Force switch to remove member FailoverRecorders."
                }
                $FailoverGroup | Get-VmsFailoverRecorder | Foreach-Object {
                    $FailoverGroup | Remove-VmsFailoverRecorder -FailoverRecorder $_ -Confirm:$false
                }
            }
            try {
                $serverTask = (Get-VmsManagementServer).FailoverGroupFolder.RemoveFailoverGroup($FailoverGroup.Path)
                while ($serverTask.Progress -lt 100) {
                    Start-Sleep -Milliseconds 100
                    $serverTask.UpdateState()
                }
                if ($serverTask.State -ne 'Success') {
                    Write-Error -Message "RemoveFailoverGroup returned with ErrorCode $($serverTask.ErrorCode). $($serverTask.ErrorText)" -TargetObject $serverTask
                }
            } catch {
                throw
            }
        }
    }
}

function Remove-VmsFailoverRecorder {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $FailoverGroup,

        [Parameter(Mandatory, Position = 0)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverRecorder]])]
        [MipItemTransformation([FailoverRecorder])]
        [FailoverRecorder]
        $FailoverRecorder
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (-not $PSCmdlet.ShouldProcess("FailoverGroup $($FailoverGroup.Name)", "Remove $($FailoverRecorder)")) {
            return
        }

        try {
            $serverTask = (Get-VmsManagementServer).FailoverGroupFolder.MoveFailoverGroup($FailoverRecorder.Path, [string]::Empty, 0)
            while ($serverTask.Progress -lt 100) {
                Start-Sleep -Milliseconds 100
                $serverTask.UpdateState()
            }
            if ($serverTask.State -ne 'Success') {
                Write-Error -Message "MoveFailoverGroup returned with ErrorCode $($serverTask.ErrorCode). $($serverTask.ErrorText)" -TargetObject $serverTask
                return
            }
        } catch {
            throw
        } finally {
            $FailoverGroup.ClearChildrenCache()
        }
    }
}

function Set-VmsFailoverGroup {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverGroup])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverGroup]])]
        [MipItemTransformation([FailoverGroup])]
        [FailoverGroup]
        $FailoverGroup,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $dirty = $false
        if (-not [string]::IsNullOrWhiteSpace($Name) -and $Name -cne $FailoverGroup.Name -and $PSCmdlet.ShouldProcess($FailoverGroup.Name, "Rename to $Name")) {
            $FailoverGroup.Name = $Name
            $dirty = $true
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Description') -and $Description -cne $FailoverGroup.Description -and $PSCmdlet.ShouldProcess($FailoverGroup.Name, "Set Description to $Description")) {
            $FailoverGroup.Description = $Description
            $dirty = $true
        }
        if ($dirty) {
            try {
                $FailoverGroup.Save()
            } catch {
                throw
            }
        }
        if ($PassThru) {
            $FailoverGroup
        }
    }
}

function Set-VmsFailoverRecorder {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.FailoverRecorder])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('21.2')]
    [RequiresVmsFeature('RecordingServerFailover')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[FailoverRecorder]])]
        [MipItemTransformation([FailoverRecorder])]
        [FailoverRecorder]
        $FailoverRecorder,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [bool]
        $Enabled,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [string]
        $DatabasePath,

        [Parameter()]
        [ValidateRange(0, 65535)]
        [int]
        $UdpPort,

        [Parameter()]
        [string]
        $MulticastServerAddress,

        [Parameter()]
        [bool]
        $PublicAccessEnabled,

        [Parameter()]
        [string]
        $PublicWebserverHostName,

        [Parameter()]
        [ValidateRange(0, 65535)]
        [int]
        $PublicWebserverPort,

        [Parameter()]
        [switch]
        $Unassigned,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($Unassigned) {
            if ($FailoverRecorder.ParentItemPath -eq '/') {
                Get-VmsRecordingServer | Where-Object {
                    $_.RecordingServerFailoverFolder.RecordingServerFailovers[0].HotStandby -eq $FailoverRecorder.Path
                } | Set-VmsRecordingServer -DisableFailover -Verbose:($VerbosePreference -eq 'Continue')
            } else {
                $group = Get-VmsFailoverGroup -Id ($FailoverRecorder.ParentItemPath -replace '\w+\[(.+)\]', '$1')
                $group | Remove-VmsFailoverRecorder -FailoverRecorder $FailoverRecorder
            }
        }

        $dirty = $false
        $settableProperties = ($FailoverRecorder | Get-Member -MemberType Property | Where-Object Definition -match 'set;').Name
        foreach ($property in $MyInvocation.BoundParameters.GetEnumerator() | Where-Object Key -in $settableProperties) {
            $key = $property.Key
            $newValue = $property.Value
            if ($FailoverRecorder.$key -cne $newValue -and $PSCmdlet.ShouldProcess("FailoverRecorder $($FailoverRecorder.Name)", "Change $key to $newValue")) {
                $FailoverRecorder.$key = $newValue
                $dirty = $true
            }
        }
        if ($dirty) {
            try {
                if ($FailoverRecorder.MulticastServerAddress -eq [string]::Empty) {
                    Write-Verbose 'Changing MulticastServerAddress to 0.0.0.0 because an empty string will not pass validation as of XProtect 2023 R1. Bug #581349.'
                    $FailoverRecorder.MulticastServerAddress = '0.0.0.0'
                }
                $FailoverRecorder.Save()
            } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
                $FailoverRecorder = Get-VmsFailoverRecorder -Id $FailoverRecorder.Id
                $_ | HandleValidateResultException -TargetObject $FailoverRecorder -ItemName $FailoverRecorder.Name
            }
        }
        if ($PassThru) {
            $FailoverRecorder
        }
    }
}

function Move-VmsHardware {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Hardware])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[Hardware]])]
        [MipItemTransformation([Hardware])]
        [Hardware[]]
        $Hardware,

        [Parameter(Mandatory, Position = 1, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer]
        $DestinationRecorder,

        [Parameter(Mandatory, Position = 2, ValueFromPipelineByPropertyName)]
        [StorageNameTransformAttribute()]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $DestinationStorage,

        [Parameter()]
        [switch]
        $AllowDataLoss,

        [Parameter()]
        [switch]
        $SkipDriverCheck,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        $recordersByPath = @{}
        $moveSucceeded = $false
    }

    process {
        $recordersByPath[$DestinationRecorder.Path] = $DestinationRecorder
        foreach ($hw in $Hardware) {
            try {
                if ($null -eq $recordersByPath[$hw.ParentItemPath]) {
                    $recordersByPath[$hw.ParentItemPath] = Get-VmsRecordingServer | Where-Object Path -EQ $hw.ParentItemPath
                }
    
                if ($DestinationRecorder.Path -eq $hw.ParentItemPath) {
                    Write-Error "Hardware '$($hw.Name)' is already assigned to recorder '$($DestinationRecorder.Name)'." -TargetObject $hw
                    continue
                }
    
                if (-not $SkipDriverCheck) {
                    $srcDriver = $recordersByPath[$hw.ParentItemPath].HardwareDriverFolder.HardwareDrivers | Where-Object Path -EQ $hw.HardwareDriverPath
                    $dstDriver = $DestinationRecorder.HardwareDriverFolder.HardwareDrivers | Where-Object Path -EQ $hw.HardwareDriverPath
                    if ($null -eq $srcDriver) {
                        Write-Error "The current driver for hardware '$($hw.Name)' can not be determined."
                        continue
                    }
                    if ($null -eq $dstDriver) {
                        Write-Error "Destination recording server '$($DestinationRecorder.Name)' does not appear to have the following driver installed: $($srcDriver.Name) ($($srcDriver.Number))."
                        continue
                    }
                    if ("$($srcDriver.DriverVersion).$($srcDriver.DriverRevision)" -cne "$($dstDriver.DriverVersion).$($dstDriver.DriverRevision)") {
                        Write-Error "Destination recording server '$($DestinationRecorder.Name)' does not have the same driver version as source recording server '$($recordersByPath[$hw.ParentItemPath].Name)': Source = '$($srcDriver.DriverVersion), $($srcDriver.DriverRevision)', Destination = '$($dstDriver.DriverVersion), $($dstDriver.DriverRevision)'."
                        continue
                    }
                    Write-Verbose "Device pack driver versions and revisions match for driver '$($srcDriver.Name)': Source = '$($srcDriver.DriverVersion), $($srcDriver.DriverRevision)', Destination = '$($dstDriver.DriverVersion), $($dstDriver.DriverRevision)'."
                }
    
                if ($PSCmdlet.ShouldProcess($hw.Name, "Move hardware to $($DestinationRecorder.Name) / $($DestinationStorage.Name)")) {
                    $taskInfo = $hw.MoveHardware()
                    $taskInfo.SetProperty('DestinationRecordingServer', $DestinationRecorder.Path)
                    $taskInfo.SetProperty('DestinationStorage', $DestinationStorage.Path)
                    $taskInfo.SetProperty('ignoreSourceRecordingServer', $AllowDataLoss)
                    $result = $taskInfo.ExecuteDefault() | Wait-VmsTask -Cleanup
                    $errorText = ($result.Properties | Where-Object Key -EQ 'ErrorText').Value
                    if (-not [string]::IsNullOrWhiteSpace($errorText)) {
                        throw $errorText
                    }
                    $moveSucceeded = $true
    
                    foreach ($property in $result.Properties) {
                        if ($property.Key -match 'Warning' -and -not [string]::IsNullOrWhiteSpace($property.Value)) {
                            Write-Warning $property.Value
                        }
                    }
                }
                if ($PassThru) {
                    Get-VmsHardware -Id $hw.Id
                }
            } catch {
                throw
            }
        }
    }

    end {
        if ($moveSucceeded) {
            foreach ($recorder in $recordersByPath.Values) {
                Write-Verbose "Clearing HardwareFolder cache for $($recorder.Name)"
                $recorder.HardwareFolder.ClearChildrenCache()
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Move-VmsHardware -ParameterName DestinationStorage -ScriptBlock {
    $recorder = $null
    if ($null -eq ($recorder = $args[4]['DestinationRecorder'] -as [VideoOS.Platform.ConfigurationItems.RecordingServer])) {
        $recorder = Get-VmsRecordingServer | Where-Object Name -eq "$($args[4]['DestinationRecorder'])"
        if ($null -eq $recorder -or $recorder.Count -ne 1) {
            return
        }
    }
    $storages = $recorder | Get-VmsStorage | Select-Object -ExpandProperty Name -Unique | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $storages
}


function Assert-VmsLicensedFeature {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (-not (Test-VmsLicensedFeature -Name $Name)) {
            $e = [VideoOS.Platform.NotSupportedMIPException]::new("The feature ""$Name"" is not enabled on your VMS.")
            Write-Error -Message $e.Message -Exception $e -Category NotEnabled -TargetObject $Name
        }
    }
}

Register-ArgumentCompleter -CommandName Assert-VmsLicensedFeature -ParameterName Name -ScriptBlock {
    $values = (Get-VmsSystemLicense).FeatureFlags | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsSystemLicense {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.License.SystemLicense])]
    [RequiresVmsConnection()]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        [MilestonePSTools.Connection.MilestoneConnection]::Instance.SystemLicense
    }
}


function Test-VmsLicensedFeature {
    [CmdletBinding()]
    [OutputType([bool])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [string]
        $Name
    )

    begin {
        Assert-VmsRequirementsMet
        $license = Get-VmsSystemLicense
    }

    process {
        $license.IsFeatureEnabled($Name)
    }
}

Register-ArgumentCompleter -CommandName Test-VmsLicensedFeature -ParameterName Name -ScriptBlock {
    $values = (Get-VmsSystemLicense).FeatureFlags | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsLicenseDetails {
    [CmdletBinding()]
    [Alias('Get-LicenseDetails')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseDetailChildItem])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param ()

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        (Get-VmsLicenseInfo).LicenseDetailFolder.LicenseDetailChildItems
    }
}


function Get-VmsLicensedProducts {
    [CmdletBinding()]
    [Alias('Get-LicensedProducts')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInstalledProductChildItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param ()

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        (Get-VmsLicenseInfo).LicenseInstalledProductFolder.LicenseInstalledProductChildItems
    }
}


function Get-VmsLicenseInfo {
    [CmdletBinding()]
    [Alias('Get-LicenseInfo')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseInformation])]
    param ()

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        $site = Get-VmsSite
        [VideoOS.Platform.ConfigurationItems.LicenseInformation]::new($site.FQID.ServerId, "LicenseInformation[$($site.FQID.ObjectId)]")
    }
}


function Get-VmsLicenseOverview {
    [CmdletBinding()]
    [Alias('Get-LicenseOverview')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseOverviewAllChildItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    param ()

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        $licenseInfo = Get-VmsLicenseInfo
        $licenseInfo.LicenseOverviewAllFolder.LicenseOverviewAllChildItems
    }
}


function Invoke-VmsLicenseActivation {
    [CmdletBinding()]
    [Alias('Invoke-LicenseActivation')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.2')]
    [OutputType([VideoOS.Platform.ConfigurationItems.LicenseDetailChildItem])]
    param (
        [Parameter(Mandatory)]
        [pscredential]
        $Credential,

        [Parameter()]
        [switch]
        $EnableAutoActivation,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        Show-DeprecationWarning $MyInvocation
    }

    process {
        try {
            $licenseInfo = Get-VmsLicenseInfo
            $result = $licenseInfo.ActivateLicense($Credential.UserName, $Credential.Password, $EnableAutoActivation) | Wait-VmsTask -Title 'Performing online license activation' -Cleanup
            $state = ($result.Properties | Where-Object Key -eq 'State').Value
            if ($state -eq 'Success') {
                if ($PassThru) {
                    Get-VmsLicenseDetails
                }
            } else {
                $errorText = ($result.Properties | Where-Object Key -eq 'ErrorText').Value
                if ([string]::IsNullOrWhiteSpace($errorText)) {
                    $errorText = "Unknown error."
                }
                Write-Error "Call to ActivateLicense failed. $($errorText.Trim('.'))."
            }
        } catch {
            Write-Error -Message $_.Message -Exception $_.Exception
        }
    }
}


function Get-MobileServerInfo {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $mobServerPath = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\WOW6432Node\Milestone\XProtect Mobile Server' -Name INSTALLATIONFOLDER
            [Xml]$doc = Get-Content "$mobServerPath.config" -ErrorAction Stop

            $xpath = "/configuration/ManagementServer/Address/add[@key='Ip']"
            $msIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            $xpath = "/configuration/ManagementServer/Address/add[@key='Port']"
            $msPort = $doc.SelectSingleNode($xpath).Attributes['value'].Value

            $xpath = "/configuration/HttpMetaChannel/Address/add[@key='Port']"
            $httpPort = [int]::Parse($doc.SelectSingleNode($xpath).Attributes['value'].Value)
            $xpath = "/configuration/HttpMetaChannel/Address/add[@key='Ip']"
            $httpIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            if ($httpIp -eq '+') { $httpIp = '0.0.0.0'}

            $xpath = "/configuration/HttpSecureMetaChannel/Address/add[@key='Port']"
            $httpsPort = [int]::Parse($doc.SelectSingleNode($xpath).Attributes['value'].Value)
            $xpath = "/configuration/HttpSecureMetaChannel/Address/add[@key='Ip']"
            $httpsIp = $doc.SelectSingleNode($xpath).Attributes['value'].Value
            if ($httpsIp -eq '+') { $httpsIp = '0.0.0.0'}
            try {
                $hash = Get-HttpSslCertThumbprint -IPPort "$($httpsIp):$($httpsPort)" -ErrorAction Stop
            } catch {
                $hash = $null
            }
            $info = [PSCustomObject]@{
                Version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($mobServerPath).FileVersion;
                ExePath = $mobServerPath;
                ConfigPath = "$mobServerPath.config";
                ManagementServerIp = $msIp;
                ManagementServerPort = $msPort;
                HttpIp = $httpIp;
                HttpPort = $httpPort;
                HttpsIp = $httpsIp;
                HttpsPort = $httpsPort;
                CertHash = $hash
            }
            $info
        } catch {
            Write-Error $_
        }
    }
}


function Set-XProtectCertificate {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection($false)]
    [RequiresElevation()]
    param (
        # Specifies the Milestone component on which to update the certificate
        # - Server: Applies to communication between Management Server and Recording Server, as well as client connections to the HTTPS port for the Management Server.
        # - StreamingMedia: Applies to all connections to Recording Servers. Typically on port 7563.
        # - MobileServer: Applies to HTTPS connections to the Milestone Mobile Server.
        [Parameter(Mandatory)]
        [ValidateSet('Server', 'StreamingMedia', 'MobileServer', 'EventServer')]
        [string]
        $VmsComponent,

        # Specifies that encryption for the specified Milestone XProtect service should be disabled
        [Parameter(ParameterSetName = 'Disable')]
        [switch]
        $Disable,

        # Specifies the thumbprint of the certificate to apply to Milestone XProtect service
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Enable')]
        [string]
        $Thumbprint,

        # Specifies the Windows user account for which read access to the private key is required
        [Parameter(ParameterSetName = 'Enable')]
        [string]
        $UserName,

        # Specifies the path to the Milestone Server Configurator executable. The default location is C:\Program Files\Milestone\Server Configurator\ServerConfigurator.exe
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $ServerConfiguratorPath = 'C:\Program Files\Milestone\Server Configurator\ServerConfigurator.exe',

        # Specifies that all certificates issued to
        [Parameter(ParameterSetName = 'Enable')]
        [switch]
        $RemoveOldCert,

        # Specifies that the Server Configurator process should be terminated if it's currently running
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet

        $certGroups = @{
            Server         = '84430eb7-847c-422d-aa00-7915cd0d7a65'
            StreamingMedia = '549df21d-047c-456b-958e-99e65dd8b3ec'
            MobileServer   = '76cfc719-a852-4210-913e-703eadab139a'
            EventServer    = '7e02e0f5-549d-4113-b8de-bda2c1f38dbf'
        }

        $knownExitCodes = @{
            0  = 'Success'
            -1 = 'Unknown error'
            -2 = 'Invalid arguments'
            -3 = 'Invalid argument value'
            -4 = 'Another instance is running'
        }
    }

    process {
        $utility = [IO.FileInfo]$ServerConfiguratorPath
        if (-not $utility.Exists) {
            $exception = [System.IO.FileNotFoundException]::new("Milestone Server Configurator not found at $ServerConfiguratorPath", $utility.FullName)
            Write-Error -Message $exception.Message -Exception $exception
            return
        }
        if ($utility.VersionInfo.FileVersion -lt [version]'20.3') {
            Write-Error "Server Configurator version 20.3 is required as the command-line interface for Server Configurator was introduced in version 2020 R3. The current version appears to be $($utility.VersionInfo.FileVersion). Please upgrade to version 2020 R3 or greater."
            return
        }
        Write-Verbose "Verified Server Configurator version $($utility.VersionInfo.FileVersion) is available at $ServerConfiguratorPath"

        $newCert = Get-ChildItem -Path "Cert:\LocalMachine\My\$Thumbprint" -ErrorAction Ignore
        if ($null -eq $newCert -and -not $Disable) {
            Write-Error "Certificate not found in Cert:\LocalMachine\My with thumbprint '$Thumbprint'. Please make sure the certificate is installed in the correct certificate store."
            return
        } elseif ($Thumbprint) {
            Write-Verbose "Located certificate in Cert:\LocalMachine\My with thumbprint $Thumbprint"
        }

        # Add read access to the private key for the specified certificate if UserName was specified
        if (-not [string]::IsNullOrWhiteSpace($UserName)) {
            try {
                Write-Verbose "Ensuring $UserName has the right to read the private key for the specified certificate"
                $newCert | Set-CertKeyPermission -UserName $UserName
            } catch {
                Write-Error -Message "Error granting user '$UserName' read access to the private key for certificate with thumbprint $Thumbprint" -Exception $_.Exception
            }
        }

        if ($Force) {
            if ($PSCmdlet.ShouldProcess("ServerConfigurator", "Kill process if running")) {
                Get-Process -Name ServerConfigurator -ErrorAction Ignore | Foreach-Object {
                    Write-Verbose 'Server Configurator is currently running. The Force switch was provided so it will be terminated.'
                    $_ | Stop-Process
                }
            }
        }

        $procParams = @{
            FilePath               = $utility.FullName
            Wait                   = $true
            PassThru               = $true
            RedirectStandardOutput = Join-Path -Path ([system.environment]::GetFolderPath([system.environment+specialfolder]::ApplicationData)) -ChildPath ([io.path]::GetRandomFileName())
        }
        if ($Disable) {
            $procParams.ArgumentList = '/quiet', '/disableencryption', "/certificategroup=$($certGroups.$VmsComponent)"
        } else {
            $procParams.ArgumentList = '/quiet', '/enableencryption', "/certificategroup=$($certGroups.$VmsComponent)", "/thumbprint=$Thumbprint"
        }
        $argumentString = [string]::Join(' ', $procParams.ArgumentList)
        Write-Verbose "Running Server Configurator with the following arguments: $argumentString"

        if ($PSCmdlet.ShouldProcess("ServerConfigurator", "Start process with arguments '$argumentString'")) {
            $result = Start-Process @procParams
            if ($result.ExitCode -ne 0) {
                Write-Error "Server Configurator exited with code $($result.ExitCode). $($knownExitCodes.$($result.ExitCode))"
                return
            }
        }

        if ($RemoveOldCert) {
            $oldCerts = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Subject -eq $newCert.Subject -and $_.Thumbprint -ne $newCert.Thumbprint }
            if ($null -eq $oldCerts) {
                Write-Verbose "No other certificates found matching the subject name $($newCert.Subject)"
                return
            }
            foreach ($cert in $oldCerts) {
                if ($PSCmdlet.ShouldProcess($cert.Thumbprint, "Remove certificate from certificate store")) {
                    Write-Verbose "Removing certificate with thumbprint $($cert.Thumbprint)"
                    $cert | Remove-Item
                }
            }
        }
    }
}


function Get-CameraRecordingStats {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param(
        # Specifies the Id's of cameras for which to retrieve recording statistics
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [guid[]]
        $Id,

        # Specifies the timestamp from which to start retrieving recording statistics. Default is 7 days prior to 12:00am of the current day.
        [Parameter()]
        [datetime]
        $StartTime = (Get-Date).Date.AddDays(-7),

        # Specifies the timestamp marking the end of the time period for which to retrieve recording statistics. The default is 12:00am of the current day.
        [Parameter()]
        [datetime]
        $EndTime = (Get-Date).Date,

        # Specifies the type of sequence to get statistics on. Default is RecordingSequence.
        [Parameter()]
        [ValidateSet('RecordingSequence', 'MotionSequence')]
        [string]
        $SequenceType = 'RecordingSequence',

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($EndTime -le $StartTime) {
            throw "EndTime must be greater than StartTime"
        }

        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose "Creating a runspace pool"
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [guid]$Id,
                [datetime]$StartTime,
                [datetime]$EndTime,
                [string]$SequenceType
            )

            $sequences = Get-SequenceData -Path "Camera[$Id]" -SequenceType $SequenceType -StartTime $StartTime -EndTime $EndTime -CropToTimeSpan
            $recordedMinutes = $sequences | Foreach-Object {
                ($_.EventSequence.EndDateTime - $_.EventSequence.StartDateTime).TotalMinutes
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            [pscustomobject]@{
                DeviceId = $Id
                StartTime = $StartTime
                EndTime = $EndTime
                SequenceCount = $sequences.Count
                TimeRecorded = [timespan]::FromMinutes($recordedMinutes)
                PercentRecorded = [math]::Round(($recordedMinutes / ($EndTime - $StartTime).TotalMinutes * 100), 1)
            }
        }

        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($cameraId in $Id) {
                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                    Id = $cameraId
                    StartTime = $StartTime
                    EndTime = $EndTime
                    SequenceType = $SequenceType
                }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                    DeviceId = $cameraId
                    PowerShell = $ps
                    Result = $asyncResult
                })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        if ($AsHashTable) {
                            $hashTable.$($thread.DeviceId.ToString()) = $null
                        }
                        else {
                            $obj = [ordered]@{
                                DeviceId = $thread.DeviceId.ToString()
                                RecordingStats = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.DeviceId.ToString()) = $result
                            }
                            else {
                                $obj.RecordingStats = $result
                            }
                        }
                        catch {
                            Write-Error $_
                        }
                        finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | Foreach-Object { [void]$threads.Remove($_)}
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break;
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        }
        finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}


function Get-CurrentDeviceStatus {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    [MilestonePSTools.RequiresVmsConnection()]
    param (
        # Specifies one or more Recording Server ID's to which the results will be limited. Omit this parameter if you want device status from all Recording Servers
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        # Specifies the type of devices to include in the results. By default only cameras will be included and you can expand this to include all device types
        [Parameter()]
        [ValidateSet('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input event', 'Output', 'Event', 'Hardware', 'All')]
        [string[]]
        $DeviceType = 'Camera',

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($DeviceType -contains 'All') {
            $DeviceType = @('Camera', 'Microphone', 'Speaker', 'Metadata', 'Input event', 'Output', 'Event', 'Hardware')
        }
        $includedDeviceTypes = $DeviceType | ForEach-Object { [videoos.platform.kind]::$_ }

        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose 'Creating a runspace pool'
            $iss = [initialsessionstate]::CreateDefault()
            $moduleManifest = (Get-Module MilestonePSTools).Path -replace 'psm1$', 'psd1'
            $iss.ImportPSModule($moduleManifest)
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1), $iss, (Get-Host))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [uri]$Uri,
                [guid[]]$DeviceIds
            )
            try {
                $client = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($Uri)
                $client.GetCurrentDeviceStatus((Get-VmsToken), $deviceIds)
            } catch {
                throw
            }
        }

        Write-Verbose 'Retrieving recording server information'
        $managementServer = [videoos.platform.configuration]::Instance.GetItems([videoos.platform.itemhierarchy]::SystemDefined) | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Server -and $_.FQID.ObjectId -eq (Get-VmsManagementServer).Id }
        $recorders = $managementServer.GetChildren() | Where-Object { $_.FQID.ServerId.ServerType -eq 'XPCORS' -and ($null -eq $RecordingServerId -or $_.FQID.ObjectId -in $RecordingServerId) }
        Write-Verbose "Retrieving video device statistics from $($recorders.Count) recording servers"
        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($recorder in $recorders) {
                Write-Verbose "Requesting device status from $($recorder.Name) at $($recorder.FQID.ServerId.Uri)"
                $folders = $recorder.GetChildren() | Where-Object { $_.FQID.Kind -in $includedDeviceTypes -and $_.FQID.FolderType -eq [videoos.platform.foldertype]::SystemDefined }
                $deviceIds = [guid[]]($folders | ForEach-Object {
                        $children = $_.GetChildren()
                        if ($null -ne $children -and $children.Count -gt 0) {
                            $children.FQID.ObjectId
                        }
                    })

                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                        Uri       = $recorder.FQID.ServerId.Uri
                        DeviceIds = $deviceIds
                    }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                        RecordingServerId   = $recorder.FQID.ObjectId
                        RecordingServerName = $recorder.Name
                        PowerShell          = $ps
                        Result              = $asyncResult
                    })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        Write-Verbose "Receiving results from recording server $($thread.RecordingServerName)"
                        if ($AsHashTable) {
                            $hashTable.$($thread.RecordingServerId.ToString()) = $null
                        } else {
                            $obj = @{
                                RecordingServerId   = $thread.RecordingServerId.ToString()
                                CurrentDeviceStatus = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.RecordingServerId.ToString()) = $result
                            } else {
                                $obj.CurrentDeviceStatus = $result
                            }
                        } catch {
                            throw
                        } finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | ForEach-Object { [void]$threads.Remove($_) }
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        } finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}


function Get-VideoDeviceStatistics {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param (
        # Specifies one or more Recording Server ID's to which the results will be limited. Omit this parameter if you want device status from all Recording Servers
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $RecordingServerId,

        # Specifies that the output should be provided in a complete hashtable instead of one pscustomobject value at a time
        [Parameter()]
        [switch]
        $AsHashTable,

        # Specifies the runspacepool to use. If no runspacepool is provided, one will be created.
        [Parameter()]
        [System.Management.Automation.Runspaces.RunspacePool]
        $RunspacePool
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $disposeRunspacePool = $true
        if ($PSBoundParameters.ContainsKey('RunspacePool')) {
            $disposeRunspacePool = $false
        }
        $pool = $RunspacePool
        if ($null -eq $pool) {
            Write-Verbose 'Creating a runspace pool'
            $pool = [runspacefactory]::CreateRunspacePool(1, ([int]$env:NUMBER_OF_PROCESSORS + 1))
            $pool.Open()
        }

        $scriptBlock = {
            param(
                [uri]$Uri,
                [guid[]]$DeviceIds,
                [string]$Token
            )
            try {
                $client = [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]::new($Uri)
                $client.GetVideoDeviceStatistics($Token, $deviceIds)
            } catch {
                throw
            }
        }

        Write-Verbose 'Retrieving recording server information'
        $managementServer = [videoos.platform.configuration]::Instance.GetItems([videoos.platform.itemhierarchy]::SystemDefined) | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Server -and $_.FQID.ObjectId -eq (Get-VmsManagementServer).Id }
        $recorders = $managementServer.GetChildren() | Where-Object { $_.FQID.ServerId.ServerType -eq 'XPCORS' -and ($null -eq $RecordingServerId -or $_.FQID.ObjectId -in $RecordingServerId) }
        Write-Verbose "Retrieving video device statistics from $($recorders.Count) recording servers"
        try {
            $threads = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($recorder in $recorders) {
                Write-Verbose "Requesting video device statistics from $($recorder.Name) at $($recorder.FQID.ServerId.Uri)"
                $folders = $recorder.GetChildren() | Where-Object { $_.FQID.Kind -eq [videoos.platform.kind]::Camera -and $_.FQID.FolderType -eq [videoos.platform.foldertype]::SystemDefined }
                $deviceIds = [guid[]]($folders | ForEach-Object {
                        $children = $_.GetChildren()
                        if ($null -ne $children -and $children.Count -gt 0) {
                            $children.FQID.ObjectId
                        }
                    })

                $ps = [powershell]::Create()
                $ps.RunspacePool = $pool
                $asyncResult = $ps.AddScript($scriptBlock).AddParameters(@{
                        Uri       = $recorder.FQID.ServerId.Uri
                        DeviceIds = $deviceIds
                        Token     = Get-VmsToken
                    }).BeginInvoke()
                $threads.Add([pscustomobject]@{
                        RecordingServerId   = $recorder.FQID.ObjectId
                        RecordingServerName = $recorder.Name
                        PowerShell          = $ps
                        Result              = $asyncResult
                    })
            }

            if ($threads.Count -eq 0) {
                return
            }

            $hashTable = @{}
            $completedThreads = New-Object System.Collections.Generic.List[pscustomobject]
            while ($threads.Count -gt 0) {
                foreach ($thread in $threads) {
                    if ($thread.Result.IsCompleted) {
                        Write-Verbose "Receiving results from recording server $($thread.RecordingServerName)"
                        if ($AsHashTable) {
                            $hashTable.$($thread.RecordingServerId.ToString()) = $null
                        } else {
                            $obj = @{
                                RecordingServerId     = $thread.RecordingServerId.ToString()
                                VideoDeviceStatistics = $null
                            }
                        }
                        try {
                            $result = $thread.PowerShell.EndInvoke($thread.Result) | ForEach-Object { Write-Output $_ }
                            if ($AsHashTable) {
                                $hashTable.$($thread.RecordingServerId.ToString()) = $result
                            } else {
                                $obj.VideoDeviceStatistics = $result
                            }
                        } catch {
                            $errorParams = @{
                                Message   = "An error occurred when calling GetVideoDeviceStatistics on recording server $($thread.RecordingServerName)"
                                Category  = 'ConnectionError'
                                ErrorId   = 'GetVideoDeviceStatisticsFailed'
                                Exception = $_.Exception
                            }
                            Write-Error @errorParams
                        } finally {
                            $thread.PowerShell.Dispose()
                            $completedThreads.Add($thread)
                            if (!$AsHashTable) {
                                Write-Output ([pscustomobject]$obj)
                            }
                        }
                    }
                }
                $completedThreads | ForEach-Object { [void]$threads.Remove($_) }
                $completedThreads.Clear()
                if ($threads.Count -eq 0) {
                    break
                }
                Start-Sleep -Milliseconds 250
            }
            if ($AsHashTable) {
                Write-Output $hashTable
            }
        } finally {
            if ($threads.Count -gt 0) {
                Write-Warning "Stopping $($threads.Count) running PowerShell instances. This may take a minute. . ."
                foreach ($thread in $threads) {
                    $thread.PowerShell.Dispose()
                }
            }
            if ($disposeRunspacePool) {
                Write-Verbose "Closing runspace pool in $($MyInvocation.MyCommand.Name)"
                $pool.Close()
                $pool.Dispose()
            }
        }
    }
}


function Get-VmsCameraReport {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        [Parameter()]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer[]]
        $RecordingServer,

        [Parameter()]
        [switch]
        $IncludePlainTextPasswords,

        [Parameter()]
        [switch]
        $IncludeRetentionInfo,

        [Parameter()]
        [switch]
        $IncludeRecordingStats,

        [Parameter()]
        [switch]
        $IncludeSnapshots,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $SnapshotTimeoutMS = 10000,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $SnapshotHeight = 300,

        [Parameter()]
        [ValidateSet('All', 'Disabled', 'Enabled')]
        [string]
        $EnableFilter = 'Enabled'
    )

    begin {
        Assert-VmsRequirementsMet -ErrorAction Stop
        try {
            $ms = Get-VmsManagementServer -ErrorAction Stop
            for ($attempt = 1; $attempt -le 2; $attempt++) {
                try {
                    $supportsFillChildren = [version]$ms.Version -ge '20.2'
                    $scs = Get-IServerCommandService -ErrorAction Stop
                    $config = $scs.GetConfiguration((Get-VmsToken))
                    $recorderCameraMap = @{}
                    $config.Recorders | ForEach-Object {
                        $deviceList = New-Object System.Collections.Generic.List[guid]
                        $_.Cameras.DeviceId | ForEach-Object { if ($_) { $deviceList.Add($_) } }
                        $recorderCameraMap.($_.RecorderId) = $deviceList
                    }
                    break
                } catch {
                    if ($attempt -ge 2) {
                        throw
                    }
                    # Typically if an error is thrown here, it's on $scs.GetConfiguration because the
                    # IServerCommandService WCF channel is cached and reused, and might be timed out.
                    # The Select-VmsSite cmdlet has a side effect of flushing all cached WCF channels.
                    Get-VmsSite | Select-VmsSite
                }
            }
            $roleMemberships = (Get-LoginSettings | Where-Object Guid -EQ (Get-VmsSite).FQID.ObjectId).GroupMembership
            $isAdmin = (Get-VmsRole -RoleType Adminstrative).Id -in $roleMemberships
            $dllFileInfo = [io.fileinfo](Get-Module MilestonePSTools)[0].Path
            $manifestPath = Join-Path $dllFileInfo.Directory.Parent.FullName 'MilestonePSTools.psd1'
            $jobRunner = [LocalJobRunner]::new($manifestPath)
            $jobRunner.JobPollingInterval = [timespan]::FromMilliseconds(500)
        } catch {
            throw
        }
    }

    process {
        try {
            if ($IncludePlainTextPasswords -and -not $isAdmin) {
                Write-Warning $script:Messages.MustBeAdminToReadPasswords
            }
            if (-not $RecordingServer) {
                Write-Verbose $script:Messages.ListingAllRecorders
                $RecordingServer = Get-VmsRecordingServer
            }
            $cache = @{
                DeviceState    = @{}
                PlaybackInfo   = @{}
                Snapshots      = @{}
                Passwords      = @{}
                RecordingStats = @{}
            }

            $ids = @()
            $RecordingServer | ForEach-Object {
                if ($null -ne $recorderCameraMap[[guid]$_.Id] -and $recorderCameraMap[[guid]$_.Id].Count -gt 0) {
                    $ids += $recorderCameraMap[[guid]$_.Id]
                }
            }

            Write-Verbose $script:Messages.CallingGetItemState
            Get-ItemState -CamerasOnly -ErrorAction Ignore | ForEach-Object {
                $cache.DeviceState[$_.FQID.ObjectId] = @{
                    ItemState = $_.State
                }
            }

            Write-Verbose $script:Messages.StartingFillChildrenThreadJob
            $fillChildrenJobs = $RecordingServer | ForEach-Object {
                $jobRunner.AddJob(
                    {
                        param([bool]$supportsFillChildren, [object]$recorder, [string]$EnableFilter, [bool]$getPasswords, [hashtable]$cache)

                        $manualMethod = {
                            param([object]$recorder)
                            $null = $recorder.HardwareDriverFolder.HardwareDrivers
                            $null = $recorder.StorageFolder.Storages.ArchiveStorageFolder.ArchiveStorages
                            $null = $recorder.HardwareFolder.Hardwares.HardwareDriverSettingsFolder.HardwareDriverSettings
                            $null = $recorder.HardwareFolder.Hardwares.CameraFolder.Cameras.StreamFolder.Streams
                            $null = $recorder.HardwareFolder.Hardwares.CameraFolder.Cameras.DeviceDriverSettingsFolder.DeviceDriverSettings
                        }
                        if ($supportsFillChildren) {
                            try {
                                $itemTypes = 'Hardware', 'HardwareDriverFolder', 'HardwareDriver', 'HardwareDriverSettingsFolder', 'HardwareDriverSettings', 'StorageFolder', 'Storage', 'StorageInformation', 'ArchiveStorageFolder', 'ArchiveStorage', 'CameraFolder', 'Camera', 'DeviceDriverSettingsFolder', 'DeviceDriverSettings', 'MotionDetectionFolder', 'MotionDetection', 'StreamFolder', 'Stream', 'StreamSettings', 'StreamDefinition', 'ClientSettings'
                                $alwaysIncludedItemTypes = @('MotionDetection', 'HardwareDriver', 'HardwareDriverSettings', 'Hardware', 'Storage', 'ArchiveStorage', 'DeviceDriverSettings')
                                $supportsPrivacyMask = (Get-IServerCommandService).GetConfiguration((Get-VmsToken)).ServerOptions | Where-Object Key -EQ 'PrivacyMask' | Select-Object -ExpandProperty Value
                                if ($supportsPrivacyMask -eq 'True') {
                                    $itemTypes += 'PrivacyProtectionFolder' , 'PrivacyProtection'
                                    $alwaysIncludedItemTypes += 'PrivacyProtectionFolder', 'PrivacyProtection'
                                }
                                $itemFilters = $itemTypes | ForEach-Object {
                                    $enableFilterSelection = if ($_ -in $alwaysIncludedItemTypes) { 'All' } else { $EnableFilter }
                                    [VideoOS.ConfigurationApi.ClientService.ItemFilter]@{
                                        ItemType        = $_
                                        EnableFilter    = $enableFilterSelection
                                        PropertyFilters = @()
                                    }
                                }
                                $recorder.FillChildren($itemTypes, $itemFilters)

                                # TODO: Remove this after TFS 447559 is addressed. The StreamFolder.Streams collection is empty after using FillChildren
                                # So this entire foreach block is only necessary to flush the children of StreamFolder and force another query for every
                                # camera so we can fill the collection up in this background task before enumerating over everything at the end.
                                foreach ($hw in $recorder.hardwarefolder.hardwares) {
                                    if ($getPasswords) {
                                        $password = $hw.ReadPasswordHardware().GetProperty('Password')
                                        $cache.Passwords[[guid]$hw.Id] = $password
                                    }
                                    foreach ($cam in $hw.camerafolder.cameras) {
                                        try {
                                            if ($null -ne $cam.StreamFolder -and $cam.StreamFolder.Streams.Count -eq 0) {
                                                $cam.StreamFolder.ClearChildrenCache()
                                                $null = $cam.StreamFolder.Streams
                                            }
                                        } catch {
                                            Write-Error $_
                                        }
                                    }
                                }
                            } catch {
                                Write-Error $_
                                $manualMethod.Invoke($recorder)
                            }
                        } else {
                            $manualMethod.Invoke($recorder)
                        }
                    },
                    @{ SupportsFillChildren = $supportsFillChildren; recorder = $_; EnableFilter = $EnableFilter; getPasswords = ($isAdmin -and $IncludePlainTextPasswords); cache = $cache }
                )
            }

            # Kick off snapshots early if requested. Pick up results at the end.
            $snapshotsById = @{}
            if ($IncludeSnapshots) {
                Write-Verbose 'Starting Get-Snapshot threadjob'
                $snapshotScriptBlock = {
                    param([guid[]]$ids, [int]$snapshotHeight, [hashtable]$snapshotsById, [hashtable]$cache, [int]$liveTimeoutMS)
                    foreach ($id in $ids) {
                        $itemState = $cache.DeviceState[$id].ItemState
                        if (-not [string]::IsNullOrWhiteSpace($itemState) -and $itemState -ne 'Responding') {
                            # Do not attempt to get a live image if the event server says the camera is not responding. Saves time.
                            continue
                        }
                        $snapshot = Get-Snapshot -CameraId $id -Live -Quality 100 -LiveTimeoutMS $liveTimeoutMS
                        if ($null -ne $snapshot) {
                            $image = $snapshot | ConvertFrom-Snapshot | Resize-Image -Height $snapshotHeight -DisposeSource
                            $snapshotsById[$id] = $image
                        }
                    }
                }
                $snapshotsJob = $jobRunner.AddJob($snapshotScriptBlock, @{ids = $ids; snapshotHeight = $SnapshotHeight; snapshotsById = $snapshotsById; cache = $cache; liveTimeoutMS = $SnapshotTimeoutMS })
            }

            if ($IncludeRetentionInfo) {
                Write-Verbose 'Starting Get-PlaybackInfo threadjob'
                $playbackInfoScriptblock = {
                    param(
                        [guid]$id,
                        [hashtable]$cache
                    )

                    $info = Get-PlaybackInfo -Path "Camera[$id]"
                    if ($null -ne $info) {
                        $cache.PlaybackInfo[$id] = $info
                    }
                }
                $playbackInfoJobs = $ids | ForEach-Object {
                    if ($null -ne $_) {
                        $jobRunner.AddJob($playbackInfoScriptblock, @{ id = $_; cache = $cache } )
                    }
                }
            }

            if ($IncludeRecordingStats) {
                Write-Verbose 'Starting recording stats threadjob'
                $recordingStatsScript = {
                    param(
                        [guid]$Id,
                        [datetime]$StartTime,
                        [datetime]$EndTime,
                        [string]$SequenceType
                    )

                    $sequences = Get-SequenceData -Path "Camera[$Id]" -SequenceType $SequenceType -StartTime $StartTime -EndTime $EndTime -CropToTimeSpan
                    $recordedMinutes = $sequences | ForEach-Object {
                        ($_.EventSequence.EndDateTime - $_.EventSequence.StartDateTime).TotalMinutes
                    } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
                    [pscustomobject]@{
                        DeviceId        = $Id
                        StartTime       = $StartTime
                        EndTime         = $EndTime
                        SequenceCount   = $sequences.Count
                        TimeRecorded    = [timespan]::FromMinutes($recordedMinutes)
                        PercentRecorded = [math]::Round(($recordedMinutes / ($EndTime - $StartTime).TotalMinutes * 100), 1)
                    }
                }
                $endTime = Get-Date
                $startTime = $endTime.AddDays(-7)
                $recordingStatsJobs = $ids | ForEach-Object {
                    $jobRunner.AddJob($recordingStatsScript, @{Id = $_; StartTime = $startTime; EndTime = $endTime; SequenceType = 'RecordingSequence' })
                }
            }

            # Get VideoDeviceStatistics for all Recording Servers in the report
            Write-Verbose 'Starting GetVideoDeviceStatistics threadjob'
            $videoDeviceStatsScriptBlock = {
                param(
                    [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]$svc,
                    [guid[]]$ids
                )
                $svc.GetVideoDeviceStatistics((Get-VmsToken), $ids)
            }
            $videoDeviceStatsJobs = $RecordingServer | ForEach-Object {
                $svc = $_ | Get-RecorderStatusService2
                if ($null -ne $svc) {
                    $jobRunner.AddJob($videoDeviceStatsScriptBlock, @{ svc = $svc; ids = $recorderCameraMap[[guid]$_.Id] })
                }
            }

            # Get Current Device Status for everything in the report
            Write-Verbose 'Starting GetCurrentDeviceStatus threadjob'
            $currentDeviceStatsJobsScriptBlock = {
                param(
                    [VideoOS.Platform.SDK.Proxy.Status2.RecorderStatusService2]$svc,
                    [guid[]]$ids
                )
                $svc.GetCurrentDeviceStatus((Get-VmsToken), $ids)
            }
            $currentDeviceStatsJobs = $RecordingServer | Where-Object { ($recorderCameraMap[[guid]$_.Id]).Count } | ForEach-Object {
                $svc = $_ | Get-RecorderStatusService2
                $jobRunner.AddJob($currentDeviceStatsJobsScriptBlock, @{svc = $svc; ids = $recorderCameraMap[[guid]$_.Id] })
            }

            Write-Verbose 'Receiving results of FillChildren threadjob'
            $jobRunner.Wait($fillChildrenJobs)
            $fillChildrenResults = $jobRunner.ReceiveJobs($fillChildrenJobs)
            foreach ($e in $fillChildrenResults.Errors) {
                Write-Error $e
            }

            if ($IncludeRetentionInfo) {
                Write-Verbose 'Receiving results of Get-PlaybackInfo threadjob'
                $jobRunner.Wait($playbackInfoJobs)
                $playbackInfoResult = $jobRunner.ReceiveJobs($playbackInfoJobs)
                foreach ($e in $playbackInfoResult.Errors) {
                    Write-Error $e
                }
            }

            if ($IncludeRecordingStats) {
                Write-Verbose 'Receiving results of recording stats threadjob'
                $jobRunner.Wait($recordingStatsJobs)
                foreach ($job in $jobRunner.ReceiveJobs($recordingStatsJobs)) {
                    if ($job.Output.DeviceId) {
                        $cache.RecordingStats[$job.Output.DeviceId] = $job.Output
                    }
                    foreach ($e in $job.Errors) {
                        Write-Error $e
                    }
                }
            }

            Write-Verbose 'Receiving results of GetVideoDeviceStatistics threadjobs'
            $jobRunner.Wait($videoDeviceStatsJobs)
            foreach ($job in $jobRunner.ReceiveJobs($videoDeviceStatsJobs)) {
                foreach ($result in $job.Output) {
                    if (-not $cache.DeviceState.ContainsKey($result.DeviceId)) {
                        $cache.DeviceState[$result.DeviceId] = @{}
                    }
                    $cache.DeviceState[$result.DeviceId].UsedSpaceInBytes = $result.UsedSpaceInBytes
                    $cache.DeviceState[$result.DeviceId].VideoStreamStatisticsArray = $result.VideoStreamStatisticsArray
                }
                foreach ($e in $job.Errors) {
                    Write-Error $e
                }
            }

            Write-Verbose 'Receiving results of GetCurrentDeviceStatus threadjobs'
            $jobRunner.Wait($currentDeviceStatsJobs)
            $currentDeviceStatsResult = $jobRunner.ReceiveJobs($currentDeviceStatsJobs)
            $currentDeviceStatsResult.Output | ForEach-Object {
                foreach ($row in $_.CameraDeviceStatusArray) {
                    if (-not $cache.DeviceState.ContainsKey($row.DeviceId)) {
                        $cache.DeviceState[$row.DeviceId] = @{}
                    }
                    $cache.DeviceState[$row.DeviceId].Status = $row
                }
            }
            foreach ($e in $currentDeviceStatsResult.Errors) {
                Write-Error $e
            }

            if ($null -ne $snapshotsJob) {
                Write-Verbose 'Receiving results of Get-Snapshot threadjob'
                $jobRunner.Wait($snapshotsJob)
                $snapshotsResult = $jobRunner.ReceiveJobs($snapshotsJob)
                $cache.Snapshots = $snapshotsById
                foreach ($e in $snapshotsResult.Errors) {
                    Write-Error $e
                }
            }

            foreach ($rec in $RecordingServer) {
                foreach ($hw in $rec.HardwareFolder.Hardwares | Where-Object { if ($EnableFilter -eq 'All') { $true } else { $_.Enabled } }) {
                    try {
                        $hwSettings = ConvertFrom-ConfigurationApiProperties -Properties $hw.HardwareDriverSettingsFolder.HardwareDriverSettings[0].HardwareDriverSettingsChildItems[0].Properties -UseDisplayNames
                        $driver = $rec.HardwareDriverFolder.HardwareDrivers | Where-Object Path -EQ $hw.HardwareDriverPath
                        foreach ($cam in $hw.CameraFolder.Cameras | Where-Object { if ($EnableFilter -eq 'All') { $true } elseif ($EnableFilter -eq 'Enabled') { $_.Enabled -and $hw.Enabled } else { !$_.Enabled -or !$hw.Enabled } }) {
                            $id = [guid]$cam.Id
                            $state = $cache.DeviceState[$id]
                            $storage = $rec.StorageFolder.Storages | Where-Object Path -EQ $cam.RecordingStorage
                            $motion = $cam.MotionDetectionFolder.MotionDetections[0]
                            if ($cam.StreamFolder.Streams.Count -gt 0) {
                                $liveStreamSettings = $cam | Get-VmsCameraStream -LiveDefault -ErrorAction Ignore
                                $liveStreamStats = $state.VideoStreamStatisticsArray | Where-Object StreamId -EQ $liveStreamSettings.StreamReferenceId
                                $recordedStreamSettings = $cam | Get-VmsCameraStream -Recorded -ErrorAction Ignore
                                $recordedStreamStats = $state.VideoStreamStatisticsArray | Where-Object StreamId -EQ $recordedStreamSettings.StreamReferenceId
                            } else {
                                Write-Warning "Live & recorded stream properties unavailable for $($cam.Name) as the camera does not support multi-streaming."
                            }
                            $obj = [ordered]@{
                                Name                         = $cam.Name
                                Channel                      = $cam.Channel
                                Enabled                      = $cam.Enabled -and $hw.Enabled
                                ShortName                    = $cam.ShortName
                                Shortcut                     = $cam.ClientSettingsFolder.ClientSettings.Shortcut
                                State                        = $state.ItemState
                                LastModified                 = $cam.LastModified
                                Id                           = $cam.Id
                                IsStarted                    = $state.Status.Started
                                IsMotionDetected             = $state.Status.Motion
                                IsRecording                  = $state.Status.Recording
                                IsInOverflow                 = $state.Status.ErrorOverflow
                                IsInDbRepair                 = $state.Status.DbRepairInProgress
                                ErrorWritingGOP              = $state.Status.ErrorWritingGop
                                ErrorNotLicensed             = $state.Status.ErrorNotLicensed
                                ErrorNoConnection            = $state.Status.ErrorNoConnection
                                StatusTime                   = $state.Status.Time
                                GpsCoordinates               = $cam.GisPoint | ConvertFrom-GisPoint

                                HardwareName                 = $hw.Name
                                HardwareId                   = $hw.Id
                                Model                        = $hw.Model
                                Address                      = $hw.Address
                                Username                     = $hw.UserName
                                Password                     = if ($cache.Passwords.ContainsKey([guid]$hw.Id)) { $cache.Passwords[[guid]$hw.Id] } else { 'NotIncluded' }
                                HTTPSEnabled                 = $hwSettings.HTTPSEnabled -eq 'yes'
                                MAC                          = $hwSettings.MacAddress
                                Firmware                     = $hwSettings.FirmwareVersion

                                DriverFamily                 = $driver.GroupName
                                Driver                       = $driver.Name
                                DriverNumber                 = $driver.Number
                                DriverVersion                = $driver.DriverVersion
                                DriverRevision               = $driver.DriverRevision

                                RecorderName                 = $rec.Name
                                RecorderUri                  = $rec.ActiveWebServerUri, $rec.WebServerUri | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                RecorderId                   = $rec.Id

                                LiveStream                   = $liveStreamSettings.Name
                                LiveStreamDescription        = $liveStreamSettings.DisplayName
                                LiveStreamMode               = $liveStreamSettings.LiveMode
                                ConfiguredLiveResolution     = $liveStreamSettings.Settings.Resolution, $liveStreamSettings.Settings.StreamProperty | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                ConfiguredLiveCodec          = $liveStreamSettings.Settings.Codec
                                ConfiguredLiveFPS            = $liveStreamSettings.Settings.FPS, $liveStreamSettings.Settings.FrameRate | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                CurrentLiveResolution        = if ($null -eq $liveStreamStats) { 'Unavailable' } else { '{0}x{1}' -f $liveStreamStats.ImageResolution.Width, $liveStreamStats.ImageResolution.Height }
                                CurrentLiveCodec             = if ($null -eq $liveStreamStats) { 'Unavailable' } else { $liveStreamStats.VideoFormat }
                                CurrentLiveFPS               = if ($null -eq $liveStreamStats) { 'Unavailable' } else { $liveStreamStats.FPS -as [int] }
                                CurrentLiveBitrate           = if ($null -eq $liveStreamStats) { 'Unavailable' } else { (($liveStreamStats.BPS -as [int]) / 1MB).ToString('N1') }

                                RecordedStream               = $recordedStreamSettings.Name
                                RecordedStreamDescription    = $recordedStreamSettings.DisplayName
                                RecordedStreamMode           = $recordedStreamSettings.LiveMode
                                ConfiguredRecordedResolution = $recordedStreamSettings.Settings.Resolution, $recordedStreamSettings.Settings.StreamProperty | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                ConfiguredRecordedCodec      = $recordedStreamSettings.Settings.Codec
                                ConfiguredRecordedFPS        = $recordedStreamSettings.Settings.FPS, $recordedStreamSettings.Settings.FrameRate | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
                                CurrentRecordedResolution    = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { '{0}x{1}' -f $recordedStreamStats.ImageResolution.Width, $recordedStreamStats.ImageResolution.Height }
                                CurrentRecordedCodec         = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { $recordedStreamStats.VideoFormat }
                                CurrentRecordedFPS           = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { $recordedStreamStats.FPS -as [int] }
                                CurrentRecordedBitrate       = if ($null -eq $recordedStreamStats) { 'Unavailable' } else { (($recordedStreamStats.BPS -as [int]) / 1MB).ToString('N1') }

                                RecordingEnabled             = $cam.RecordingEnabled
                                RecordKeyframesOnly          = $cam.RecordKeyframesOnly
                                RecordOnRelatedDevices       = $cam.RecordOnRelatedDevices
                                PrebufferEnabled             = $cam.PrebufferEnabled
                                PrebufferSeconds             = $cam.PrebufferSeconds
                                PrebufferInMemory            = $cam.PrebufferInMemory

                                RecordingStorageName         = $storage.Name
                                RecordingPath                = [io.path]::Combine($storage.DiskPath, $storage.Id)
                                ExpectedRetentionDays        = ($storage | Get-VmsStorageRetention).TotalDays
                                PercentRecordedOneWeek       = if ($IncludeRecordingStats) { $cache.RecordingStats[$id].PercentRecorded -as [double] } else { 'NotIncluded' }

                                MediaDatabaseBegin           = if ($null -eq $cache.PlaybackInfo[$id].Begin) { if ($IncludeRetentionInfo) { 'Unavailable' } else { 'NotIncluded' } } else { $cache.PlaybackInfo[$id].Begin }
                                MediaDatabaseEnd             = if ($null -eq $cache.PlaybackInfo[$id].End) { if ($IncludeRetentionInfo) { 'Unavailable' } else { 'NotIncluded' } } else { $cache.PlaybackInfo[$id].End }
                                UsedSpaceInGB                = if ($null -eq $state.UsedSpaceInBytes) { 'Unavailable' } else { ($state.UsedSpaceInBytes / 1GB).ToString('N2') }

                            }
                            if ($IncludeRetentionInfo) {
                                $obj.ActualRetentionDays  = ($cache.PlaybackInfo[$id].End - $cache.PlaybackInfo[$id].Begin).TotalDays
                                $obj.MeetsRetentionPolicy = $obj.ActualRetentionDays -gt $obj.ExpectedRetentionDays
                                $obj.MediaDatabaseBegin   = $cache.PlaybackInfo[$id].Begin
                                $obj.MediaDatabaseEnd     = $cache.PlaybackInfo[$id].End
                            }

                            $obj.MotionEnabled = $motion.Enabled
                            $obj.MotionKeyframesOnly = $motion.KeyframesOnly
                            $obj.MotionProcessTime = $motion.ProcessTime
                            $obj.MotionManualSensitivityEnabled = $motion.ManualSensitivityEnabled
                            $obj.MotionManualSensitivity = [int]($motion.ManualSensitivity / 3)
                            $obj.MotionThreshold = $motion.Threshold
                            $obj.MotionMetadataEnabled = $motion.GenerateMotionMetadata
                            $obj.MotionExcludeRegions = $motion.UseExcludeRegions
                            $obj.MotionHardwareAccelerationMode = $motion.HardwareAccelerationMode

                            $obj.PrivacyMaskEnabled = ($cam.PrivacyProtectionFolder.PrivacyProtections | Select-Object -First 1).Enabled -eq $true

                            if ($IncludeSnapshots) {
                                $obj.Snapshot = $cache.Snapshots[$id]
                            }
                            Write-Output ([pscustomobject]$obj)
                        }
                    } catch {
                        Write-Error $_
                    }
                }
            }
        } finally {
            if ($jobRunner) {
                $jobRunner.Dispose()
            }
        }
    }
}


function Get-VmsRestrictedMedia {
    [CmdletBinding()]
    [Alias('Get-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMedia])]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMediaLive])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Media can also be singular.')]
    param (
        [Parameter()]
        [switch]
        $Live
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($Live) {
            { (Get-IServerCommandService).RestrictedMediaLiveQueryAll() } | ExecuteWithRetry -ClearVmsCache
        } else {
            { (Get-IServerCommandService).RestrictedMediaQueryAll() } | ExecuteWithRetry -ClearVmsCache
        }
    }
}



function New-VmsRestrictedMedia {
    [CmdletBinding()]
    [Alias('New-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMedia])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Media can also be singular.')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $DeviceId,

        [Parameter(Mandatory)]
        [datetime]
        $StartTime,

        [Parameter(Mandatory)]
        [datetime]
        $EndTime,

        [Parameter(Mandatory)]
        [string]
        $Header,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $IgnoreRelatedDevices
    )
    
    begin {
        Assert-VmsRequirementsMet
        $deviceIds = [collections.generic.list[guid]]::new()
    }
    
    process {
        if ($EndTime -gt (Get-Date)) {
            $EndTime = (Get-Date).AddMinutes(-1)
            Write-Warning "EndTime cannot be in the future and will be updated to $($EndTime.ToString('o'))"
        }

        foreach ($id in $DeviceId) {
            $deviceIds.Add($id)
            if ($IgnoreRelatedDevices) {
                continue
            }

            if ($null -eq ($item = Find-VmsVideoOSItem -SearchText $id.ToString().ToLower())) {
                continue
            }

            foreach ($relatedItem in $item.GetRelated()) {
                $deviceIds.Add($relatedItem.FQID.ObjectId)
            }
        }
    }
    
    end {
        $result = { (Get-IServerCommandService).RestrictedMediaCreate(
            (New-Guid),
            $deviceIds,
            $Header,
            $Description,
            $StartTime.ToUniversalTime(),
            $EndTime.ToUniversalTime()
        ) } | ExecuteWithRetry -ClearVmsCache
        foreach ($fault in $result.FaultDevices) {
            Write-Error -Message "$($fault.Message) DeviceId = '$($fault.DeviceId)'." -ErrorId 'RestrictedMediaLive.Fault' -Category InvalidResult
        }
        foreach ($warning in $result.WarningDevices) {
            Write-Warning -Message "$($warning.Message) DeviceId = '$($warning.DeviceId)'."
        }
        $result.RestrictedMedia
    }
}


function Remove-VmsRestrictedMedia {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [Alias('Remove-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Media can also be singular.')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'RestrictedMedia')]
        [VideoOS.Common.Proxy.Server.WCF.RestrictedMedia]
        $RestrictedMedia,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'RestrictedMediaLive')]
        [VideoOS.Common.Proxy.Server.WCF.RestrictedMediaLive]
        $RestrictedMediaLive,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'DeviceId')]
        [Alias('Id')]
        [guid[]]
        $DeviceId
    )

    begin {
        Assert-VmsRequirementsMet
        $ids = [collections.generic.list[guid]]::new()
        $idToNameMap = @{}
    }
    
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'RestrictedMedia' {
                $ids.Add($RestrictedMedia.Id)
                $idToNameMap[$RestrictedMedia.Id] = $RestrictedMedia.Header
                break
            }

            'DeviceId' {
                $ids.AddRange($DeviceId)
                break
            }

            'RestrictedMediaLive' {
                $ids.Add($RestrictedMediaLive.DeviceId)
                break
            }
        }
    }

    end {
        $results = [collections.generic.list[object]]::new()
        switch ($PSCmdlet.ParameterSetName) {
            'RestrictedMedia' {
                foreach ($id in $ids) {
                    if ($PSCmdlet.ShouldProcess($idToNameMap[$id], "Remove recorded media restriction")) {
                        $results.Add(({ (Get-IServerCommandService).RestrictedMediaDelete($id) } | ExecuteWithRetry -ClearVmsCache))
                    }
                }
            }

            { $_ -in @('RestrictedMediaLive', 'DeviceId') } {
                if ($PSCmdlet.ShouldProcess("$($ids.Count) devices", "Remove live media restriction")) {
                    $results.Add(({ (Get-IServerCommandService).RestrictedMediaLiveDelete($ids) } | ExecuteWithRetry -ClearVmsCache))
                }
            }
        }
        foreach ($result in $results) {
            Write-Verbose "Removed $($result.RestrictedMedia.Count) $($result.GetType().Name) records."
            foreach ($fault in $result.FaultDevices) {
                Write-Error -Message "$($fault.Message) DeviceId = '$($fault.DeviceId)'." -ErrorId 'RestrictedMediaLive.Fault' -Category InvalidResult
            }
            foreach ($warning in $result.WarningDevices) {
                Write-Warning -Message "$($warning.Message) DeviceId = '$($warning.DeviceId)'."
            }
        }
    }
}


function Set-VmsRestrictedMedia {
    [CmdletBinding()]
    [Alias('Set-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMedia])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Media can also be singular.')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Common.Proxy.Server.WCF.RestrictedMedia]
        $InputObject,

        [Parameter()]
        [guid[]]
        $IncludeDeviceId = [guid[]]::new(0),

        [Parameter()]
        [guid[]]
        $ExcludeDeviceId = [guid[]]::new(0),

        [Parameter()]
        [string]
        $Header,
        
        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [datetime]
        $StartTime,

        [Parameter()]
        [datetime]
        $EndTime,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        if (!$PSBoundParameters.ContainsKey('Header')) {
            $Header = $InputObject.Header
        }
        if (!$PSBoundParameters.ContainsKey('Description')) {
            $Description = $InputObject.Description
        }
        if (!$PSBoundParameters.ContainsKey('StartTime')) {
            $StartTime = $InputObject.StartTime
        }
        if (!$PSBoundParameters.ContainsKey('EndTime')) {
            $EndTime = $InputObject.EndTime
        }
        $result = { (Get-IServerCommandService).RestrictedMediaUpdate(
            $InputObject.Id,
            $IncludeDeviceId,
            $ExcludeDeviceId,
            $Header,
            $Description,
            $StartTime,
            $EndTime
        ) } | ExecuteWithRetry -ClearVmsCache
        foreach ($fault in $result.FaultDevices) {
            Write-Error -Message "$($fault.Message) DeviceId = '$($fault.DeviceId)'." -ErrorId 'RestrictedMedia.Fault' -Category InvalidResult
        }
        foreach ($warning in $result.WarningDevices) {
            Write-Warning -Message "$($warning.Message) DeviceId = '$($warning.DeviceId)'."
        }
        if ($PassThru -and $result.RestrictedMedia) {
            $result.RestrictedMedia
        }
    }
}


function Start-VmsRestrictedLiveMode {
    [CmdletBinding()]
    [Alias('Start-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMediaLive])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $DeviceId,

        [Parameter()]
        [datetime]
        $StartTime = (Get-Date).AddMinutes(-1),

        [Parameter()]
        [switch]
        $IgnoreRelatedDevices
    )
    
    begin {
        Assert-VmsRequirementsMet
        $deviceIds = [collections.generic.list[guid]]::new()
    }

    process {
        foreach ($id in $DeviceId) {
            $deviceIds.Add($id)

            if ($IgnoreRelatedDevices) {
                continue
            }

            if ($null -eq ($item = Find-VmsVideoOSItem -SearchText $id.ToString().ToLower())) {
                continue
            }

            foreach ($relatedItem in $item.GetRelated()) {
                $deviceIds.Add($relatedItem.FQID.ObjectId)
            }
        }
    }

    end {
        $result = {
            (Get-IServerCommandService).RestrictedMediaLiveModeEnter(
                $deviceIds,
                $StartTime.ToUniversalTime()
            )
        } | ExecuteWithRetry -ClearVmsCache
        foreach ($fault in $result.FaultDevices) {
            Write-Error -Message "$($fault.Message) DeviceId = '$($fault.DeviceId)'." -ErrorId 'RestrictedMediaLive.Fault' -Category InvalidResult
        }
        foreach ($warning in $result.WarningDevices) {
            Write-Warning -Message "$($warning.Message) DeviceId = '$($warning.DeviceId)'."
        }
        foreach ($restrictedMedia in $result.RestrictedMedia) {
            $restrictedMedia
        }
    }
}


function Stop-VmsRestrictedLiveMode {
    [CmdletBinding()]
    [Alias('Stop-VmsRm')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.2')]
    [RequiresVmsFeature('RestrictedMedia')]
    [OutputType([VideoOS.Common.Proxy.Server.WCF.RestrictedMedia])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Id')]
        [guid[]]
        $DeviceId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [datetime]
        $StartTime,

        [Parameter()]
        [datetime]
        $EndTime = (Get-Date).AddMinutes(-1),

        [Parameter(Mandatory)]
        [string]
        $Header,

        [Parameter()]
        [string]
        $Description
    )

    begin {
        Assert-VmsRequirementsMet
        $deviceIds = [collections.generic.list[guid]]::new()
        $startTimeValue = [datetime]::MinValue
    }
    
    process {
        $deviceIds.AddRange($DeviceId)
        $startTimeValue = $StartTime
    }

    end {
        $result = { (Get-IServerCommandService).RestrictedMediaLiveModeExit(
            (New-Guid),
            $deviceIds,
            $Header,
            $Description,
            $startTimeValue.ToUniversalTime(),
            $EndTime.ToUniversalTime()
        ) } | ExecuteWithRetry -ClearVmsCache
        foreach ($fault in $result.FaultDevices) {
            Write-Error -Message "$($fault.Message) DeviceId = '$($fault.DeviceId)'." -ErrorId 'RestrictedMediaLive.Fault' -Category InvalidResult
        }
        foreach ($warning in $result.WarningDevices) {
            Write-Warning -Message "$($warning.Message) DeviceId = '$($warning.DeviceId)'."
        }
        foreach ($restrictedMedia in $result.RestrictedMedia) {
            $restrictedMedia
        }
    }
}


function Add-VmsRoleClaim {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('RoleName')]
        [ValidateNotNull()]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 1)]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [VideoOS.Platform.ConfigurationItems.LoginProvider]
        $LoginProvider,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 2)]
        [string]
        $ClaimName,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 3)]
        [string]
        $ClaimValue
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($r in $Role) {
            if ($PSCmdlet.ShouldProcess("$($Role.Name)", "Add claim '$ClaimName' with value '$ClaimValue'")) {
                $null = $r.ClaimFolder.AddRoleClaim($LoginProvider.Id, $ClaimName, $ClaimValue)
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Add-VmsRoleClaim -ParameterName ClaimName -ScriptBlock {
    $values = (Get-VmsLoginProvider | Get-VmsLoginProviderClaim).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Add-VmsRoleMember {
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'ByAccountName')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'ByAccountName')]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'BySid')]
        [Alias('RoleName')]
        [ValidateNotNull()]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 1, ParameterSetName = 'ByAccountName')]
        [string[]]
        $AccountName,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 2, ParameterSetName = 'BySid')]
        [string[]]
        $Sid
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'ByAccountName') {
            $Sid = $AccountName | ConvertTo-Sid
        }
        foreach ($r in $Role) {
            foreach ($s in $Sid) {
                try {
                    if ($PSCmdlet.ShouldProcess($Role.Name, "Add member with SID $s to role")) {
                        $null = $r.UserFolder.AddRoleMember($s)
                    }
                }
                catch {
                    Write-Error -ErrorRecord $_
                }
            }
        }
    }
}

function Copy-VmsRole {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.Role])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role]
        $Role,

        [Parameter(Mandatory, Position = 0)]
        [string]
        $NewName
    )

    begin {
        Assert-VmsRequirementsMet
        if (Get-VmsRole -Name $NewName -ErrorAction SilentlyContinue) {
            throw "Role with name '$NewName' already exists."
            return
        }
    }

    process {
        $roleDefinition = $Role | Export-VmsRole -PassThru
        $roleDefinition.Name = $NewName
        $roleDefinition | Import-VmsRole
    }
}

function Export-VmsRole {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline)]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter()]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
        if ($MyInvocation.BoundParameters.ContainsKey('Path')) {
            $resolvedPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue -ErrorVariable rpError).Path
            if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
                $resolvedPath = $rpError.TargetObject
            }
            $Path = $resolvedPath
            $fileInfo = [io.fileinfo]$Path
            if (-not $fileInfo.Directory.Exists) {
                throw ([io.directorynotfoundexception]::new("Directory not found: $($fileInfo.Directory.FullName)"))
            }
            if (($fi = [io.fileinfo]$Path).Extension -ne '.json') {
                Write-Verbose "A .json file extension will be added to the file '$($fi.Name)'"
                $Path += ".json"
            }
        } elseif (-not $MyInvocation.BoundParameters.ContainsKey('PassThru') -or -not $PassThru.ToBool()) {
            throw "Either or both of Path, or PassThru parameters must be specified."
        }

        $roles = [system.collections.generic.list[pscustomobject]]::new()

        $providers = @{}
        $supportsOidc = [version](Get-VmsManagementServer).Version -ge '22.1'
        if ($supportsOidc) {
            Get-VmsLoginProvider | Foreach-Object {
                $providers[$_.Id] = $_
            }
        }

        $clientProfiles = @{}
        (Get-VmsManagementServer).ClientProfileFolder.ClientProfiles | ForEach-Object {
            if ($null -eq $_) { return }
            $clientProfiles[$_.Path] = $_
        }

        $timeProfiles = @{
            'TimeProfile[11111111-1111-1111-1111-111111111111]' = [pscustomobject]@{
                Name        = 'Always'
                DisplayName = 'Always'
                Path        = 'TimeProfile[11111111-1111-1111-1111-111111111111]'
            }
            'TimeProfile[00000000-0000-0000-0000-000000000000]' = [pscustomobject]@{
                Name        = 'Default'
                DisplayName = 'Default'
                Path        = 'TimeProfile[00000000-0000-0000-0000-000000000000]'
            }
        }
        (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | ForEach-Object {
            if ($null -eq $_) { return }
            $timeProfiles[$_.Path] = $_
        }
    }

    process {
        if ($Role.Count -eq 0) {
            $Role = Get-VmsRole
        }

        foreach ($r in $Role) {
            $item = $r | Get-ConfigurationItem
            $clientProfile = $item | Get-ConfigurationItemProperty -Key ClientProfile -ErrorAction SilentlyContinue
            if ($clientProfile -and $clientProfiles.ContainsKey($clientProfile)) {
                $clientProfile = $clientProfiles[$clientProfile].Name
            }
            $defaultTimeProfile = $item | Get-ConfigurationItemProperty -Key RoleDefaultTimeProfile -ErrorAction SilentlyContinue
            if ($defaultTimeProfile -and $timeProfiles.ContainsKey($defaultTimeProfile)) {
                $defaultTimeProfile = $timeProfiles[$defaultTimeProfile].Name
            }
            $logonTimeProfile = $item | Get-ConfigurationItemProperty -Key RoleClientLogOnTimeProfile -ErrorAction SilentlyContinue
            if ($logonTimeProfile -and $timeProfiles.ContainsKey($logonTimeProfile)) {
                $logonTimeProfile = $timeProfiles[$logonTimeProfile].Name
            }
            $roleDto = [pscustomobject]@{
                Name                               = $r.Name
                Description                        = $r.Description
                AllowMobileClientLogOn             = $r.AllowMobileClientLogOn
                AllowSmartClientLogOn              = $r.AllowSmartClientLogOn
                AllowWebClientLogOn                = $r.AllowWebClientLogOn
                DualAuthorizationRequired          = $r.DualAuthorizationRequired
                MakeUsersAnonymousDuringPTZSession = $r.MakeUsersAnonymousDuringPTZSession
                ClientProfile                      = $clientProfile
                DefaultTimeProfile                 = $defaultTimeProfile
                ClientLogOnTimeProfile             = $logonTimeProfile
                Claims                             = [system.collections.generic.list[pscustomobject]]::new()
                Users                              = [system.collections.generic.list[pscustomobject]]::new()
                OverallSecurity                    = [system.collections.generic.list[pscustomobject]]::new()
            }
            $r.UserFolder.Users | Foreach-Object {
                $roleDto.Users.Add([pscustomobject]@{
                        Sid          = $_.Sid
                        IdentityType = $_.IdentityType
                        DisplayName  = $_.DisplayName
                        AccountName  = $_.AccountName
                        Domain       = $_.Domain
                    })
            }
            if ($supportsOidc) {
                $r | Get-VmsRoleClaim | ForEach-Object {
                    $roleDto.Claims.Add([pscustomobject]@{
                            LoginProvider = $providers[$_.ClaimProvider].Name
                            ClaimName     = $_.ClaimName
                            ClaimValue    = $_.ClaimValue
                        })
                }
            }
            
            if ($r.RoleType -eq 'UserDefined') {
                $r | Get-VmsRoleOverallSecurity | Sort-Object DisplayName | ForEach-Object {
                    $obj = [ordered]@{
                        DisplayName       = $_.DisplayName
                        SecurityNamespace = $_.SecurityNamespace
                    }
                    foreach ($key in $_.Keys | Where-Object { $_ -notin 'DisplayName', 'SecurityNamespace', 'Role' } | Sort-Object) {
                        $obj[$key] = $_[$key]
                    }
                    $roleDto.OverallSecurity.Add($obj)
                }
            }

            $roles.Add($roleDto)
            if ($PassThru) {
                $roleDto
            }
        }
    }

    end {
        if ($roles.Count -gt 0 -and $Path) {
            [io.file]::WriteAllText($Path, (ConvertTo-Json -InputObject $roles -Depth 10 -Compress), [system.text.encoding]::UTF8)
        }
    }
}

function Get-VmsRole {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    [RequiresVmsConnection()]
    [OutputType([VideoOS.Platform.ConfigurationItems.Role])]
    param (
        [Parameter(Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'ByName')]
        [ArgumentCompleter([MilestonePSTools.Utility.MipItemNameCompleter[VideoOS.Platform.ConfigurationItems.Role]])]
        [string]
        $Name = '*',

        [Parameter(ParameterSetName = 'ByName')]
        [string]
        $RoleType = '*',

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'ById')]
        [Alias('RoleId')]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            try {
                [VideoOS.Platform.ConfigurationItems.Role]::new((Get-VmsManagementServer).ServerId, "Role[$Id]")
            } catch [VideoOS.Platform.PathNotFoundMIPException] {
                Write-Error -Message "No item found with ID matching $Id" -Exception $_.Exception
            }
        } else {
            $matchFound = $false
            foreach ($role in (Get-VmsManagementServer).RoleFolder.Roles) {
                if ($role.Name -notlike $Name -or $role.RoleType -notlike $RoleType) {
                    continue
                }
                if ($null -eq $role.ClientProfile) {
                    # TODO: Added because the ClientProfile, RoleDefaultTimeProfile, and RoleClientLogOnTimeProfile are $null
                    # when enumerating a role from the RoleFolder.Roles collection. If it's not null, then the MIP SDK
                    # behavior will have improved and we can avoid extra API calls by returning cached values.
                    [VideoOS.Platform.ConfigurationItems.Role]::new($role.ServerId, $role.Path)
                } else {
                    $role
                }
                $matchFound = $true
            }
            if (-not $matchFound -and -not [management.automation.wildcardpattern]::ContainsWildcardCharacters($Name)) {
                Write-Error "Role '$Name' not found."
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsRole -ParameterName Id -ScriptBlock {
    $values = (Get-VmsManagementServer).RoleFolder.Roles.Id
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

Register-ArgumentCompleter -CommandName Get-VmsRole -ParameterName RoleType -ScriptBlock {
    $values = (Get-VmsManagementServer).RoleFolder.Roles[0].RoleTypeValues.Values | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Get-VmsRoleClaim {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.ClaimChildItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [Alias('RoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(Position = 1)]
        [string[]]
        $ClaimName,

        [Parameter()]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        if ($null -eq $Role) {
            $Role = Get-VmsRole
        }
        foreach ($r in $Role) {
            $matchFound = $false
            foreach ($claim in $r.ClaimFolder.ClaimChildItems) {
                if ($MyInvocation.BoundParameters.ContainsKey('ClaimName') -and $claim.ClaimName -notin $ClaimName) {
                    continue
                }
                if ($MyInvocation.BoundParameters.ContainsKey('LoginProvider') -and $claim.ClaimProvider -ne $LoginProvider.Id) {
                    continue
                }
                $claim
                $matchFound = $true
            }
            if ($MyInvocation.BoundParameters.ContainsKey('ClaimName') -and -not $matchFound) {
                Write-Error "No claim found matching the name '$ClaimName' in role '$($r.Name)'."
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsRoleClaim -ParameterName ClaimName -ScriptBlock {
    $values = (Get-VmsLoginProvider | Get-VmsLoginProviderClaim).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Get-VmsRoleMember {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.User])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, Position = 0)]
        [Alias('RoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Role) {
            $Role = Get-VmsRole
        }
        foreach ($record in $Role) {
            foreach ($user in $record.UserFolder.Users) {
                $user
            }
        }
    }
}

function Get-VmsRoleOverallSecurity {
    [CmdletBinding()]
    [OutputType([hashtable])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('RoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role]
        $Role,

        [Parameter()]
        [SecurityNamespaceTransformAttribute()]
        [guid[]]
        $SecurityNamespace
    )

    begin {
        Assert-VmsRequirementsMet
        $namespacevalues = Get-SecurityNamespaceValues
        if ($SecurityNamespace.Count -eq 0) {
            $SecurityNamespace = $namespacevalues.SecurityNamespacesById.Keys
        }
    }

    process {
        if ($Role.RoleType -ne 'UserDefined') {
            Write-Error 'Overall security settings do not apply to the Administrator role.'
            return
        }

        try {
            foreach ($namespace in $SecurityNamespace) {
                $response = $Role.ChangeOverallSecurityPermissions($namespace)
                $result = @{
                    Role        = $Role.Path
                    DisplayName = $namespacevalues.SecurityNamespacesById[$namespace]
                }
                foreach ($key in $response.GetPropertyKeys()) {
                    $result[$key] = $response.GetProperty($key)
                }
                # :: milestonesystemsinc/powershellsamples/issue-81
                # Older VMS versions may not include a SecurityNamespace value
                # in the ChangeOverallSecurityPermissions properties which means
                # you can't pass this hashtable into Set-VmsRoleOverallSecurity
                # without explicity including the namespace parameter. So we'll
                # manually add it here just in case it's not already set.
                $result['SecurityNamespace'] = $namespace.ToString()
                $result
            }
        } catch {
            Write-Error -ErrorRecord $_
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsRoleOverallSecurity -ParameterName SecurityNamespace -ScriptBlock {
    $values = (Get-SecurityNamespaceValues).SecurityNamespacesByName.Keys | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Import-VmsRole {
    [CmdletBinding(DefaultParameterSetName = 'Path', SupportsShouldProcess)]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
        [object[]]
        $InputObject,

        [Parameter(Mandatory, ParameterSetName = 'Path')]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $Force,

        [Parameter()]
        [switch]
        $RemoveUndefinedClaims,

        [Parameter()]
        [switch]
        $RemoveUndefinedUsers
    )

    begin {
        Assert-VmsRequirementsMet
        $null = Get-VmsManagementServer -ErrorAction Stop

        if ($MyInvocation.BoundParameters.ContainsKey('Path')) {
            $resolvedPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue -ErrorVariable rpError).Path
            if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
                $resolvedPath = $rpError.TargetObject
            }
            $Path = $resolvedPath
            $fileInfo = [io.fileinfo]$Path
            if (-not $fileInfo.Directory.Exists) {
                throw ([io.directorynotfoundexception]::new("Directory not found: $($fileInfo.Directory.FullName)"))
            }
            if (($fi = [io.fileinfo]$Path).Extension -ne '.json') {
                Write-Verbose "A .json file extension will be added to the file '$($fi.Name)'"
                $Path += ".json"
            }
        }


        $roles = @{}
        (Get-VmsManagementServer).RoleFolder.ClearChildrenCache()
        Get-VmsRole | Foreach-Object {
            if ($roles.ContainsKey($_.Name)) {
                throw "There are multiple existing roles with the same case-insensitive name '$($_.Name)'. The VMS may allow this, but this cmdlet does not. Please consider renaming roles so that they all have unique names."
            }
            $roles[$_.Name] = $_
        }

        

        $providers = @{}
        $supportsOidc = [version](Get-VmsManagementServer).Version -ge '22.1'
        if ($supportsOidc) {
            Get-VmsLoginProvider | Foreach-Object {
                if ($null -eq $_) { return }
                $providers[$_.Name] = $_
            }
        }

        $clientProfiles = @{}
        (Get-VmsManagementServer).ClientProfileFolder.ClientProfiles | ForEach-Object {
            if ($null -eq $_) { return }
            $clientProfiles[$_.Name] = $_
        }

        $timeProfiles = @{
            'Always' = [pscustomobject]@{
                Name        = 'Always'
                DisplayName = 'Always'
                Path        = 'TimeProfile[11111111-1111-1111-1111-111111111111]'
            }
            'Default' = [pscustomobject]@{
                Name        = 'Default'
                DisplayName = 'Default'
                Path        = 'TimeProfile[00000000-0000-0000-0000-000000000000]'
            }
        }
        (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles | ForEach-Object {
            if ($null -eq $_) { return }
            $timeProfiles[$_.Name] = $_
        }

        $basicUsers = @{}
        Get-VmsBasicUser -External:$false | ForEach-Object {
            $basicUsers[$_.Name] = $_
        }
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $InputObject = [io.file]::ReadAllText($Path, [text.encoding]::UTF8) | ConvertFrom-Json -ErrorAction Stop
        }

        foreach ($dto in $InputObject) {
            if ([string]::IsNullOrWhiteSpace($dto.Name)) {
                Write-Error -Message "Record does not have a 'Name' property, the minimum required information to create a new role." -TargetObject $dto
                continue
            }
            $role = $roles[$dto.Name]
            if ($role -and -not $Force) {
                Write-Warning "Role '$($dto.Name)' already exists. To import changes to existing roles, use the -Force switch."
                continue
            }

            $roleParams = @{
                ErrorAction = 'Stop'
            }
            foreach ($propertyName in 'Name', 'Description', 'AllowSmartClientLogOn', 'AllowMobileClientLogOn', 'AllowWebClientLogOn', 'DualAuthorizationRequired', 'MakeUsersAnonymousDuringPTZSession', 'ClientLogOnTimeProfile', 'DefaultTimeProfile', 'ClientProfile') {
                $propertyValue = $dto.$propertyName
                if ($propertyName -in @('DefaultTimeProfile', 'ClientLogOnTimeProfile')) {
                    if ($propertyValue -ne 'Always' -and $propertyValue -ne 'Default') {
                        # The default "Always" and "<default>" time profiles are not actually a time profile defined in (Get-VmsManagementServer).TimeProfileFolder.TimeProfiles
                        # but the TimeProfileNameTransformAttribute class will accept 'Always' or 'Default' as a value and mock up a TimeProfile object for us.
                        $propertyValue = $timeProfiles[$propertyValue]
                    }
                }
                if ($propertyName -eq 'ClientProfile' -and -not $clientProfiles.ContainsKey($dto.ClientProfile)) {
                    $propertyValue = $null
                }
                if ($null -ne $propertyValue -or $propertyName -eq 'Description') {
                    $roleParams[$propertyName] = $propertyValue
                } else {
                    Write-Warning "Skipping property '$propertyName'. Unable to resolve the value '$($dto.$propertyName)'."
                }
            }

            # Create/update the main role properties
            if ($role) {
                $roleParams.Role = $role
                $roleParams.PassThru = $true
                $role = Set-VmsRole @roleParams
            }
            else {
                $role = New-VmsRole @roleParams
            }

            # Update overall security for all roles except default admin role
            if ($role.RoleType -eq 'UserDefined') {
                foreach ($definition in $dto.OverallSecurity) {
                    $permissions = $definition
                    if ($permissions -isnot [System.Collections.IDictionary]) {
                        $permissions = @{}
                        ($definition | Get-Member -MemberType NoteProperty).Name | ForEach-Object {
                            $permissions[$_] = $definition.$_
                        }
                    }
                    $role | Set-VmsRoleOverallSecurity -Permissions $permissions
                }
            }

            # Update the role members, and claims
            if ($supportsOidc) {
                $existingClaims = @()
                $role | Get-VmsRoleClaim | ForEach-Object {
                    $existingClaims += $_
                }
                foreach ($claim in $dto.Claims) {
                    if ([string]::IsNullOrWhiteSpace($claim.LoginProvider) -or -not $providers.ContainsKey($claim.LoginProvider)) {
                        Write-Warning "Skipping claim '$($claim.ClaimName)'. Unable to resolve LoginProvider value '$($claim.LoginProvider)'."
                        continue
                    }
                    $provider = $providers[$claim.LoginProvider]
                    $registeredClaims = ($provider | Get-VmsLoginProviderClaim).Name
                    if ($claim.ClaimName -notin $registeredClaims) {
                        Write-Verbose "Adding '$($claim.ClaimName)' as a new registered claim."
                        $provider | Add-VmsLoginProviderClaim -Name $claim.ClaimName
                    }
                    if ($null -eq ($existingClaims | Where-Object {$_.ClaimProvider -eq $provider.Id -and $_.ClaimName -eq $claim.ClaimName -and $_.ClaimValue -eq $claim.ClaimValue })) {
                        $role | Add-VmsRoleClaim -LoginProvider $provider -ClaimName $claim.ClaimName -ClaimValue $claim.ClaimValue
                        $existingClaims += [pscustomobject]@{
                            ClaimProvider = $provider.Id
                            ClaimName     = $claim.ClaimName
                            ClaimValue    = $claim.ClaimValue
                        }
                    }
                }
                if ($RemoveUndefinedClaims) {
                    foreach ($claim in $existingClaims) {
                        $provider = Get-VmsLoginProvider | Where-Object Id -eq $claim.ClaimProvider
                        $definedClaims = $dto.Claims | Where-Object { $_.LoginProvider -eq $provider.Name -and $_.ClaimName -eq $claim.ClaimName -and $_.ClaimValue -eq $claim.ClaimValue }
                        if ($null -eq $definedClaims) {
                            $role | Remove-VmsRoleClaim -LoginProvider $provider -ClaimName $claim.ClaimName -ClaimValue $claim.ClaimValue
                        }
                    }
                }
            }

            $existingUsers = @{}
            $role | Get-VmsRoleMember | ForEach-Object {
                $existingUsers[$_.Sid] = $null
            }
            foreach ($user in $dto.Users) {
                if ($user.Sid -and -not $existingUsers.ContainsKey($user.Sid)) {
                    if ($user.IdentityType -eq 'BasicUser') {
                        if ($basicUsers.ContainsKey($user.AccountName)) {
                            $user.Sid = $basicUsers[$user.AccountName].Sid
                        } else {
                            try {
                                $passwordChars = [System.Web.Security.Membership]::GeneratePassword(26, 10).ToCharArray() + (Get-Random -Minimum 1000 -Maximum 10000).ToString().ToCharArray()
                                $randomPassword = [securestring]::new()
                                ($passwordChars | Get-Random -Count ($passwordChars.Length)) | ForEach-Object { $randomPassword.AppendChar($_) }
                                $newUser = New-VmsBasicUser -Name $user.AccountName -Password $randomPassword -Status LockedOutByAdmin
                                $basicUsers[$newUser.Name] = $newUser
                                $user.Sid = $newUser.Sid
                            } finally {
                                0..($passwordChars.Length - 1) | ForEach-Object { $passwordChars[$_] = 0 }
                                Remove-Variable -Name passwordChars
                            }
                        }
                    }
                    $role | Add-VmsRoleMember -Sid $user.Sid
                    $existingUsers[$user.Sid] = $null
                }
            }
            if ($RemoveUndefinedUsers) {
                foreach ($sid in $existingUsers.Keys | Where-Object { $_ -notin $dto.Users.Sid}) {
                    $role | Remove-VmsRoleMember -Sid $sid
                }
            }

            $role
        }
    }
}


function New-VmsRole {
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('Add-Role')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Role])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $Description,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $AllowSmartClientLogOn,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $AllowMobileClientLogOn,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $AllowWebClientLogOn,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $DualAuthorizationRequired,

        [Parameter(ValueFromPipelineByPropertyName)]
        [switch]
        $MakeUsersAnonymousDuringPTZSession,

        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('RoleClientLogOnTimeProfile')]
        [ArgumentCompleter([MipItemNameCompleter[TimeProfile]])]
        [MipItemTransformation([TimeProfile])]
        [TimeProfile]
        $ClientLogOnTimeProfile,

        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('RoleDefaultTimeProfile')]
        [ArgumentCompleter([MipItemNameCompleter[TimeProfile]])]
        [MipItemTransformation([TimeProfile])]
        [TimeProfile]
        $DefaultTimeProfile,

        [Parameter(ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $ms = Get-VmsManagementServer -ErrorAction Stop
            if (-not $PSCmdlet.ShouldProcess("$($ms.Name) ($($ms.ServerId.Uri))", "Create role '$Name'")) {
                return
            }

            $serverTask = $ms.RoleFolder.AddRole(
                $Name, $Description,
                $DualAuthorizationRequired,
                $MakeUsersAnonymousDuringPTZSession,
                $AllowMobileClientLogOn, $AllowSmartClientLogOn, $AllowWebClientLogOn,
                $DefaultTimeProfile.Path, $ClientLogOnTimeProfile.Path)

            if ($serverTask.State -ne 'Success') {
                throw "RoleFolder.AddRole(..) state: $($serverTask.State). Error: $($serverTask.ErrorText)"
            }

            $newRole = [VideoOS.Platform.ConfigurationItems.Role]::new($ms.ServerId, $serverTask.Path)
            if ($MyInvocation.BoundParameters.ContainsKey('ClientProfile')) {
                $newRole | Set-VmsRole -ClientProfile $ClientProfile
            }

            <#
                TFS 540814 / 577523: On 2022 R2 and earlier, time profile paths were ignored during role creation and you needed to set these after creating the role.
            #>
            $dirty = $false
            if ($MyInvocation.BoundParameters.ContainsKey('ClientLogOnTimeProfile') -and $newRole.RoleClientLogOnTimeProfile -ne $ClientLogOnTimeProfile.Path) {
                $newRole.RoleClientLogOnTimeProfile = $ClientLogOnTimeProfile.Path
                $dirty = $true
            }
            if ($MyInvocation.BoundParameters.ContainsKey('DefaultTimeProfile') -and $newRole.RoleDefaultTimeProfile -ne $DefaultTimeProfile.Path) {
                $newRole.RoleDefaultTimeProfile = $DefaultTimeProfile.Path
                $dirty = $true
            }
            if ($dirty) {
                $null = $newRole.Save()
            }


            $newRole
            if ($PassThru) {
                Write-Verbose "NOTICE: The PassThru parameter is deprecated as of MilestonePSTools v23.1.2. The new role is now always returned."
            }
        } catch {
            if ($_.Exception.Message) {
                Write-Error -Message $_.Exception.Message -Exception $_.Exception
            } else {
                Write-Error -ErrorRecord $_
            }
        }
    }
}

function Remove-VmsRole {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'ByName')]
    [Alias('Remove-Role')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName, ParameterSetName = 'ByName')]
        [Alias('RoleName', 'Name')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role]
        $Role,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'ById')]
        [Alias('RoleId')]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Role) {
            $Role = Get-VmsRole -Id $Id -ErrorAction Stop
        }
        if (-not $PSCmdlet.ShouldProcess("Role: $($Role.Name)", "Delete")) {
            return
        }
        try {
            $folder = (Get-VmsManagementServer).RoleFolder
            $invokeResult = $folder.RemoveRole($Role.Path)
            if ($invokeResult.State -ne 'Success') {
                throw "Error removing role '$($Role.Name)'. $($invokeResult.GetProperty('ErrorText'))"
            }
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsRole -ParameterName Id -ScriptBlock {
    $values = (Get-VmsRole | Sort-Object Name).Id
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Remove-VmsRoleClaim {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('22.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('RoleName')]
        [ValidateNotNull()]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(ValueFromPipelineByPropertyName, Position = 1)]
        [Alias('ClaimProvider')]
        [ArgumentCompleter([MipItemNameCompleter[LoginProvider]])]
        [MipItemTransformation([LoginProvider])]
        [LoginProvider]
        $LoginProvider,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 2)]
        [string[]]
        $ClaimName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $ClaimValue
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        foreach ($r in $Role) {
            $claims = $r | Get-VmsRoleClaim | Where-Object ClaimName -in $ClaimName
            if ($claims.Count -eq 0) {
                Write-Error "No matching claims found on role $($r.Name)."
                continue
            }
            foreach ($c in $claims) {
                if (-not [string]::IsNullOrWhiteSpace($ClaimValue) -and $c.ClaimValue -ne $ClaimValue) {
                    continue
                }
                if ($null -ne $LoginProvider -and $c.ClaimProvider -ne $LoginProvider.Id) {
                    continue
                }
                try {
                    if ($PSCmdlet.ShouldProcess("Claim '$($c.ClaimName)' on role '$($r.Name)'", "Remove")) {
                        $null = $r.ClaimFolder.RemoveRoleClaim($c.ClaimProvider, $c.ClaimName, $c.ClaimValue)
                    }
                } catch {
                    Write-Error -Message $_.Exception.Message -TargetObject $c
                }
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsRoleClaim -ParameterName ClaimName -ScriptBlock {
    $values = (Get-VmsLoginProvider | Get-VmsLoginProviderClaim).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Remove-VmsRoleMember {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'ByUser')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'ByUser')]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'BySid')]
        [Alias('RoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'ByUser')]
        [VideoOS.Platform.ConfigurationItems.User[]]
        $User,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 2, ParameterSetName = 'BySid')]
        [string[]]
        $Sid
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $removeRoleMember = {
            param($role, $member)
            if ($PSCmdlet.ShouldProcess("$($member.Domain)\$($member.AccountName)", "Remove member from role '$($role.Name)'")) {
                $null = $role.UserFolder.RemoveRoleMember($member.Path)
            }
        }
        foreach ($r in $Role) {
            switch ($PSCmdlet.ParameterSetName) {
                'ByUser' {
                    foreach ($u in $User) {
                        try {
                            $removeRoleMember.Invoke($r, $u)
                        }
                        catch {
                            Write-Error -ErrorRecord $_
                        }
                    }
                }

                'BySid' {
                    foreach ($u in $r | Get-VmsRoleMember | Where-Object Sid -in $Sid) {
                        try {
                            $removeRoleMember.Invoke($r, $u)
                        }
                        catch {
                            Write-Error -ErrorRecord $_
                        }
                    }
                }
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsRoleMember -ParameterName Role -ScriptBlock {
    $values = (Get-VmsRole).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Set-VmsRole {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.Role])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role[]]
        $Role,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter()]
        [switch]
        $AllowSmartClientLogOn,

        [Parameter()]
        [switch]
        $AllowMobileClientLogOn,

        [Parameter()]
        [switch]
        $AllowWebClientLogOn,

        [Parameter()]
        [switch]
        $DualAuthorizationRequired,

        [Parameter()]
        [switch]
        $MakeUsersAnonymousDuringPTZSession,

        [Parameter()]
        [Alias('RoleClientLogOnTimeProfile')]
        [ArgumentCompleter([MipItemNameCompleter[TimeProfile]])]
        [MipItemTransformation([TimeProfile])]
        [TimeProfile]
        $ClientLogOnTimeProfile,

        [Parameter()]
        [Alias('RoleDefaultTimeProfile')]
        [ArgumentCompleter([MipItemNameCompleter[TimeProfile]])]
        [MipItemTransformation([TimeProfile])]
        [TimeProfile]
        $DefaultTimeProfile,

        [Parameter()]
        [ArgumentCompleter([MipItemNameCompleter[ClientProfile]])]
        [MipItemTransformation([ClientProfile])]
        [ClientProfile]
        $ClientProfile,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $dirty = $false
        foreach ($r in $Role) {
            try {
                foreach ($property in $r | Get-Member -MemberType Property | Where-Object Definition -like '*set;*' | Select-Object -ExpandProperty Name) {
                    $parameterName = $property
                    switch ($property) {
                        # We would just use the $property variable, but these properties are prefixed with "Role" which is
                        # redundant and doesn't match the New-VmsRole function.
                        'RoleClientLogOnTimeProfile' { $parameterName = 'ClientLogOnTimeProfile' }
                        'RoleDefaultTimeProfile'     { $parameterName = 'DefaultTimeProfile' }
                    }
                    if (-not $PSBoundParameters.ContainsKey($parameterName)) {
                        continue
                    }

                    $newValue = $PSBoundParameters[$parameterName]
                    if ($parameterName -like '*Profile') {
                        $newValue = $newValue.Path
                    }
                    if ($PSBoundParameters[$parameterName] -ceq $r.$property) {
                        continue
                    }
                    if ($PSCmdlet.ShouldProcess($r.Name, "Set $property to $($PSBoundParameters[$parameterName])")) {
                        $r.$property = $newValue
                        $dirty = $true
                    }
                }

                if ($MyInvocation.BoundParameters.ContainsKey('ClientProfile') -and $PSCmdlet.ShouldProcess($r.Name, "Set ClientProfile to $($ClientProfile.Name)")) {
                    try {
                        $serverTask = $r.SetClientProfile($ClientProfile.Path)
                        if ($serverTask.State -ne 'Success') {
                            Write-Error -Message "Failed to update ClientProfile. $($serverTask.ErrorText)" -TargetObject $r
                        }
                    } catch {
                        Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $r
                    }
                }

                if ($dirty) {
                    $r.Save()
                }
                if ($PassThru) {
                    $r
                }
            } catch {
                Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $r
            }
        }
    }
}

function Set-VmsRoleOverallSecurity {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([hashtable])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('RoleName')]
        [ArgumentCompleter([MipItemNameCompleter[Role]])]
        [MipItemTransformation([Role])]
        [Role]
        $Role,

        [Parameter(ValueFromPipelineByPropertyName)]
        [SecurityNamespaceTransformAttribute()]
        [guid]
        $SecurityNamespace,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [hashtable]
        $Permissions
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Role) {
            $roleId = Split-VmsConfigItemPath -Path $Permissions.Role
            if ([string]::IsNullOrEmpty($roleId)) {
                Write-Error "Role must be provided either using the Role parameter, or by including a key of 'Role' in the Permissions hashtable with the Configuration Item path of an existing role."
                return
            }
            $Role = Get-VmsRole -Id $roleId
        }

        if ($Role.RoleType -eq 'Adminstrative') {
            Write-Error 'Overall security settings do not apply to the Administrator role.'
            return
        }

        if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('SecurityNamespace') -and $null -eq ($SecurityNamespace = $Permissions.SecurityNamespace -as [guid])) {
            Write-Error "SecurityNamespace must be provided either using the SecurityNamespace parameter, or by including a key of 'SecurityNamespace' in the Permissions hashtable with a GUID value matching the ID of an existing overall security namespace."
            return
        }

        try {
            $invokeInfo = $Role.ChangeOverallSecurityPermissions($SecurityNamespace)
            $attributes = @{}
            $invokeInfo.GetPropertyKeys() | ForEach-Object { $attributes[$_] = $invokeInfo.GetProperty($_) }
            if ($attributes.Count -eq 0) {
                Write-Error "No security attribute key/value pairs were returned for namespace ID '$SecurityNamespace'." -TargetObject $invokeInfo
                return
            }
            $dirty = $false
            foreach ($key in $Permissions.Keys) {
                if ($key -in 'DisplayName', 'SecurityNamespace', 'Role') {
                    continue
                }
                if (-not $attributes.ContainsKey($key)) {
                    Write-Warning "Attribute '$key' not found in SecurityNamespace"
                    continue
                } elseif ($attributes[$key] -cne $Permissions[$key]) {
                    if ($PSCmdlet.ShouldProcess($Role.Name, "Set $key to $($Permissions[$key])")) {
                        $invokeInfo.SetProperty($key, $Permissions[$key])
                        $dirty = $true
                    }
                }
            }
            if ($dirty) {
                $null = $invokeInfo.ExecuteDefault()
            }
        } catch [VideoOS.Platform.Proxy.ConfigApi.ValidateResultException] {
            $_ | HandleValidateResultException -TargetObject $Role
        } catch {
            Write-Error -ErrorRecord $_
        }
    }
}


Register-ArgumentCompleter -CommandName Set-VmsRoleOverallSecurity -ParameterName Role -ScriptBlock {
    $values = ((Get-VmsManagementServer).RoleFolder.Roles | Where-Object RoleType -EQ 'UserDefined').Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

Register-ArgumentCompleter -CommandName Set-VmsRoleOverallSecurity -ParameterName SecurityNamespace -ScriptBlock {
    $values = (Get-SecurityNamespaceValues).SecurityNamespacesByName.Keys | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Export-VmsRule {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline)]
        [RuleNameTransformAttribute()]
        [ValidateVmsItemType('Rule')]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem[]]
        $Rule,

        [Parameter(Position = 0)]
        [string]
        $Path,

        [Parameter()]
        [switch]
        $PassThru,

        [Parameter()]
        [switch]
        $Force
    )

    begin {
        Assert-VmsRequirementsMet
        if ($MyInvocation.BoundParameters.ContainsKey('Path')) {
            $resolvedPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue -ErrorVariable rpError).Path
            if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
                $resolvedPath = $rpError.TargetObject
            }
            $Path = $resolvedPath
            $fileInfo = [io.fileinfo]$Path
            if (-not $fileInfo.Directory.Exists) {
                throw ([io.directorynotfoundexception]::new("Directory not found: $($fileInfo.Directory.FullName)"))
            }
            if ($fileInfo.Extension -ne '.json') {
                Write-Verbose "A .json file extension will be added to the file '$($fileInfo.Name)'"
                $Path += ".json"
            }
            if ((Test-Path -Path $Path) -and -not $Force) {
                throw ([System.IO.IOException]::new("The file '$Path' already exists. Include the -Force switch to overwrite an existing file."))
            }
        } elseif (-not $MyInvocation.BoundParameters.ContainsKey('PassThru') -or -not $PassThru.ToBool()) {
            throw "Either or both of Path, or PassThru parameters must be specified."
        }
        $rules = @{}
    }

    process {
        if ($Rule.Count -eq 0) {
            $Rule = Get-VmsRule
        }
        foreach ($currentRule in $Rule) {
            $obj = [pscustomobject]@{
                DisplayName = $currentRule.DisplayName
                Enabled     = $currentRule.EnableProperty.Enabled
                Id          = [guid]$currentRule.Path.Substring(5, 36)
                Properties  = [pscustomobject[]]@($currentRule.Properties | Foreach-Object {
                        $prop = $_
                        [pscustomobject]@{
                            DisplayName    = $prop.DisplayName
                            Key            = $prop.Key
                            Value          = $prop.Value
                            ValueType      = $prop.ValueType
                            ValueTypeInfos = [pscustomobject[]]@($prop.ValueTypeInfos | Select-Object @{Name = 'Key'; Expression = { $prop.Key } }, Name, Value)
                            IsSettable     = $prop.IsSettable
                        }
                    })
            }

            $duplicateCount = 0
            $baseName = $obj.DisplayName -replace ' DUPLICATE \d+$', ''
            while ($rules.ContainsKey($obj.DisplayName)) {
                $duplicateCount++
                $obj.DisplayName = $baseName + " DUPLICATE $duplicateCount"
                $obj.Properties | Where-Object Key -eq 'Name' | ForEach-Object { $_.Value = $obj.DisplayName }
            }
            $rules[$obj.DisplayName] = $obj
            if ($duplicateCount) {
                Write-Warning "There are multiple rules named '$baseName'. Duplicates will be renamed."
            }

            if ($PassThru) {
                $obj
            }
        }
    }

    end {
        if ($rules.Count -and $Path) {
            Write-Verbose "Saving $($rules.Count) exported rules in JSON format to $Path"
            [io.file]::WriteAllText($Path, (ConvertTo-Json -InputObject $rules.Values -Depth 10 -Compress), [system.text.encoding]::UTF8)
        }
    }
}


function Get-VmsRule {
    [CmdletBinding(DefaultParameterSetName = 'Name')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.1')]
    param (
        [Parameter(ParameterSetName = 'Name', ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('DisplayName')]
        [SupportsWildcards()]
        [string]
        $Name = '*',

        [Parameter(Mandatory, ParameterSetName = 'Id')]
        [guid]
        $Id
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'Name' {
                    $matchFound = $false
                    Get-ConfigurationItem -Path /RuleFolder -ChildItems -ErrorAction Stop | Where-Object DisplayName -like $Name | Foreach-Object {
                        $matchFound = $true
                        $_
                    }
                    if (-not $matchFound -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
                        Write-Error "Rule with DisplayName '$($Name)' not found."
                    }
                }

                'Id' {
                    Get-ConfigurationItem -Path "Rule[$Id]" -ErrorAction Stop
                }
            }

        } catch {
            if ($null -eq (Get-ConfigurationItem -Path / -ChildItems | Where-Object Path -eq '/RuleFolder')) {
                Write-Error "The current VMS version does not support management of rules using configuration api."
            } elseif ($_.FullyQualifiedErrorId -match 'PathNotFoundExceptionFault') {
                Write-Error "Rule with Id '$Id' not found."
            } else {
                Write-Error -Message $_.Exception.Message -Exception $_.Exception
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsRule -ParameterName Name -ScriptBlock {
    $values = (Get-VmsRule).DisplayName | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Import-VmsRule {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.1')]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromObject')]
        [ValidateScript({
                $members = $_ | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                foreach ($member in @('DisplayName', 'Enabled', 'Id', 'Properties')) {
                    if ($member -notin $members) {
                        throw "InputObject is missing member named '$member'"
                    }
                }
                $true
            })]
        [object[]]
        $InputObject,

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'FromFile')]
        [string]
        $Path
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $progressParams = @{
                Activity        = 'Importing rules'
                PercentComplete = 0
            }
            Write-Progress @progressParams
            if ($PSCmdlet.ParameterSetName -eq 'FromFile') {
                $Path = (Resolve-Path -Path $Path -ErrorAction Stop).Path
                $InputObject = [io.file]::ReadAllText($Path, [text.encoding]::UTF8) | ConvertFrom-Json
            }
            $total = $InputObject.Count
            $processed = 0
            foreach ($exportedRule in $InputObject) {
                try {
                    $progressParams.CurrentOperation = "Importing rule '$($exportedRule.DisplayName)'"
                    $progressParams.PercentComplete = $processed / $total * 100
                    $progressParams.Status = ($progressParams.PercentComplete / 100).ToString('p0')
                    Write-Progress @progressParams

                    if ($PSCmdlet.ShouldProcess($exportedRule.DisplayName, "Create rule")) {
                        $newRule = $exportedRule | New-VmsRule -ErrorAction Stop
                        $newRule
                    }
                } catch {
                    Write-Error -ErrorRecord $_
                } finally {
                    $processed++
                }
            }
        } finally {
            $progressParams.Completed = $true
            $progressParams.PercentComplete = 100
            Write-Progress @progressParams
        }
    }
}


function New-VmsRule {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.1')]
    param (
        [Parameter(Mandatory, Position = 0, ValueFromPipelineByPropertyName)]
        [Alias('DisplayName')]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [PropertyCollectionTransformAttribute()]
        [hashtable]
        $Properties,

        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('EnableProperty')]
        [BooleanTransformAttribute()]
        [bool]
        $Enabled = $true
    )

    begin {
        Assert-VmsRequirementsMet
        $ruleFolder = Get-ConfigurationItem -Path /RuleFolder
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($Name, "Create rule")) {
            return
        }
        $invokeInfo = $null
        try {
            $Properties['Name'] = $Name
            $invokeInfo = $ruleFolder | Invoke-Method -MethodId AddRule
            $lastPropertyCount = $invokeInfo.Properties.Count
            $iteration = 0
            $maxIterations = 20
            $filledProperties = @{ Id = $null }
            do {
                if ((++$iteration) -ge $maxIterations) {
                    $propertyDump = ($invokeInfo.Properties | Select-Object Key, Value, @{Name = 'ValueTypeInfos'; Expression = { $_.ValueTypeInfos.Value -join '|'}}) | Format-Table | Out-String
                    Write-Verbose "InvokeInfo Properties:`r`n$propertyDump"

                    $exception = [invalidoperationexception]::new("Maximum request/response iterations reached while creating rule. This can happen when the supplied properties hashtable is missing important key/value pairs or when a provided value is incorrect. Inspect the 'Properties' collection on the TargetObject property on this ErrorRecord.")
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, $exception.Message, [System.Management.Automation.ErrorCategory]::InvalidData, $invokeInfo)
                    throw $errorRecord
                }
                try {
                    foreach ($key in $invokeInfo.Properties.Key) {
                        # Skip key if already set in a previous iteration
                        if ($filledProperties.ContainsKey($key)) {
                            continue
                        } else {
                            $filledProperties[$key] = $null
                        }

                        # If imported rule definition doesn't have a property that the configuration api has,
                        # we might be able to finish creating the rule, or we might end up in a perpetual loop
                        # until we reach $maxIterations and fail.
                        if (-not $Properties.ContainsKey($key)) {
                            Write-Verbose "Property with key '$key' not provided in Properties hashtable for new rule '$($Name)'."
                            continue
                        }

                        # Protect against null or empty property values
                        if ([string]::IsNullOrWhiteSpace($Properties[$key])) {
                            continue
                        }
                        $newRuleProperty = $invokeInfo.Properties | Where-Object Key -eq $key
                        switch ($newRuleProperty.ValueType) {
                            'Enum' {
                                # Use the enum value with the same supplied value using case-insensitive comparison
                                $newValue = ($newRuleProperty.ValueTypeInfos | Where-Object Value -eq $Properties[$key]).Value
                                if ($null -eq $newValue) {
                                    # The user-supplied value doesn't match any enum values so compare against the enum value display names
                                    $newValue = ($newRuleProperty.ValueTypeInfos | Where-Object Name -eq $Properties[$key]).Value
                                    if ($null -eq $newValue) {
                                        Write-Warning "Value for user-supplied property '$key' does not match the available options: $($newRuleProperty.ValueTypeInfos.Value -join ', ')."
                                        $newValue = $Properties[$key]
                                    } else {
                                        Write-Verbose "Value for user-supplied property '$key' has been mapped from '$($Properties[$key])' to '$newValue'"
                                    }
                                }
                                $Properties[$key] = $newValue
                            }
                        }
                        $invokeInfo | Set-ConfigurationItemProperty -Key $key -Value $Properties[$key]
                    }

                    $response = $invokeInfo | Invoke-Method AddRule -ErrorAction Stop
                    $invokeInfo = $response
                    $newPropertyCount = $invokeInfo.Properties.Count
                    if ($lastPropertyCount -ge $newPropertyCount -and $null -eq ($invokeInfo | Get-ConfigurationItemProperty -Key 'State' -ErrorAction SilentlyContinue)) {
                        $exception = [invalidoperationexception]::new("Invalid rule definition. Inspect the properties of the InvokeInfo object in this error's TargetObject property. This is commonly a result of creating a rule using the ID of an object that does not exist.")
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, $exception.Message, [System.Management.Automation.ErrorCategory]::InvalidData, $invokeInfo)
                        throw $errorRecord
                    }
                    $lastPropertyCount = $newPropertyCount
                } catch {
                    throw
                }
            } while ($invokeInfo.ItemType -eq 'InvokeInfo')

            if (($invokeInfo | Get-ConfigurationItemProperty -Key State) -ne 'Success') {
                $exception = [invalidoperationexception]::new("Error in New-VmsRule: $($invokeInfo | Get-ConfigurationItemProperty -Key 'ErrorText' -ErrorAction SilentlyContinue)")
                $errorRecord = [System.Management.Automation.ErrorRecord]::new($_.Exception, $_.Exception.Message, [System.Management.Automation.ErrorCategory]::InvalidData, $invokeInfo)
                throw $errorRecord
            }

            $newRuleId = ($invokeInfo | Get-ConfigurationItemProperty -Key Path) -replace 'Rule\[(.+)\]', '$1'
            $newRule = Get-VmsRule -Id $newRuleId -ErrorAction Stop

            if ($Enabled -ne $newRule.EnableProperty.Enabled) {
                $newRule.EnableProperty.Enabled = $Enabled
                $null = $newRule | Set-ConfigurationItem
            }

            $newRule
        } catch {
            $exception = [invalidoperationexception]::new("An error occurred while creating the rule: $($_.Exception.Message)", $_.Exception)
            $errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, $exception.Message, [System.Management.Automation.ErrorCategory]::InvalidData, $invokeInfo)
            Write-Error -Message $exception.Message -Exception $exception -TargetObject $invokeInfo
        }
    }
}


function Remove-VmsRule {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.1')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [RuleNameTransformAttribute()]
        [ValidateVmsItemType('Rule')]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $Rule
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        if (-not $PSCmdlet.ShouldProcess($Rule.DisplayName, 'Remove')) {
            return
        }

        try {
            $invokeInfo = Get-ConfigurationItem -Path /RuleFolder | Invoke-Method RemoveRule
            $invokeInfo | Set-ConfigurationItemProperty -Key 'RemoveRulePath' -Value $Rule.Path
            $invokeInfo = $invokeInfo | Invoke-Method RemoveRule -ErrorAction Stop
            if (($invokeInfo | Get-ConfigurationItemProperty -Key State) -ne 'Success') {
                throw "Configuration API response did not indicate success."
            }
        } catch {
            Write-Error -Message $_.Exception.Message -Exception $_.Exception -TargetObject $invokeInfo
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsRule -ParameterName Rule -ScriptBlock {
    $values = (Get-VmsRule).DisplayName | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Set-VmsRule {
    [CmdletBinding(SupportsShouldProcess)]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('20.1')]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [RuleNameTransformAttribute()]
        [ValidateVmsItemType('Rule')]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        $Rule,

        [Parameter()]
        [string]
        $Name,

        [Parameter()]
        [BooleanTransformAttribute()]
        [bool]
        $Enabled,

        [Parameter()]
        [PropertyCollectionTransformAttribute()]
        [hashtable]
        $Properties,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $dirty = $false
        if ($MyInvocation.BoundParameters.ContainsKey('Name')) {
            $currentValue = $Rule | Get-ConfigurationItemProperty -Key Name
            if ($Name -cne $currentValue -and $PSCmdlet.ShouldProcess("Rule '$($Rule.DisplayName)'", "Set DisplayName to $Name")) {
                $Rule.DisplayName = $Name
                $Rule | Set-ConfigurationItemProperty -Key Name -Value $Name
                $dirty = $true
            }
        }
        if ($MyInvocation.BoundParameters.ContainsKey('Enabled')) {
            if ($Enabled -ne $Rule.EnableProperty.Enabled -and $PSCmdlet.ShouldProcess("Rule '$($Rule.DisplayName)'", "Set Enabled to $Enabled")) {
                $Rule.EnableProperty.Enabled = $Enabled
                $dirty = $true
            }
        }

        if ($MyInvocation.BoundParameters.ContainsKey('Properties') -and $PSCmdlet.ShouldProcess("Rule '$($Rule.DisplayName)'", "Update properties")) {
            $currentProperties = @{}
            $Rule.Properties | ForEach-Object {
                $currentProperties[$_.Key] = $_.Value
            }
            foreach ($newProperty in $Properties.GetEnumerator()) {
                if ($currentProperties.ContainsKey($newProperty.Key)) {
                    if ($newProperty.Value -cne $currentProperties[$newProperty.Key]) {
                        $Rule | Set-ConfigurationItemProperty -Key $newProperty.Key -Value $newProperty.Value
                        $dirty = $true
                    }
                } else {
                    $Rule.Properties += [VideoOS.ConfigurationApi.ClientService.Property]@{ Key = $newProperty.Key; Value = $newProperty.Value.ToString() }
                    $dirty = $true
                }
            }
        }

        if ($dirty -and $PSCmdlet.ShouldProcess("Rule '$($Rule.DisplayName)'", 'Save changes')) {
            $null = $Rule | Set-ConfigurationItem
        }

        if ($PassThru) {
            $Rule
        }
    }
}

Register-ArgumentCompleter -CommandName Set-VmsRule -ParameterName Rule -ScriptBlock {
    $values = (Get-VmsRule).DisplayName | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Add-VmsArchiveStorage {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.ArchiveStorage])]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter()]
        [string]
        $Description,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter()]
        [ValidateTimeSpanRange('00:01:00', '365000.00:00:00')]
        [timespan]
        $Retention,

        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaximumSizeMB,

        [Parameter()]
        [switch]
        $ReduceFramerate,

        [Parameter()]
        [ValidateRange(0.00028, 100)]
        [double]
        $TargetFramerate = 5
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $archiveFolder = $Storage.ArchiveStorageFolder
        if ($PSCmdlet.ShouldProcess("Recording storage '$($Storage.Name)'", "Add new archive storage named '$($Name)' with retention of $($Retention.TotalHours) hours and a maximum size of $($MaximumSizeMB) MB")) {
            try {
                $taskInfo = $archiveFolder.AddArchiveStorage($Name, $Description, $Path, $TargetFrameRate, $Retention.TotalMinutes, $MaximumSizeMB)
                if ($taskInfo.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    Write-Error -Message $taskInfo.ErrorText
                    return
                }

                $archive = [VideoOS.Platform.ConfigurationItems.ArchiveStorage]::new((Get-VmsManagementServer).ServerId, $taskInfo.Path)

                if ($ReduceFramerate) {
                    $invokeInfo = $archive.SetFramerateReductionArchiveStorage()
                    $invokeInfo.SetProperty('FramerateReductionEnabled', 'True')
                    [void]$invokeInfo.ExecuteDefault()
                }

                $storage.ClearChildrenCache()
                Write-Output $archive
            }
            catch {
                Write-Error $_
                return
            }
        }
    }
}


function Add-VmsStorage {
    [CmdletBinding(DefaultParameterSetName = 'WithoutEncryption', SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.Storage])]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'WithEncryption')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [string]
        $Description,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [ValidateTimeSpanRange('00:01:00', '365000.00:00:00')]
        [timespan]
        $Retention,

        [Parameter(Mandatory, ParameterSetName = 'WithoutEncryption')]
        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaximumSizeMB,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [switch]
        $Default,

        [Parameter(ParameterSetName = 'WithoutEncryption')]
        [Parameter(ParameterSetName = 'WithEncryption')]
        [switch]
        $EnableSigning,

        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [ValidateSet('Light', 'Strong', IgnoreCase = $false)]
        [string]
        $EncryptionMethod,

        [Parameter(Mandatory, ParameterSetName = 'WithEncryption')]
        [securestring]
        $Password
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $storageFolder = $RecordingServer.StorageFolder
        if ($PSCmdlet.ShouldProcess("Recording Server '$($RecordingServer.Name)' at $($RecordingServer.HostName)", "Add new storage named '$($Name)' with retention of $($Retention.TotalHours) hours and a maximum size of $($MaximumSizeMB) MB")) {
            try {
                $taskInfo = $storageFolder.AddStorage($Name, $Description, $Path, $EnableSigning, $Retention.TotalMinutes, $MaximumSizeMB)
                if ($taskInfo.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    Write-Error -Message $taskInfo.ErrorText
                    return
                }
                $storageFolder.ClearChildrenCache()
            }
            catch {
                Write-Error $_
                return
            }

            $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-VmsManagementServer).ServerId, $taskInfo.Path)
        }

        if ($PSCmdlet.ParameterSetName -eq 'WithEncryption' -and $PSCmdlet.ShouldProcess("Recording Storage '$Name'", "Enable '$EncryptionMethod' Encryption")) {
            try {
                $invokeResult = $storage.EnableEncryption($Password, $EncryptionMethod)
                if ($invokeResult.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    throw $invokeResult.ErrorText
                }

                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-VmsManagementServer).ServerId, $taskInfo.Path)
            }
            catch {
                [void]$storageFolder.RemoveStorage($taskInfo.Path)
                Write-Error $_
                return
            }
        }

        if ($Default -and $PSCmdlet.ShouldProcess("Recording Storage '$Name'", "Set as default storage configuration")) {
            try {
                $invokeResult = $storage.SetStorageAsDefault()
                if ($invokeResult.State -ne [videoos.platform.configurationitems.stateenum]::Success) {
                    throw $invokeResult.ErrorText
                }

                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-VmsManagementServer).ServerId, $taskInfo.Path)
            }
            catch {
                [void]$storageFolder.RemoveStorage($taskInfo.Path)
                Write-Error $_
                return
            }
        }

        if (!$PSBoundParameters.ContainsKey('WhatIf')) {
            Write-Output $storage
        }
    }
}

function Get-VmsArchiveStorage {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.ConfigurationItems.ArchiveStorage])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string]
        $Name = '*'
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $storagesMatched = 0
        $Storage.ArchiveStorageFolder.ArchiveStorages | ForEach-Object {
            if ($_.Name -like $Name) {
                $storagesMatched++
                Write-Output $_
            }
        }

        if ($storagesMatched -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
            Write-Error "No recording storages found matching the name '$Name'"
        }
    }
}


function Get-VmsStorage {
    [CmdletBinding(DefaultParameterSetName = 'FromName')]
    [OutputType([VideoOS.Platform.ConfigurationItems.Storage])]
    [RequiresVmsConnection()]
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'FromName')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer[]]
        $RecordingServer,

        [Parameter(ParameterSetName = 'FromName')]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string]
        $Name = '*',

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromPath')]
        [ValidateScript({
            if ($_ -match 'Storage\[.{36}\]') {
                $true
            }
            else {
                throw "Invalid storage item path. Expected format: Storage[$([guid]::NewGuid())]"
            }
        })]
        [Alias('RecordingStorage', 'Path')]
        [string]
        $ItemPath
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'FromName' {
                if ($null -eq $RecordingServer -or $RecordingServer.Count -eq 0) {
                    $RecordingServer = Get-VmsRecordingServer
                }
                $storagesMatched = 0
                $RecordingServer.StorageFolder.Storages | ForEach-Object {
                    if ($_.Name -like $Name) {
                        $storagesMatched++
                        Write-Output $_
                    }
                }

                if ($storagesMatched -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name)) {
                    Write-Error "No recording storages found matching the name '$Name'"
                }
            }
            'FromPath' {
                [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-VmsManagementServer).ServerId, $ItemPath)
            }
            Default {
                throw "ParameterSetName $($PSCmdlet.ParameterSetName) not implemented"
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsStorage -ParameterName Name -ScriptBlock {
    $values = (Get-VmsRecordingServer | Get-VmsStorage).Name | Select-Object -Unique | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Remove-VmsArchiveStorage {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByName')]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage,

        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByStorage')]
        [VideoOS.Platform.ConfigurationItems.ArchiveStorage]
        $ArchiveStorage
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                foreach ($archiveStorage in $Storage | Get-VmsArchiveStorage -Name $Name) {
                    $archiveStorage | Remove-VmsArchiveStorage
                }
            }

            'ByStorage' {
                $recorder = [VideoOS.Platform.ConfigurationItems.RecordingServer]::new((Get-VmsManagementServer).ServerId, $Storage.ParentItemPath)
                $storage = [VideoOS.Platform.ConfigurationItems.Storage]::new((Get-VmsManagementServer).ServerId, $ArchiveStorage.ParentItemPath)
                if ($PSCmdlet.ShouldProcess("Recording server $($recorder.Name)", "Delete archive $($ArchiveStorage.Name) from $($storage.Name)")) {
                    $folder = [VideoOS.Platform.ConfigurationItems.ArchiveStorageFolder]::new((Get-VmsManagementServer).ServerId, $ArchiveStorage.ParentPath)
                    [void]$folder.RemoveArchiveStorage($ArchiveStorage.Path)
                }
            }
            Default {
                throw 'Unknown parameter set'
            }
        }
    }
}


function Remove-VmsStorage {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByName')]
        [ArgumentCompleter([MipItemNameCompleter[RecordingServer]])]
        [MipItemTransformation([RecordingServer])]
        [RecordingServer]
        $RecordingServer,

        [Parameter(Mandatory, ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'ByStorage')]
        [VideoOS.Platform.ConfigurationItems.Storage]
        $Storage
    )

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                foreach ($vmsStorage in $RecordingServer | Get-VmsStorage -Name $Name) {
                    $vmsStorage | Remove-VmsStorage
                }
            }

            'ByStorage' {
                $recorder = [VideoOS.Platform.ConfigurationItems.RecordingServer]::new((Get-VmsManagementServer).ServerId, $Storage.ParentItemPath)
                if ($PSCmdlet.ShouldProcess("Recording server $($recorder.Name)", "Delete $($Storage.Name) and all archives")) {
                    $folder = [VideoOS.Platform.ConfigurationItems.StorageFolder]::new((Get-VmsManagementServer).ServerId, $Storage.ParentPath)
                    [void]$folder.RemoveStorage($Storage.Path)
                }
            }
            Default {
                throw 'Unknown parameter set'
            }
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsStorage -ParameterName Name -ScriptBlock {
    $values = (Get-VmsRecordingServer | Get-VmsStorage).Name | Select-Object -Unique | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

function Set-VmsDeviceStorage {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([VideoOS.Platform.ConfigurationItems.IConfigurationItem])]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.IConfigurationItem[]]
        $Device,

        [Parameter(Mandatory)]
        [string]
        $Destination,

        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet -ErrorAction Stop
    }
    
    process {
        foreach ($currentDevice in $Device) {
            try {
                $taskInfo = $currentDevice.ChangeDeviceRecordingStorage()
                $itemSelection =  $taskInfo.ItemSelectionValues.GetEnumerator() | Where-Object { $_.Value -eq $Destination -or $_.Key -eq $Destination }
                if ($itemSelection.Count -eq 0) {
                    Write-Error -TargetObject $currentDevice "No storage destination available for device '$currentDevice' named '$Destination'" -RecommendedAction "Use one of the available destinations: $($taskInfo.ItemSelectionValues.Keys -join ', ')"
                    continue
                } elseif ($itemSelection.Count -gt 1) {
                    Write-Error -TargetObject $currentDevice "More than one storage destination matching '$Destination' for device '$currentDevice'." -RecommendedAction "Check your recording server storage configuration. The only way you should see this error is if a storage configuration display name matches a storage configuration ID on that recording server."
                    continue
                }
                
                if ($PSCmdlet.ShouldProcess($currentDevice, "Set storage to $($itemSelection.Key)")) {
                    $taskInfo.ItemSelection = $itemSelection.Value
                    $task = $taskInfo.ExecuteDefault()
                    $null = $task | Wait-VmsTask -Title "Change device recording storage: $currentDevice" -Cleanup
                    if ($PassThru) {
                        $currentDevice
                    }
                }
            } catch {
                Write-Error -TargetObject $currentDevice -Exception $_.Exception -Message $_.Exception.Message -Category $_.CategoryInfo.Category
            }
        }
    }
}


function ConvertFrom-ConfigurationApiProperties {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Platform.ConfigurationItems.ConfigurationApiProperties]
        $Properties,

        [Parameter()]
        [switch]
        $UseDisplayNames
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $languageId = (Get-Culture).Name
        $result = @{}
        foreach ($key in $Properties.Keys) {
            if ($key -notmatch '^.+/(?<Key>.+)/(?:[0-9A-F\-]{36})$') {
                Write-Warning "Failed to parse property with key name '$key'"
                continue
            }
            $propertyInfo = $Properties.GetValueTypeInfoCollection($key)
            $propertyValue = $Properties.GetValue($key)

            if ($UseDisplayNames) {
                $valueTypeInfo = $propertyInfo | Where-Object Value -eq $propertyValue
                $displayName = $valueTypeInfo.Name
                if ($propertyInfo.Count -gt 0 -and $displayName -and $displayName -notin @('true', 'false', 'MinValue', 'MaxValue', 'StepValue')) {
                    if ($valueTypeInfo.TranslationId -and $languageId -and $languageId -ne 'en-US') {
                        $translatedName = (Get-Translations -LanguageId $languageId).($valueTypeInfo.TranslationId)
                        if (![string]::IsNullOrWhiteSpace($translatedName)) {
                            $displayName = $translatedName
                        }
                    }
                    $result[$Matches.Key] = $displayName
                }
                else {
                    $result[$Matches.Key] = $propertyValue
                }
            }
            else {
                $result[$Matches.Key] = $propertyValue
            }
        }

        Write-Output $result
    }
}


function ConvertFrom-GisPoint {
    [CmdletBinding()]
    [OutputType([system.device.location.geocoordinate])]
    [RequiresVmsConnection($false)]
    param (
        # Specifies the GisPoint value to convert to a GeoCoordinate. Milestone stores GisPoint data in the format "POINT ([longitude] [latitude])" or "POINT EMPTY".
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ValueFromPipeline)]
        [string]
        $GisPoint
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($GisPoint -eq 'POINT EMPTY') {
            Write-Output ([system.device.location.geocoordinate]::Unknown)
        }
        else {
            $temp = $GisPoint.Substring(7, $GisPoint.Length - 8)
            $long, $lat, $null = $temp -split ' '
            Write-Output ([system.device.location.geocoordinate]::new($lat, $long))
        }
    }
}


function ConvertFrom-Snapshot {
    [CmdletBinding()]
    [OutputType([system.drawing.image])]
    [RequiresVmsConnection($false)]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('Bytes')]
        [byte[]]
        $Content
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Content -or $Content.Length -eq 0) {
            return $null
        }
        $ms = [io.memorystream]::new($Content)
        Write-Output ([system.drawing.image]::FromStream($ms))
    }
}


function ConvertTo-GisPoint {
    [CmdletBinding()]
    [OutputType([string])]
    [RequiresVmsConnection($false)]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromGeoCoordinate')]
        [system.device.location.geocoordinate]
        $Coordinate,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Latitude,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Longitude,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'FromValues')]
        [double]
        $Altitude = [double]::NaN,

        [Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'FromString')]
        [string]
        $Coordinates
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {

        switch ($PsCmdlet.ParameterSetName) {
            'FromValues' {
                break
            }

            'FromGeoCoordinate' {
                $Latitude = $Coordinate.Latitude
                $Longitude = $Coordinate.Longitude
                $Altitude = $Coordinate.Altitude
                break
            }

            'FromString' {
                $values = $Coordinates -split ',' | Foreach-Object {
                    [double]$_.Trim()
                }
                if ($values.Count -lt 2 -or $values.Count -gt 3) {
                    Write-Error "Failed to parse coordinates into latitude, longitude and optional altitude."
                    return
                }
                $Latitude = $values[0]
                $Longitude = $values[1]
                if ($values.Count -gt 2) {
                    $Altitude = $values[2]
                }
                break
            }
        }

        if ([double]::IsNan($Altitude)) {
            Write-Output ('POINT ({0} {1})' -f $Longitude, $Latitude)
        }
        else {
            Write-Output ('POINT ({0} {1} {2})' -f $Longitude, $Latitude, $Altitude)
        }
    }
}


function Get-BankTable {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param (
        [Parameter()]
        [string]
        $Path,
        [Parameter()]
        [string[]]
        $DeviceId,
        [Parameter()]
        [DateTime]
        $StartTime = [DateTime]::MinValue,
        [Parameter()]
        [DateTime]
        $EndTime = [DateTime]::MaxValue.AddHours(-1)
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $di = [IO.DirectoryInfo]$Path
        foreach ($table in $di.EnumerateDirectories()) {
            if ($table.Name -match "^(?<id>[0-9a-fA-F\-]{36})(_(?<tag>\w+)_(?<endTime>\d\d\d\d-\d\d-\d\d_\d\d-\d\d-\d\d).*)?") {
                $tableTimestamp = if ($null -eq $Matches["endTime"]) { (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss") } else { $Matches["endTime"] }
                $timestamp = [DateTime]::ParseExact($tableTimestamp, "yyyy-MM-dd_HH-mm-ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal)
                if ($timestamp -lt $StartTime -or $timestamp -gt $EndTime.AddHours(1)) {
                    # Timestamp of table is outside the requested timespan
                    continue
                }
                if ($null -ne $DeviceId -and [cultureinfo]::InvariantCulture.CompareInfo.IndexOf($DeviceId, $Matches["id"], [System.Globalization.CompareOptions]::IgnoreCase) -eq -1) {
                    # Device ID for table is not requested
                    continue
                }
                [pscustomobject]@{
                    DeviceId = [Guid]$Matches["id"]
                    EndTime = $timestamp
                    Tag = $Matches["tag"]
                    IsLiveTable = $null -eq $Matches["endTime"]
                    Path = $table.FullName
                }
            }
        }
    }
}


function Get-ConfigurationItemProperty {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        [ValidateNotNullOrEmpty()]
        $InputObject,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Key
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $property = $InputObject.Properties | Where-Object Key -eq $Key
        if ($null -eq $property) {
            Write-Error -Message "Key '$Key' not found on configuration item $($InputObject.Path)" -TargetObject $InputObject -Category InvalidArgument
            return
        }
        $property.Value
    }
}


function Get-StreamProperties {
    [CmdletBinding()]
    [OutputType([VideoOS.ConfigurationApi.ClientService.Property[]])]
    [RequiresVmsConnection()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Command has already been published.')]
    param (
        # Specifies the camera to retrieve stream properties for
        [Parameter(ValueFromPipeline, Mandatory, ParameterSetName = 'ByName')]
        [Parameter(ValueFromPipeline, Mandatory, ParameterSetName = 'ByNumber')]
        [ArgumentCompleter([MipItemNameCompleter[Camera]])]
        [MipItemTransformation([Camera])]
        [Camera]
        $Camera,

        # Specifies a StreamUsageChildItem from Get-Stream
        [Parameter(ParameterSetName = 'ByName')]
        [ValidateNotNullOrEmpty()]
        [string]
        $StreamName,

        # Specifies the stream number starting from 0. For example, "Video stream 1" is usually in the 0'th position in the StreamChildItems collection.
        [Parameter(ParameterSetName = 'ByNumber')]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $StreamNumber
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' {
                $stream = (Get-ConfigurationItem -Path "DeviceDriverSettings[$($Camera.Id)]").Children | Where-Object { $_.ItemType -eq 'Stream' -and $_.DisplayName -like $StreamName }
                if ($null -eq $stream -and ![system.management.automation.wildcardpattern]::ContainsWildcardCharacters($StreamName)) {
                    Write-Error "No streams found on $($Camera.Name) matching the name '$StreamName'"
                    return
                }
                foreach ($obj in $stream) {
                    Write-Output $obj.Properties
                }
            }
            'ByNumber' {
                $streams = (Get-ConfigurationItem -Path "DeviceDriverSettings[$($Camera.Id)]").Children | Where-Object { $_.ItemType -eq 'Stream' }
                if ($StreamNumber -lt $streams.Count) {
                    Write-Output ($streams[$StreamNumber].Properties)
                }
                else {
                    Write-Error "There are $($streams.Count) streams available on the camera and stream number $StreamNumber does not exist. Remember to index the streams from zero."
                }
            }
            Default {}
        }
    }
}


function Install-StableFPS {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    [RequiresElevation()]
    param (
        [Parameter()]
        [string]
        $Source = "C:\Program Files\Milestone\MIPSDK\Tools\StableFPS",
        [Parameter()]
        [int]
        [ValidateRange(1, 200)]
        $Cameras = 32,
        [Parameter()]
        [int]
        [ValidateRange(1, 5)]
        $Streams = 1,
        [Parameter()]
        [string]
        $DevicePackPath
    )

    begin {
        Assert-VmsRequirementsMet
        if (!(Test-Path (Join-Path $Source "StableFPS_DATA"))) {
            throw "Path not found: $((Join-Path $Source "StableFPS_DATA"))"
        }
        if (!(Test-Path (Join-Path $Source "vLatest"))) {
            throw "Path not found: $((Join-Path $Source "vLatest"))"
        }
    }

    process {
        $serviceStopped = $false
        try {
            $dpPath = if ([string]::IsNullOrWhiteSpace($DevicePackPath)) { (Get-RecorderConfig).DevicePackPath } else { $DevicePackPath }
            if (!(Test-Path $dpPath)) {
                throw "DevicePackPath not valid"
            }
            if ([string]::IsNullOrWhiteSpace($DevicePackPath)) {
                $service = Get-Service "Milestone XProtect Recording Server"
                if ($service.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running) {
                    $service | Stop-Service -Force
                    $serviceStopped = $true
                }
            }

            $srcData = Join-Path $Source "StableFPS_Data"
            $srcDriver = Join-Path $Source "vLatest"
            Copy-Item $srcData -Destination $dpPath -Container -Recurse -Force
            Copy-Item "$srcDriver\*" -Destination $dpPath -Recurse -Force

            $tempXml = Join-Path $dpPath "resources\StableFPS_TEMP.xml"
            $newXml = Join-Path $dpPath "resources\StableFPS.xml"
            $content = Get-Content $tempXml -Raw
            $content = $content.Replace("{CAM_NUM_REQUESTED}", $Cameras)
            $content = $content.Replace("{STREAM_NUM_REQUESTED}", $Streams)
            $content | Set-Content $newXml
            Remove-Item $tempXml
        }
        catch {
            throw
        }
        finally {
            if ($serviceStopped -and $null -ne $service) {
                $service.Refresh()
                $service.Start()
            }
        }
    }
}


function Invoke-ServerConfigurator {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    [RequiresElevation()]
    param(
        # Enable encryption for the CertificateGroup specified
        [Parameter(ParameterSetName = 'EnableEncryption', Mandatory)]
        [switch]
        $EnableEncryption,

        # Disable encryption for the CertificateGroup specified
        [Parameter(ParameterSetName = 'DisableEncryption', Mandatory)]
        [switch]
        $DisableEncryption,

        # Specifies the CertificateGroup [guid] identifying which component for which encryption
        # should be enabled or disabled
        [Parameter(ParameterSetName = 'EnableEncryption')]
        [Parameter(ParameterSetName = 'DisableEncryption')]
        [guid]
        $CertificateGroup,

        # Specifies the thumbprint of the certificate to be used to encrypt communications with the
        # component designated by the CertificateGroup id.
        [Parameter(ParameterSetName = 'EnableEncryption', Mandatory)]
        [string]
        $Thumbprint,

        # List the available certificate groups on the local machine. Output will be a [hashtable]
        # where the keys are the certificate group names (which may contain spaces) and the values
        # are the associated [guid] id's.
        [Parameter(ParameterSetName = 'ListCertificateGroups')]
        [switch]
        $ListCertificateGroups,

        # Register all local components with the optionally specified AuthAddress. If no
        # AuthAddress is provided, the last-known address will be used.
        [Parameter(ParameterSetName = 'Register', Mandatory)]
        [switch]
        $Register,

        # Specifies the address of the Authorization Server which is usually the Management Server
        # address. A [uri] value is expected, but only the URI host value will be used. The scheme
        # and port will be inferred based on whether encryption is enabled/disabled and is fixed to
        # port 80/443 as this is how Server Configurator is currently designed.
        [Parameter(ParameterSetName = 'Register')]
        [uri]
        $AuthAddress,

        [Parameter(ParameterSetName = 'Register')]
        [switch]
        $OverrideLocalManagementServer,

        # Specifies the path to the Server Configurator utility. Omit this path and the path will
        # be discovered using Get-RecorderConfig or Get-ManagementServerConfig by locating the
        # installation path of the Management Server or Recording Server and assuming the Server
        # Configurator is located in the same path.
        [Parameter()]
        [string]
        $Path,

        # Specifies that the standard output from the Server Configurator utility should be written
        # after the operation is completed. The output will include the following properties:
        # - StandardOutput
        # - StandardError
        # - ExitCode
        [Parameter(ParameterSetName = 'EnableEncryption')]
        [Parameter(ParameterSetName = 'DisableEncryption')]
        [Parameter(ParameterSetName = 'Register')]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $exePath = $Path
        if ([string]::IsNullOrWhiteSpace($exePath)) {
            # Find ServerConfigurator.exe by locating either the Management Server or Recording Server installation path
            $configurationInfo = try {
                Get-ManagementServerConfig
            }
            catch {
                try {
                    Get-RecorderConfig
                }
                catch {
                    $null
                }
            }
            if ($null -eq $configurationInfo) {
                Write-Error "Could not find a Management Server or Recording Server installation"
                return
            }
            $fileInfo = [io.fileinfo]::new($configurationInfo.InstallationPath)
            $exePath = Join-Path $fileInfo.Directory.Parent.FullName "Server Configurator\serverconfigurator.exe"
            if (-not (Test-Path $exePath)) {
                Write-Error "Expected to find Server Configurator at '$exePath' but failed."
                return
            }
        }


        # Ensure version is 20.3 (2020 R3) or newer
        $fileInfo = [io.fileinfo]::new($exePath)
        if ($fileInfo.VersionInfo.FileVersion -lt [version]"20.3") {
            Write-Error "Invoke-ServerConfigurator requires Milestone version 2020 R3 or newer as this is when command-line options were introduced. Found Server Configurator version $($fileInfo.VersionInfo.FileVersion)"
            return
        }

        $exitCode = @{
            0 = 'Success'
            -1 = 'Unknown error'
            -2 = 'Invalid arguments'
            -3 = 'Invalid argument value'
            -4 = 'Another instance is running'
        }

        # Get Certificate Group list for either display to user or verification
        $output = Get-ProcessOutput -FilePath $exePath -ArgumentList /listcertificategroups
        if ($output.ExitCode -ne 0) {
            Write-Error "Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
            Write-Error $output.StandardOutput
            return
        }
        Write-Information $output.StandardOutput
        $groups = @{}
        foreach ($line in $output.StandardOutput -split ([environment]::NewLine)) {
            if ($line -match "Found '(?<groupName>.+)' group with ID = (?<groupId>.{36})") {
                $groups.$($Matches.groupName) = [guid]::Parse($Matches.groupId)
            }
        }


        switch ($PSCmdlet.ParameterSetName) {
            'EnableEncryption' {
                if ($MyInvocation.BoundParameters.ContainsKey('CertificateGroup') -and $CertificateGroup -notin $groups.Values) {
                    Write-Error "CertificateGroup value '$CertificateGroup' not found. Use the ListCertificateGroups switch to discover valid CertificateGroup values"
                    return
                }

                $enableArgs = @('/quiet', '/enableencryption', "/thumbprint=$Thumbprint")
                if ($MyInvocation.BoundParameters.ContainsKey('CertificateGroup')) {
                    $enableArgs += "/certificategroup=$CertificateGroup"
                }
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $enableArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "EnableEncryption failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }
            }

            'DisableEncryption' {
                if ($MyInvocation.BoundParameters.ContainsKey('CertificateGroup') -and $CertificateGroup -notin $groups.Values) {
                    Write-Error "CertificateGroup value '$CertificateGroup' not found. Use the ListCertificateGroups switch to discover valid CertificateGroup values"
                    return
                }
                $disableArgs = @('/quiet', '/disableencryption')
                if ($MyInvocation.BoundParameters.ContainsKey('CertificateGroup')) {
                    $disableArgs += "/certificategroup=$CertificateGroup"
                }
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $disableArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "EnableEncryption failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }
            }

            'ListCertificateGroups' {
                Write-Output $groups
                return
            }

            'Register' {
                $registerArgs = @('/register', '/quiet')
                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('AuthAddress')) {
                    $registerArgs += '/managementserveraddress={0}' -f $AuthAddress.ToString()
                    if ($OverrideLocalManagementServer) {
                        $registerArgs += '/overridelocalmanagementserver'
                    }
                }
                $output = Get-ProcessOutput -FilePath $exePath -ArgumentList $registerArgs
                if ($output.ExitCode -ne 0) {
                    Write-Error "Registration failed. Server Configurator exited with code $($output.ExitCode). $($exitCode.($output.ExitCode))."
                    Write-Error $output.StandardOutput
                }
            }

            Default {
            }
        }

        Write-Information $output.StandardOutput
        if ($PassThru) {
            Write-Output $output
        }
    }
}


function Resize-Image {
    [CmdletBinding()]
    [OutputType([System.Drawing.Image])]
    [RequiresVmsConnection($false)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Drawing.Image]
        $Image,

        [Parameter(Mandatory)]
        [int]
        $Height,

        [Parameter()]
        [long]
        $Quality = 95,

        [Parameter()]
        [ValidateSet('BMP', 'JPEG', 'GIF', 'TIFF', 'PNG')]
        [string]
        $OutputFormat,

        [Parameter()]
        [switch]
        $DisposeSource
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Image -or $Image.Width -le 0 -or $Image.Height -le 0) {
            Write-Error 'Cannot resize an invalid image object.'
            return
        }

        [int]$width = $image.Width / $image.Height * $Height
        $bmp = [system.drawing.bitmap]::new($width, $Height)
        $graphics = [system.drawing.graphics]::FromImage($bmp)
        $graphics.InterpolationMode = [system.drawing.drawing2d.interpolationmode]::HighQualityBicubic
        $graphics.DrawImage($Image, 0, 0, $width, $Height)
        $graphics.Dispose()

        try {
            $formatId = if ([string]::IsNullOrWhiteSpace($OutputFormat)) {
                    $Image.RawFormat.Guid
                }
                else {
                    ([system.drawing.imaging.imagecodecinfo]::GetImageEncoders() | Where-Object FormatDescription -eq $OutputFormat).FormatID
                }
            $encoder = [system.drawing.imaging.imagecodecinfo]::GetImageEncoders() | Where-Object FormatID -eq $formatId
            $encoderParameters = [system.drawing.imaging.encoderparameters]::new(1)
            $qualityParameter = [system.drawing.imaging.encoderparameter]::new([system.drawing.imaging.encoder]::Quality, $Quality)
            $encoderParameters.Param[0] = $qualityParameter
            Write-Verbose "Saving resized image as $($encoder.FormatDescription) with $Quality% quality"
            $ms = [io.memorystream]::new()
            $bmp.Save($ms, $encoder, $encoderParameters)
            $resizedImage = [system.drawing.image]::FromStream($ms)
            Write-Output ($resizedImage)
        }
        finally {
            $qualityParameter.Dispose()
            $encoderParameters.Dispose()
            $bmp.Dispose()
            if ($DisposeSource) {
                $Image.Dispose()
            }
        }

    }
}


function Select-Camera {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresInteractiveSession()]
    param(
        [Parameter()]
        [string]
        $Title = "Select Camera(s)",
        [Parameter()]
        [switch]
        $SingleSelect,
        [Parameter()]
        [switch]
        $AllowFolders,
        [Parameter()]
        [switch]
        $AllowServers,
        [Parameter()]
        [switch]
        $RemoveDuplicates,
        [Parameter()]
        [switch]
        $OutputAsItem
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $items = Select-VideoOSItem -Title $Title -Kind ([VideoOS.Platform.Kind]::Camera) -AllowFolders:$AllowFolders -AllowServers:$AllowServers -SingleSelect:$SingleSelect -FlattenOutput
        $processed = @{}
        if ($RemoveDuplicates) {
            foreach ($item in $items) {
                if ($processed.ContainsKey($item.FQID.ObjectId)) {
                    continue
                }
                $processed.Add($item.FQID.ObjectId, $null)
                if ($OutputAsItem) {
                    Write-Output $item
                }
                else {
                    Get-VmsCamera -Id $item.FQID.ObjectId
                }
            }
        }
        else {
            if ($OutputAsItem) {
                Write-Output $items
            }
            else {
                Write-Output ($items | ForEach-Object { Get-VmsCamera -Id $_.FQID.ObjectId })
            }
        }
    }
}


function Select-VideoOSItem {
    [CmdletBinding()]
    [RequiresVmsConnection()]
    [RequiresInteractiveSession()]
    param (
        [Parameter()]
        [string]
        $Title = "Select Item(s)",
        [Parameter()]
        [guid[]]
        $Kind,
        [Parameter()]
        [VideoOS.Platform.Admin.Category[]]
        $Category,
        [Parameter()]
        [switch]
        $SingleSelect,
        [Parameter()]
        [switch]
        $AllowFolders,
        [Parameter()]
        [switch]
        $AllowServers,
        [Parameter()]
        [switch]
        $KindUserSelectable,
        [Parameter()]
        [switch]
        $CategoryUserSelectable,
        [Parameter()]
        [switch]
        $FlattenOutput,
        [Parameter()]
        [switch]
        $HideGroupsTab,
        [Parameter()]
        [switch]
        $HideServerTab
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $form = [MilestonePSTools.UI.CustomItemPickerForm]::new();
        $form.KindFilter = $Kind
        $form.CategoryFilter = $Category
        $form.AllowFolders = $AllowFolders
        $form.AllowServers = $AllowServers
        $form.KindUserSelectable = $KindUserSelectable
        $form.CategoryUserSelectable = $CategoryUserSelectable
        $form.SingleSelect = $SingleSelect
        $form.GroupTabVisable = -not $HideGroupsTab
        $form.ServerTabVisable = -not $HideServerTab
        $form.Icon = [System.Drawing.Icon]::FromHandle([VideoOS.Platform.UI.Util]::ImageList.Images[[VideoOS.Platform.UI.Util]::SDK_GeneralIx].GetHicon())
        $form.Text = $Title
        $form.TopMost = $true
        $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $form.BringToFront()
        $form.Activate()

        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($FlattenOutput) {
                Write-Output $form.ItemsSelectedFlattened
            }
            else {
                Write-Output $form.ItemsSelected
            }
        }
    }
}


function Set-ConfigurationItemProperty {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    [OutputType([VideoOS.ConfigurationApi.ClientService.ConfigurationItem])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.ConfigurationApi.ClientService.ConfigurationItem]
        [ValidateNotNullOrEmpty()]
        $InputObject,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Key,
        [Parameter(Mandatory)]
        [string]
        [ValidateNotNullOrEmpty()]
        $Value,
        [Parameter()]
        [switch]
        $PassThru
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $property = $InputObject.Properties | Where-Object Key -eq $Key
        if ($null -eq $property) {
            Write-Error -Message "Key '$Key' not found on configuration item $($InputObject.Path)" -TargetObject $InputObject -Category InvalidArgument
            return
        }
        $property.Value = $Value
        if ($PassThru) {
            $InputObject
        }
    }
}


function Get-VmsTrustedIssuer {
    <#
    .SYNOPSIS
    Gets the specified, or all TrustedIssuer records from the current Milestone XProtect VMS.
    
    .DESCRIPTION
    Gets the specified, or all TrustedIssuer records from the current Milestone XProtect VMS.
    
    .PARAMETER Id
    Specifies the integer ID value for the TrustedIssuer record to retrieve.
    
    .PARAMETER Refresh
    Specifies that any previously cached copies of the TrustedIssuer(s) should be refreshed.
    
    .EXAMPLE
    Get-VmsTrustedIssuer | Select-Object Id, Issuer, Address
    
    Gets a list of existing TrustedIssuer records and returns the Id, Issuer, and Address properties.
    #>
    [CmdletBinding()]
    [OutputType([VideoOS.Management.VmoClient.TrustedIssuer])]
    [MilestonePSTools.RequiresVmsConnection()]
    [MilestonePSTools.RequiresVmsWindowsUser()]
    [MilestonePSTools.RequiresVmsFeature('FederatedSites')]
    param (
        [Parameter(Position = 0)]
        [ValidateRange(0, [int]::MaxValue)]
        [int]
        $Id,

        [Parameter()]
        [switch]
        $Refresh
    )
    
    begin {
        Assert-VmsRequirementsMet
    }

    process {
        try {
            $client = Get-VmsVmoClient
            if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Id')) {
                $method = $client.Repositories.GetType().GetMethod('GetObjectById', [type[]]@([int], [boolean])).MakeGenericMethod([VideoOS.Management.VmoClient.TrustedIssuer])
                $method.Invoke($client.Repositories, @($Id, $Refresh.ToBool()))
            } else {
                $method = $client.Repositories.GetType().GetMethod('GetObjectsByParentId', [type[]]@([guid], [boolean])).MakeGenericMethod([VideoOS.Management.VmoClient.ManagementServer], [VideoOS.Management.VmoClient.TrustedIssuer])
                $method.Invoke($client.Repositories, @($client.ManagementServer.Id, $Refresh.ToBool()))
            }
        } catch {
            throw
        }
    }
}

function New-VmsTrustedIssuer {
    <#
    .SYNOPSIS
    Creates a new Trusted Issuer on the current Milestone XProtect VMS.
    
    .DESCRIPTION
    This command is used on a child site in a Milestone XProtect VMS to add a parent site as a trusted issuer of tokens.
    This is currently necessary to allow external OIDC identities from Azure or other identity providers to access all
    sites in a Milestone Federated Hierarchy.
    
    .PARAMETER Address
    Specifies the base address of the trusted Milestone Identity Provider (IDP). Normally this will be a URI like
    "https://parentsite.domain/IDP".
    
    .PARAMETER Issuer
    Specifies the OpenID Connect "issuer" string found in the "/IDP/.well-known/openid-configuration" JSON document of
    the new trusted issuer. If this is not provided, it will be discovered automatically. Under normal circumstances the
    value of "issuer" is the same as "Address".
    
    .PARAMETER Force
    Skips validation of the Issuer if provided.
    
    .EXAMPLE
    New-VmsTrustedIssuer -Address https://parentsite.domain/IDP

    Creates a new TrustedIssuer record for the management server at "https://parentsite.domain".
    
    .NOTES
    You must be logged in to the child site using a Windows account. Trusted Issuer records currently cannot be managed
    using a basic user account or an external identity.
    #>
    [CmdletBinding()]
    [OutputType([VideoOS.Management.VmoClient.TrustedIssuer])]
    [MilestonePSTools.RequiresVmsConnection()]
    [MilestonePSTools.RequiresVmsWindowsUser()]
    [MilestonePSTools.RequiresVmsFeature('FederatedSites')]
    param (
        [Parameter()]
        [uri]
        $Address
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (!$PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Address')) {
            if (![uri]::TryCreate((Get-VmsManagementServer).MasterSiteAddress, [urikind]::Absolute, [ref]$Address)) {
                Write-Error "No address was provided and no valid MasterSiteAddress value is available on this management server. Ensure this management server is added as a child site in a Milestone Federated Architecture hierarchy."
                return
            }
        }

        $issuerUri = (Get-VmsOpenIdConfig -Address $Address -ErrorAction Stop).issuer
        $client = Get-VmsVmoClient
        try {
            $trustedIssuer = [VideoOS.Management.VmoClient.TrustedIssuer]::new($client.ManagementServer, $issuerUri, $issuerUri)
            $trustedIssuer.Create()
            $trustedIssuer
        } catch {
            Write-Error -Message "Failed to create a TrustedIssuer record for $issuerUri. See the exception for more information." -Exception $_.Exception
        }
    }
}

function Remove-VmsTrustedIssuer {
    <#
    .SYNOPSIS
    Removes an existing TrustedIssuer record.
    
    .DESCRIPTION
    The Remove-VmsTrustedIssuer command is used to remove or delete an existing TrustedIssuer.
    
    .PARAMETER TrustedIssuer
    Specifies a TrustedIssuer record returned by the Get-VmsTrustedIssuer command.
    
    .EXAMPLE
    Get-VmsTrustedIssuer -Id 4 | Remove-VmsTrustedIssuer
    
    Deletes the TrustedIssuer with Id "4".

    .EXAMPLE
    Get-VmsTrustedIssuer | Remove-VmsTrustedIssuer

    Deletes all TrustedIssuer records.
    #>
    [CmdletBinding()]
    [MilestonePSTools.RequiresVmsConnection()]
    [MilestonePSTools.RequiresVmsFeature('FederatedSites')]
    [MilestonePSTools.RequiresVmsWindowsUser()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [VideoOS.Management.VmoClient.TrustedIssuer]
        $TrustedIssuer
    )
    
    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $TrustedIssuer.Delete()
    }
}

function Assert-VmsRequirementsMet {
    [CmdletBinding()]
    [RequiresVmsConnection($false)]
    param ()
    
    process {
        $frame = Get-PSCallStack | Select-Object -Skip 1 -First 1
        if ($frame.InvocationInfo.MyCommand.CommandType -ne 'Function') {
            return
        }
        # Re-using this cmdlet for telemetry calls since it is already called by all functions.
        # The alternative is to create add a separate call to telemetry in every function begin block.
        [MilestonePSTools.Telemetry.AppInsightsTelemetry]::SendInvokeCommandTelemetry($frame.InvocationInfo, 'NotSpecified')
        foreach ($attribute in $frame.InvocationInfo.MyCommand.ScriptBlock.Attributes) {
            try {
                if (($requirement = $attribute -as [MilestonePSTools.IVmsRequirementValidator])) {
                    $requirement.Source = $frame.FunctionName
                    $requirement.Validate()
                }
            } catch {
                throw
            }
        }
    }
}

function Get-VmsOpenIdConfig {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [uri]
        $Address
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Address')) {
            $loginSettings = Get-LoginSettings
            $Address = $loginSettings.UriCorporate
        }
        $builder = [uribuilder]$Address
        $builder.Path = '/IDP/.well-known/openid-configuration'
        $builder.Query = ''
        try {
            Invoke-RestMethod -Uri $builder.Uri -ErrorAction Stop
        } catch {
            Write-Error -Message "Failed to retrieve the openid configuration from $($builder.Uri)" -Exception $_.Exception -TargetObject $builder.Uri
        }
    }
}

function  Get-VmsVmoClient {
    <#
    .SYNOPSIS
    Gets a Milestone VMO Client used to access and configure a management server.
    
    .DESCRIPTION
    The VMO Client is used internally by Milestone's MIP SDK, but is not supported
    for external use unless otherwise specified. The raw VMO client is provided
    here to enable configuration of trusted issuers from PowerShell.
    
    .EXAMPLE
    $client = Get-VmsVmoClient
    
    Creates a VMO client and stores it in the $client variable.

    .NOTES
    Direct use of the VMO client is not necessary for configuring trusted issuers, but the
    client can be useful for troubleshooting, diagnostics, and experimentation.
    #>
    [CmdletBinding()]
    [OutputType([VideoOS.Management.VmoClient.VmoClient])]
    [MilestonePSTools.RequiresVmsConnection()]
    param ()

    begin {
        Assert-VmsRequirementsMet
    }
    
    process {
        $loginSettings = Get-LoginSettings
        $connection = [VideoOS.Management.VmoClient.ServerConnection]::new($loginSettings.UriCorporate.Host, $true, $loginSettings.IdentityTokenCache.TokenCache)
        [VideoOS.Management.VmoClient.VmoClient]::new($connection)
    }
}

function Split-VmsConfigItemPath {
    [CmdletBinding(DefaultParameterSetName = 'Id')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'Id')]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'ItemType')]
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0, ParameterSetName = 'ParentItemType')]
        [AllowNull()]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [string[]]
        $Path,

        [Parameter(ParameterSetName = 'Id')]
        [switch]
        $Id,

        [Parameter(ParameterSetName = 'ItemType')]
        [switch]
        $ItemType,

        [Parameter(ParameterSetName = 'ParentItemType')]
        [switch]
        $ParentItemType
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        if ($null -eq $Path) { $Path = '' }
        foreach ($record in $Path) {
            try {
                [videoos.platform.proxy.ConfigApi.ConfigurationItemPath]::new($record).($PSCmdlet.ParameterSetName)
            } catch {
                throw
            }
        }
    }
}

function Find-VmsVideoOSItem {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.Item])]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, Position = 0)]
        [string[]]
        $SearchText,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaxCount = [int]::MaxValue,

        [Parameter()]
        [ValidateRange(1, [int]::MaxValue)]
        [int]
        $MaxSeconds = 30,

        [Parameter()]
        [ArgumentCompleter([KindArgumentCompleter])]
        [KindNameTransform()]
        [guid]
        $Kind,

        [Parameter()]
        [VideoOS.Platform.FolderType]
        $FolderType
    )

    begin {
        Assert-VmsRequirementsMet
        $config = [VideoOS.Platform.Configuration]::Instance
    }

    process {
        foreach ($text in $SearchText) {
            $result = [VideoOS.Platform.SearchResult]::OK
            $items = $config.GetItemsBySearch($text, $MaxCount, $MaxSeconds, [ref]$result)

            foreach ($item in $items) {
                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Kind') -and $item.FQID.Kind -ne $Kind) {
                    continue
                }
                if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('FolderType') -and $FolderType -ne $item.FQID.FolderType) {
                    continue
                }
                $item
            }

            if ($result -ne [VideoOS.Platform.SearchResult]::OK) {
                Write-Warning "Search result: $result"
            }
        }
    }
}

function Get-VmsVideoOSItem {
    [CmdletBinding()]
    [OutputType([VideoOS.Platform.Item])]
    [RequiresVmsConnection()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'GetItemByFQID')]
        [VideoOS.Platform.FQID]
        $Fqid,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'GetItem')]
        [VideoOS.Platform.ServerId]
        $ServerId,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'GetItem')]
        [guid]
        $Id,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'GetItem')]
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'GetItems')]
        [ArgumentCompleter([KindArgumentCompleter])]
        [KindNameTransform()]
        [guid]
        $Kind,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'GetItems')]
        [VideoOS.Platform.ItemHierarchy]
        $ItemHierarchy = [VideoOS.Platform.ItemHierarchy]::SystemDefined,

        [Parameter(ParameterSetName = 'GetItems')]
        [VideoOS.Platform.FolderType]
        $FolderType
    )

    begin {
        Assert-VmsRequirementsMet
        $config = [VideoOS.Platform.Configuration]::Instance
    }

    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'GetItemByFQID' {
                    $config.GetItem($Fqid)
                }

                'GetItem' {
                    if ($ServerId) {
                        $config.GetItem($ServerId, $Id, $Kind)
                    } else {
                        $config.GetItem($Id, $Kind)
                    }
                }

                'GetItems' {
                    $checkKind = $false
                    $checkFolderType = $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('FolderType')

                    $stack = [system.collections.generic.stack[VideoOS.Platform.Item]]::new()
                    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Kind')) {
                        $checkKind = $true
                        $config.GetItemsByKind($Kind, $ItemHierarchy) | Foreach-Object {
                            if ($null -ne $_) {
                                $stack.Push($_)
                            }
                        }
                    } else {
                        $config.GetItems($ItemHierarchy) | Foreach-Object {
                            if ($null -ne $_) {
                                $stack.Push($_)
                            }
                        }
                    }
                    while ($stack.Count -gt 0) {
                        $item = $stack.Pop()
                        if (-not $checkKind -or $item.FQID.Kind -eq $Kind) {
                            if (-not $checkFolderType -or $item.FQID.FolderType -eq $FolderType) {
                                $item
                            }
                        }
                        if ($item.HasChildren -ne 'No') {
                            $item.GetChildren() | ForEach-Object {
                                $stack.Push($_)
                            }
                        }
                    }
                }
                Default {
                    throw "ParameterSet '$_' not implemented."
                }
            }
        } catch {
            Write-Error -ErrorRecord $_
        }
    }
}

function Get-VmsWebhook {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [OutputType([MilestonePSTools.Webhook])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.1')]
    param (
        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Name', Position = 0)]
        [SupportsWildcards()]
        [Alias('DisplayName')]
        [string]
        $Name,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'LiteralName')]
        [string]
        $LiteralName,

        [Parameter(ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [string]
        $Path,

        # Any unrecognized parameters and their values will be ignored when splatting a hashtable with keys that do not match a parameter name.
        [Parameter(ValueFromRemainingArguments, DontShow)]
        [object[]]
        $ExtraParams
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $folderPath = 'MIPKind[b9a5bc9c-e9a5-4a15-8453-ffa41f2815ac]/MIPItemFolder'

        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            if ([string]::IsNullOrWhiteSpace($Path)) {
                Get-ConfigurationItem -Path $folderPath -ChildItems | ConvertTo-Webhook
            } else {
                Get-ConfigurationItem -Path $Path -ErrorAction Stop | ConvertTo-Webhook
            }
            return
        }

        $notFound = $true
        Get-ConfigurationItem -Path $folderPath -ChildItems -PipelineVariable webhook | ForEach-Object {
            switch ($PSCmdlet.ParameterSetName) {
                'Name' {
                    if ($webhook.DisplayName -like $Name) {
                        $notFound = $false
                        $webhook | ConvertTo-Webhook
                    }
                }

                'LiteralName' {
                    if ($webhook.DisplayName -eq $LiteralName) {
                        $notFound = $false
                        $webhook | ConvertTo-Webhook
                    }
                }
            }
        }
        if ($notFound -and ($PSCmdlet.ParameterSetName -eq 'LiteralName' -or -not [Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Name))) {
            $Name = if ($PSCmdlet.ParameterSetName -eq 'Name') { $Name } else { $LiteralName }
            Write-Error -Message "Webhook with name matching '$Name' not found." -TargetObject $Name
        }
    }
}

Register-ArgumentCompleter -CommandName Get-VmsWebhook -ParameterName Name -ScriptBlock {
    $values = (Get-VmsWebhook).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

Register-ArgumentCompleter -CommandName Get-VmsWebhook -ParameterName LiteralName -ScriptBlock {
    $values = (Get-VmsWebhook).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function New-VmsWebhook {
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([MilestonePSTools.Webhook])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.1')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, Position = 0)]
        [Alias('DisplayName')]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [uri]
        $Address,

        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]
        $Token,

        # Any unrecognized parameters and their values will be ignored when splatting a hashtable with keys that do not match a parameter name.
        [Parameter(ValueFromRemainingArguments, DontShow)]
        [object[]]
        $ExtraParams
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $folder = Get-ConfigurationItem -Path 'MIPKind[b9a5bc9c-e9a5-4a15-8453-ffa41f2815ac]/MIPItemFolder'
        $invokeInfo = $folder | Invoke-Method -MethodId AddMIPItem
        'ApiVersion', 'Address', 'Token' | ForEach-Object {
            $invokeInfo.Properties += [VideoOS.ConfigurationApi.ClientService.Property]@{
                Key         = $_
                DisplayName = $_
                ValueType   = 'String'
                IsSettable  = $true
            }
        }
        $action = 'Create webhook {0}' -f $Name
        if ($PSCmdlet.ShouldProcess((Get-VmsSite).Name, $action)) {
            $invokeInfo | Set-ConfigurationItemProperty -Key Name -Value $Name
            $invokeInfo | Set-ConfigurationItemProperty -Key Address -Value $Address
            $invokeInfo | Set-ConfigurationItemProperty -Key ApiVersion -Value 'v1.0'
            if (-not [string]::IsNullOrWhiteSpace($Token)) {
                $invokeInfo | Set-ConfigurationItemProperty -Key Token -Value $Token
            }
            $invokeInfo | Invoke-Method -MethodId AddMIPItem | Get-ConfigurationItem | ConvertTo-Webhook
        }
    }
}


function Remove-VmsWebhook {
    [CmdletBinding(DefaultParameterSetName = 'Path', SupportsShouldProcess)]
    [RequiresVmsVersion('23.1')]
    [RequiresVmsConnection()]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Name', Position = 0)]
        [Alias('DisplayName')]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [string]
        $Path,

        # Any unrecognized parameters and their values will be ignored when splatting a hashtable with keys that do not match a parameter name.
        [Parameter(ValueFromRemainingArguments, DontShow)]
        [object[]]
        $ExtraParams
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $folder = Get-ConfigurationItem -Path 'MIPKind[b9a5bc9c-e9a5-4a15-8453-ffa41f2815ac]/MIPItemFolder'
        $invokeInfo = $folder | Invoke-Method -MethodId RemoveMIPItem
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $valueTypeInfo = $invokeInfo.Properties[0].ValueTypeInfos | Where-Object Name -EQ $Name
            if ($null -eq $valueTypeInfo) {
                Write-Error -Message "Webhook with name '$Name' not found." -TargetObject $Name
                return
            }
            if ($valueTypeInfo.Count -gt 1) {
                Write-Error -Message "Multiple webhooks found with name '$Name'. To remove a specific webhook, use 'Get-VmsWebhook -Name ''$Name'' | Remove-VmsWebhook'." -TargetObject $Name
                return
            }
            $Path = $valueTypeInfo.Value
        } else {
            $Name = ($invokeInfo.Properties[0].ValueTypeInfos | Where-Object Value -EQ $Path).Name
        }
        
        $action = 'Remove webhook {0}' -f $Name
        if ($PSCmdlet.ShouldProcess((Get-VmsSite).Name, $action)) {
            $invokeInfo | Set-ConfigurationItemProperty -Key ItemSelection -Value $Path
            $null = $invokeInfo | Invoke-Method -MethodId RemoveMIPItem
        }
    }
}

Register-ArgumentCompleter -CommandName Remove-VmsWebhook -ParameterName Name -ScriptBlock {
    $values = (Get-VmsWebhook).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


function Set-VmsWebhook {
    [CmdletBinding(DefaultParameterSetName = 'Path', SupportsShouldProcess)]
    [OutputType([MilestonePSTools.Webhook])]
    [RequiresVmsConnection()]
    [RequiresVmsVersion('23.1')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Name', Position = 0)]
        [Alias('DisplayName')]
        [string]
        $Name,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Path')]
        [string]
        $Path,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $NewName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [uri]
        $Address,

        [Parameter(ValueFromPipelineByPropertyName)]
        [AllowEmptyString()]
        [AllowNull()]
        [string]
        $Token,

        [Parameter()]
        [switch]
        $PassThru,

        # Any unrecognized parameters and their values will be ignored when splatting a hashtable with keys that do not match a parameter name.
        [Parameter(ValueFromRemainingArguments, DontShow)]
        [object[]]
        $ExtraParams
    )

    begin {
        Assert-VmsRequirementsMet
    }

    process {
        $folder = Get-ConfigurationItem -Path 'MIPKind[b9a5bc9c-e9a5-4a15-8453-ffa41f2815ac]/MIPItemFolder'
        $invokeInfo = $folder | Invoke-Method -MethodId RemoveMIPItem
        if ([string]::IsNullOrWhiteSpace($Path)) {
            $valueTypeInfo = $invokeInfo.Properties[0].ValueTypeInfos | Where-Object Name -EQ $Name
            if ($null -eq $valueTypeInfo) {
                Write-Error -Message "Webhook with name '$Name' not found." -TargetObject $Name
                return
            }
            if ($valueTypeInfo.Count -gt 1) {
                Write-Error -Message "Multiple webhooks found with name '$Name'. Use 'Get-VmsWebhook' to find the one to update, and pipe it to Set-VmsWebhook instead to use the Path parameter rather than the Name." -TargetObject $Name
                return
            }
            $Path = $valueTypeInfo.Value
        }

        $webhook = Get-ConfigurationItem -Path $Path -ErrorAction Stop
        $dirty = $false
        'NewName', 'Address', 'Token' | ForEach-Object {
            if (-not $PSCmdlet.MyInvocation.BoundParameters.ContainsKey($_)) {
                return
            }
            $key = $_ -replace 'New', ''
            $property = $webhook.Properties | Where-Object Key -EQ $key
            if ($null -eq $property) {
                $dirty = $false
                throw "Property with key '$key' not found."
            }
            $currentValue = $property.Value
            $newValue = (Get-Variable -Name $_).Value
            if ($currentValue -cne $newValue) {
                Write-Verbose "Changing $key from '$currentValue' to '$newValue' on webhook '$($webhook.DisplayName)'"
                $dirty = $true
                $property.Value = $newValue
            }
        }

        $action = 'Update webhook {0}' -f $webhook.DisplayName
        if ($dirty -and $PSCmdlet.ShouldProcess((Get-VmsSite).Name, $action)) {
            $null = $webhook | Set-ConfigurationItem -ErrorAction Stop
        }
        if ($PassThru) {
            $webhook | Get-VmsWebhook
        }
    }
}

Register-ArgumentCompleter -CommandName Set-VmsWebhook -ParameterName Name -ScriptBlock {
    $values = (Get-VmsWebhook).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}


<#
Functions in this module are written as independent PS1 files, and to improve module load time they
are "compiled" into this PSM1 file. If you're looking at this file prior to build, now you know how
all the functions will be loaded later. If you're looking at this file after build, now you know
why this file has so many lines :)
#>

#region Argument Completers
# The default place for argument completers is within the same .PS1 as the function
# but argument completers for C# cmdlets can be placed here if needed.

Register-ArgumentCompleter -CommandName Get-VmsSite, Select-VmsSite -ParameterName Name -ScriptBlock {
    $values = (Get-VmsSite -ListAvailable).Name | Sort-Object
    Complete-SimpleArgument -Arguments $args -ValueSet $values
}

Register-ArgumentCompleter -CommandName Start-Export -ParameterName Codec -ScriptBlock {
    $location = [environment]::CurrentDirectory
    try {
        Push-Location -Path $MipSdkPath
        [environment]::CurrentDirectory = $MipSdkPath
        $exporter = [VideoOS.Platform.Data.AVIExporter]::new()
        $values = $exporter.CodecList | Sort-Object
        Complete-SimpleArgument -Arguments $args -ValueSet $values
    } finally {
        [environment]::CurrentDirectory = $location
        Pop-Location
        if ($exporter) {
            $exporter.Close()
        }
    }
}


#endregion

# Enable the use of any TLS protocol version greater than or equal to TLS 1.2
$protocol = [Net.SecurityProtocolType]::SystemDefault
[enum]::GetNames([Net.SecurityProtocolType]) | Where-Object {
    # Match any TLS version greater than 1.1
            ($_ -match 'Tls(\d)(\d+)?') -and ([version]("$($Matches[1]).$([int]$Matches[2])")) -gt 1.1
} | ForEach-Object { $protocol = $protocol -bor [Net.SecurityProtocolType]::$_ }
[Net.ServicePointManager]::SecurityProtocol = $protocol

$script:Deprecations = Import-PowerShellDataFile -Path "$PSScriptRoot\deprecations.psd1"
$script:Messages = @{}
Import-LocalizedData -BindingVariable 'script:Messages' -FileName 'messages'
Export-ModuleMember -Cmdlet * -Alias * -Function *

if ((Get-VmsModuleConfig).Mip.ConfigurationApiManager.UseRestApiWhenAvailable) {
    Write-Warning @'

Experimental Feature: UseRestApiWhenAvailable
MilestonePSTools is configured to use the API Gateway REST API when available. Some features may not yet be implemented in the API Gateway.
If you experience unexpected errors, try disabling this behavior with the following commands:

  $config = Get-VmsModuleConfig
  $config.Mip.ConfigurationApiManager.UseRestApiWhenAvailable = $false
  $config | Set-VmsModuleConfig

'@
}

if ((Get-VmsModuleConfig).ApplicationInsights.Enabled -and -not [MilestonePSTools.Telemetry.AppInsightsTelemetry]::HasDisplayedTelemetryNotice) {
    $null = New-Item -ItemType Directory -Path ([MilestonePSTools.Module]::AppDataDirectory) -Force -ErrorAction Ignore
    (Get-Date).ToUniversalTime().ToString('o') | Set-Content -Path (Join-Path ([MilestonePSTools.Module]::AppDataDirectory) "telemetry_notice_displayed.txt") -ErrorAction Ignore
    Write-Host @'
MilestonePSTools may send telemetry data using Azure Application Insights. This
data is anonymous and helps us to prioritize new features, fixes, and
performance improvements.

You may opt-out using the command `Set-VmsModuleConfig -EnableTelemetry $false`

Read more at https://www.milestonepstools.com/commands/en-US/about_Telemetry/
'@
}



# SIG # Begin signature block
# MIIuzgYJKoZIhvcNAQcCoIIuvzCCLrsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCW/x0XeLKHnYK
# Qo46MQnMwkN3oBKy50MJ4/oOC6Rk7KCCE24wggVyMIIDWqADAgECAhB2U/6sdUZI
# k/Xl10pIOk74MA0GCSqGSIb3DQEBDAUAMFMxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENvZGUgU2ln
# bmluZyBSb290IFI0NTAeFw0yMDAzMTgwMDAwMDBaFw00NTAzMTgwMDAwMDBaMFMx
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQD
# EyBHbG9iYWxTaWduIENvZGUgU2lnbmluZyBSb290IFI0NTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALYtxTDdeuirkD0DcrA6S5kWYbLl/6VnHTcc5X7s
# k4OqhPWjQ5uYRYq4Y1ddmwCIBCXp+GiSS4LYS8lKA/Oof2qPimEnvaFE0P31PyLC
# o0+RjbMFsiiCkV37WYgFC5cGwpj4LKczJO5QOkHM8KCwex1N0qhYOJbp3/kbkbuL
# ECzSx0Mdogl0oYCve+YzCgxZa4689Ktal3t/rlX7hPCA/oRM1+K6vcR1oW+9YRB0
# RLKYB+J0q/9o3GwmPukf5eAEh60w0wyNA3xVuBZwXCR4ICXrZ2eIq7pONJhrcBHe
# OMrUvqHAnOHfHgIB2DvhZ0OEts/8dLcvhKO/ugk3PWdssUVcGWGrQYP1rB3rdw1G
# R3POv72Vle2dK4gQ/vpY6KdX4bPPqFrpByWbEsSegHI9k9yMlN87ROYmgPzSwwPw
# jAzSRdYu54+YnuYE7kJuZ35CFnFi5wT5YMZkobacgSFOK8ZtaJSGxpl0c2cxepHy
# 1Ix5bnymu35Gb03FhRIrz5oiRAiohTfOB2FXBhcSJMDEMXOhmDVXR34QOkXZLaRR
# kJipoAc3xGUaqhxrFnf3p5fsPxkwmW8x++pAsufSxPrJ0PBQdnRZ+o1tFzK++Ol+
# A/Tnh3Wa1EqRLIUDEwIrQoDyiWo2z8hMoM6e+MuNrRan097VmxinxpI68YJj8S4O
# JGTfAgMBAAGjQjBAMA4GA1UdDwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBQfAL9GgAr8eDm3pbRD2VZQu86WOzANBgkqhkiG9w0BAQwFAAOCAgEA
# Xiu6dJc0RF92SChAhJPuAW7pobPWgCXme+S8CZE9D/x2rdfUMCC7j2DQkdYc8pzv
# eBorlDICwSSWUlIC0PPR/PKbOW6Z4R+OQ0F9mh5byV2ahPwm5ofzdHImraQb2T07
# alKgPAkeLx57szO0Rcf3rLGvk2Ctdq64shV464Nq6//bRqsk5e4C+pAfWcAvXda3
# XaRcELdyU/hBTsz6eBolSsr+hWJDYcO0N6qB0vTWOg+9jVl+MEfeK2vnIVAzX9Rn
# m9S4Z588J5kD/4VDjnMSyiDN6GHVsWbcF9Y5bQ/bzyM3oYKJThxrP9agzaoHnT5C
# JqrXDO76R78aUn7RdYHTyYpiF21PiKAhoCY+r23ZYjAf6Zgorm6N1Y5McmaTgI0q
# 41XHYGeQQlZcIlEPs9xOOe5N3dkdeBBUO27Ql28DtR6yI3PGErKaZND8lYUkqP/f
# obDckUCu3wkzq7ndkrfxzJF0O2nrZ5cbkL/nx6BvcbtXv7ePWu16QGoWzYCELS/h
# AtQklEOzFfwMKxv9cW/8y7x1Fzpeg9LJsy8b1ZyNf1T+fn7kVqOHp53hWVKUQY9t
# W76GlZr/GnbdQNJRSnC0HzNjI3c/7CceWeQIh+00gkoPP/6gHcH1Z3NFhnj0qinp
# J4fGGdvGExTDOUmHTaCX4GUT9Z13Vunas1jHOvLAzYIwggbmMIIEzqADAgECAhB3
# vQ4DobcI+FSrBnIQ2QRHMA0GCSqGSIb3DQEBCwUAMFMxCzAJBgNVBAYTAkJFMRkw
# FwYDVQQKExBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDEyBHbG9iYWxTaWduIENv
# ZGUgU2lnbmluZyBSb290IFI0NTAeFw0yMDA3MjgwMDAwMDBaFw0zMDA3MjgwMDAw
# MDBaMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8w
# LQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANZCTfnjT8Yj9GwdgaYw90g9
# z9DljeUgIpYHRDVdBs8PHXBg5iZU+lMjYAKoXwIC947Jbj2peAW9jvVPGSSZfM8R
# Fpsfe2vSo3toZXer2LEsP9NyBjJcW6xQZywlTVYGNvzBYkx9fYYWlZpdVLpQ0LB/
# okQZ6dZubD4Twp8R1F80W1FoMWMK+FvQ3rpZXzGviWg4QD4I6FNnTmO2IY7v3Y2F
# QVWeHLw33JWgxHGnHxulSW4KIFl+iaNYFZcAJWnf3sJqUGVOU/troZ8YHooOX1Re
# veBbz/IMBNLeCKEQJvey83ouwo6WwT/Opdr0WSiMN2WhMZYLjqR2dxVJhGaCJedD
# CndSsZlRQv+hst2c0twY2cGGqUAdQZdihryo/6LHYxcG/WZ6NpQBIIl4H5D0e6lS
# TmpPVAYqgK+ex1BC+mUK4wH0sW6sDqjjgRmoOMieAyiGpHSnR5V+cloqexVqHMRp
# 5rC+QBmZy9J9VU4inBDgoVvDsy56i8Te8UsfjCh5MEV/bBO2PSz/LUqKKuwoDy3K
# 1JyYikptWjYsL9+6y+JBSgh3GIitNWGUEvOkcuvuNp6nUSeRPPeiGsz8h+WX4VGH
# aekizIPAtw9FbAfhQ0/UjErOz2OxtaQQevkNDCiwazT+IWgnb+z4+iaEW3VCzYkm
# eVmda6tjcWKQJQ0IIPH/AgMBAAGjggGuMIIBqjAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU
# 2rONwCSQo2t30wygWd0hZ2R2C3gwHwYDVR0jBBgwFoAUHwC/RoAK/Hg5t6W0Q9lW
# ULvOljswgZMGCCsGAQUFBwEBBIGGMIGDMDkGCCsGAQUFBzABhi1odHRwOi8vb2Nz
# cC5nbG9iYWxzaWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUwRgYIKwYBBQUHMAKG
# Omh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2NvZGVzaWduaW5n
# cm9vdHI0NS5jcnQwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5nbG9iYWxz
# aWduLmNvbS9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3JsMFYGA1UdIARPME0wQQYJKwYB
# BAGgMgEyMDQwMgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29t
# L3JlcG9zaXRvcnkvMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEACIhyJsav
# +qxfBsCqjJDa0LLAopf/bhMyFlT9PvQwEZ+PmPmbUt3yohbu2XiVppp8YbgEtfjr
# y/RhETP2ZSW3EUKL2Glux/+VtIFDqX6uv4LWTcwRo4NxahBeGQWn52x/VvSoXMNO
# Ca1Za7j5fqUuuPzeDsKg+7AE1BMbxyepuaotMTvPRkyd60zsvC6c8YejfzhpX0FA
# Z/ZTfepB7449+6nUEThG3zzr9s0ivRPN8OHm5TOgvjzkeNUbzCDyMHOwIhz2hNab
# XAAC4ShSS/8SS0Dq7rAaBgaehObn8NuERvtz2StCtslXNMcWwKbrIbmqDvf+28rr
# vBfLuGfr4z5P26mUhmRVyQkKwNkEcUoRS1pkw7x4eK1MRyZlB5nVzTZgoTNTs/Z7
# KtWJQDxxpav4mVn945uSS90FvQsMeAYrz1PYvRKaWyeGhT+RvuB4gHNU36cdZytq
# tq5NiYAkCFJwUPMB/0SuL5rg4UkI4eFb1zjRngqKnZQnm8qjudviNmrjb7lYYuA2
# eDYB+sGniXomU6Ncu9Ky64rLYwgv/h7zViniNZvY/+mlvW1LWSyJLC9Su7UpkNpD
# R7xy3bzZv4DB3LCrtEsdWDY3ZOub4YUXmimi/eYI0pL/oPh84emn0TCOXyZQK8ei
# 4pd3iu/YTT4m65lAYPM8Zwy2CHIpNVOBNNwwggcKMIIE8qADAgECAgxwaYJwgzTu
# /lOAy3UwDQYJKoZIhvcNAQELBQAwWTELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEds
# b2JhbFNpZ24gbnYtc2ExLzAtBgNVBAMTJkdsb2JhbFNpZ24gR0NDIFI0NSBDb2Rl
# U2lnbmluZyBDQSAyMDIwMB4XDTI1MDEyOTE1NDU0MloXDTI2MDMwMTE1NDU0Mlow
# eDELMAkGA1UEBhMCVVMxDzANBgNVBAgTBk9yZWdvbjEUMBIGA1UEBxMLTGFrZSBP
# c3dlZ28xIDAeBgNVBAoTF01pbGVzdG9uZSBTeXN0ZW1zLCBJbmMuMSAwHgYDVQQD
# ExdNaWxlc3RvbmUgU3lzdGVtcywgSW5jLjCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAIfjLyk2i3fyML6J7+LzzwNUZc9/22fhwSDPtkRfLjy8yYgTxbSu
# TgJi/twyUkts2qjbnGqD/maukegK3uZZHHOV2dWL1fjTOa4Mc0Vw3jBELmvrW/CO
# TSvPQMS3u2WVGGJBqVu+IVwWkC2dc2qnNeN4K1VXr4TqCjE0st/K2cJdRbEavMpz
# BCqYdfA41feH7s22C3C0JE5aXVuCaeXbJB0uBfH4fcLn/Wt3wQuT6sMsjSpFj5gH
# iwt4PPP4jWt12kB1LzOzMVUMNfIXrfECqEdRN7lfyZPWTsEiiamjXZEmsxH7Q0rn
# QVvfeh12vf5SrUZFKyi5Gj98AkrWJIciDohd3zRxB/tdSTKF8cVkYwT80/6IOyK1
# y8WaBEO6myXNixx8VDZp20teUSxQex0WgTNQ0/raNWT21ZbKe1SHg/zziHaize6K
# ZM/TikiFz9ibppZR87/lxWw8Oy8eMckjWNDJYitpELcJ78PkBQGISOtGDbH6JbNW
# 2gJ/idTb6hasv/Dvu9QIbyOu+gTm3pJ2Ot5E0nlLGEqem3F2VBoOPLfgsmXfw0S0
# 8IiYGm9QTPy5T15rJNhQHnX0Q7WgPLaictxsOvO5vHXlKln6MRv4dYSYAWUx5Kgy
# 6cptfVXqH1MSf66781BjX3W7a5ySps2O8eWzd/H80tsNF64drkn4kAfdAgMBAAGj
# ggGxMIIBrTAOBgNVHQ8BAf8EBAMCB4AwgZsGCCsGAQUFBwEBBIGOMIGLMEoGCCsG
# AQUFBzAChj5odHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc2dj
# Y3I0NWNvZGVzaWduY2EyMDIwLmNydDA9BggrBgEFBQcwAYYxaHR0cDovL29jc3Au
# Z2xvYmFsc2lnbi5jb20vZ3NnY2NyNDVjb2Rlc2lnbmNhMjAyMDBWBgNVHSAETzBN
# MEEGCSsGAQQBoDIBMjA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxz
# aWduLmNvbS9yZXBvc2l0b3J5LzAIBgZngQwBBAEwCQYDVR0TBAIwADBFBgNVHR8E
# PjA8MDqgOKA2hjRodHRwOi8vY3JsLmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29k
# ZXNpZ25jYTIwMjAuY3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB8GA1UdIwQYMBaA
# FNqzjcAkkKNrd9MMoFndIWdkdgt4MB0GA1UdDgQWBBQOXMn2R0Y1ZfbEfoeRqdTg
# piHkQjANBgkqhkiG9w0BAQsFAAOCAgEAat2qjMH8QzC3NJg+KU5VmeEymyAxDaqX
# VUrlYjMR9JzvTqcKrXvhTO8SwM37TW4UR7kmG1TKdVkKvbBAThjEPY5xH+TMgxSQ
# TrOTmBPx2N2bMUoRqNCuf4RmJXGk2Qma26aU/gSKqC2hq0+fstuRQjINUVTp00VS
# 9XlAK0zcXFrZrVREaAr9p4U606oYWT5oen7Bi1M7L8hbbZojw9N4WwuH6n0Ctwlw
# ZUHDbYzXWnKALkPWNHZcqZX4zHFDGKzn/wgu8VDkhaPrmn9lRW+BDyI/EE0iClGJ
# KXMidK71y7BT+DLGBvga6xfbInj5e+n7Nu9gU8D9RkjQqdoq7mO/sXUIsHMJQPkD
# qnyy20PsMYlEbuRTJ8v9c7HN+oWlttEq4iJ24EerC/1Lamr55L/F9rcJ3XRYLKLR
# E5K9pOsq7oJpbsOuYfpv61WISX3ewy5v1tY9VHjn7NQxoMgczSuAf97qbQNpblt0
# h+9KTiwWmLnw1jP1/vwNBYwZk3mtL0Z7Z/l2qqVawrT2W3/EwovP0DWcQr9idTAI
# WLbnWRUHlUZv4rCeoIwXWGgCUOF+BHU1sacps1V3kK1OpNvRWYs+mk2tzGyxoIEB
# 9whM6Vxzik4Q7ciXG7G2ZzYR9f2J4kbr4hTxZEB6ysCPT8DTdqmUwUcxDc2i2j2p
# AkA8F4fhc3Yxghq2MIIasgIBATBpMFkxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZHbG9iYWxTaWduIEdDQyBSNDUgQ29k
# ZVNpZ25pbmcgQ0EgMjAyMAIMcGmCcIM07v5TgMt1MA0GCWCGSAFlAwQCAQUAoIGk
# MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgor
# BgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC9JhpLTa3W5m8tliNmuFmIqqXBjIC/
# hjqzVd9ZTZrXwjA4BgorBgEEAYI3AgEMMSowKKACgAChIoAgaHR0cHM6Ly93d3cu
# bWlsZXN0b25lcHN0b29scy5jb20wDQYJKoZIhvcNAQEBBQAEggIAhmCJEPkF2TZb
# DFm+Fqs3W30m5h1XZSTvWOhvHdpw7nF6lFqE9q+PSuz7t/D6/ASFGg8oGPPhrTjT
# 9c+gEWpeq7s4zv9Q9AJCA9wSMW9Arv9yWHRR2ct248QTVmKaxcE5CyrzF4Q1L0NA
# iTU7VdujHlxqMeC/j+Jp5b9rgvZ3ZJlfQHhwelgV8LsH9tRRDMUEVj63iKEiNYFY
# FSI6BpTVSxggxPEPy7flwyhmmxhKOduSqgVMThJylqhM8u/lNJZSrjOZUV/ljiEn
# y89HwaMAd9PbZy0EL4WRCWd+aD15RAYxIDzE1BKir08YLuCz+OJAcW3IyuKOg0vZ
# HrCrCI5gItzBcyJ0IDV/2M+Sv/jT85+mZsIpIWx7lkdDNrKvKoCehQcPoycncrSf
# pwzknw+eQz5/34q80J57FXPOL6MqVLaozkE9t1xuccFsgb5TsaRUl8yw7V/VoIcg
# 1pbO6TWfIj7qgvkCzYeJNc0P1LIBsdvrgva9lPe2NOXBIcQ76a6b0zqUIIgySF+u
# /NpbHkLyuIP4+Frd/H91g1wwXVyX5ZQhtmL6oE8yFeZk+RxnENQIQUxCr/FE56sx
# oNTtFMBcQtDB4qzEnmv5hhIvL0UtvYs/Xov95Qs8W4aVdiaI04INKg4jsi+Ue/zm
# BCFZvJfvHKv/a25IuFGC//f1LtRJ7tChghd3MIIXcwYKKwYBBAGCNwMDATGCF2Mw
# ghdfBgkqhkiG9w0BBwKgghdQMIIXTAIBAzEPMA0GCWCGSAFlAwQCAQUAMHgGCyqG
# SIb3DQEJEAEEoGkEZzBlAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# Td6XpsgGXiXeaMdI8b60rusep+rb88eGvVnxjOYZIUACEQD4GAU78ioHnhiK7xOR
# hk8SGA8yMDI1MDYxODE4NTczMlqgghM6MIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC
# 0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMO
# RGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGlt
# ZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAw
# MFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lD
# ZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1l
# c3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA
# 69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6w
# W2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00
# Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOM
# A3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmot
# uQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1O
# pbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeH
# VZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1r
# oSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSURO
# wnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW0
# 0aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGV
# MIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM
# 6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMC
# B4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKG
# UWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRp
# bWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSg
# UqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRU
# aW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAI
# BgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcE
# ua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/Ym
# RDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8
# AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/E
# ABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQ
# VTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gV
# utDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85
# EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hg
# gt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJ
# gKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLv
# UxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7P
# OGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCBrQwggScoAMCAQICEA3HrFcF/yGZ
# LkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UE
# AxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4
# MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBS
# U0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81
# FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUil
# jNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9
# +0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8NF5v
# p7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUUFREm
# DrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8
# SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCy
# QEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespYMQmU
# iote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrPV6/7
# umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+zoFp
# p4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFd
# MIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW
# 9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8B
# Af8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290
# RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQC
# MAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZxML2+
# C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97frPBtI
# j+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+NMvE
# uBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYAgwPy
# WLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSj
# jf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+BFo+
# z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06VXxy
# KkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284NHNb
# oDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDezooIs
# 8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7W
# P/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpSM9LH
# JmyrxaFtoza2zNaQ9k+5t1wwggWNMIIEdaADAgECAhAOmxiO+dAt5+/bUOIIQBha
# MA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lD
# ZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0zMTExMDky
# MzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo
# 3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutW
# xpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQ
# RBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nP
# zaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6P
# gNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEB
# fCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3
# wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2
# mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2
# Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF1
# 3nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+d
# IPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAPBgNVHRMB
# Af8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYweQYIKwYB
# BQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDARBgNV
# HSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4X6CjdBs9
# thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc47/T/gLn4
# offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0mcudT6cG
# AxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk1T3osdz9
# HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLBsln1VWvP
# J6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGpn1eqXiji
# uZQxggN8MIIDeAIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJ
# YIZIAWUDBAIBBQCggdEwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwGCSqG
# SIb3DQEJBTEPFw0yNTA2MTgxODU3MzJaMCsGCyqGSIb3DQEJEAIMMRwwGjAYMBYE
# FN1iMKyGCi0wa9o4sWh5UjAH+0F+MC8GCSqGSIb3DQEJBDEiBCDPn2pmnqc3UM/C
# qFo5a+0n3A9DY3ebty4fewutkGxnvTA3BgsqhkiG9w0BCRACLzEoMCYwJDAiBCBK
# oD+iLNdchMVck4+CjmdrnK7Ksz/jbSaaozTxRhEKMzANBgkqhkiG9w0BAQEFAASC
# AgBAEdR1NDOW4CZUmwqUGlfTqX/Ysx74txdKrLOtRs0AOWZ42sS3ATUVTffFnjzq
# 9uo+d5U8Cpjsjj5uTzGZxUPRgC4ExcYXn7e/Qo78sJhz/MR6zZZTezew4Spqeikw
# cJMp98aE6aj8S3xQwdx1r+UAlfD5HmWWPPIOrdA4vJQQomInrxWSrWMM68XjQvrP
# jYT2TDmQMbd6brxCOtXUPI1mywEKws/xNBS70T5xmeBnEWPAr0V8qE30Rdtrr6Wr
# dQWz/VcsYbqUjKfTLT0O8xxFHnj7egGIstskPgrB3Xj4LydpL4tSae3/PghdSTLH
# BStfcX4J0tAxCEzh+jdi5SlbO/KEvO8VX9su2HbbjE4pzHzdKJLJenpy3DW/pQEk
# O1DfT9Lcp1t1IK1si4TEdXYCBhIITHf9fkMBnsziyYZwn/FvM8brmaEPAGeEXrNS
# KAQM88uLks1P+ra+vH+zgvpzRMCRkbe7UyQztpj5uim1qbHNEHweTOFDwc/GQdF4
# 89TjIxpBnIE2jCG7Cl8EA70AEC52TgqU4ARkzhEfm/upNU/gU8SwABYew6tBmCi9
# Pn7pAEJNBhvoYT64NQ0C7ndXFLl0Nc3gjeCLpNeJ9j8cSLZTlAZBQWGRlSY5a9lo
# 3HnhE1S1FR/RUrWpXT3oqLWVPBn4FOBdwZLxC0v92Z48Rg==
# SIG # End signature block
