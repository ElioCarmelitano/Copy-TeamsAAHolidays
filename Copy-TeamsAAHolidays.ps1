<#
    .SYNOPSIS
    Copies Teams Auto Attendant (AA) Holiday configuration (holidays + actions) from one AA to one or more other AAs.

    .DESCRIPTION
    This function duplicates the "Holiday" behaviour from a source Auto Attendant to one or more target Auto Attendants by copying:

      1) Holiday Call Handling Associations (Type = Holiday)
         - These determine which call flow is executed when a holiday schedule is in effect.

      2) Call Flows referenced by those Holiday associations
         - Call flows contain the greetings, menus, menu options and actions (transfer targets, disconnect, etc.).
         - Call flows are CLONED to each target (new CallFlow IDs), so holiday actions can diverge per AA later.

      3) Schedules referenced by those Holiday associations
         - IMPORTANT: This function REUSES the existing Schedule IDs from the source.
           Teams treats schedules as service-managed objects; inventing new schedule GUIDs will fail.
         - Reusing schedule IDs means the same holiday schedule can be shared across multiple AAs.

    Resulting behaviour:
      - Holiday dates/times (the schedule) are SHARED between source and targets (same ScheduleId).
      - Holiday actions (call flows) are COPIED per target (new CallFlowId), so actions can be edited independently.

    .REQUIREMENTS
    - MicrosoftTeams PowerShell module installed.
    - Connected session to Teams PowerShell:
        Connect-MicrosoftTeams
    - Account running the function must have permissions to read and modify Auto Attendants.
      (Typically Teams Administrator / Teams Communications Administrator, or equivalent delegated rights.)

    .PREREQUISITES / ASSUMPTIONS
    - Source AA must already contain at least one Holiday call handling association.
    - Target AAs exist in the same tenant.
    - If -Interactive is used:
        - Out-GridView must be available (typically Windows PowerShell / Windows desktop with ISE/RSAT components).

    .SAFETY / CHANGE CONTROL
    - Supports -WhatIf and -Confirm via SupportsShouldProcess.
      Example:
        Copy-TeamsAAHolidays -Interactive -WhatIf

    - Overwrite behaviour:
      By default (-OverwriteTargetHolidays = $true), the function will:
        - Remove existing Holiday call handling associations from each target.
        - Remove any existing target call flows whose names match the source holiday call flow names
          (prevents duplicate call flows when re-running).
      It will NOT remove:
        - Default call flow
        - AfterHours associations/call flows
        - Non-Holiday associations/call flows

    .LIMITATIONS
    - Because schedules are reused, editing the holiday schedule dates later will affect ALL AAs sharing that schedule.
      If you require fully independent schedules per AA, you must create new schedules using New-CsOnlineSchedule
      and build new Holiday associations referencing those new schedule IDs.

    .PARAMETER SourceAAName
    Source Auto Attendant Name (exact match). Used when not running interactively.

    .PARAMETER TargetAANames
    One or more target Auto Attendant Names (exact match). Used when not running interactively.

    .PARAMETER Interactive
    When specified, displays Out-GridView pickers to select:
      - One Source AA
      - One or more Target AAs

    .PARAMETER OverwriteTargetHolidays
    When true (default), overwrites Holiday config on the target:
      - Removes Holiday call handling associations
      - Removes call flows with names matching the source holiday call flows (dedupe)

    .EXAMPLE
    # Interactive dry run (no changes)
    Copy-TeamsAAHolidays -Interactive -WhatIf

    .EXAMPLE
    # Copy holidays from "Test1" to "Test2" and "Test3"
    Copy-TeamsAAHolidays -SourceAAName "Test1" -TargetAANames @("Test2","Test3")

    #>

function Copy-TeamsAAHolidays {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        # Use either names OR interactive picker
        [Parameter(Mandatory = $false)]
        [string]$SourceAAName,

        [Parameter(Mandatory = $false)]
        [string[]]$TargetAANames,

        # If set, uses Out-GridView to pick Source + Targets
        [switch]$Interactive,

        # Overwrite target holiday associations/callflows (recommended)
        [switch]$OverwriteTargetHolidays = $true
    )

    $ErrorActionPreference = "Stop"

    # ---- Get all auto attendants once ----
    $allAAs = @(Get-CsAutoAttendant)

    if (-not $allAAs -or $allAAs.Count -eq 0) {
        throw "No Auto Attendants returned. Are you connected to Teams PowerShell?"
    }

    # ---- Interactive selection ----
    if ($Interactive) {

        if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
            throw "Out-GridView is not available on this machine. Install the needed components or run without -Interactive."
        }

        $sourcePick = $allAAs |
            Select-Object Name, Identity |
            Out-GridView -Title "Select SOURCE Auto Attendant (single selection)" -PassThru

        if (-not $sourcePick) { throw "No source Auto Attendant selected." }
        $srcId = $sourcePick.Identity

        $targetPick = $allAAs |
            Where-Object { $_.Identity -ne $srcId } |
            Select-Object Name, Identity |
            Out-GridView -Title "Select TARGET Auto Attendant(s) (multi-select)" -PassThru

        if (-not $targetPick) { throw "No target Auto Attendant(s) selected." }
        $tgtIds = @($targetPick.Identity)
    }
    else {
        if ([string]::IsNullOrWhiteSpace($SourceAAName)) { throw "Provide -SourceAAName or use -Interactive." }
        if (-not $TargetAANames -or $TargetAANames.Count -eq 0) { throw "Provide -TargetAANames or use -Interactive." }

        $src = $allAAs | Where-Object { $_.Name -eq $SourceAAName } | Select-Object -First 1
        if (-not $src) { throw "Source AA with Name '$SourceAAName' not found (name match is case-insensitive but must be exact)." }
        $srcId = $src.Identity

        $tgtIds = foreach ($n in $TargetAANames) {
            $t = $allAAs | Where-Object { $_.Name -eq $n } | Select-Object -First 1
            if (-not $t) { throw "Target AA with Name '$n' not found (name match must be exact)." }
            $t.Identity
        }
    }

    # ---- Load source AA details ----
    $srcAA = Get-CsAutoAttendant -Identity $srcId

    # Holiday associations only
    $srcHolidayAssocs = @($srcAA.CallHandlingAssociations | Where-Object { $_.Type -eq "Holiday" })
    if ($srcHolidayAssocs.Count -eq 0) { throw "No Holiday associations found on source AA '$($srcAA.Name)'." }

    # Lookups
    $srcSchedulesById = @{}
    foreach ($s in @($srcAA.Schedules)) { $srcSchedulesById[[string]$s.Id] = $s }

    $srcCallFlowsById = @{}
    foreach ($cf in @($srcAA.CallFlows)) { $srcCallFlowsById[[string]$cf.Id] = $cf }

    $srcHolidayScheduleIds = $srcHolidayAssocs | ForEach-Object { [string]$_.ScheduleId } | Select-Object -Unique
    $srcHolidayCallFlowIds  = $srcHolidayAssocs | ForEach-Object { [string]$_.CallFlowId } | Select-Object -Unique

    foreach ($sid in $srcHolidayScheduleIds) {
        if (-not $srcSchedulesById.ContainsKey($sid)) { throw "Holiday ScheduleId not found in source Schedules: $sid" }
    }
    foreach ($cid in $srcHolidayCallFlowIds) {
        if (-not $srcCallFlowsById.ContainsKey($cid)) { throw "Holiday CallFlowId not found in source CallFlows: $cid" }
    }

    # Names of holiday call flows (used to dedupe target call flows on overwrite)
    $srcHolidayCallFlowNames = $srcHolidayCallFlowIds | ForEach-Object { $srcCallFlowsById[$_].Name }

    # ---- Apply to each target ----
    foreach ($tid in $tgtIds) {

        $tgtAA = Get-CsAutoAttendant -Identity $tid

        if ($PSCmdlet.ShouldProcess($tgtAA.Name, "Copy Holiday schedules/call flows/associations from '$($srcAA.Name)'")) {

            if (-not $tgtAA.Schedules) { $tgtAA.Schedules = @() }
            if (-not $tgtAA.CallFlows) { $tgtAA.CallFlows = @() }
            if (-not $tgtAA.CallHandlingAssociations) { $tgtAA.CallHandlingAssociations = @() }

            # Overwrite holiday config on target (remove only Holiday associations)
            if ($OverwriteTargetHolidays) {
                $tgtAA.CallHandlingAssociations = @(
                    $tgtAA.CallHandlingAssociations | Where-Object { $_.Type -ne "Holiday" }
                )

                # Remove existing call flows with same names as source holiday call flows (prevents duplicates)
                $tgtAA.CallFlows = @(
                    $tgtAA.CallFlows | Where-Object { $srcHolidayCallFlowNames -notcontains $_.Name }
                )
            }

            # Ensure schedules exist on target (REUSE schedule IDs; do not invent new)
            $existingTgtScheduleIds = @($tgtAA.Schedules | ForEach-Object { [string]$_.Id })
            foreach ($sid in $srcHolidayScheduleIds) {
                if ($existingTgtScheduleIds -notcontains $sid) {
                    $tgtAA.Schedules += $srcSchedulesById[$sid]
                }
            }

            # Clone holiday call flows (new IDs), map source -> target
            $callFlowIdMap = @{}
            foreach ($cid in $srcHolidayCallFlowIds) {
                $newCf = $srcCallFlowsById[$cid] | ConvertTo-Json -Depth 50 | ConvertFrom-Json
                $newCf.Id = ([guid]::NewGuid().ToString())
                $tgtAA.CallFlows += $newCf
                $callFlowIdMap[$cid] = $newCf.Id
            }

            # Recreate holiday associations
            foreach ($assoc in $srcHolidayAssocs) {
                $oldCfId = [string]$assoc.CallFlowId
                $newCfId = $callFlowIdMap[$oldCfId]
                if ([string]::IsNullOrWhiteSpace($newCfId)) { throw "No mapped CallFlowId found for source CallFlowId: $oldCfId" }

                $tgtAA.CallHandlingAssociations += New-CsAutoAttendantCallHandlingAssociation `
                    -Type "Holiday" `
                    -ScheduleId ([string]$assoc.ScheduleId) `
                    -CallFlowId $newCfId
            }

            Set-CsAutoAttendant -Instance $tgtAA

            Write-Host "Updated '$($tgtAA.Name)' from source '$($srcAA.Name)'" -ForegroundColor Green
        }
    }
}