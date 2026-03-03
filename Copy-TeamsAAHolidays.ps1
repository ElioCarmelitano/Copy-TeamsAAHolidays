#Requires -Modules MicrosoftTeams

<#
.SYNOPSIS
    Copies Teams Auto Attendant (AA) Holiday configuration from a source AA to one or more target AAs.

.DESCRIPTION
    Duplicates Holiday call handling behaviour from a source Auto Attendant to target AAs by copying:
      1. Holiday Call Handling Associations (Type = Holiday)
      2. Call Flows referenced by those associations (cloned with new IDs per target)
      3. Schedules referenced by those associations (reused; same ScheduleId shared across AAs)

    Schedule dates/times are SHARED (same ScheduleId).
    Call flows are COPIED per target (new CallFlowId), allowing independent edits later.

.NOTES
    SCHEDULE SHARING : Editing holiday dates later will affect ALL AAs sharing that schedule.
                       Use New-CsOnlineSchedule to create independent schedules if needed.
    OVERWRITE        : By default, existing Holiday associations and matching call flows are
                       removed from each target before copying. Non-holiday config is never touched.

.REQUIREMENTS
    - MicrosoftTeams PowerShell module installed.
    - Active Teams PowerShell session: Connect-MicrosoftTeams
    - Teams Administrator / Teams Communications Administrator role (or equivalent).

.PARAMETER SourceAAName
    Source Auto Attendant name (exact match). Required when -Interactive is not used.

.PARAMETER TargetAANames
    One or more target Auto Attendant names (exact match). Required when -Interactive is not used.

.PARAMETER Interactive
    Uses Out-GridView pickers to select Source and Target AAs.
    Requires a Windows desktop environment with Out-GridView available.

.PARAMETER NoOverwrite
    When specified, preserves existing Holiday configuration on targets (no replacement).
    Default behaviour removes Holiday associations and matching call flows before copying.

.EXAMPLE
    # Interactive dry run - no changes made
    Copy-TeamsAAHolidays -Interactive -WhatIf

.EXAMPLE
    # Copy holidays from "Main AA" to two targets
    Copy-TeamsAAHolidays -SourceAAName "Main AA" -TargetAANames "Reception", "Support"

.EXAMPLE
    # Capture results for logging
    $results = Copy-TeamsAAHolidays -SourceAAName "Main AA" -TargetAANames "Reception"
    $results | Where-Object Status -eq 'Failed'
#>
function Copy-TeamsAAHolidays {
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Named')]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Named')]
        [ValidateNotNullOrEmpty()]
        [string]$SourceAAName,

        [Parameter(Mandatory, ParameterSetName = 'Named')]
        [ValidateNotNullOrEmpty()]
        [string[]]$TargetAANames,

        [Parameter(Mandatory, ParameterSetName = 'Interactive')]
        [switch]$Interactive,

        # Flip of the original $OverwriteTargetHolidays = $true;
        # absence of this switch means "do overwrite" which is the safe default.
        [switch]$NoOverwrite
    )

    $ErrorActionPreference = 'Stop'

    # ── Retrieve all Auto Attendants once ──────────────────────────────────────
    Write-Verbose "Retrieving all Auto Attendants..."
    $allAAs = @(Get-CsAutoAttendant)

    if ($allAAs.Count -eq 0) {
        throw "No Auto Attendants returned. Verify your session: Connect-MicrosoftTeams"
    }

    # ── Resolve source and target identities ───────────────────────────────────
    if ($PSCmdlet.ParameterSetName -eq 'Interactive') {

        if (-not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
            throw "Out-GridView is unavailable. Use -SourceAAName / -TargetAANames instead."
        }

        $sourcePick = $allAAs |
            Select-Object Name, Identity |
            Out-GridView -Title "Select SOURCE Auto Attendant (single selection)" -PassThru |
            Select-Object -First 1

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
        $src = $allAAs | Where-Object Name -eq $SourceAAName | Select-Object -First 1
        if (-not $src) { throw "Source AA '$SourceAAName' not found. Name must be an exact match." }
        $srcId = $src.Identity

        # foreach returns values; collected into array implicitly
        $tgtIds = foreach ($name in $TargetAANames) {
            $t = $allAAs | Where-Object Name -eq $name | Select-Object -First 1
            if (-not $t) { throw "Target AA '$name' not found. Name must be an exact match." }
            $t.Identity
        }
    }

    # ── Load and validate source AA ────────────────────────────────────────────
    Write-Verbose "Loading source AA '$srcId'..."
    $srcAA = Get-CsAutoAttendant -Identity $srcId

    $srcHolidayAssocs = @($srcAA.CallHandlingAssociations | Where-Object Type -eq 'Holiday')
    if ($srcHolidayAssocs.Count -eq 0) {
        throw "Source AA '$($srcAA.Name)' has no Holiday associations to copy."
    }

    # Build lookup tables for fast O(1) access
    $srcScheduleMap = @{}
    foreach ($s  in @($srcAA.Schedules )) { $srcScheduleMap[[string]$s.Id]  = $s  }

    $srcCallFlowMap = @{}
    foreach ($cf in @($srcAA.CallFlows  )) { $srcCallFlowMap[[string]$cf.Id] = $cf }

    $srcHolidayScheduleIds = @($srcHolidayAssocs.ScheduleId | Select-Object -Unique)
    $srcHolidayCallFlowIds  = @($srcHolidayAssocs.CallFlowId  | Select-Object -Unique)

    # Validate all referenced IDs actually exist on the source
    foreach ($sid in $srcHolidayScheduleIds) {
        if (-not $srcScheduleMap.ContainsKey([string]$sid)) {
            throw "Holiday ScheduleId '$sid' not found in source AA schedules."
        }
    }
    foreach ($cid in $srcHolidayCallFlowIds) {
        if (-not $srcCallFlowMap.ContainsKey([string]$cid)) {
            throw "Holiday CallFlowId '$cid' not found in source AA call flows."
        }
    }

    # Used to deduplicate target call flows when overwriting
    $srcHolidayCallFlowNames = $srcHolidayCallFlowIds | ForEach-Object { $srcCallFlowMap[$_].Name }

    # ── Apply to each target ───────────────────────────────────────────────────
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i = 0

    foreach ($tid in $tgtIds) {
        $i++
        Write-Progress -Activity "Copying holiday configuration from '$($srcAA.Name)'" `
                       -Status   "Target $i of $($tgtIds.Count): $tid" `
                       -PercentComplete ([math]::Round(($i / $tgtIds.Count) * 100))

        try {
            $tgtAA = Get-CsAutoAttendant -Identity $tid

            if (-not $PSCmdlet.ShouldProcess($tgtAA.Name, "Copy Holiday config from '$($srcAA.Name)'")) {
                continue
            }

            # Ensure collections are never null
            if (-not $tgtAA.Schedules              ) { $tgtAA.Schedules                = @() }
            if (-not $tgtAA.CallFlows              ) { $tgtAA.CallFlows                = @() }
            if (-not $tgtAA.CallHandlingAssociations) { $tgtAA.CallHandlingAssociations = @() }

            # ── Strip existing Holiday-only config (safe; non-holiday config untouched) ──
            if (-not $NoOverwrite) {
                $tgtAA.CallHandlingAssociations = @(
                    $tgtAA.CallHandlingAssociations | Where-Object Type -ne 'Holiday'
                )
                # Remove call flows whose names match source holiday call flow names (dedupe on re-run)
                $tgtAA.CallFlows = @(
                    $tgtAA.CallFlows | Where-Object { $_.Name -notin $srcHolidayCallFlowNames }
                )
            }

            # ── Add missing schedules (reuse source IDs; new IDs will be rejected by Teams) ──
            # Use a HashSet for O(1) existence checks
            $existingScheduleIds = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]@($tgtAA.Schedules | ForEach-Object { [string]$_.Id })
            )
            foreach ($sid in $srcHolidayScheduleIds) {
                if (-not $existingScheduleIds.Contains([string]$sid)) {
                    Write-Verbose "  Adding schedule '$sid' to '$($tgtAA.Name)'"
                    $tgtAA.Schedules += $srcScheduleMap[$sid]
                }
            }

            # ── Clone call flows with new GUIDs (so targets can diverge independently later) ──
            $cfIdMap = @{}
            foreach ($cid in $srcHolidayCallFlowIds) {
                $clonedCf    = $srcCallFlowMap[$cid] | ConvertTo-Json -Depth 50 | ConvertFrom-Json
                $clonedCf.Id = [guid]::NewGuid().ToString()
                $tgtAA.CallFlows += $clonedCf
                $cfIdMap[[string]$cid] = $clonedCf.Id
                Write-Verbose "  Cloned call flow '$($srcCallFlowMap[$cid].Name)' -> $($clonedCf.Id)"
            }

            # ── Recreate Holiday associations pointing at cloned call flows ──
            foreach ($assoc in $srcHolidayAssocs) {
                $newCfId = $cfIdMap[[string]$assoc.CallFlowId]
                if ([string]::IsNullOrWhiteSpace($newCfId)) {
                    throw "No mapped CallFlowId for source CallFlowId '$($assoc.CallFlowId)'."
                }

                $tgtAA.CallHandlingAssociations += New-CsAutoAttendantCallHandlingAssociation `
                    -Type       Holiday `
                    -ScheduleId ([string]$assoc.ScheduleId) `
                    -CallFlowId $newCfId
            }

            Set-CsAutoAttendant -Instance $tgtAA

            $results.Add([PSCustomObject]@{ Target = $tgtAA.Name; Status = 'Success'; Error = $null })
            Write-Verbose "Successfully updated '$($tgtAA.Name)'"
        }
        catch {
            # Capture failure but continue with remaining targets
            $results.Add([PSCustomObject]@{ Target = $tid; Status = 'Failed'; Error = $_.Exception.Message })
            Write-Warning "Failed to update target '$tid': $_"
        }
    }

    Write-Progress -Activity "Copying holiday configuration" -Completed

    # ── Summary ────────────────────────────────────────────────────────────────
    $succeeded = ($results | Where-Object Status -eq 'Success').Count
    $failed    = ($results | Where-Object Status -eq 'Failed' ).Count

    Write-Host "`nHoliday copy complete — Source: '$($srcAA.Name)'" -ForegroundColor Cyan
    Write-Host "  Succeeded : $succeeded" -ForegroundColor Green

    if ($failed -gt 0) {
        Write-Host "  Failed    : $failed" -ForegroundColor Red
        $results | Where-Object Status -eq 'Failed' | ForEach-Object {
            Write-Host "    ✗ $($_.Target) — $($_.Error)" -ForegroundColor Red
        }
    }

    # Return structured results for pipeline / logging use
    return $results
}
