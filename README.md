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
