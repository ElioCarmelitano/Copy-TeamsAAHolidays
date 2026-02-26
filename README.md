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
