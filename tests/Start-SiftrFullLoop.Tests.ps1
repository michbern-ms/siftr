$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '..\scripts\Start-SiftrFullLoop.ps1'
$parseErrors = $null
$tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors) {
    throw ($parseErrors | ForEach-Object Message | Out-String)
}

function Import-LoopFunction {
    param([Parameter(Mandatory)][string]$Name)

    $functionAst = $ast.Find(
        {
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
        },
        $true
    )

    if (-not $functionAst) {
        throw "Function '$Name' not found in $scriptPath"
    }

    $definition = $functionAst.Extent.Text -replace ("^function\s+{0}\b" -f [regex]::Escape($Name)), ("function global:{0}" -f $Name)
    Invoke-Expression $definition
}

@(
    'Get-UtcDateTime',
    'Get-CurrentTimeZone',
    'Get-CurrentLocalDateTime',
    'Convert-UtcToCurrentLocalDateTime',
    'Convert-CurrentLocalToUtc',
    'Test-ContinuousMode',
    'Get-LoopRecentCutoffLocal',
    'Filter-LoopMessagesBySince',
    'Get-LoopBacklogSince',
    'Assert-SiftrInboxActionSummary',
    'Get-Addressing',
    'Get-ToCount',
    'Test-InternalSender',
    'Get-SenderRelation',
    'Test-ManagerIncluded',
    'Test-ThreadHasUserReply',
    'Test-MessageFromUser',
    'Test-TextMatch',
    'Get-SenderIdentity',
    'Test-ExplicitMention',
    'Test-DirectAsk',
    'Test-CompletionReply',
    'Test-DeadlineUrgent',
    'Test-AutomatedApproval',
    'Test-ExternalSpam',
    'Get-TierInfo',
    'Test-TrustedExternalContentSender',
    'Get-HeuristicDecision'
) | ForEach-Object { Import-LoopFunction -Name $_ }

function global:New-TestRecord {
    param(
        [string]$SenderName,
        [string]$SenderAddress,
        [string]$SenderSmtp,
        [string]$Subject,
        [string]$To = '',
        [string]$CC = '',
        [string]$BodyPreview = '',
        [string]$FullBody = '',
        [string]$MessageClass = 'IPM.Note',
        [string]$Importance = 'Normal'
    )

    [pscustomobject]@{
        From = [pscustomobject]@{
            Name = $SenderName
            Address = $SenderAddress
        }
        SenderSmtp = $SenderSmtp
        Subject = $Subject
        To = $To
        CC = $CC
        BodyPreview = $BodyPreview
        FullBody = $FullBody
        MessageClass = $MessageClass
        Importance = $Importance
    }
}

Describe 'Start-SiftrFullLoop heuristic classifier' {
    BeforeAll {
        $script:user = [pscustomobject]@{
            Smtp = 'user.person@example.test'
            Alias = 'user.person'
            DisplayName = 'User Person'
            Tokens = @('@user.person', '@user')
        }

        $script:config = [pscustomobject]@{
            orgDomain = 'example.test'
        }

        $script:org = [pscustomobject]@{
            manager = [pscustomobject]@{ name = 'Manager Person'; email = 'manager.person@example.test' }
            directs = @()
            peers = @(
                [pscustomobject]@{ name = 'Peer Person'; email = 'peer.person@example.test' }
            )
            slt = $null
        }
    }

    It 'keeps user-authored mail out of Low Priority' {
        $latest = New-TestRecord `
            -SenderName 'User Person' `
            -SenderAddress 'user.person@example.test' `
            -SenderSmtp 'user.person@example.test' `
            -Subject 'RE: [Action: Please Review] WinHEC 2026 Conference Recap' `
            -To 'Somebody Else <someone@example.test>' `
            -CC 'User Person <user.person@example.test>' `
            -BodyPreview 'Sharing the recap and next steps.' `
            -FullBody 'Please review the recap when you have time.'

        $decision = Get-HeuristicDecision -Latest $latest -ThreadRecords @($latest) -User $user -Config $config -Org $org

        $decision.tier | Should -Be 'INFORMED'
        $decision.reason | Should -Be 'Fallback Phase 2: user-authored message'
    }

    It 'treats peer status updates as informed even when only the display name matches org context' {
        $latest = New-TestRecord `
            -SenderName 'Peer Person' `
            -SenderAddress 'Peer Person <peer.person2@example.test>' `
            -SenderSmtp 'peer.person2@example.test' `
            -Subject 'Compete Watch 5/15' `
            -To 'Compete DL <compete@example.test>' `
            -CC 'User Person <user.person@example.test>' `
            -BodyPreview 'Weekly status update for compete watch.' `
            -FullBody 'FYI: weekly compete watch status update with no actions needed.'

        $decision = Get-HeuristicDecision -Latest $latest -ThreadRecords @($latest) -User $user -Config $config -Org $org

        $decision.tier | Should -Be 'INFORMED'
        $decision.reason | Should -Be 'Fallback Phase 2: org sender FYI'
    }

    It 'round-trips stored UTC bookmarks through the current local timezone' {
        $utc = Get-UtcDateTime -Timestamp '2026-05-18T21:00:18.6740055Z'

        $local = Convert-UtcToCurrentLocalDateTime -Timestamp $utc
        $roundTrip = Convert-CurrentLocalToUtc -Timestamp $local

        $local.Kind | Should -Be 'Local'
        $roundTrip.ToString('o') | Should -Be $utc.ToString('o')
    }
}

Describe 'Start-SiftrFullLoop backlog and action reliability' {
    BeforeAll {
        $global:LoopRecentWindowHours = 72
    }

    It 'keeps recent uncategorized backlog eligible across bookmark advances' {
        $bookmark = (Get-CurrentLocalDateTime).AddHours(-1)
        $backlogSince = Get-LoopBacklogSince -Since $bookmark -State ([pscustomobject]@{ mode = 'continuous' })
        $messages = @(
            [pscustomobject]@{ ReceivedTime = (Get-CurrentLocalDateTime).AddHours(-2); Subject = 'missed earlier cycle' }
            [pscustomobject]@{ ReceivedTime = (Get-CurrentLocalDateTime).AddHours(-80); Subject = 'historical backlog' }
        )

        $filtered = @(Filter-LoopMessagesBySince -Messages $messages -Since $backlogSince)

        $filtered.Count | Should -Be 1
        $filtered[0].Subject | Should -Be 'missed earlier cycle'
    }

    It 'honors an older continuous-mode recovery bookmark' {
        $bookmark = (Get-CurrentLocalDateTime).AddHours(-100)

        $backlogSince = Get-LoopBacklogSince -Since $bookmark -State ([pscustomobject]@{ mode = 'continuous' })

        $backlogSince.ToString('o') | Should -Be $bookmark.ToString('o')
    }

    It 'fails the cycle when Outlook actions report errors' {
        $summary = [pscustomobject]@{
            Errors = 1
            Details = @(
                [pscustomobject]@{
                    Action = 'Failed'
                    Subject = 'Newsletter'
                    Error = 'Move failed'
                }
            )
        }

        { Assert-SiftrInboxActionSummary -Summary $summary } |
            Should -Throw 'Outlook actions reported 1 error(s): Newsletter: Move failed'
    }

    It 'allows nonfatal Outlook skips' {
        $summary = [pscustomobject]@{
            Errors = 0
            Skipped = 1
            Details = @(
                [pscustomobject]@{
                    Action = 'Skipped'
                    Subject = 'Moved elsewhere'
                    Error = 'Message not found in Inbox root'
                }
            )
        }

        { Assert-SiftrInboxActionSummary -Summary $summary } | Should -Not -Throw
    }
}
