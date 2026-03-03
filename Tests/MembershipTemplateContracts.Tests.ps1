Describe 'Membership Template Contracts' {
    BeforeAll {
        $repoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
        $script:m3007TemplatePath = Join-Path -Path $repoRoot -ChildPath 'Online/Modify/M3007-Set-EntraSecurityGroupMembers.input.csv'
        $script:m3115TemplatePath = Join-Path -Path $repoRoot -ChildPath 'Online/Modify/M3115-Set-ExchangeOnlineDistributionListMembers.input.csv'
    }

    It 'M3007 template includes MemberAction with Add/Remove values' {
        Test-Path -LiteralPath $m3007TemplatePath | Should -BeTrue
        $rows = @(Import-Csv -LiteralPath $m3007TemplatePath)
        $rows.Count | Should -BeGreaterThan 0

        $rows[0].PSObject.Properties.Name | Should -Contain 'MemberAction'

        foreach ($row in $rows) {
            @('Add', 'Remove') | Should -Contain $row.MemberAction
        }
    }

    It 'M3115 template includes MemberAction with Add/Remove values' {
        Test-Path -LiteralPath $m3115TemplatePath | Should -BeTrue
        $rows = @(Import-Csv -LiteralPath $m3115TemplatePath)
        $rows.Count | Should -BeGreaterThan 0

        $rows[0].PSObject.Properties.Name | Should -Contain 'MemberAction'

        foreach ($row in $rows) {
            @('Add', 'Remove') | Should -Contain $row.MemberAction
        }
    }
}
