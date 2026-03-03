Describe 'Repository Contracts' {
    BeforeAll {
        $repoRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)
        $script:contractScriptPath = Join-Path -Path $repoRoot -ChildPath 'Build/Test-RepositoryContracts.ps1'
    }

    It 'passes repository contract validation' {
        Test-Path -LiteralPath $contractScriptPath | Should -BeTrue
        { & $contractScriptPath } | Should -Not -Throw
    }
}
