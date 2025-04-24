function Get-UniqieRules( [Analyser.Domain.Rule[]] $rules )
{
    [Linq.Enumerable]::ToArray(
        [Linq.Enumerable]::Distinct(
            $rules,
            $SCRIPT:RuleIdentityComparer
        )
    )
}
