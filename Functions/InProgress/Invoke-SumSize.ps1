#Invoke-SumSize function from: https://gist.github.com/Szeraax/c077faba75aa577a7790852afbb1d25c
Function Invoke-SumSize
{
    [alias("iss")]
    param (
        [parameter(
            ValueFromPipeline
        )]
        [Alias('Length')]
        [ValidateNotNullorEmpty()]
        # Removed type enforcement on the input object because there are several possible types they send in on
        # and if their input object doesn't cast to double and propertyname and propertyname alias also doesn't cast
        # to double, then we'd just rather keep going without big errors.
        $InputObject
    )

    begin
    {
        [double]$Total = 0
    }

    process
    {
        # Because people may pass in a single object containing several values that you actually want to operate on
        foreach ($Item in $InputObject)
        {
            if ($Item -as [double])
            {
                $Total += $Item
            }
            # Use PSObject.Properties method of getting Length so that we avoid cases where powershell auto-returns
            # a number on an array of X length and also check that we are not dealing with a string of X length
            elseif ($Item.PSObject.Properties['Length'].Value -as [double] -and -not ($Item -is "String"))
            {
                $Total += $Item.Length
            }
            else
            {
                # If we can't cast it, then let the non-terminating error float on up
                [double]$Item
            }
        }
    }

    end
    {
        switch -Regex ([math]::truncate([math]::log([double]$Total, 1024)))
        {
            '^0' { "$Total Bytes" ; Break }
            '^1' { "{0:n2} KB" -f ($Total / 1KB) ; Break }
            '^2' { "{0:n2} MB" -f ($Total / 1MB) ; Break }
            '^3' { "{0:n2} GB" -f ($Total / 1GB) ; Break }
            '^4' { "{0:n2} TB" -f ($Total / 1TB) ; Break }
            '^5' { "{0:n2} PB" -f ($Total / 1PB) ; Break }
            # When we fail to have any matches, 0 Bytes is more clear than 0.00 PB (since <5GB would be 0.00 PB still)
            Default { "0 Bytes" }
        }
    }
}