Function GetArrayIndexForProperty
{
    [cmdletbinding()]
    [OutputType([System.Int32])]
    param(
        [parameter(mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        $array #The array for which you want to find a value's index
        ,
        [parameter(mandatory = $true)]
        $value #The Value for which you want to find an index
        ,
        [parameter(Mandatory)]
        $property #The property name for the value for which you want to find an index
    )
    if ($null -ne $array -and $array.count -ne 0)
    {
        [array]::indexof(($array.$property), $value)
    }
    else
    {
        -1
    }

}
