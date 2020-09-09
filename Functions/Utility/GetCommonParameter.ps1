Function GetCommonParameter
{

    [cmdletbinding(SupportsShouldProcess)]
    param()
    if ($PSCmdlet.ShouldProcess("ShouldProcess?"))
    {
        #The impactful code
    }
    $MyInvocation.MyCommand.Parameters.Keys

}
