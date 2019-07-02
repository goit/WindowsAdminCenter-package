<#
.Synopsis
   Get product and version information from MSI file.
.DESCRIPTION
   Use the MsiInstaller.Installer ComObject to enumerate MSI database specific information
   There are only 5 properties for MSI's that are mandatory.  (According to https://docs.microsoft.com/en-us/windows/desktop/Msi/property-reference )
   These are:
       ProductCode     - A unique identifier for a specific product release.
       Manufacturer    - Name of the application manufacturer.
       ProductName     - Human readable name of an application.
       ProductVersion  - String format of the product version as a numeric value.
       ProductLanguage - Numeric language identifier (LANGID) for the database.

   By default all of these are returned. This can be modified by using the [-Property] parameter.

.EXAMPLE
PS C:\> Get-MsiInformation -Path "$env:Temp\Installer.msi"

Path            : C:\Users\username\AppData\Local\Temp\Installer.msi
ProductCode     : {75BDEFC7-6E84-55FF-C326-CE14E3C889EC}
ProductVersion  : 1.9.492.0
ProductName     : Installer v1.9.0
Manufacturer    : My Company, Inc.
ProductLanguage : 1033

This example takes the path as a parameter and returns all fields

.EXAMPLE
Get-ChildItem -Path "$env:Temp\1.0.0" -Recurse -File -Include "*.msi" | Get-MsiInformation -Property ProductVersion

Path                                                        ProductVersion
----                                                        --------------
C:\Users\username\AppData\Local\Temp\Build456\Installer.msi 1.0.456.0     
C:\Users\username\AppData\Local\Temp\Build457\Installer.msi 1.0.451.0     

This example takes multiple paths from a Get-ChildItem query and extracts the information.

.INPUTS
   [System.IO.File[]] - Single or Array of Paths to interrogate

.OUTPUTS
   [System.Management.Automation.PSCustomObject[]] - Contains the Msi File Object and Associated Properties

.LINK
   https://blog.kmsigma.com/2016/11/30/getting-information-about-msi-files/

.LINK
   https://docs.microsoft.com/en-us/windows/desktop/Msi/property-reference

.LINK
   http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/

.NOTES
   Heavily Infuenced by http://www.scconfigmgr.com/2014/08/22/how-to-get-msi-file-information-with-powershell/

.FUNCTIONALITY
   Uses ComObjects to Enumerate specific fields in the MSI database
#>
function Get-MsiInformation
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                   PositionalBinding=$false,
                   ConfirmImpact='Medium')]
    [Alias("gmsi")]
    Param(
        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage = "Provide the path to an MSI")]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$Path,
 
        [parameter(Mandatory=$false)]
        [ValidateSet( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )]
        [string[]]$Property = ( "ProductCode", "Manufacturer", "ProductName", "ProductVersion", "ProductLanguage" )
    )

    Begin
    {
        # Do nothing for prep
    }
    Process
    {
        
        ForEach ( $P in $Path )
        {
            if ($pscmdlet.ShouldProcess($P, "Get MSI Properties"))
            {            
                try
                {
                    Write-Verbose -Message "Resolving file information for $P"
                    $MsiFile = Get-Item -Path $P
                    Write-Verbose -Message "Executing on $P"
                    
                    # Read property from MSI database
                    $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
                    $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($MsiFile.FullName, 0))
                    
                    # Build hashtable for retruned objects properties
                    $PSObjectPropHash = [ordered]@{File = $MsiFile.FullName}
                    ForEach ( $Prop in $Property )
                    {
                        Write-Verbose -Message "Enumerating Property: $Prop"
                        $Query = "SELECT Value FROM Property WHERE Property = '$( $Prop )'"
                        $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
                        $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                        $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                        $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
 
                        # Return the value to the Property Hash
                        $PSObjectPropHash.Add($Prop, $Value)

                    }
                    
                    # Build the Object to Return
                    $Object = @( New-Object -TypeName PSObject -Property $PSObjectPropHash )
                    
                    # Commit database and close view
                    $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
                    $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
                    $MSIDatabase = $null
                    $View = $null
                }
                catch
                {
                    Write-Error -Message $_.Exception.Message
                }
                finally
                {
                    Write-Output -InputObject @( $Object )
                }
            } # End of ShouldProcess If
        } # End For $P in $Path Loop

    }
    End
    {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}
