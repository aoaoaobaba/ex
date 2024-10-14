function Get-FormattedFileList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, HelpMessage="Specify the folder path.")]
        [string]$Path
    )

    Get-ChildItem -Path $Path -Recurse -File | 
        ForEach-Object {
            "{0}`t{1}`t{2}" -f $_.Name, $_.Directory.Name, $_.FullName
        }
}
Set-Alias -Name GetFileList -Value Get-FormattedFileList