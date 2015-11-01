Import-Module Simplex

New-PSDrive mail -psprovider simplex -root "$PSScriptRoot\OutlookProvider.ps1"

function Send-MailReply
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]
        $To
    )

    Process
    {
        $newMessage = $InputObject.Reply()
        $newMessage.To = $To
        $newMessage.Send()
    }
}

function Send-MailForward
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]
        $To
    )

    Process
    {
        $newMessage = $InputObject.Forward()
        $newMessage.To = $To
        $newMessage.Send()
    }
}