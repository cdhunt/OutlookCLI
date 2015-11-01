Import-Module Simplex

New-PSDrive mail -psprovider simplex -root "$PSScriptRoot\OutlookProvider.ps1"

function Send-MailReply
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

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
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

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

function Set-MailRead
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject
    )

    Process
    {
        $InputObject.Unread = $false
    }
}

function Set-MailTask
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=1)]
        $InputObject,

		[Parameter(Position=0)]
		[ValidateSet('Today','Tomorrow','ThisWeek','NextWeek','NoDate','Complete')]
		[String]
		$Interval

    )
	
	Begin
	{
		$olMarkInterval = "Microsoft.Office.Interop.Outlook.OlMarkInterval" -as [type] 
		$intervalValue = $olMarkInterval::olMarkNoDate
		switch ($Interval)
		{
			'Today' {$intervalValue = $olMarkInterval::olMarkToday}
			'Tomorrow' {$intervalValue = $olMarkInterval::olMarkTomorrow}
			'ThisWeek' {$intervalValue = $olMarkInterval::olMarkThisWeek}
			'NextWeek' {$intervalValue = $olMarkInterval::olMarkNextWeek}
			'NoDate' {$intervalValue = $olMarkInterval::olMarkNoDate}
			'Complete' {$intervalValue = $olMarkInterval::olMarkComplete}
		}
	}
    Process
    {
        $InputObject.MarkAsTask($intervalValue)
    }
}