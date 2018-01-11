param([Parameter(Mandatory=$true)]$commitMessage)

Start-Job -RunAs32 -ArgumentList $PSScriptRoot\JordanView.hdf -ScriptBlock {
	param($path)

	$connection = new-object -comobject ADODB.Connection
	$connection.Open("Provider = Microsoft.Jet.OLEDB.4.0;Data Source = $path")

	$connection.Execute("UPDATE [Zones] SET Low_End = 1, High_End = 99") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'Dining Room Lights' WHERE Parent_ID = 32 AND Position = 17") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Foyer Lights' WHERE Parent_ID = 32 AND Position = 18") | Out-Null


	$connection.Execute("UPDATE EngravingText SET [Text] = 'Stair Lights' WHERE Parent_ID = 32 AND Position = 1") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bedroom Lights' WHERE Parent_ID = 32 AND Position = 9") | Out-Null

	$connection.Close()
} | Wait-Job | Receive-Job

git commit -a -m $commitMessage
git push