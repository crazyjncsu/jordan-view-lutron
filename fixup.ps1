Start-Job -RunAs32 -ArgumentList $PSScriptRoot\JordanView.hdf -ScriptBlock {
	param($path)

	$connection = new-object -comobject ADODB.Connection
	$connection.Open("Provider = Microsoft.Jet.OLEDB.4.0;Data Source = $path")

	#$connection.Execute("UPDATE [Zones] SET Low_End = 1, High_End = 99") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'Dining Room Lights' WHERE Parent_ID = 32 AND Position = 15") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Foyer Lights' WHERE Parent_ID = 32 AND Position = 16") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Game Room Lights' WHERE Parent_ID = 32 AND Position = 17") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Kitchen Lights' WHERE Parent_ID = 32 AND Position = 18") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Living Room Lights' WHERE Parent_ID = 32 AND Position = 19") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Hallway Lights' WHERE Parent_ID = 32 AND Position = 20") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Bridge Lights' WHERE Parent_ID = 32 AND Position = 21") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bedroom Lights' WHERE Parent_ID = 32 AND Position = 8") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bathroom Lights' WHERE Parent_ID = 32 AND Position = 9") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Closet Lights' WHERE Parent_ID = 32 AND Position = 10") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Solarium Lights' WHERE Parent_ID = 32 AND Position = 11") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Carport Lights' WHERE Parent_ID = 32 AND Position = 12") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Landscape Lights' WHERE Parent_ID = 32 AND Position = 13") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Flood Lights' WHERE Parent_ID = 32 AND Position = 14") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'Stair Lights' WHERE Parent_ID = 32 AND Position = 1") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Balcony Lights' WHERE Parent_ID = 32 AND Position = 2") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Patio Lights' WHERE Parent_ID = 32 AND Position = 3") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bedroom Shades' WHERE Parent_ID = 32 AND Position = 4") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Main Shades' WHERE Parent_ID = 32 AND Position = 5") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Garage Outdoor Lights' WHERE Parent_ID = 32 AND Position = 6") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Alarm Lights' WHERE Parent_ID = 32 AND Position = 7") | Out-Null



	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 35 AND Position = 15") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Hot Water Circulator' WHERE Parent_ID = 35 AND Position = 16") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Indoor Fireplace' WHERE Parent_ID = 35 AND Position = 17") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 35 AND Position = 18") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Pool Lights' WHERE Parent_ID = 35 AND Position = 19") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Alarm Siren' WHERE Parent_ID = 35 AND Position = 20") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Living Room Door Lights' WHERE Parent_ID = 35 AND Position = 21") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'Balcony Bedroom Lights' WHERE Parent_ID = 35 AND Position = 8") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Balcony Bathroom Lights' WHERE Parent_ID = 35 AND Position = 9") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Balcony Closet Lights' WHERE Parent_ID = 35 AND Position = 10") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bedroom Tray Lights' WHERE Parent_ID = 35 AND Position = 11") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Master Bedroom Nightstand Lights' WHERE Parent_ID = 35 AND Position = 12") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Kitchen Gathering Floor Receptacles' WHERE Parent_ID = 35 AND Position = 13") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Dining Room Receptacles' WHERE Parent_ID = 35 AND Position = 14") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = 'HVAC Main Downstairs Interrupt' WHERE Parent_ID = 35 AND Position = 1") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'HVAC Master Downstairs Interrupt' WHERE Parent_ID = 35 AND Position = 2") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'HVAC Upstairs Interrupt' WHERE Parent_ID = 35 AND Position = 3") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 35 AND Position = 4") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Roof Bedroom Lights' WHERE Parent_ID = 35 AND Position = 5") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Pond Bedroom Lights' WHERE Parent_ID = 35 AND Position = 6") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Living Room Entertainment Center Lights' WHERE Parent_ID = 35 AND Position = 7") | Out-Null



	$connection.Execute("UPDATE EngravingText SET [Text] = 'Indoor Fireplace Damper' WHERE Parent_ID = 55 AND Position = 15") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Solarium Receptacles' WHERE Parent_ID = 55 AND Position = 16") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Driveway Gate Hold' WHERE Parent_ID = 55 AND Position = 17") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = 'Driveway Gate Lights' WHERE Parent_ID = 55 AND Position = 18") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 19") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 20") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 21") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 8") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 9") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 10") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 11") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 12") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 13") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 14") | Out-Null

	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 1") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 2") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 3") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 4") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 5") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 6") | Out-Null
	$connection.Execute("UPDATE EngravingText SET [Text] = '' WHERE Parent_ID = 55 AND Position = 7") | Out-Null



	$connection.Close()
} | Wait-Job | Receive-Job
