
$formPPTXConnector_Load = {
	#TODO: Initialize Form Controls here
}


$buttonStart_Click = {
	$listener = $null
	$clientSocket = $null
	$powerPoint = $null
	$previousSpeakerNotes = ""
	
	try
	{
		# Retrieve input values from the form controls
		$address = $textBoxIP.Text
		$port = [int]$textBoxPort.Text
		$pptxPath = $textBoxPPTXPath.Text
		Write-Host $address
		Write-Host $port
		
		# Create a PowerPoint application object
		$powerPoint = New-Object -ComObject PowerPoint.Application
		
		# Open the presentation in a visible window
		$presentation = $powerPoint.Presentations.Open($pptxPath, $null, $null, $true)
		
		$clientSocket = New-Object System.Net.Sockets.TcpClient
		$clientSocket.Connect($address, $port)
		
		$statusLabel.Text = "Connected to " + $address + ":" + $port.ToString()
		
		while ($clientSocket.Connected)
		{
			$statusLabel.Text = "Connected"
			Start-Sleep -Milliseconds 1000
			
			$currentSlideNumber = GetSelectedSlideNumber
			
			if ($currentSlideNumber -ne $previousSlideNumber)
			{
				Write-Host "Current Slide Number: $currentSlideNumber"
				$newSpeakerNotes = GetSpeakerNotes $currentSlideNumber
				Write-Host "Speaker Notes:"
				Write-Host $newSpeakerNotes
				
				$speakerNotes = $newSpeakerNotes
				$previousSlideNumber = $currentSlideNumber
			}
			
			if ($speakerNotes -ne "" -and $speakerNotes -ne $previousSpeakerNotes)
			{
				Write-Host "Sending Speaker Notes:"
				Write-Host $speakerNotes
				
				$data = $speakerNotes
				$clientStream = $clientSocket.GetStream()
				$dataBytes = [System.Text.Encoding]::UTF8.GetBytes($data)
				$clientStream.Write($dataBytes, 0, $dataBytes.Length)
				
				$previousSpeakerNotes = $speakerNotes
				$speakerNotes = ""
				$statusLabel.Text = "Sent speaker note"
			}
			
			if ($powerPoint.SlideShowWindows.Count -gt 0)
			{
				$presenterView = $powerPoint.SlideShowWindows.Item(1).View.Type
				if ($presenterView -eq 2)
				{
					# Presenter mode
					Write-Host "Presenter Mode"
					$nextSlide = $currentSlideNumber + 1
					if ($nextSlide -le $presentation.Slides.Count)
					{
						$powerPoint.SlideShowWindows.Item(1).View.Next
					}
					else
					{
						$powerPoint.SlideShowWindows.Item(1).View.Exit
					}
				}
			}
		}
	}
	catch
	{
		$statusLabel.Text = "An error occurred: $_"
		Write-Host "An error occurred: $_"
	}
	finally
	{
		
	}
}

$buttonStop_Click = {
	if ($listener -ne $null)
	{
		$listener.Stop()
	}
	if ($powerPoint -ne $null)
	{
		$powerPoint.Quit()
	}
	$statusLabel.Text = "TCP connection stopped"
}

$buttonClose_Click = {
	# Close the form
	$formPPTXConnector.Close()
}

# Close the TCP connection and PowerPoint application when the form is closing
$formPPTXConnector_FormClosing = {
	if ($listener -ne $null)
	{
		$listener.Stop()
	}
	if ($clientSocket -ne $null)
	{
		$clientSocket.Close()
	}
	if ($powerPoint -ne $null)
	{
		$powerPoint.Quit()
	}
}

function Start-TcpListener
{
	param (
		[string]$serverIP,
		[int]$port
	)
	
	# Create a TCP listener object and start listening
	$listener = [System.Net.Sockets.TcpListener]::new($serverIP, $port)
	$listener.Start()
	
	Write-Host "Server listening on $($listener.LocalEndpoint)"
	
	# Wait for client connection
	Write-Host "Waiting for client connection..."
	$clientSocket = $listener.AcceptTcpClient()
	$clientAddress = $clientSocket.Client.RemoteEndPoint.Address.ToString()
	Write-Host "Client connected: $clientAddress"
	
	return $listener, $clientSocket
}

function GetSelectedSlideNumber()
{
	# Check if PowerPoint is in presenter mode
	if ($powerPoint.SlideShowWindows.Count -gt 0)
	{
		$slideNumber = $powerPoint.SlideShowWindows.Item(1).View.Slide.SlideNumber
		return $slideNumber
	}
	
	# Default case for normal view mode
	$slide = $powerPoint.ActiveWindow.View.Slide
	$slideNumber = $slide.SlideNumber
	return $slideNumber
}

function GetSpeakerNotes($slideNumber)
{
	$slide = $presentation.Slides.Item($slideNumber)
	$notesPage = $slide.NotesPage
	
	$speakerNotes = ""
	foreach ($shape in $notesPage.Shapes)
	{
		if ($shape.HasTextFrame -and $shape.TextFrame.HasText)
		{
			$paragraphs = $shape.TextFrame.TextRange.Paragraphs()
			foreach ($paragraph in $paragraphs)
			{
				$text = $paragraph.Text
				$speakerNotes += "$text`r`n"
			}
		}
	}
	
	return $speakerNotes
}