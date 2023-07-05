
# Create a PowerPoint application object
$powerPoint = New-Object -ComObject PowerPoint.Application

# Open the presentation in a visible window
$presentation = $powerPoint.Presentations.Open("C:\Users\Joshua.Tiffany\Downloads\Intern Workshop - Financial Literacy 06.21.23.pptx", $null, $null, $true)

# Define a function to retrieve the current slide number
function GetSelectedSlideNumber() {
    $slide = $powerPoint.ActiveWindow.View.Slide
    $slideNumber = $slide.SlideNumber
    return $slideNumber
}

# Define a function to retrieve the speaker notes of the specified slide
function GetSpeakerNotes($slideNumber) {
    $slide = $presentation.Slides.Item($slideNumber)
    $notesPage = $slide.NotesPage

    $speakerNotes = ""
    foreach ($paragraph in $notesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Paragraphs()) {
        $text = $paragraph.Text
        $speakerNotes += "$text`r`n"
    }

    return $speakerNotes
}

# Define the server IP address and port
$serverIP = [System.Net.IPAddress]::Parse("127.0.0.1")
$port = 12345

# Create a TCP listener object and start listening
$listener = [System.Net.Sockets.TcpListener]::new($serverIP, $port)
$listener.Start()

Write-Host "Server listening on $($listener.LocalEndpoint)"

# Wait for client connection
Write-Host "Waiting for client connection..."
$clientSocket = $listener.AcceptTcpClient()
$clientAddress = $clientSocket.Client.RemoteEndPoint.Address.ToString()
Write-Host "Client connected: $clientAddress"

# Initialize the previous slide number
$previousSlideNumber = GetSelectedSlideNumber

# Periodically check for slide changes and send data to the client
while ($true) {
    Start-Sleep -Milliseconds 1000

    $currentSlideNumber = GetSelectedSlideNumber

    # Check if the slide number has changed
    if ($currentSlideNumber -ne $previousSlideNumber) {
        Write-Host "Current Slide Number: $currentSlideNumber"
        $speakerNotes = GetSpeakerNotes($currentSlideNumber)
        Write-Host "Speaker Notes:"
        Write-Host $speakerNotes
        $previousSlideNumber = $currentSlideNumber

        # Send slide data to the client
        $data = "Current Slide Number: $currentSlideNumber`r`n"
        $notesData = "Speaker Notes:`r`n$speakerNotes"
        $clientStream = $clientSocket.GetStream()
        $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($data)
        $notesDataBytes = [System.Text.Encoding]::UTF8.GetBytes($notesData)
        $clientStream.Write($dataBytes, 0, $dataBytes.Length)
        $clientStream.Write($notesDataBytes, 0, $notesDataBytes.Length)
    }
}

# Close the client connection and listener
$clientSocket.Close()
$listener.Stop()