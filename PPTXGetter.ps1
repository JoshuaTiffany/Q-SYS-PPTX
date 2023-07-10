
# Create a PowerPoint application object
$powerPoint = New-Object -ComObject PowerPoint.Application

# Open the presentation in a visible window
$presentation = $powerPoint.Presentations.Open("C:\Users\Joshua.Tiffany\Downloads\FY23 Q2 Town Hall (OCT2022)-for demo.pptx", $null, $null, $true)

function GetSelectedSlideNumber() {
    # Check if PowerPoint is in presenter mode
    if ($powerPoint.SlideShowWindows.Count -gt 0) {
        $slideNumber = $powerPoint.SlideShowWindows.Item(1).View.Slide.SlideNumber
        return $slideNumber
    }

    # Default case for normal view mode
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

# Initialize the previous slide number and speaker notes
$previousSlideNumber = GetSelectedSlideNumber
$speakerNotes = ""

# Periodically check for slide changes and send data to the client
while ($true) {
    Start-Sleep -Milliseconds 1000

    $currentSlideNumber = GetSelectedSlideNumber

    # Check if the slide number has changed
    if ($currentSlideNumber -ne $previousSlideNumber) {
        Write-Host "Current Slide Number: $currentSlideNumber"
        $newSpeakerNotes = GetSpeakerNotes($currentSlideNumber)
        Write-Host "Speaker Notes:"
        Write-Host $newSpeakerNotes
        
        # Append the new speaker notes to the existing notes
        $speakerNotes += $newSpeakerNotes

        $previousSlideNumber = $currentSlideNumber
    }

    # Send the complete speaker notes to the client
    if ($speakerNotes -ne "") {
        Write-Host "Sending Speaker Notes:"
        Write-Host $speakerNotes
        
        $data = $speakerNotes
        $clientStream = $clientSocket.GetStream()
        $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($data)
        $clientStream.Write($dataBytes, 0, $dataBytes.Length)
        
        # Reset the speaker notes
        $speakerNotes = ""
    }
    
    # Check if PowerPoint is in presenter mode
    if ($powerPoint.SlideShowWindows.Count -gt 0) {
        $presenterView = $powerPoint.SlideShowWindows.Item(1).View.Type
        if ($presenterView -eq 2) {  # Presenter mode
            Write-Host "Presenter Mode"
            $nextSlide = $currentSlideNumber + 1
            if ($nextSlide -le $presentation.Slides.Count) {
                $powerPoint.SlideShowWindows.Item(1).View.Next
            } else {
                $powerPoint.SlideShowWindows.Item(1).View.Exit
            }
        }
    }
}

# Close the client connection and listener
$clientSocket.Close()
$listener.Stop()