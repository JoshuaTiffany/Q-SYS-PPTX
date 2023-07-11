
function Start-TcpListener {
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

function Stop-TcpListener {
    param (
        [System.Net.Sockets.TcpListener]$listener,
        [System.Net.Sockets.TcpClient]$clientSocket
    )

    # Close the client connection and listener
    $clientSocket.Close()
    $listener.Stop()

    Write-Host "Client disconnected"
    Write-Host "Server stopped"
}

try {
    # Create a PowerPoint application object
    $powerPoint = New-Object -ComObject PowerPoint.Application

    # Open the presentation in a visible window
    $presentation = $powerPoint.Presentations.Open("C:\Users\Joshua.Tiffany\Downloads\FY23 Q2 Town Hall (OCT2022)-for demo.pptx", $null, $null, $true)

    while ($true) {
        try {
            $listener, $clientSocket = Start-TcpListener -serverIP "127.0.0.1" -port 12345

            # Main code block
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
            
            $previousSlideNumber = GetSelectedSlideNumber
            $speakerNotes = ""
            
            while ($clientSocket.Connected) {
                Start-Sleep -Milliseconds 1000
            
                $currentSlideNumber = GetSelectedSlideNumber
            
                if ($currentSlideNumber -ne $previousSlideNumber) {
                    Write-Host "Current Slide Number: $currentSlideNumber"
                    $newSpeakerNotes = GetSpeakerNotes($currentSlideNumber)
                    Write-Host "Speaker Notes:"
                    Write-Host $newSpeakerNotes
            
                    $speakerNotes += $newSpeakerNotes
            
                    $previousSlideNumber = $currentSlideNumber
                }
            
                if ($speakerNotes -ne "") {
                    Write-Host "Sending Speaker Notes:"
                    Write-Host $speakerNotes
            
                    $data = $speakerNotes
                    $clientStream = $clientSocket.GetStream()
                    $dataBytes = [System.Text.Encoding]::UTF8.GetBytes($data)
                    $clientStream.Write($dataBytes, 0, $dataBytes.Length)
            
                    $speakerNotes = ""
                }
            
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
            # End of main code block

            # Close the client connection
            $clientSocket.Close()
            Write-Host "Client disconnected"
        }
        catch {
            Write-Host "An error occurred: $_"
            if ($clientSocket -ne $null) {
                Stop-TcpListener -listener $listener -clientSocket $clientSocket
            }
        }
        finally {
            $listener.Stop()
        }
    }
}
catch {
    Write-Host "PowerPoint error occurred: $_"
}
finally {
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
    }
}