# Q-SYS-PPTX-Module
- Uses TCP Socket to connect data from powerpoint to Q-SYS for automation
- Inlcudes automatic speaker notes, timer, current speaker and next speaker 
# How To Use:
Install the PPTXProject.qsys and the PPTXConnector.exe
the PPTXProject.qsys is a base template with all of the code implemeted into a Text Controller with controls setup and UCI setup - all fully functional.
PPTXProject.qsys has to be used with PPTXConnector.exe to establish the TCP connection needed to send data from powerpoint to Q-SYS.
All LUA inside of the text controller can be edited at anytime.

IP is setup for localhost(127.0.0.1) and port is configurable, by default port is 12345 in Q-SYS.
In the GUI for PPTXConnector put the IP in, 127.0.0.1 in this instance, put the corresponding port in and the hard URL to your PPTX then click start.
In the Q-SYS designer there is a section called "Controls" if the connection was made and working, you should see the console populate with info based on the slide.
Similarly, this can be seen in the text controller console.

Warning:
Sometimes the TCP connection between Powershell and Q-SYS can bug out, in some instances you might have to restart your core and/or the powershell script to get the connection to work.
Always make sure you close the TCP connection properly when finished by closing all instances that use TCP connection, I.E. powerpoint that opens, the GUI and the Console window.



