# Q-SYS-PPTX-Module
- Uses TCP Socket to connect data from powerpoint to Q-SYS for automation
- Inlcudes automatic speaker notes, timer, current speaker and next speaker

Link to QSC Code Exchange: https://developers.qsc.com/s/exchange/developer-repo/a3Y4X000001uCt2UAE/qsys-pptx-module
# How To Use:
Install the PPTXProject.qsys and the PPTXConnector.exe
the PPTXProject.qsys is a base template with all of the code implemeted into a Text Controller with controls setup and UCI setup - all fully functional.
PPTXProject.qsys has to be used with PPTXConnector.exe to establish the TCP connection needed to send data from powerpoint to Q-SYS.
All LUA inside of the text controller can be edited at anytime.

With UCI Interface 2, or with the controls section in the design, make sure the IP is correct and your Port is correct than click the start button on the UCI. You must do this first as, this will start a server porting from your core. Make sure you are not using a port that is already in use by your core.

IP is setup for core IP, so you need to make sure you know your core's IP, and port is configurable, by default port is 1703 in Q-SYS. DO NOT USE A PORT ALREADY IN USE BY YOUR CORE, I.E: PORT TCP:1702
In the GUI for PPTXConnector put your core's IP in the text box for IP: - put the corresponding port in, the same one you have in Q-SYS, and the hard URL to your PPTX then click start.
In the Q-SYS designer there is a section called "Controls" if the connection was made and working, you should see the console populate with info based on the slide.(Same as interface 2) This section is also where you put the port in, you can change the IP by going to Text Controller(the one in design, inside this github) script in Q-SYS and by just changing the variable
Similarly, this can be seen in the text controller console.

Warning:
Sometimes the TCP connection between Powershell and Q-SYS can bug out, in some instances you might have to restart your core and/or the powershell script to get the connection to work.
Always make sure you close the TCP connection properly when finished by closing all instances that use TCP connection, I.E. powerpoint that opens, the GUI and the Console window.

Make sure the core Data/Time is setup to the correct and current data/time so the Timer can work.
Equally, make sure the device running the Powershell script, PPTXConnector, is the correct and current data/time



**TLDR:** Use interface 2 to setup the server by using your core's IP and an open port on your core > Start - Open PPTXConnector.exe > put in same IP as core and correspoding port(If using a seperate PC, you have to connect it to the same network as the Q-SYS core), put in hard path to the PPTX you are wanting to display > Start. You should see the speaker notes populate in both console windows in the UCI and in the PPTXConnector console window. Most errors can be fixed by stoping the server and starting it.


# How it works:
Powershell connects the PPTX that the user provides, opens an instance of it and gets speaker notes, current slide - and for automation: Custom Command Lines(Which will be discussed below).
Once it grabs that info from powerpoint it sends the data VIA a TCP Socket to Q-SYS and from there Q-SYS using LUA extracts the data and processes it.
The data sent is cleaned, I.E removing command lines from the speaker notes displayed, and then populated on the UCI.

Q-SYS core = TCP server
PPTXConnector = Client sending data

There is a 2m warning system for the speaker to see, which when happens will turn the background of the timer to turn an orangeish color and once the speaker goes over the time the clock will start to count up instead of down, plus the background will turn red.

# Custom Command Lines:
Command lines are lines of text that you put into the speaker notes of any slide you want.
- Info(currentTalker(tilde)endTime(tilde)nextTalker) EX. Info(Josh Tiffany(tilde)13:45:00(tilde)Stan Nice) - WARNING: In the speaker notes it should not be (tilde) but just ~ to seperate the variables : This is because README syntax makes words between ~ crossout 
- [set] Is a control pins(1-16) that can be declared in the speaker notes and can be used for in connection in Q-SYS(I.E: Changing lighting in the room automatically when the presentation starts) - [set] will remain On(True) until the program is closed or TCP server stopped
- [trigger] Similar to [set], control pins(1-16) that can be setup to connect to Q-SYS objects, main difference is trigger will be set back to default(false) once a slide is changed.  

The Info[] variables, aka currentTalker(tilde)endTime(tilde)nextTalker, will stay the same until another slide has another info[] command line in it. So, for example the first slide will have an info[] command and the next 3 do not - what will happen is that the current talker, timer and next talker will not change until a slide with another info[] command is present.

Command Lines must be spelled properly and exactly as presented above - and in the correct syntax.
endTime is in military time/Zulu, "~" is used to seperate variables, ONLY USE ~
Make sure the command lines you put into any speaker notes is at the top of the speaker notes - this is so the program does not have the chance to get confused and so blank lines created by parsing out the comamnd lines from the speaker notes does not appear weird in the UCI dispaly.

