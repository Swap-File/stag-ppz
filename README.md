stag-ppz
========

This project was origionally hosted on TinyAppz back in 2006-2007 before they went down.

It has been a long time since I looked at or touched the code, but I'm sure someone will still find it useful.

This program was made in VB Express 2008 to interface with the Stag PPZ EPROM Programmer.
If you do not have it already, you will need the .NET Framework to run this program.

The first time you run the program it will create PPZ.ini.  Edit this to change your settings.


PPZ.ini examples:

Example:
[Settings]
ComPort=2   (Enter Com Port Number Here)
baud=4800   (Enter Baud Speed  Here - 4800 was maximum stable speed for my PPZ)
databits=8  (Enter Number of Data Bits Here - 8 is normal)
stopbits=1  (Enter Number of Stop Bits Here - 1 is normal)
parity=0    (Change Parity Settings.  0 is None, 1 is Odd, 2 is Even)


Settings On the Programmer Side:

Menu A 1: Serial Port Settings
Interface Serial 1
Baud Rate 4800 (higher speeds were unstable for me)
Word Length 8
Parity Off
Stop Bits 1

Menu A 2: I/O Software Settings
Format: DEC-Binary
Cntrl Z:Off (I do not know what this does, I just left it on)

Menu A 4:
Remote I/O: Serial I
Pass-Through: Off



To Receive a file, go into I/O on the StagPPZ, select output and the RAM limits.
Then hit enter so it says "Dec-Binary Load in Progress".
Now, on the computer side press "Receive Data To Buffer"

To receive a file, go into I/O on the StagPPZ and select input. 
You can leave the ram limit at FFFF no matter how much you are planning to upload.
Then hit enter so it says "Dec-Binary Dump in Progress".
Now, on the computer side press "Send Data From Buffer"


Do not do anything that will max out your CPU usage while transferring files.  Bits may be lost.

If you find any bugs, feel free to report them.
