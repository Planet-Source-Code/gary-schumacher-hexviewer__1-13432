Hex File Viewer

Program will display any file as Hex (Left Side)/Ascii (Right Side). It works in all display resolutions (although 640 X 480 is very tight).  I had to fudge a little on the Column Header, it was not aligning to the Rich Text Box.  I had experimented with converting each byte in the file to hex...but I found the lookup table was about 5 times faster.

Currently, the program only reads a file (no search or update capability).

Large Files (3MB and up) take a lot of time to be formatted - Because entire file is loaded into Rich Text File Box.  I have a Pentium III 500MHz with 128M of RAM and the Microsoft Biblio.MDB took a couple of minutes to load. (Hard Drive was working overtime with all of the swapping).  Users with less RAM may want to stick to smaller files.

Important:  Font is Mono-spaced (All chars same size)  I used New Courier.  If not Mono-spaced...Display will not be formatted properly.

From the Projects|Components Menu...The Following Components must be included in the Project:
Microsoft Common Dialog Control 6.0
Microsoft Rich TextBox Control 6.0
Microsoft Windows Common Controls 6.0

Code is well-documented and demonstrates uses of the Status Bar, Progress Bar & Rich Text Box controls.

12/13/2000
Added C Dll to speed text formatting process
Removed ProgressBar
Commented all Routines/Code that were no longer needed because of the Dll
