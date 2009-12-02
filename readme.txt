This is a YouTube video downloader written in VBScript. I wrote it because
when I was a child, I did a lot of programming in Visual Basic and I wanted to
remember what it was like.

It was written by Peteris Krumins (peter@catonmat.net).
His blog is at http://www.catonmat.net  --  good coders code, great reuse.

The code is licensed under the MIT license.

I also wrote a tutorial on how I created this program. The tutorial is called
"Writing a YouTube Video Downloader in VBScript", and I explain what Windows
Scripting Host (WSH) is, what cscript and wscript are, how to parse command
line arguments in VBScript, and how to use XmlHttp COM object. Read the
article here:

 http://www.catonmat.net/blog/writing-a-youtube-video-downloader-in-vbscript/

------------------------------------------------------------------------------

The program is called "ytdown.vbs". It can be either be used from command line
via cscript, or as a dialog-based application via wscript.

To run it as a dialog-based application, just double click the "ytdown.vbs"
file and it will ask you to enter the address of a YouTube video (see the
article for a screenshot).

To run it from command line, run it via cscript as following:

    cscript ytdown.vbs "http://www.youtube.com/watch?v=ID1" "..."

You may specify multiple video URLs and it will download all of them.

Here is an example run:

    C:\>cscript ytdown.vbs "http://www.youtube.com/watch?v=2mTLO2F_ERY"
    Microsoft (R) Windows Script Host Version 5.7
    Copyright (C) Microsoft Corporation. All rights reserved.

    Downloading video 'Mr. W'
    Done downloading video. Saved to Mr__W.flv.

------------------------------------------------------------------------------

Happy downloading!


Sincerely,
Peteris Krumins
http://www.catonmat.net

