'
' Peteris Krumins (peter@catonmat.net)
' http://www.catonmat.net  -  good coders code, great reuse
'
' 2007.08.03 v1.0 - initial release
' 2007.10.21 v1.1 - youtube changed the way it displays vids
' 2008.03.01 v1.2 - youtube changed the way it displays vids
' 2009.12.02 v1.3 - youtube changed the way it displays vids
'
Option Explicit

Dim WscriptMode

' Detect if we are running in WScript or CScript
If UCase(Right(WScript.Fullname, 11)) = "WSCRIPT.EXE" Then
    WScriptMode = True
Else 
    WScriptMode = False
End If 

Dim Args: Set Args = WScript.Arguments

If Args.Count = 0 And WScriptMode Then
    ' If running in WScript and no command line args are provided
    ' ask the user for a URL to the YouTube video
    Dim Url: Url = InputBox("Enter a YouTube video URL to download" & vbCrLf & _
                   "For example, http://youtube.com/watch?v=G1ynTV_E-5s", _
                   "YouTube Downloader, http://www.catonmat.net")
    If Len(Url) = 0 Then: WScript.Quit 1
    DownloadVideo Url
ElseIf Args.Count = 0 And Not WScriptMode Then
    ' If running in CScript and no command line args are provided
    ' show the usage and quit
    WScript.Echo "Usage: " & WScript.ScriptName & " <video url 1> [video url 2] ..."
    WScript.Quit 1
Else 
    ' Download all videos
    Dim I

    For I = 0 to args.Count - 1
        DownloadVideo args(I)
    Next
End If

' Downloads a YouTube video and saves it to a file
Sub DownloadVideo(Url)
    Dim Http, VideoTitle, VideoName, Req

    Set Http = CreateObject("Microsoft.XmlHttp")
    Http.open "GET", Url, False
    Http.send

    If Http.status <> 200 Then
        WScript.Echo "Failed getting video page at: " & Url & vbCrLf & _
                     "Error: " & Http.statusText
        Exit Sub
    End If

    Dim VideoId: VideoId = ExtractMatch(Url, "v=([A-Za-z0-9-_]+)")
    If Len(VideoID) = 0 Then
        WScript.Echo "Could not extract video ID from " & Url
        Exit Sub
    End If

    VideoTitle = GetVideoTitle(Http.responseText)
    If Len(VideoTitle) = 0 Then
        WScript.Echo "Failed extracting video title from video at URL: " & Url & vbCrLf & _
                     "Will use the video ID '" & VideoID & "' for the filename."
        VideoName = VideoID
    Else
        VideoName = VideoTitle
    End If

    Dim FmtMap: FmtMap = GetFmtMap(Http.responseText)
    If Len(FmtMap) = 0 Then
        WScript.Echo "Could not extract fmt_url_map from the video page."
        Exit Sub
    End If

    Dim VideoURL: VideoURL = Find_Video_5(FmtMap)
    If Len(VideoURL) = 0 Then
        WScript.Echo "Could not extract fmt_url_map from the video page."
        Exit Sub
    End If

    If WScriptMode = False Then: WScript.Echo "Downloading video '" & VideoName & "'"
    Http.open "GET", VideoURL, False
    Http.send

    If Http.status <> 200 Then
        WScript.Echo "Failed getting the flv video: " & Url & vbCrLf & _
                     "Error: " & Http.statusText
        Exit Sub
    End If

    Dim SaneFilename
    SaneFilename = MkFileName(VideoName)

    SaveVideo SaneFilename, Http.ResponseBody
    WScript.Echo "Done downloading video. Saved to " & SaneFilename & "."
End Sub

' Given fmt_url_map, url-escapes it, and finds the video url for video
' with id 5, which is the regular quality flv video.
Function Find_Video_5(FmtMap)
    FmtMap = Unescape(FmtMap)
    Find_Video_5 = ExtractMatch(FmtMap, ",?5\|([^,]+)")
End Function

' Given YouTube Html page, extract the fmt_url_map parameter that contains
' the URL to the .flv video
Function GetFmtMap(Html)
    GetFmtMap = ExtractMatch(Html, """fmt_url_map"": ""([^""]+)""")
End Function

' Given YouTube Html page, the function extracts the title from <title> tag
Function GetVideoTitle(Html)
    ' get rid of all tabs
    Html = Replace(Html, Chr(9), "")

    ' get rid of all newlines (vbscript regex engine doesn't like them)
    Html = Replace(Html, vbCrLf, "")
    Html = Replace(Html, vbLf, "")
    Html = Replace(Html, vbCr, "")

    GetVideoTitle = ExtractMatch(Html, "<title>YouTube ?- ?([^<]+)<")
End Function

' Given the Title of a video, function creates a usable filename for a video by
' sanitizing it - stripping parenthesis, changing non alphanumeric characters
' to _ and adding .flv extension
Function MkFileName(Title)
    Title = Replace(Title, "(", "")
    Title = Replace(Title, ")", "")

    Dim Regex
    Set Regex = New RegExp
    With Regex
        .Pattern = "[^A-Za-z0-9-_]"
        .Global = True
    End With

    Title = Regex.Replace(Title, "_")
    MkFileName = Title & ".flv"
End Function

' Given Text and a regular expression Pattern, the function extracts
' the first submatch
Function ExtractMatch(Text, Pattern)
    Dim Regex, Matches

    Set Regex = New RegExp
    Regex.Pattern = Pattern

    Set Matches = Regex.Execute(Text)
    If Matches.Count = 0 Then
        ExtractMatch = ""
        Exit Function
    End If

    ExtractMatch = Matches(0).SubMatches(0)
End Function

' Function saves Data to FileName
Function SaveVideo(FileName, Data)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  
  Stream.Type = adTypeBinary
  Stream.Open
  Stream.Write Data
  Stream.SaveToFile FileName, adSaveCreateOverWrite
End Function

'
' ==========================================================================
' The following code saves binary data to file using FileSystemObject
' It is so slow that even on a 3.2Ghz computer saving 1 MB takes 10 minutes!
' Don't use it! I put it here just to illustrate the wrong solution!
' ==========================================================================
'

' Given a Filename and Data, the function saves Data to File
'Sub SaveVideo(File, Data)
'    Dim Fso: Set Fso = CreateObject("Scripting.FileSystemObject")
'    Dim TextStream: Set TextStream = Fso.CreateTextFile(File, True)
'
'    WScript.Echo LenB(Data)
'    TextStream.Write BinaryToString(Data)
'End Sub

' Given Binary data, converts it to a string
'Function BinaryToString(Binary)
'  Dim I, S
'  For I = 1 To LenB(Binary)
'    S = S & Chr(AscB(MidB(Binary, I, 1)))
'  Next
'  BinaryToString = S
'End Function


'
' ==========================================================================
' The following is an implementation of UrlUnescape. It turned out VBScript
' has Unescape() function built in already, that does it!
'
'Function UrlUnescape(Str)
'    Dim Regex, Match, Matches
'
'    Set Regex = New RegExp
'    With Regex
'        .Pattern = "%([0-9a-f][0-9a-f])"
'        .IgnoreCase = True
'        .Global = True
'    End With
'    ' Wanted to do this, but it wasn't quite possible
'    ' UrlUnescape = Regex.Replace(Str, Chr(CInt("&H" & $0)))
'
'    Set Matches = Regex.Execute(Str)
'    For Each Match in Matches
'        Str = Replace(Str, Match, Chr(CInt("&H" & Match.SubMatches(0))))
'    Next
'
'    UrlUnescape = Str
'End Function

