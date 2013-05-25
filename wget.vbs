'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' wget.vbs
'' v1.0, 2012-03-30, 16:33:49 CST
'' 
'' AUTHOR: Mark R. Gollnick <mark.r.gollnick@gmail.com>
'' HOMEPAGE: http://home.engineering.iastate.edu/~mrgoll12/
''
'' DESCRIPTION:
''   An HTTP file downloader similar to WGET for Windows
''   Visual Basic Scripting engines (cscript.exe or wscript.exe).
''
'' USAGE:
''   cscript wget.vbs <url> [save_to_file] [[/Y]|[/NC]]
''   wscript wget.vbs <url> [save_to_file] [[/Y]|[/NC]]
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' MAIN
If WScript.Arguments.Count < 1 Then
  WScript.Echo(_
    "VBS Wget 1.00, a non-interactive HTTP retriever." & vbCrLf & _
    "Usage: [c|w]script " & WScript.ScriptName & " <url> [save_to_file] " & _
    "[[/Y]|[/NC]]" & vbCrLf & vbCrLf & _
    "    /Y     Suppresses prompting to confirm you want to overwrite an " & _
    "existing destination file." & vbCrLf & _
    "    /NC    No clobber: skip downloads that would download to " & _
    "existing files (overwriting them)." & vbCrLf _
  )
  WScript.Quit
Else '' WScript.Arguments.Count >= 1
  Dim arg0, arg1, arg2
  arg0 = WScript.Arguments(0)
  arg1 = ""
  arg2 = ""
  If WScript.Arguments.Count = 1 Then
    Call HttpGet(arg0, "", False)
  Else '' WScript.Arguments.Count >= 2
    arg1 = WScript.Arguments(1)
    Dim boolOverwrite
    boolOverwrite = -1 '' Don't know; prompt user
    If WScript.Arguments.Count = 3 Then
      arg2 = WScript.Arguments(2)
      If ((InStrRev(arg2, "/Y") = 1) And (Len(arg2) = 2)) Then
        boolOverwrite = 1 '' Yes; overwrite
      ElseIf ((InStrRev(arg2, "/NC") = 1) And (Len(arg2) = 3)) Then
        boolOverwrite = 0 '' No; don't clobber
      End If
    End If
    Call HttpGet(arg0, arg1, boolOverwrite)
  End If
  WScript.Quit
End If

'' This Subroutine does all the work.
Sub HttpGet(strUrl, strSaveToFileName, boolOverwrite)
  Dim strWorkingDir, strUrlFileName
  
  '' Stores the working directory
  strWorkingDir = _
    Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
  
  '' Stores the name of the file to download
  strUrlFileName = _
    Mid(strUrl, InStrRev(strUrl, "/") + 1)
  
  '' Instantiate necessary objects
  Dim NewFile
  Dim FileObject:  Set FileObject  = CreateObject("Scripting.FileSystemObject")
  Dim HttpRequest: Set HttpRequest = CreateObject("MSXML2.XMLHTTP")
  Dim FileStream:  Set FileStream  = CreateObject("ADODB.Stream")
  
  '' Determine file write location
  If strSaveToFileName = "" Then
    NewFile = FileObject.BuildPath(strWorkingDir, strUrlFileName)
  Else
    If InStrRev(strSaveToFileName, "\") = 0 Then
      NewFile = FileObject.BuildPath(strWorkingDir, strSaveToFileName)
    Else
      NewFile = _
        FileObject.BuildPath( _
          Left(strSaveToFileName, InStrRev(strSaveToFileName, "\") - 1), _
          Mid(strSaveToFileName,  InStrRev(strSaveToFileName, "\") + 1) _
        )
    End If
  End If
  
  '' Determine if file already exists, if so, determine if user overwrites
  If FileObject.FileExists(NewFile) Then
    If boolOverwrite = -1 Then
      boolOverwrite = YesNoPrompt("WARNING", "File exists! Overwrite?")
    End If
    If boolOverwrite = 1 Then
      FileObject.DeleteFile(NewFile)
    Else
      Exit Sub
    End If
  ElseIf FileObject.FolderExists(NewFile) Then
    NewFile = FileObject.BuildPath(NewFile, strUrlFileName)
  End If
  
  '' Create and send the HTTP Request header
  Call HttpRequest.Open("GET", strUrl, False)
  HttpRequest.Send
  
  '' If the HTTP Response comes back 200 ("OK"), then save results to a file
  If HttpRequest.Status = 200 Then
    With FileStream
      .Type = 1                        '' stream is a binary file
      .Open                            '' open stream for writing
      .Write HttpRequest.ResponseBody  '' write http response to stream
      .SaveToFile NewFile, 2           '' save stream to file (overwrite)
      .Close                           '' close stream
    End With
  End If
  
  '' Destroy objects
  Set FileStream = Nothing
  Set HttpRequest = Nothing
  Set FileObject = Nothing
  Set NewFile = Nothing
End Sub

'' YesNoPrompt
'' Written by Rob van der Woude
'' http://www.robvanderwoude.com
'' Modified by Mark R. Gollnick
'' http://home.eng.iastate.edu/~mrgoll12
Function YesNoPrompt(header, text)
  Dim result
  
  '' If running from command line (cscript.exe)... use command line interface
  If UCase(Right(WScript.FullName, 12)) = "\CSCRIPT.EXE" Then
    WScript.StdOut.Write(header & ": " & text & " (y/n) ")
    result = WScript.StdIn.ReadLine
    If (result = "y" Or result = "Y") Then
      result = True
    Else
      result = False
    End If
    
  '' If running from windows (wscript.exe)... use graphical user interface
  Else
    result = MsgBox(text, vbYesNo, header)
    If result = vbYes Then
      result = True
    Else
      result = False
    End If
  End If
  
  '' Return user's response
  YesNoPrompt = result
End Function
