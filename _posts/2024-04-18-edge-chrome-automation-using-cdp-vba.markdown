---
layout: post
title:  "Edge/Chrome automation with CDP on VBA"
date:   2024-04-18 16:36:00 +0900
categories: vba cdp chrome edge automation
---

Automating Internet Explorer via Excel/VBA for scraping or any task that could be automated was quite easy because it was ActiveX based and there were an exclusive ActiveX control to do so.

As the end of suport of IE, this became a problem because neither Edge or Chrome has this flexibility for scripting. ChrisK presented a solution for that at codeproject.com using only VBA interfacing with the Chrome DevTools Protocol a.k.a. CDP.

Here below are the necessary code to make it work.

Class Module: clsEdge.cls
{% highlight vb %}
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Basic implementation of Chrome Devtools Protocol (CDP)
' using VBA






Private objBrowser As clsExec

' every message sent to edge has an id
Private lngLastID As Long

Private strSessionId As String

'this buffer holds messages that are not yet processed
Private strBuffer As String

Public Function serialize() As String

    Dim objSerialize As Dictionary
    Set objSerialize = New Dictionary
    Call objSerialize.Add("objBrowser", objBrowser.serialize())
    Call objSerialize.Add("lngLastID", lngLastID)
    Call objSerialize.Add("strSessionId", strSessionId)
    serialize = ConvertToJson(objSerialize)


End Function

Public Sub deserialize(strSerialized As String)
    Dim objSerialize As Dictionary
    Set objSerialize = ParseJson(strSerialized)
    
    Set objBrowser = New clsExec
    Call objBrowser.deserialize(objSerialize.Item("objBrowser"))
    
    lngLastID = objSerialize.Item("lngLastID")
    strSessionId = objSerialize.Item("strSessionId")

End Sub

' CDP messages received from chrome are null-terminated
' It seemed to me you cant search for vbnull in a string
' in vba. Thats why i re-implemented the search function

Private Function searchNull() As Long
    Dim i As Long
    
    Dim lngBufferLength As Long
    lngBufferLength = Len(strBuffer)
    searchNull = 0
    
    If lngBufferLength = 0 Then
        Exit Function
    End If
    
    For i = 1 To lngBufferLength
        If Mid(strBuffer, i, 1) = vbNullChar Then
            searchNull = i
            Exit Function
        End If
    Next i

End Function

Private Function sendMessage(strMessage As String, Optional objAllMessages As Dictionary) As Dictionary

    Dim intRes As Long
    Dim strRes As String
    
    Dim lngCurrentId As Long
    lngCurrentId = lngLastID
    
    'We increase the global ID counter
    lngLastID = lngLastID + 1
    
    
    ' Before sending a message the messagebuffer is emptied
    ' All messages that we have received sofar cannot be an answer
    ' to the message that we will send
    ' So they can be safely discarded
    
    If objAllMessages Is Nothing Then
        intRes = 1
        Do Until intRes < 1
            intRes = objBrowser.readProcCDP(strRes)
            
            If intRes > 0 Then
                strBuffer = strBuffer & strRes
            End If
        Loop
        
        Dim lngNullCharPos As Long
        lngNullCharPos = searchNull()
        
        Do Until lngNullCharPos = 0
        
            'Debug.Print (Left(strBuffer, lngNullCharPos))
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
                        
            lngNullCharPos = searchNull()
    
        Loop
    End If

    ' sometimes edge writes to stdout
    ' we clear stdout here, too.
    
    intRes = objBrowser.readProcSTD(strRes)
    

    ' We add the currentID and sessionID to the message
    
    strMessage = Left(strMessage, Len(strMessage) - 1)
    
    If strSessionId <> "" Then
        strMessage = strMessage & ", ""sessionId"":""" & strSessionId & """"
    End If
    
    strMessage = strMessage & ", ""id"":" & lngCurrentId & "}" & vbNullChar
    
    Call objBrowser.writeProc(strMessage)
    
    
    
    ' We have some failsafe counter in order not to
    ' loop forever
    Dim intCounter As Integer
    intCounter = 0
    
    
    Do Until intCounter > 1000
        intRes = 1
        ' We read from edge and process messages until we receive a
        ' message with our ID
        Do Until intRes < 1
            intRes = objBrowser.readProcCDP(strRes)
            
            If intRes > 0 Then
                strBuffer = strBuffer & strRes
            End If
        Loop
        
        lngNullCharPos = searchNull()
        
        Do Until lngNullCharPos = 0
        
            
            strRes = Left(strBuffer, lngNullCharPos - 1)
            
            strBuffer = Right(strBuffer, Len(strBuffer) - lngNullCharPos)
                        
            Dim objRes As Dictionary
            Dim objDic2 As Dictionary
            Dim objDic3 As Dictionary
            
            Dim boolFound As Boolean
            boolFound = False
            
            'Debug.Print (strRes)
            
            
            If strRes <> "" Then
        
                'Set objRes = getJSONCollection(strRes)
                Set objRes = JsonConverter.ParseJson(strRes)
                
                If Not (objAllMessages Is Nothing) Then
                    objAllMessages.Add CStr(objAllMessages.Count), objRes
                End If
                
                
                If objRes.Exists("id") Then
        
                
                    If objRes.Item("id") = lngCurrentId Then
                        Set sendMessage = objRes
                        Exit Function
                    End If
                End If
            End If
            lngNullCharPos = searchNull()

        Loop
        DoEvents
        Call Sleep(0.1)
        intCounter = intCounter + 1
    Loop
    
    Debug.Print ("-----")
    Debug.Print ("timeout")
    Debug.Print (strMessage)
    Debug.Print ("-----")
    
    Call err.Raise(-900, , "timeout")


End Function

' This function allows to evaulate a javascript expression
' in the context of the  page
Public Function jsEval(strString As String, Optional boolRetry = True) As Variant

    Dim objRes As Dictionary
    
    Dim strMessage As String
    
    strMessage = "{""method"":""Runtime.evaluate"",""params"":{""expression"":""1+1;""}}"
    
    Dim objMessage As Dictionary
    Set objMessage = JsonConverter.ParseJson(strMessage)
    objMessage.Item("params").Item("expression") = strString & ";"
    
    strMessage = JsonConverter.ConvertToJson(objMessage)
    
    Set objRes = sendMessage(strMessage)
    
    If objRes Is Nothing And boolRetry Then
        Stop
        Set objRes = sendMessage(strMessage)
    End If
    
    
    If (objRes.Exists("error")) Then
        ' Oops, there was an error in out javascript expression
        Stop
        Exit Function
    End If
    
    ' If the return type has a specific type
    ' we can return the result
    
    If objRes.Item("result").Item("result").Item("type") = "string" Or objRes.Item("result").Item("result").Item("type") = "boolean" Or objRes.Item("result").Item("result").Item("type") = "number" Then
        jsEval = objRes.Item("result").Item("result").Item("value")
    End If

End Function


' This function must be calles after start and before all other methods
' This function attaches to a session of the browser
Public Function attach(strUrl As String) As Integer
    

    Dim objRes As Dictionary
    
    Dim objAllMessages As Dictionary
    Set objAllMessages = New Dictionary
     
    Set objRes = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":true}}", objAllMessages)
    
    Dim i As Integer
    Dim boolFound As Boolean
    
    Dim strKey As Variant
    
    Dim objDic2 As Dictionary
    Dim objDic3 As Dictionary
    
    
    For Each strKey In objAllMessages.Keys
    
        Set objRes = objAllMessages.Item(strKey)
        
        
        If Not objRes.Exists("params") Then GoTo nextloop1
        Set objDic2 = objRes.Item("params")
        
        If Not objDic2.Exists("targetInfo") Then GoTo nextloop1
        Set objDic3 = objDic2.Item("targetInfo")
        
        If objDic3.Item("type") <> "page" Then GoTo nextloop1
        
        If objDic3.Item("url") <> strUrl And strUrl <> "" Then
            GoTo nextloop1
        End If
        
        boolFound = True
        Exit For

nextloop1:
    Next strKey
    
    If Not boolFound Then
        attach = -1
        Exit Function
    End If
    
    'Stop
    
    Set objRes = sendMessage("{""method"":""Target.attachToTarget"",""params"":{""targetId"":""" & objDic3.Item("targetId") & """,""flatten"":true}}")
    
    strSessionId = objRes.Item("result").Item("sessionId")
    
    Set objRes = sendMessage("{""method"":""Runtime.enable"",""params"":{}}")
    
    Set objRes = sendMessage("{""method"":""Target.setDiscoverTargets"",""params"":{""discover"":false}}")

    
    attach = 0
    
    Call Sleep
    

End Function

' This function makes edhe naviagte to a given URL
Public Sub navigate(strUrl As String)
    Dim objRes As Dictionary
    Set objRes = sendMessage("{""method"":""Page.navigate"",""params"":{""url"":""" & strUrl & """}}")
    Call Sleep
End Sub



' This method starts up the browser
Public Sub start(Optional boolSerializable As Boolean = False)
    Set objBrowser = New clsExec
    
    Dim strCall As String
    strCall = """C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"" --remote-debugging-pipe --enable-automation --enable-logging"
    
    Dim intRes As Integer
    
    intRes = objBrowser.init(strCall, boolSerializable)
    
    If intRes <> 0 Then
       Call err.Raise(-99, , "error start browser")
    End If

    Call Sleep
    lngLastID = 1
    
    Dim strRes As String
    
    intRes = 0
    
    Dim intCounter As Integer
    intCounter = 0
    
    Do Until intRes > 0 Or intCounter > 1000
        intRes = objBrowser.readProcSTD(strRes)
        DoEvents
        Call Sleep(0.1)
        intCounter = intCounter + 1
    Loop
    
End Sub

Public Sub closeBrowser()

    Dim objRes As Dictionary
    
    On Error Resume Next
    Set objRes = sendMessage("{""method"":""Browser.close"",""params"":{}}")
    
    'it seems without waitng a bit the browser crashes and the next time wants ro recover from a crash
    Call Sleep(5)

End Sub

Public Function connectionAlive() As Boolean
    On Error GoTo err
    Dim strLoc As String
    strLoc = jsEval("window.location.href")
    
    connectionAlive = True
    Exit Function
    
err:

    connectionAlive = False
    
End Function

Public Sub waitCompletion()
    Dim strState As String
    strState = "x"
    Call Sleep(1)
    Do Until strState = "complete"
        strState = Me.jsEval("document.readyState")
        Call Sleep(1)
    Loop
    
End Sub

{% endhighlight %}

Class Module: clsExec.cls
{% highlight vb %}
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' derived from  https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit



Private hStdOutWr As LongPtr
Private hStdOutRd As LongPtr
Private hStdInWr As LongPtr
Private hStdInRd As LongPtr
Private hCDPOutWr As LongPtr
Private hCDPOutRd As LongPtr
Private hCDPInWr As LongPtr
Private hCDPInRd As LongPtr
Private boolSerializable As Boolean

Private hProcess As LongPtr

Public Function serialize() As String

    If Not boolSerializable Then
        Call err.Raise(-904, , "this instance is not serializable")
    End If

    Dim objSerialize As Dictionary
    Set objSerialize = New Dictionary
    Call objSerialize.Add("hStdOutRd", hStdOutRd)
    Call objSerialize.Add("hStdInWr", hStdInWr)
    Call objSerialize.Add("hCDPOutRd", hCDPOutRd)
    Call objSerialize.Add("hCDPInWr", hCDPInWr)
    serialize = ConvertToJson(objSerialize)


End Function

Public Sub deserialize(strSerialized As String)
    Dim objSerialize As Dictionary
    Set objSerialize = ParseJson(strSerialized)
    
    hStdOutRd = objSerialize.Item("hStdOutRd")
    hStdInWr = objSerialize.Item("hStdInWr")
    hCDPOutRd = objSerialize.Item("hCDPOutRd")
    hCDPInWr = objSerialize.Item("hCDPInWr")

End Sub

Public Function init(strExec As String, Optional aboolSerializable As Boolean = False) As Integer
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As LongPtr, hWritePipe As LongPtr
    Dim L As Long, result As Long, bSuccess As Long
    Dim Buffer As String
    Dim k As Long
    
    Dim pipes As STDIO_BUFFER
    Dim pipes2 As STDIO_BUFFER2
    
    boolSerializable = aboolSerializable


    ' First we create all 4 pipes
    
    ' We start with stdout of the edge process
    ' This pipe is used for stderr, too
    
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    result = CreatePipe(hStdOutRd, hStdOutWr, sa, 0)
    
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then stdin
    
    result = CreatePipe(hStdInRd, hStdInWr, sa, 0)


    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then the out pipe for the CDP Protocol
    
    result = CreatePipe(hCDPOutRd, hCDPOutWr, sa, 2 ^ 20)
    
    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' and finally the in pipe

    
    result = CreatePipe(hCDPInRd, hCDPInWr, sa, 0)


    If result = 0 Then
        init = -2
        Exit Function
    End If
    
    ' then we fill the special structure for passing arbitrary pipes (i.e. fds)
    ' to a process
    
    pipes.number_of_fds = 5
    
    pipes.os_handle(0) = hStdInRd
    pipes.os_handle(1) = hStdOutWr
    pipes.os_handle(2) = hStdOutWr
    pipes.os_handle(3) = hCDPInRd
    pipes.os_handle(4) = hCDPOutWr
    
    pipes.crt_flags(0) = 9
    pipes.crt_flags(1) = 9
    pipes.crt_flags(2) = 9
    pipes.crt_flags(3) = 9
    pipes.crt_flags(4) = 9
    
    ' pipes2 is filled by copying memory from pipes
    
    pipes2.number_of_fds = pipes.number_of_fds
    
    Call MoveMemory(pipes2.raw_bytes(0), pipes.crt_flags(0), 5)
    Call MoveMemory(pipes2.raw_bytes(5), pipes.os_handle(0), UBound(pipes2.raw_bytes) - 4)


    With start
        .cb = Len(start)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .hStdOutput = hStdOutWr
        .hStdInput = hStdInRd
        .hStdError = hStdOutWr
        .wShowWindow = vbHide ' hide console window, seems not to work
        .cbReserved2 = Len(pipes2)
        .lpReserved2 = VarPtr(pipes2)
    End With
    

    result = CreateProcessA(0&, strExec, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    If result = 0 Then
        init = -1
    End If
    
    ' We close the sides of the handles that we dont need anymore
    
    Call CloseHandle(hStdOutWr)
    Call CloseHandle(hStdInRd)
    Call CloseHandle(hCDPOutWr)
    Call CloseHandle(hCDPInRd)
    
'    Call Sleep(10)
'
'    EnumThreadWindows proc.dwThreadId, AddressOf EnumThreadWndProc, 0

    hProcess = proc.hProcess
    
    init = 0
    
End Function

' This function tries to read from the CDP out pipe
' Reading is non-blocking, if there are no bytes ro read the function returns 0
' otherwise the number of bytes read
Public Function readProcCDP(ByRef strData As String) As Long
    Dim lPeekData As Long
    
    Dim lngRes As Long
    
    lngRes = PeekNamedPipe(hCDPOutRd, ByVal 0&, 0&, ByVal 0&, _
        lPeekData, ByVal 0&)
        
    If lngRes = 0 Then
        Call err.Raise(901, , "Error PeekNamedPipe in readProcCDP")
    End If
    
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hCDPOutRd, Buffer, Len(Buffer), L, 0&)
        If bSuccess = 1 Then
            strData = Buffer
            
            readProcCDP = Len(strData)
        Else
            readProcCDP = -2
        End If
    Else
        readProcCDP = -1
    End If

End Function

' Same as ReadProcCDP

Public Function readProcSTD(ByRef strData As String) As Integer
    Dim lPeekData As Long
    
    Call PeekNamedPipe(hStdOutRd, ByVal 0&, 0&, ByVal 0&, _
    lPeekData, ByVal 0&)
    
    
    If lPeekData > 0 Then
        Dim Buffer As String
        Dim L As Long
        Dim bSuccess As Long
        Buffer = Space$(lPeekData)
        bSuccess = ReadFile(hStdOutRd, Buffer, Len(Buffer), L, 0&)
        If bSuccess = 1 Then
            strData = Buffer
            readProcSTD = Len(strData)
        Else
            readProcSTD = -2
        End If
    Else
        readProcSTD = -1
    End If

End Function

' This functions sends a CDP message to edge

Public Function writeProc(ByVal strData As String) As Integer
    Dim lngWritten As Long
    writeProc = WriteFile(hCDPInWr, ByVal strData, Len(strData), lngWritten, ByVal 0&)

End Function

Private Sub Class_Terminate()

    If boolSerializable Then Exit Sub

    Call CloseHandle(hStdOutRd)
    Call CloseHandle(hStdOutWr)
    Call CloseHandle(hStdInRd)
    Call CloseHandle(hStdInWr)

    Call CloseHandle(hCDPOutRd)
    Call CloseHandle(hCDPOutWr)
    Call CloseHandle(hCDPInRd)
    Call CloseHandle(hCDPInWr)
    
End Sub
{% endhighlight %}

Module: JsonConverter.bas
{% highlight vb %}
Attribute VB_Name = "JsonConverter"
'from https://github.com/VBA-tools/VBA-JSON


'Attribute VB_Name = "JsonConverter"
''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        ' Error: Invalid JSON string
        err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ConvertToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))

        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ConvertToJson = JsonValue
        Else
            ConvertToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
            End If
        End If

        ' Array
        json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength

        On Error Resume Next

        json_LBound = LBound(JsonValue, 1)
        json_UBound = UBound(JsonValue, 1)
        json_LBound2D = LBound(JsonValue, 2)
        json_UBound2D = UBound(JsonValue, 2)

        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    ' Append comma to previous line
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    ' 2D Array
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If
                    json_BufferAppend json_Buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength

                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If

                        json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                json_Converted = "null"
                            End If
                        End If

                        If json_PrettyPrint Then
                            json_Converted = vbNewLine & json_InnerIndentation & json_Converted
                        End If

                        json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                    Next json_Index2D

                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If

                    json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    ' 1D Array
                    json_Converted = ConvertToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_Converted = "" Then
                        ' (nest to only check if converted = "")
                        If json_IsUndefined(JsonValue(json_Index)) Then
                            json_Converted = "null"
                        End If
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If

        On Error GoTo 0

        If json_PrettyPrint Then
            json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
            Else
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If
        End If

        json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength

        ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)

    ' Dictionary or Collection
    Case VBA.vbObject
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
            End If
        End If

        ' Dictionary
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In JsonValue.Keys
                ' For Objects, undefined (Empty/Nothing) is not added to object
                json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" Then
                    json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                Else
                    json_SkipItem = False
                End If

                If Not json_SkipItem Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                    Else
                        json_Converted = """" & json_Key & """:" & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Key

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

        ' Collection
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1)

                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                If json_Converted = "" Then
                    ' (nest to only check if converted = "")
                    If json_IsUndefined(json_Value) Then
                        json_Converted = "null"
                    End If
                End If

                If json_PrettyPrint Then
                    json_Converted = vbNewLine & json_Indentation & json_Converted
                End If

                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
            Next json_Value

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
        End If

        ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ConvertToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ConvertToJson = JsonValue
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
        Case 34
            ' " -> 34 -> \"
            json_Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_Char = "\\"
        Case 47
            ' / -> 47 -> \/ (optional)
            If JsonOptions.EscapeSolidus Then
                json_Char = "\/"
            End If
        Case 8
            ' backspace -> 8 -> \b
            json_Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            ' Non-ascii characters -> convert to 4-digit hex
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

''
' VBA-UTC v1.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & err.Number & " - " & err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & err.Number & " - " & err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling

    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")

                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)

        If utc_HasOffset Then
            ParseIso = ParseIso - utc_Offset
        End If
    End If

    Exit Function

utc_ErrorHandling:
    err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & err.Number & " - " & err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & err.Number & " - " & err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then
'
'Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
'    Dim utc_ShellCommand As String
'    Dim utc_Result As utc_ShellResult
'    Dim utc_Parts() As String
'    Dim utc_DateParts() As String
'    Dim utc_TimeParts() As String
'
'    If utc_ConvertToUtc Then
'        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
'            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
'            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
'    Else
'        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
'            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
'            "+'%Y-%m-%d %H:%M:%S'"
'    End If
'
'    utc_Result = utc_ExecuteInShell(utc_ShellCommand)
'
'    If utc_Result.utc_Output = "" Then
'        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
'    Else
'        utc_Parts = Split(utc_Result.utc_Output, " ")
'        utc_DateParts = Split(utc_Parts(0), "-")
'        utc_TimeParts = Split(utc_Parts(1), ":")
'
'        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
'            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
'    End If
'End Function
'
'Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
'#If VBA7 Then
'    Dim utc_File As LongPtr
'    Dim utc_Read As LongPtr
'#Else
'    Dim utc_File As Long
'    Dim utc_Read As Long
'#End If
'
'    Dim utc_Chunk As String
'
'    On Error GoTo utc_ErrorHandling
'    utc_File = utc_popen(utc_ShellCommand, "r")
'
'    If utc_File = 0 Then: Exit Function
'
'    Do While utc_feof(utc_File) = 0
'        utc_Chunk = VBA.Space$(50)
'        utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
'        If utc_Read > 0 Then
'            utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
'            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
'        End If
'    Loop
'
'utc_ErrorHandling:
'    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
'End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If



{% endhighlight %}

Module: modEdge.cls
{% highlight vb %}
Attribute VB_Name = "modEdge"
Option Explicit

'This is anexample of how to use the classes
Sub runedge()

    'Start Browser
    Dim objBrowser As clsEdge
    Set objBrowser = New clsEdge
    Call objBrowser.start
    
    'Attach to any ("") or a specific page
    Call objBrowser.attach("")
    
    'navigate
    Call objBrowser.navigate("https://google.de")
    
    Call objBrowser.waitCompletion
    
    'evaluate javascript
    Call objBrowser.jsEval("alert(""hi"")")
    
    'fill search form (textbox is named q)
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].value=""automate edge vba""")
    
    'run search
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].form.submit()")
    
    'wait till search has finished
    Call objBrowser.waitCompletion
    

    'click on codeproject link
    Call objBrowser.jsEval("document.evaluate("".//h3[text()='Automate Chrome / Edge using VBA - CodeProject']"", document).iterateNext().click()")
    
    Call objBrowser.waitCompletion
    
    Dim strVotes As String
    strVotes = objBrowser.jsEval("ctl00_RateArticle_VountCountHist.innerText")
    
    MsgBox ("finish! Vote count is " & strVotes)
    
    objBrowser.closeBrowser
    
    
End Sub


'the following two snippets show the serialization of the object
Sub runedge2()

    'Start Browser
    Dim objBrowser As clsEdge
    Set objBrowser = New clsEdge
    Call objBrowser.start(True)
    
    'Attach to any ("") or a specific page
    Call objBrowser.attach("")
    
    'navigate
    Call objBrowser.navigate("https://google.de")
    
    'evaluate javascript
    Call objBrowser.jsEval("alert(""hi"")")
    
    MsgBox ("finish1!")
    
    Dim strSerialized As String
    strSerialized = objBrowser.serialize()
    Tabelle1.Cells(1, 1) = strSerialized
End Sub

Sub runedge3()
    
    Dim objBrowser2 As clsEdge
    Set objBrowser2 = New clsEdge
    
    
    Call objBrowser2.deserialize(Tabelle1.Cells(1, 1))
    
    If Not objBrowser2.connectionAlive Then Stop
    
    Call objBrowser2.jsEval("alert(""hi again"")")
   
    MsgBox ("finish2!")
    
    
End Sub


{% endhighlight %}

Module: modExec.cls
{% highlight vb %}
Attribute VB_Name = "modExec"
Option Explicit

'from https://stackoverflow.com/questions/62172551/error-with-createpipe-in-vba-office-64bit


Public Declare PtrSafe Function CreatePipe Lib "kernel32" ( _
    phReadPipe As LongPtr, _
    phWritePipe As LongPtr, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, _
    ByVal nSize As Long) As Long

Public Declare PtrSafe Function ReadFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Public Declare PtrSafe Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As Any, _
    lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As Any, _
    lpProcessInformation As Any) As Long

Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As Long

Public Declare PtrSafe Function PeekNamedPipe Lib "kernel32" ( _
    ByVal hNamedPipe As LongPtr, _
    lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    lpBytesRead As Long, _
    lpTotalBytesAvail As Long, _
    lpBytesLeftThisMessage As Long) As Long


Declare PtrSafe Function WriteFile Lib "kernel32" ( _
    ByVal hFile As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Long) As Long
    
Declare PtrSafe Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As LongPtr, ByVal lParam As LongPtr) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type


Public Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

'this is the structure to pass more than 3 fds to a child process

'see https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c
Public Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type

' the fields crt_flags and os_handle must lie contigously in memory
' i.e. should not be aligned to byte boundaries
' you cannot define a packed struct in VBA
' thats why we need to have a second struct

#If Win64 Then
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 44) As Byte
End Type
#Else
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 24) As Byte
End Type
#End If

Public Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Const STARTF_USESTDHANDLES = &H100&
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const STARTF_USESHOWWINDOW As Long = &H1&
'Public Const STARTF_CREATE_NO_WINDOW As Long = &H8000000

' we need to move memory

Public Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

' Using the following defintions I tried to hide the console windows
' This does not yet work, so it is commented out
'
'Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'
' Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Boolean
'
'Public Function getNameFromHwnd(hWnd As Long) As String
'Dim title As String * 255
'Dim tLen As Long
'tLen = GetWindowTextLength(hWnd)
'GetWindowText hWnd, title, 255
'getNameFromHwnd = Left(title, tLen)
'End Function
'
'
'
'
'Public Function EnumThreadWndProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'    Dim Ret As Long, sText As String
'
'    'CloseWindow hwnd ' This is the handle to your process window which you created.
'
'    sText = getNameFromHwnd(hWnd)
'
'    If sText = "" Then
'        Call ShowWindow(hWnd, 0)
'    End If
'
'    EnumThreadWndProc = 1
'
'End Function


{% endhighlight %}

Module: modSleep.cls
{% highlight vb %}
Attribute VB_Name = "modSleep"
Option Explicit

Private Declare PtrSafe Sub sleep2 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)



'Custom sleep function
'change sleep period if processing is not robust
Public Const cnlngSleepPeriod As Long = 1000

Public Sub Sleep(Optional dblFrac As Double = 1)
    DoEvents
    Call sleep2(cnlngSleepPeriod * dblFrac)
    DoEvents
End Sub
    
{% endhighlight %}

Reference: [Automate Chrome / Edge using VBA by ChrisK][codeproject]

[codeproject]: https://www.codeproject.com/script/Articles/ViewDownloads.aspx?aid=5307593