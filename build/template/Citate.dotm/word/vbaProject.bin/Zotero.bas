Attribute VB_Name = "Zotero"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Option Explicit

' Version of the Citate VBA template code, stamped on every log line. Bump
' whenever the .bas sources change so logs identify which build produced them.
Public Const CITATE_VBA_VERSION As String = "1.0.0"

Private Const CP_UTF8 = 65001
Private Const WM_COPYDATA = &H4A

' Message-queue constants for the macro-pill keystroke buffer
Private Const PM_REMOVE = &H1
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYDOWN = &H100
Private Const WM_CHAR = &H102
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_KEYLAST = &H109

Private Type POINTAPI
    x As Long
    y As Long
End Type
#If VBA7 Then
    Global ZotWnd As LongPtr
#Else
    Global ZotWnd As Long
#End If
Global IsZotero7 As Boolean

#If VBA7 Then
    Type COPYDATASTRUCT
        dwData As LongPtr
        cbData As Long
        lpData As LongPtr
    End Type
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Boolean
    Private Declare PtrSafe Function EnumThreadWindows Lib "user32" _
        (ByVal dwThreadId As Long, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Boolean
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
        ByVal dwflags As Long, ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
    Private Type MSGTYPE
        hwnd As LongPtr
        message As Long
        wParam As LongPtr
        lParam As LongPtr
        time As Long
        pt As POINTAPI
    End Type
    Private Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageW" _
        (lpMsg As MSGTYPE, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, _
        ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Private Declare PtrSafe Function TranslateMessage Lib "user32" _
        (lpMsg As MSGTYPE) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function AllowSetForegroundWindow Lib "user32" _
        (ByVal dwProcessId As Long) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
    End Type
    Private Declare Function FindWindow Lib "user32" Alias _
        "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias _
        "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
        ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
        wParam As Long, lParam As Any) As Integer
    Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As Long) As Boolean
    Private Declare Function EnumThreadWindows Lib "user32" _
        (ByVal dwThreadId As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Private Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
        ByVal dwflags As Long, ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, lpMultiByteStr As Any, _
        ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
    Private Type MSGTYPE
        hwnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
    End Type
    Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" _
        (lpMsg As MSGTYPE, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, _
        ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
    Private Declare Function TranslateMessage Lib "user32" _
        (lpMsg As MSGTYPE) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function AllowSetForegroundWindow Lib "user32" _
        (ByVal dwProcessId As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' Runs when Word loads the template. The keybinding is session-only: it is re-added
' on every load and the template is immediately marked saved, so Word neither
' prompts to save the template nor persists the binding into the .dotm file.
Public Sub AutoExec()
    On Error Resume Next
    Call CitateLog("template loaded (Word " & Application.Version & ")")
    CustomizationContext = ThisDocument
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="ZoteroExpandMacrosOrCite", _
        KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyZ)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
        Command:="CitateMacroPillTrigger", _
        KeyCode:=BuildKeyCode(wdKeyBackSlash)
    ThisDocument.Saved = True
End Sub

' "\" typed in the document: ask Citate to show the inline macro-entry pill at the
' caret, then capture any keystrokes typed before the pill takes keyboard focus and
' forward them. Falls back to typing a literal backslash when the feature is
' disabled for this document or Citate is not running. No state outlives this
' handler: the buffer loop below is hard-limited by a timeout.
Public Sub CitateMacroPillTrigger()
    Dim x As Long, y As Long, w As Long, h As Long
    Dim pid As Long
    Call CitateLog("pill: trigger fired")
    On Error GoTo TypeLiteral
    If Not CitateInlineMacrosEnabled() Then
        Call CitateLog("pill: disabled for this document")
        GoTo TypeLiteral
    End If
    Call FindZoteroWindow
    Call CitateLog("pill: ZotWnd=" & ZotWnd)
    If ZotWnd = 0 Then GoTo TypeLiteral

    ' Caret position in screen pixels; the pill appears just below the caret line
    ActiveWindow.GetPoint x, y, w, h, Selection.Range
    Call CitateLog("pill: point x=" & x & " y=" & y & " h=" & h)

    ' Grant the Citate process the right to bring its pill window to the foreground
    Call GetWindowThreadProcessId(ZotWnd, pid)
    Call AllowSetForegroundWindow(pid)

    Call CitateLog("pill: sending showMacroPill")
    Call ZoteroCommand("showMacroPill", False, _
        " -ZoteroIntegrationMacroContext """ & x & "," & (y + h + 2) & """")

    Call CitateForwardPillBuffer
    Call CitateLog("pill: done")
    Exit Sub
TypeLiteral:
    Call CitateLog("pill: TypeLiteral err=" & Err.Number & " " & Err.Description)
    On Error Resume Next
    Selection.TypeText "\"
End Sub

' Appends a timestamped, version-stamped line to %TEMP%\citate-word.log.
' Best-effort: logging must never break the operation that called it. The log
' is wiped once it grows past 1 MB rather than rotated.
Public Sub CitateLog(msg As String)
    On Error Resume Next
    Dim logPath$
    Dim f As Integer
    logPath$ = Environ$("TEMP") & "\citate-word.log"
    If FileLen(logPath$) > 1048576 Then Kill logPath$
    f = FreeFile
    Open logPath$ For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " v" & CITATE_VBA_VERSION & " " & msg
    Close #f
End Sub

' While waiting for the pill to take keyboard focus, keystrokes queue on Word's UI
' thread (this handler blocks it). Drain them from our own message queue so they
' never reach the document, and forward them to the pill as char codes. Stops as
' soon as the foreground window changes (the pill took focus) or after a hard
' timeout, so no state can persist beyond this call.
Private Sub CitateForwardPillBuffer()
    Dim m As MSGTYPE
    Dim buffer$
    Dim startTick As Long
    #If VBA7 Then
        Dim wordWnd As LongPtr
    #Else
        Dim wordWnd As Long
    #End If
    On Error GoTo Send
    wordWnd = GetForegroundWindow()
    startTick = GetTickCount()
    Do
        Do While PeekMessage(m, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
            If m.message = WM_CHAR Then
                buffer = buffer & CStr(CLng(m.wParam And &HFFFF&)) & ","
            ElseIf m.message = WM_KEYDOWN Or m.message = WM_SYSKEYDOWN Then
                ' Posts the corresponding WM_CHAR, retrieved on the next pass
                Call TranslateMessage(m)
            End If
        Loop
        If GetForegroundWindow() <> wordWnd Then Exit Do
        If GetTickCount() - startTick > 600 Then Exit Do
        Sleep 5
    Loop
    ' Final sweep for messages translated but not yet retrieved
    Do While PeekMessage(m, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0
        If m.message = WM_CHAR Then
            buffer = buffer & CStr(CLng(m.wParam And &HFFFF&)) & ","
        ElseIf m.message = WM_KEYDOWN Or m.message = WM_SYSKEYDOWN Then
            Call TranslateMessage(m)
        End If
    Loop
Send:
    On Error Resume Next
    If Len(buffer) > 0 Then
        Call ZoteroCommand("pillBuffer", False, _
            " -ZoteroIntegrationMacroContext """ & Left$(buffer, Len(buffer) - 1) & """")
    End If
End Sub

' The inline-macro doc pref, read from the Citate/Zotero preference chunks stored
' as custom document properties, so no round-trip to the app is needed. Absent or
' unparseable data means enabled (the global default).
Private Function CitateInlineMacrosEnabled() As Boolean
    Dim json$
    Dim i As Long
    On Error GoTo Done
    For i = 1 To 100
        json = json & ActiveDocument.CustomDocumentProperties("ZOTERO_PREF_" & i).Value
    Next i
Done:
    On Error GoTo 0
    CitateInlineMacrosEnabled = (InStr(json, """inlineMacroCapture"":false") = 0)
End Function

' Ctrl+Shift+Z: if the paragraph containing the cursor may hold unexpanded citation
' macro tokens (backslash-prefixed), ask Citate to expand them; Citate falls back to
' the add/edit citation dialog when nothing matches. Without a backslash in the
' paragraph, behaves exactly like Add/Edit Citation.
Public Sub ZoteroExpandMacrosOrCite()
    Dim paraText$
    On Error Resume Next
    paraText$ = Selection.Paragraphs(1).Range.Text
    On Error GoTo 0
    If InStr(paraText$, "\") = 0 Then
        Call ZoteroCommand("addEditCitation", True)
        Exit Sub
    End If
    ' Strip paragraph/line breaks and tabs; escape quotes for the command line
    paraText$ = Replace(paraText$, vbCr, " ")
    paraText$ = Replace(paraText$, vbLf, " ")
    paraText$ = Replace(paraText$, Chr$(11), " ")
    paraText$ = Replace(paraText$, Chr$(9), " ")
    paraText$ = Replace(paraText$, """", """""")
    ' Cap the context we send; tokens beyond this are ignored
    If Len(paraText$) > 8000 Then paraText$ = Left$(paraText$, 8000)
    ' Expansion should not steal focus from Word, so no bringToFront. The trailing
    ' space keeps a paragraph-final backslash from escaping the closing quote.
    Call ZoteroCommand("expandMacros", False, _
        " -ZoteroIntegrationMacroContext """ & paraText$ & " """)
End Sub

Public Sub ZoteroInsertCitation()
    Call ZoteroCommand("addCitation", True)
End Sub

Public Sub ZoteroInsertBibliography()
    Call ZoteroCommand("addBibliography", False)
End Sub

Public Sub ZoteroEditCitation()
    Call ZoteroCommand("editCitation", True)
End Sub

Public Sub ZoteroEditBibliography()
    Call ZoteroCommand("editBibliography", True)
End Sub

Public Sub ZoteroAddEditCitation()
    Call ZoteroCommand("addEditCitation", True)
End Sub

Public Sub ZoteroAddNote()
    Call ZoteroCommand("addNote", True)
End Sub

Public Sub ZoteroAddAnnotation()
    Call ZoteroCommand("addAnnotation", True)
End Sub

Public Sub ZoteroAddEditBibliography()
    Call ZoteroCommand("addEditBibliography", True)
End Sub

Public Sub ZoteroSetDocPrefs()
    Call ZoteroCommand("setDocPrefs", True)
End Sub

Public Sub ZoteroCitationExplorer()
    Call ZoteroCommand("citationExplorer", True)
End Sub

Public Sub ZoteroRefresh()
    Call ZoteroCommand("refresh", False)
End Sub

Public Sub ZoteroRemoveCodes()
    Call ZoteroCommand("removeCodes", False)
End Sub

Private Sub FindZoteroWindow()
    ZotWnd = 0
    #If VBA7 Then
        Dim ThWnd As LongPtr
    #Else
        Dim ThWnd As Long
    #End If
    ' Zotero 6 / FX60+
    ThWnd = FindWindow("ZoteroMessageWindow", vbNullString)
    If ThWnd <> 0 Then
        ZotWnd = ThWnd
        Exit Sub
    End If
    
    IsZotero7 = True
    
    ' Zotero 7 / FX102+
    Dim lpdwThreadId As Long
    ThWnd = FindWindow("MozillaWindowClass", vbNullString)
    Do While ThWnd <> 0
        lpdwThreadId = GetWindowThreadProcessId(ThWnd, 0)
        Call EnumThreadWindows(lpdwThreadId, AddressOf EnumWindowsCallback, ByVal 0&)
        If ZotWnd <> 0 Then
            Exit Do
        End If
        ThWnd = FindWindowEx(0, ThWnd, "MozillaWindowClass", vbNullString)
    Loop
End Sub

Function EnumWindowsCallback(ByVal hwnd As Long, ByVal lParams As Long) As Long ' {
    Dim windowClass As String * 256
    Dim retVal      As Long
    Dim zoteroPosition As Long
    Dim remoteWindowPosition As Long

    retVal = GetClassName(hwnd, windowClass, 255)
    windowClass = Left$(windowClass, retVal)
    zoteroPosition = InStr(windowClass, "Mozilla_citate_")
    remoteWindowPosition = InStr(windowClass, "RemoteWindow")
    ' Looking for window name like `Mozilla_citate_%profileName%_RemoteWindow`
    ' which is not configurable and used to be much simpler in Z6 - `ZoteroMessageWindow`
    If zoteroPosition <> 0 And remoteWindowPosition <> 0 Then
        ZotWnd = hwnd
        EnumWindowsCallback = False
    Else
        '
        ' Return true to indicate that we want to continue
        ' with the enumeration of the windows:
        '
        EnumWindowsCallback = True
    End If
End Function ' }


Sub ZoteroCommand(cmd As String, bringToFront As Boolean, Optional extraArgs As String = "")
    Dim cds As COPYDATASTRUCT
    Dim a$, args$, name$, templateVersion$
    Dim i As Long
    Dim ignore As Long
    Dim sBuffer$
    Dim lLength As Long
    Dim buf() As Byte
    
    Call FindZoteroWindow
    If ZotWnd = 0 Then
        Call CitateLog("command " & cmd & " failed: Citate window not found")
        MsgBox ("Word could not communicate with Citate. Please ensure Citate is running and try again. If this problem persists, see https://www.zotero.org/support/word_processor_plugin_troubleshooting")
        Exit Sub
    End If
    Call CitateLog("command " & cmd)
    
    ' Allow Firefox to bring a window to the front
    If bringToFront Then Call SetForegroundWindow(ZotWnd)
    
    ' Get path to active document
    If ActiveDocument.Path <> "" Then
        name$ = ActiveDocument.Path & Application.PathSeparator & ActiveDocument.name
    Else
        name$ = ActiveDocument.name
    End If
    
    templateVersion$ = 1
    
    ' Set up command line arguments
    name$ = Replace(name$, """", """""")
    args$ = "-silent -ZoteroIntegrationAgent WinWord -ZoteroIntegrationCommand " & cmd & " -ZoteroIntegrationDocument """ & name$ & """ -ZoteroIntegrationTemplateVersion " & templateVersion$ & extraArgs
    a$ = "citate.exe " & args$ & Chr$(0) & "C:\"
    
    If IsZotero7 Then
        ' With FX128+ WM_COPYDATA either has to be in UTF16 (native VBA string encoding)
        ' or the UTF8 WM_COPYDATA must start with Fire + Fox emojis. We cannot do that
        ' because the VBA editor doesn't support unicode at all.
        ' Either way this works with FX115 too which is now released
        ' so we do this for Zotero 7
        cds.dwData = 2
        cds.cbData = LenB(a$)
        cds.lpData = StrPtr(a$)
    Else
        ' Do some UTF-8 magic for Zotero 6
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, ByVal 0, 0, 0, 0)
        ReDim buf(lLength) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(a$), -1, buf(1), lLength, 0, 0)
    
        cds.dwData = 1
        cds.cbData = lLength
        cds.lpData = VarPtr(buf(1))
    End If
    
    ' Send message to Firefox
    i = SendMessage(ZotWnd, WM_COPYDATA, 0, cds)
    
    ' Handle error
    If Err.LastDllError = 5 Then
        If Dir("C:\Program Files\Citate\citate.exe") <> "" Then
            Call Shell("""C:\Program Files\Citate\citate.exe"" " & args$, vbNormalFocus)
        ElseIf Dir("C:\Program Files (x86)\Citate\citate.exe") <> "" Then
            Call Shell("""C:\Program Files (x86)\Citate\citate.exe"" " & args$, vbNormalFocus)
        End If
    End If
End Sub

