Attribute VB_Name = "modLog"
Option Explicit

' --- MODULE CONSTANTS ---
Private Const m_LOG_FOLDER          As String = "logs"
Private Const m_LOG_FILE            As String = "CompBot.log"

' Purpose: Append one line to logs\CompBot.log beside CompBot. Returns True if the line was written.
'          Never raises, never shows a dialog. If CompBot's folder cannot be resolved to a local
'          path, nothing is written. Stores nothing personal: paths are resolved at run time.
Public Function LogError(ByVal strProc As String, ByVal lngNumber As Long, ByVal strDescription As String) As Boolean

    Dim strBase As String
    Dim strFolder As String
    Dim strBook As String
    Dim lngFile As Long

    On Error GoTo ErrHandler

    strBase = ResolveLocalPath(ThisWorkbook.Path)
    If Len(strBase) = 0 Then Exit Function

    strFolder = JoinPath(strBase, m_LOG_FOLDER)
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then MkDir strFolder

    If Not ActiveWorkbook Is Nothing Then strBook = ActiveWorkbook.Name

    lngFile = FreeFile
    Open JoinPath(strFolder, m_LOG_FILE) For Append As #lngFile
    Print #lngFile, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & strBook & " | " & strProc & _
                    " | " & lngNumber & " | " & strDescription
    Close #lngFile
    LogError = True
    Exit Function

ErrHandler:
    ' Logging must never cascade a second error out of a handler.
    On Error Resume Next
    If lngFile <> 0 Then Close #lngFile
    LogError = False
End Function

' Purpose: Where the log lives, short enough for the StatusBar: "<CompBot's folder>\logs\CompBot.log".
'          Derived at run time from wherever CompBot actually is.
Public Function LogLocation() As String

    Dim strBase As String

    On Error GoTo ErrHandler

    strBase = ResolveLocalPath(ThisWorkbook.Path)
    If Len(strBase) = 0 Then
        LogLocation = m_LOG_FOLDER & "\" & m_LOG_FILE
    Else
        LogLocation = Mid$(strBase, InStrRev(strBase, "\") + 1) & "\" & m_LOG_FOLDER & "\" & m_LOG_FILE
    End If
    Exit Function

ErrHandler:
    LogLocation = m_LOG_FOLDER & "\" & m_LOG_FILE
End Function

' Purpose: Convert a OneDrive / SharePoint URL to its local synced path.
'          Returns "" if a URL cannot be resolved - callers must check before writing.
'          Never raises; never calls LogError (LogError depends on this function).
Public Function ResolveLocalPath(ByVal strRawPath As String) As String

    ' --- CONSTANTS (local to function) ---
    Const HKCU          As Long = &H80000001
    Const PROVIDER_KEY  As String = "Software\SyncEngines\Providers\OneDrive"
    Const PERSONAL_HOST As String = "d.docs.live.net"

    Dim objReg       As Object
    Dim varSubKeys   As Variant
    Dim varKey       As Variant
    Dim strUrlSpace  As String
    Dim strMount     As String
    Dim strBestUrl   As String
    Dim strBestMount As String
    Dim strTail      As String
    Dim strLocalRoot As String

    On Error GoTo ErrHandler

    ' Already a drive letter or UNC path - nothing to resolve.
    If InStr(1, strRawPath, "https://", vbTextCompare) = 0 Then
        ResolveLocalPath = strRawPath
        Exit Function
    End If

    ' Longest matching UrlNamespace wins: a team site and one of its sub-libraries
    ' can both match, and the more specific mount point is the correct one.
    Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    objReg.EnumKey HKCU, PROVIDER_KEY, varSubKeys

    If IsArray(varSubKeys) Then
        For Each varKey In varSubKeys
            strUrlSpace = vbNullString
            strMount = vbNullString
            objReg.GetStringValue HKCU, PROVIDER_KEY & "\" & varKey, "UrlNamespace", strUrlSpace
            objReg.GetStringValue HKCU, PROVIDER_KEY & "\" & varKey, "MountPoint", strMount

            If Len(strUrlSpace) > 0 And Len(strMount) > 0 Then
                If Right$(strUrlSpace, 1) = "/" Then strUrlSpace = Left$(strUrlSpace, Len(strUrlSpace) - 1)
                ' Trailing separators on both sides stop a partial segment matching.
                If InStr(1, strRawPath & "/", strUrlSpace & "/", vbTextCompare) = 1 Then
                    If Len(strUrlSpace) > Len(strBestUrl) Then
                        strBestUrl = strUrlSpace
                        strBestMount = strMount
                    End If
                End If
            End If
        Next varKey
    End If

    If Len(strBestMount) > 0 Then
        strTail = Replace(Mid$(strRawPath, Len(strBestUrl) + 1), "/", "\")
        ' Personal OneDrive registers its namespace as the bare host (no CID), but the
        ' workbook URL still carries the CID segment, which has no local folder. Drop it.
        If StrComp(strBestUrl, "https://" & PERSONAL_HOST, vbTextCompare) = 0 Then
            Do While Left$(strTail, 1) = "\"
                strTail = Mid$(strTail, 2)
            Loop
            If InStr(1, strTail, "\") > 0 Then
                strTail = Mid$(strTail, InStr(1, strTail, "\") + 1)
            Else
                strTail = vbNullString
            End If
        End If
        ResolveLocalPath = JoinPath(strBestMount, strTail)
        Exit Function
    End If

    ' Fallback: personal OneDrive with no usable register entry. Strips the scheme,
    ' host and CID segment, leaving the path relative to the local OneDrive root.
    If InStr(1, strRawPath, PERSONAL_HOST, vbTextCompare) > 0 Then
        strLocalRoot = Environ$("OneDrive")
        If Len(strLocalRoot) > 0 Then
            strTail = Replace(strRawPath, "/", "\")
            strTail = Mid$(strTail, InStr(1, strTail, PERSONAL_HOST, vbTextCompare) + Len(PERSONAL_HOST))
            Do While Left$(strTail, 1) = "\"
                strTail = Mid$(strTail, 2)
            Loop
            ' Drop the CID segment.
            If InStr(1, strTail, "\") > 0 Then
                strTail = Mid$(strTail, InStr(1, strTail, "\") + 1)
            Else
                strTail = vbNullString
            End If
            ResolveLocalPath = JoinPath(strLocalRoot, strTail)
            Exit Function
        End If
    End If

    ' Unresolvable. Return "" rather than a plausible-looking sync root, so callers
    ' cannot silently write to the wrong folder.
    ResolveLocalPath = vbNullString
    Exit Function

ErrHandler:
    ResolveLocalPath = vbNullString
End Function

' Purpose: Join two path fragments with exactly one separator between them.
Public Function JoinPath(ByVal strLeft As String, ByVal strRight As String) As String
    Do While Right$(strLeft, 1) = "\"
        strLeft = Left$(strLeft, Len(strLeft) - 1)
    Loop
    Do While Left$(strRight, 1) = "\"
        strRight = Mid$(strRight, 2)
    Loop

    If Len(strRight) = 0 Then
        JoinPath = strLeft
    Else
        JoinPath = strLeft & "\" & strRight
    End If
End Function


