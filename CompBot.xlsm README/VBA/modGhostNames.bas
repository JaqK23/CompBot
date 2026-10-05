Attribute VB_Name = "modGhostNames"
Option Explicit

'==============================================================================
' modGhostNames - find and clear orphaned _xlop.* parameter ghosts.
'
' WHAT THEY ARE: Excel registers a hidden defined name "_xlop.<name>" for every
' OPTIONAL LAMBDA parameter name used anywhere in the workbook. They all show
' =#NAME? and they are Excel's own bookkeeping, NOT junk - modLambdas.NormaliseName
' strips the prefix for exactly this reason.
'
' Confirmed 2026-09-22 by reading the stored formulas out of a closed .xlsm
' (commandtools\dump_lambda_full.py). Excel's on-disk form of a lambda carries
' one prefix per parameter kind, and they are how a signature is read back:
'     _xlpm.Name   parameter, REQUIRED
'     _xlop.Name   parameter, OPTIONAL  - i.e. it was written [Name]
' So an _xlop entry is not evidence of anything deleted; it just means some
' lambda, at some point, had an optional parameter of that name.
'
' WHY ANY ARE STALE: when a lambda is deleted or a parameter renamed, the ghost
' is left behind. CompBot still carries _xlop.ShowAll_0 (GridToCol's parameter
' before it became HideZeros/ShowBlanks) and _xlop.Flatten (from the lambda the
' 2026-09 review retired).
'
' THE SAFETY RULE, and it is deliberately conservative: a ghost is only stale if
' its name appears NOWHERE in ANY live LAMBDA definition - not as a parameter,
' not in a body, not as a LET variable. Anything that appears anywhere is kept.
' That errs towards leaving a ghost in place, which costs nothing, rather than
' deleting one Excel still wants, which could break a lambda.
'
' ALWAYS RUN THE PREVIEW FIRST. _xleta.* names are NEVER touched - those map to
' built-in functions used inside lambdas.
'
' *** THEY CANNOT BE DELETED. MEASURED 2026-09-22, DO NOT TRY AGAIN. ***
' Every delete fails with 1004 "The syntax of this name isn't correct", because
' the dot makes "_xlop.foo" invalid under Excel's own naming rules. Tried both
' ways and both fail the same:
'     ThisWorkbook.Names("_xlop.foo").Delete   - fails at the LOOKUP
'     nm.Delete via the Name object itself      - fails at the DELETE
' Excel created these for itself and will not hand them back through the object
' model. The only remaining route is editing xl/workbook.xml inside the closed
' file, which is real surgery on a macro-enabled workbook and is not worth it for
' names that are hidden, inert and cost nothing.
'
' Jaq's own Ghostbuster.xlsm was checked (2026-09-22) and does NOT help here: it
' hunts phantom EXTERNAL LINK references to one named workbook - LinkSources,
' names, formulas, chart series, pivot caches, query tables, connections, CF,
' OLE objects - and only reports them to the Immediate window. Different problem.
'
' RemoveStaleGhosts is therefore kept only so the failure is reproducible.
' PreviewStaleGhosts is still genuinely useful: it tells you which parameter
' names no live lambda uses any more, which is a good signal when auditing.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_DEBUG_MODE As Boolean = False
Private Const m_PREFIX     As String = "_xlop."


'------------------------------------------------------------------------------
' List the stale ghosts without deleting anything.
'------------------------------------------------------------------------------
Public Function PreviewStaleGhosts(ByVal strOutPath As String) As String
    PreviewStaleGhosts = ScanGhosts(strOutPath, False)
End Function


'------------------------------------------------------------------------------
' Delete them. Run PreviewStaleGhosts first and read the list.
'------------------------------------------------------------------------------
Public Function RemoveStaleGhosts(ByVal strOutPath As String) As String
    RemoveStaleGhosts = ScanGhosts(strOutPath, True)
End Function


Private Function ScanGhosts(ByVal strOutPath As String, _
                            ByVal blnDelete As Boolean) As String

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ScanGhosts"

    Dim nm          As Name
    Dim strHaystack As String
    Dim colGhosts   As Collection
    Dim colStale    As Collection
    Dim varGhost    As Variant
    Dim strBare     As String
    Dim lngFile     As Long
    Dim lngKept     As Long
    Dim lngGone     As Long
    Dim strFirstFail As String
    Dim lngErrNum   As Long
    Dim strErrDesc  As String

    On Error GoTo ErrHandler

    Set colGhosts = New Collection
    Set colStale = New Collection

    ' One pass: collect the ghosts, and build one big haystack of every live
    ' LAMBDA definition to test their names against.
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.Name, m_PREFIX, vbTextCompare) = 1 Then
            colGhosts.Add nm.Name
        Else
            On Error Resume Next
            If InStr(1, nm.RefersTo, "LAMBDA(", vbTextCompare) > 0 Then
                strHaystack = strHaystack & "|" & nm.RefersTo
            End If
            On Error GoTo ErrHandler
        End If
    Next nm

    For Each varGhost In colGhosts
        strBare = Mid$(CStr(varGhost), Len(m_PREFIX) + 1)
        If Len(strBare) > 0 Then
            If Not TokenAppears(strHaystack, strBare) Then
                colStale.Add CStr(varGhost)
            Else
                lngKept = lngKept + 1
            End If
        End If
    Next varGhost

    ' Write the list before touching anything, so there is always a record.
    lngFile = FreeFile
    Open strOutPath For Output As #lngFile
    Print #lngFile, "STALE _xlop ghosts - name appears in no live LAMBDA definition"
    For Each varGhost In colStale
        Print #lngFile, CStr(varGhost)
    Next varGhost
    Close #lngFile

    If blnDelete Then
        ' Delete through the Name OBJECT, never by string key. Looking one up as
        ' ThisWorkbook.Names("_xlop.foo") raises 1004 "The syntax of this name
        ' isn't correct" - the dot makes it an invalid key - so the delete never
        ' even runs. Walking the collection backwards by index sidesteps that and
        ' is also safe against the collection reindexing as entries are removed.
        Dim lngI As Long
        Dim nmDel As Name
        For lngI = ThisWorkbook.Names.Count To 1 Step -1
            Set nmDel = Nothing
            On Error Resume Next
            Set nmDel = ThisWorkbook.Names(lngI)
            On Error GoTo ErrHandler
            If Not nmDel Is Nothing Then
                If InList(colStale, nmDel.Name) Then
                    On Error Resume Next
                    nmDel.Delete
                    If Err.Number = 0 Then
                        lngGone = lngGone + 1
                    ElseIf Len(strFirstFail) = 0 Then
                        strFirstFail = nmDel.Name & " -> " & Err.Number & " " & Err.Description
                    End If
                    Err.Clear
                    On Error GoTo ErrHandler
                End If
            End If
        Next lngI

        ScanGhosts = "ghosts " & colGhosts.Count & ", in use " & lngKept & _
                     ", stale " & colStale.Count & ", DELETED " & lngGone
        If Len(strFirstFail) > 0 Then _
            ScanGhosts = ScanGhosts & " | first failure: " & strFirstFail
    Else
        ScanGhosts = "ghosts " & colGhosts.Count & ", in use " & lngKept & _
                     ", stale " & colStale.Count & " (nothing deleted)"
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    On Error Resume Next
    Close #lngFile
    LogError PROC_NAME, lngErrNum, strErrDesc
    ScanGhosts = "FAILED " & lngErrNum & ": " & strErrDesc
    Resume Cleanup

End Function


'------------------------------------------------------------------------------
' True if strToken appears in strText as a whole token - not as part of a longer
' identifier. "Fill" must not match inside "HideBlanksFill" or "Fills".
'------------------------------------------------------------------------------
Private Function TokenAppears(ByVal strText As String, _
                              ByVal strToken As String) As Boolean

    Dim lngAt   As Long
    Dim strPrev As String
    Dim strNext As String

    lngAt = InStr(1, strText, strToken, vbTextCompare)
    Do While lngAt > 0
        strPrev = vbNullString
        strNext = vbNullString
        If lngAt > 1 Then strPrev = Mid$(strText, lngAt - 1, 1)
        If lngAt + Len(strToken) <= Len(strText) Then _
            strNext = Mid$(strText, lngAt + Len(strToken), 1)

        If Not IsIdentChar(strPrev) And Not IsIdentChar(strNext) Then
            TokenAppears = True
            Exit Function
        End If

        lngAt = InStr(lngAt + 1, strText, strToken, vbTextCompare)
    Loop

End Function


Private Function InList(ByVal colItems As Collection, ByVal strWanted As String) As Boolean

    Dim varItem As Variant

    For Each varItem In colItems
        If StrComp(CStr(varItem), strWanted, vbBinaryCompare) = 0 Then
            InList = True
            Exit Function
        End If
    Next varItem

End Function


Private Function IsIdentChar(ByVal strChar As String) As Boolean
    If Len(strChar) = 0 Then Exit Function
    IsIdentChar = (strChar Like "[A-Za-z0-9_.]")
End Function


