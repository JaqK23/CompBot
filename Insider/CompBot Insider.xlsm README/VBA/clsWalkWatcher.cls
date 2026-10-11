VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWalkWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==============================================================================
' clsWalkWatcher - the event sink for "Record Walk Route".
'
' Application-level SheetSelectionChange, so the route keeps recording whichever
' sheet the user is on. All logic lives in modWalk; this class only forwards.
'
' NOTE: this whole mechanism depends on Application.EnableEvents being TRUE.
' The house vbaInit/vbaFin wrapper turns events OFF, so the walk command must
' NOT use it. Ruled by Jaq, 2026-09-22. See modWalk's header.
'==============================================================================

Private WithEvents m_app As Application
Attribute m_app.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set m_app = Application
End Sub

Private Sub Class_Terminate()
    Set m_app = Nothing
End Sub

Private Sub m_app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    modWalk.HandleSelection Target
End Sub

