VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim explorerEventClass As New ExplorerEventClassModule
Dim myExplorers() As ExplorerEventClassModule

Public oApp As Outlook.Application
Public WithEvents oExplorers As Outlook.explorers
Attribute oExplorers.VB_VarHelpID = -1

Sub Register_Event_Handler()
    Set oApp = GetObject(, "Outlook.Application")
    Set oExplorers = oApp.explorers
    ReDim Preserve myExplorers(oApp.explorers.Count - 1)
    
    Dim e As Integer
    For e = 0 To explorers.Count - 1
        Set myExplorers(e) = New ExplorerEventClassModule
        myExplorers(e).RegisterExplorer explorers(e + 1)
    Next e
End Sub

    
Private Sub oExplorers_NewExplorer(ByVal newExplorer As explorer)
    Debug.Print ("New Explorer detected.")
    
    Dim explorerCnt As Integer
    explorerCnt = oApp.explorers.Count - 1
    ReDim Preserve myExplorers(explorerCnt)
    Set myExplorers(explorerCnt) = New ExplorerEventClassModule
    myExplorers(explorerCnt).RegisterExplorer newExplorer
End Sub
