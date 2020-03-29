VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExplorerEventClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents oExplorer As Outlook.explorer
Attribute oExplorer.VB_VarHelpID = -1

Sub RegisterExplorer(ByRef theExplorer As Outlook.explorer)
    Set oExplorer = theExplorer
End Sub

Private Sub oExplorer_BeforeItemPaste(ClipboardContent As Variant, ByVal Target As Folder, Cancel As Boolean)
    Debug.Print ("Item paste event caught!")
    
    If (CheckOgcsRunning) Then
        Debug.Print ("OGCS is running and will handle it.")
        GoTo endSub
    End If
    
    Dim selection As selection
    Set selection = ClipboardContent
    
    Debug.Print ("We've got " & selection.Count & " items selected for copy.")
    For Each anItem In selection
 '       On Error GoTo cleanUp
        Dim copiedAi As AppointmentItem
        If (TypeName(anItem) = "AppointmentItem") Then
            Set copiedAi = anItem.CopyTo(Target, olCreateAppointment)
        Else
            Debug.Print ("The item is not an appointment item.");
            GoTo cleanUp
        End If
        
        Debug.Print (copiedAi.Start & " """ & copiedAi.Subject & """")
        
        RemoveCustomProperty copiedAi
        
        Dim currentCalView As CalendarView
        Set currentCalView = oExplorer.CurrentView
        copiedAi.Start = currentCalView.SelectedStartTime
        
        copiedAi.Save
        
        Cancel = True
        
cleanUp:
        Set copiedAi = Nothing
        GoTo nextItem
errHandler:
        Debug.Print ("Not able to process copy and pasted event.");
        Debug.Print (Err.Description)
        GoTo cleanUp
nextItem:

    Next
    
endSub:
End Sub

Function CheckOgcsRunning() As Boolean
    Dim objWMIcimv2
    Dim objList
    
    Debug.Print ("Checking for process " & ThisOutlookSession.ogcsProcessName)
    
    Set objWMIcimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set objList = objWMIcimv2.execQuery("select * from win32_process where name like '" & ThisOutlookSession.ogcsProcessName & "%'")
    
    Debug.Print ("Found " & objList.Count & " such process(es)")
    CheckOgcsRunning = (objList.Count > 0)
End Function

Sub RemoveCustomProperty(ai As AppointmentItem)
    Dim ups As UserProperties
    Dim i As Integer
    'On Error Resume Next
    On Error GoTo cleanUp
    Set ups = ai.UserProperties
    
    If (ups.Count > 0) Then
        Debug.Print ("Checking " & ups.Count & " properties...")
        For i = ups.Count To 1 Step -1
            Set up = ups(i)
            If (up.Name Like "googleEventID*" Or _
                up.Name Like "googleCalendarID*" Or _
                up.Name Like "forceSave*" Or _
                up.Name Like "locallyCopied*" Or _
                up.Name Like "ogcsModified*" _
            ) Then
                Debug.Print ("Deleting " & up.Name)
                up.Delete
            Else
                Debug.Print ("Other: " & up.Name)
            End If
            Set up = Nothing
        Next i
        Debug.Print (ups.Count & " items remaining.")
    End If
    
    Dim newUp As UserProperty
    Debug.Print ("Adding locallyCopied property.")
    Set newUp = ups.Add("locallyCopied", olText)
    newUp.Value = "True"
    Debug.Print (ups.Count & " items remaining.")
    
cleanUp:
    Set ups = Nothing
    Set newUp = Nothing
End Sub
