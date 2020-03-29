VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
'' ADJUST THE VARIABLES IN THE CONFIGURE() SUB BELOW AS APPROPRIATE FOR YOUR INSTALLATION OF OGCS
''

Dim eventClass As New EventClassModule
Public ogcsDirectory, ogcsExecutable As String
Public ogcsProcessName As String
Public ogcsStartWithOutlook As Boolean

'Load the right driver bitness at runtime
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long _
    ) As Long
#Else
    'Don't worry if compiler highlights in red - it isn't instantiated at runtime for 64-bit systems
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long _
    ) As Long
#End If

Private Sub Configure()
    'Set these variables as appropriate for your installation of OGCS
    ThisOutlookSession.ogcsDirectory = Environ("LOCALAPPDATA") & "\OutlookGoogleCalendarSync\app-2.8.0-beta"
    ThisOutlookSession.ogcsExecutable = "OutlookGoogleCalendarSync.exe"
    ThisOutlookSession.ogcsProcessName = Replace(ThisOutlookSession.ogcsExecutable, ".exe", "", 1, 1, vbTextCompare)
    ThisOutlookSession.ogcsStartWithOutlook = False
    
    'Log settings
    Debug.Print ("ogcsDirectory = " & ThisOutlookSession.ogcsDirectory)
    Debug.Print ("ogcsExecutable = " & ThisOutlookSession.ogcsExecutable)
    Debug.Print ("ogcsProcessName = " & ThisOutlookSession.ogcsProcessName)
    Debug.Print ("ogcsStartWithOutlook = " & ThisOutlookSession.ogcsStartWithOutlook)
End Sub

Private Sub Application_Startup()
    Configure
    
    If (ThisOutlookSession.ogcsStartWithOutlook) Then
        'Start OGCS
        Dim RetVal As Long
        On Error Resume Next
        RetVal = ShellExecute(0, "open", ThisOutlookSession.ogcsDirectory & "\" & ThisOutlookSession.ogcsProcessName, " ", ThisOutlookSession.ogcsDirectory, 3)
    End If
    
    InitiateClass
End Sub

Sub InitiateClass()
    Debug.Print ("Initialiasing detection of user's copy & paste of OGCS managed items")
    eventClass.Register_Event_Handler
End Sub


