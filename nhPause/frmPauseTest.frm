VERSION 5.00
Begin VB.Form frmPauseTest 
   Caption         =   "Pause Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Pause"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDoEvents 
      Caption         =   "DoEvents Pause"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdPauseAPI 
      Caption         =   "API Pause"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtPauseInterval 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "3000"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblInterval 
      Caption         =   "Interval (MS)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmPauseTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : PauseTest
' DateTime  : 2/6/2004 10:43
' Author    : nhenderson
' Purpose   : demonstrate a way to implement a pause function in iFIX that
'               does not cause the CPU to hit 100%. As a comparison, it also allows
'               the same test to run against the pause function from the
'               iFIX migration tools, which does peg the CPU while pausing.
'               It also demonstrates a slightly more efficient DoEvents implementation,
'               that only processes if there is something in the user input queue of the Workspace
'---------------------------------------------------------------------------------------
Private m_blnUseRTPause As Boolean
Private m_blnCancel As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long


' --------------------------------------------------
' Sub        : CancelPause
' Parameters :
' Created    : n - 01/29/2004 19:52:15
' Modified   :
' --------------------------------------------------
' Comments
' winds up CPU
'
' --------------------------------------------------
Private Sub CancelPause()
    m_blnCancel = True
End Sub

'***************************************************************************************
'This subroutine supports the runtime conversion of PAUSE.
Private Sub vbPause(iPause As Integer)
    On Error GoTo PROC_ERR
    'Debug.Print "vbPause - Entering"
    Dim CurDateTime As Date
    'Get Current Time
    CurDateTime = Now()
    'Loop until Pause achieved
    While (CurDateTime + iPause / (86400000)) > Now()
        DoEvents
    Wend
    Exit Sub
PROC_EXIT:
    On Error Resume Next
    'Debug.Print "vbPause - Exiting"
    Exit Sub
PROC_ERR:
    Debug.Print "vbPause - Error : " & Err.Number & " (" & Err.Description & ")"
    Resume PROC_EXIT
    'Resume Next
End Sub

' --------------------------------------------------
' Sub        : apiPause
' Parameters :
' Created    : n - 01/29/2004 19:06:35
' Modified   :
' --------------------------------------------------
' Comments
' safe pause, with minimal CPU impact
' --------------------------------------------------
Private Sub apiPause(ByVal lngMS As Long)
    On Error GoTo PROC_ERR
    'Debug.Print "apiPause - Entering"
    Dim lngHandle As Long   ' handle of dummy object
    Dim lngRetVal As Long   ' status of API call
    Dim lngTotalMS As Long  ' number of milliseconds we have waited
    ' number of MS to wait each iteration
    ' the smaller the number, the smoother the UI
    ' the smaller the number, the more times this loop gets processed
    ' somewhere between 5-10 is nirvana for my machine
    Const WAIT_INCREMEMENT = 5
    ' explicit initialization is always good
    lngTotalMS = 0
    ' create a dummy event that will never fire
    lngHandle = CreateEvent(ByVal 0&, False, False, ByVal 0&)
    ' loop in WAIT_INCREMEMENT MS increments until we accumulate requested delay
    Do While lngTotalMS < lngMS
        ' this will wait the requested milliseconds for the object to fire
        ' since we never fire, it will return in the number of milliseconds
        lngRetVal = WaitForSingleObject(lngHandle, WAIT_INCREMEMENT)
        ' keep track of how long we have waited so far
        lngTotalMS = lngTotalMS + WAIT_INCREMEMENT
        ' process UI thread, to allow for cancel events, and other updates on UI thread
        ' this makes everthing smooth
        DoEvents
        If m_blnCancel Then
            ' I am done with THIS witness
            Exit Do
        End If
    Loop
PROC_EXIT:
    On Error Resume Next
    'Debug.Print "apiPause - Exiting"
    ' make sure we clean up
    lngRetVal = CloseHandle(lngHandle)
    Exit Sub
PROC_ERR:
    Debug.Print "apiPause - Error : " & Err.Number & " (" & Err.Description & ")"
    Resume PROC_EXIT
End Sub

Private Sub cmdCancel_Click()
    m_blnCancel = True
End Sub

Private Sub cmdDoEvents_Click()
    m_blnCancel = False
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngTotal As Long
    lngStart = GetTickCount()
    vbPause CInt(txtPauseInterval.Text)
    lngEnd = GetTickCount()
    lngTotal = lngEnd - lngStart
    'MsgBox "Time Paused: " & lngTotal, vbOKOnly, "Pause Test"
    
End Sub

Private Sub cmdPauseAPI_Click()
    m_blnCancel = False
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngTotal As Long
    lngStart = GetTickCount()
    apiPause CInt(txtPauseInterval.Text)
    lngEnd = GetTickCount()
    lngTotal = lngEnd - lngStart
    'MsgBox "Time Paused: " & lngTotal, vbOKOnly, "Pause Test"
End Sub
