VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test File"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton btnView 
      Caption         =   "View Log"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Add to Log"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox tbText 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Send Sample Test to Log:"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' LOG FILE CREATOR AND VIEWER
' WITH SEARCH CAPABILITIES
' by     : Michael Marzilli
' date   : November 10, 2000
' site   : www.anthro-morphs.com
' license: You are free to use this code in your own projects,
'          so long as you do not try to pass it off as your own.
'          Send comments to programmer@anthro-morphs.com.

'This is a Test Form to:
'  1) Add lines to the Log File.
'  2) Press a button to view the Log File.
'This Form is not needed.


Public Sub Form_Load()
'Create and Populate the Type Combobox with the Types.

  Dim i As Integer
  
  cmbType.Clear
  For i = 1 To MAX_TYPES
    cmbType.AddItem (logType(i))
  Next i
  cmbType.ListIndex = 0
End Sub

Public Sub Form_Unload(cancel As Integer)
'Quit the Program

  Unload Me
End Sub

Public Sub btnSend_Click()
'If the Add to Log button is pressed, make sure the
'textbox is not empty, and add it to the log if it
'is not empty.

  If (Trim(tbText.Text) <> "") Then
  
  ' This is the Call to add to the Log File
    Call frmLog.AddLog(cmbType.ListIndex + 1, 0, Trim(tbText.Text))
    
    tbText.Text = ""
    tbText.SetFocus
  Else
    Beep
  End If
End Sub

Public Sub btnView_Click()
'Open the View Log Form

  Set fLog = New frmLog
  Call fLog.Form_Load
  fLog.Show
End Sub

Public Sub btnQuit_Click()
'The User is Quitting the Program

  Call Form_Unload(0)
End Sub
