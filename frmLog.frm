VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Log"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox tbOutput 
      Height          =   3495
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLog.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox tbQuery 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton btnDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3720
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox tbData 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cmbModule 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox tbSocket 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox tbTimeTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tbTimeFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tbDateTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbDateFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "( use '0' to list All Sockets )"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Data : "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   " to "
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   " to "
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Module : "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Socket : "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Time : "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date : "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Your Query :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLog"
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


Public Sub Form_Load()
'Perform these Operations when the Window Opens
  
  Dim i As Integer
  
 'Copy the active file so we can search on a Copy of the Log File.
 'We don't want to use the active Log File because, theoretically
 'more data will be written to it...which could cause problems.
  FileCopy App.Path & "/" & LOG_FILE, App.Path & "/server.bak"
  
 'Add the Type Names to a ComboBox; this is data for Queries
  cmbModule.Clear
  cmbModule.AddItem ("ALL")
  
  For i = 1 To MAX_TYPES
    cmbModule.AddItem (logType(i))
  Next i
  cmbModule.ListIndex = 0
End Sub

Private Sub Form_Unload(cancel As Integer)
'Perform these Operations when the Window Closes
 
 'Delete the Copy of the Log File
  On Error Resume Next
  Kill App.Path & "/server.bak"
  If (Err.Number <> 0) Then Err.Clear
  
 'The Close this Log Window
  Unload Me
End Sub

'
' ***** PROCEDURES *****
'
Public Sub WriteLog(st As String)
'Open a Text File (LOG_FILE), Write a String (st) to it, and Close the File

  st = Trim(st)
  Open LOG_FILE For Append As #9
  Print #9, st
  Close #9
End Sub

Public Sub AddLog(typ As Integer, skt As Integer, st As String)
'Formats the String, then adds it to the Text File (LOG_FILE)
'skt  is used as the Socket Number in my application.
'You can use this for anything; example: Usernumber, etc.
'Otherwise, send a Zero to skt.

  Dim ns As String
  
 'Add Leading Zeros to the SKT number to make the length 3.
  ns = Trim(CStr(skt))
  If (Len(ns) < 3) Then ns = Left("000", 3 - Len(ns)) & ns
  
  st = Trim(st)
  st = ns & ": [" & logType(typ) & "] " & st
  st = Format(Now, "MM/DD/YYYY HH:MM") & " " & st
  
 'You can add the following lines to display this string in a List Box
 'for a Real-Time view of the Log. You must have a ListBox control on
 'your Form in order for this to work. Example is Below:
 '  frmMain.lbErrorLog.AddItem st
 '  frmMain.lbErrorLog.ListIndex = (frmMain.lbErrorLog.ListCount - 1)
 '  frmMain.lbErrorLog.ListIndex = -1
 '  If (frmMain.lbErrorLog.ListCount > 500) Then frmMain.lbErrorLog.RemoveItem (0)

  Call WriteLog(st)
End Sub

Private Sub QueryLog()
'Takes the data from the various Textboxes and Combobox, and
'searches the Text File (LOG_FILE) for lines that match.

  Dim srchSdate As Date
  Dim srchEdate As Date
  Dim srchSock As Integer
  Dim srchMod As String
  Dim srchData As String
  Dim lineDate As Date
  Dim lineSock As Integer
  Dim lineMod As String
  Dim lineData As String
  Dim st, temp As String
  Dim CRLF As String
  Dim foundLin As Boolean
  Dim ctr As Integer
  Dim i As Integer
  Dim e As Integer
  
  CRLF = Chr(13) & Chr(10)
 'CREATE QUERY
  If (tbTimeFrom.Text = "") Then
    st = "00:00"
  Else
    st = tbTimeFrom.Text
  End If
  
  If (tbDateFrom.Text = "") Then
    temp = "01/01/2000"
  Else
    temp = tbDateFrom.Text
  End If
  
  temp = temp & " " & st
  srchSdate = CDate(Format(temp, "MM/DD/YYYY HH:MM"))
  
  If (tbTimeTo.Text = "") Then
    If (tbDateTo.Text = "") Then
      st = Format(Now, "HH:MM")
    Else
      st = "23:59"
    End If
  Else
    st = tbTimeTo.Text
  End If
  
  If (tbDateTo.Text = "") Then
    temp = Format(Now, "MM/DD/YYYY")
  Else
    temp = tbDateTo.Text
  End If
  
  temp = temp & " " & st
  srchEdate = CDate(Format(temp, "MM/DD/YYYY HH:MM"))
  
  If (tbSocket.Text = "") Then
    srchSock = 0
  Else
    srchSock = Trim(CStr(tbSocket.Text))
  End If
  
  srchMod = cmbModule.List(cmbModule.ListIndex)
  srchData = Trim(tbData.Text)
  
 'DISPLAY QUERY
  tbQuery.Text = ""
  tbQuery.Text = tbQuery.Text & "Start Date: " & Format(srchSdate, "MM/DD/YYYY HH:MM") & CRLF
  tbQuery.Text = tbQuery.Text & "  End Date: " & Format(srchEdate, "MM/DD/YYYY HH:MM") & CRLF
  
  If (srchock = 0) Then
    tbQuery.Text = tbQuery.Text & "   Sockets: ALL" & CRLF
  Else
    tbQuery.Text = tbQuery.Text & "    Socket: " & Trim(CStr(srchSock)) & CRLF
  End If
  
  tbQuery.Text = tbQuery.Text & "    Module: " & srchMod & CRLF
  tbQuery.Text = tbQuery.Text & "      Data: " & srchData & CRLF
  tbQuery.Visible = True
  tbQuery.Refresh
  
  tbOutput.Text = ""
  tbOutput.Visible = False
  foundline = False
  ctr = 0
  
 'DISPLAY QUERY RESULTS
  Open App.Path & "/server.bak" For Input As #8
  Do While (Not EOF(8))
    st = ""
    temp = ""
    Do While (temp <> Chr(10)) And (Not EOF(8))
      temp = Input(1, #8)
      st = st & temp
    Loop
    
    If (InStr(1, st, Chr(13)) > 0) Then st = Mid(st, 1, InStr(1, st, Chr(13)) - 1)
    
    If (Trim(CStr(Mid(st, 1, 2))) > 0) Then
      lineDate = CDate(Format(Mid(st, 1, 16), "MM/DD/YYYY HH:MM"))
    Else
      lineDate = CDate(Format("01/01/1990 01:00", "MM/DD/YYYY HH:MM"))
    End If
    
    lineSock = Trim(CStr(Mid(st, 18, 3)))
    lineMod = Mid(st, 24, MAX_LEN)
    lineData = Trim(Mid(st, 26 + MAX_LEN, Len(st)))
    
    If (lineDate >= srchSdate) And (lineDate <= srchEdate) And ((lineSock = srchSock) Or (srchSock = 0)) And ((lineMod = srchMod) Or (srchMod = "ALL")) And (InStr(1, lineData, srchData) > 0) Then
      tbOutput.Text = tbOutput.Text & st & CRLF
      foundline = True
    End If
    
    DoEvents
  Loop
  Close #8
  If (Not foundline) Then tbOutput.Text = "No Lines match your Query."
  tbOutput.Visible = True
End Sub

'
' ***** BUTTONS
'
Private Sub btnDone_Click()
  Call Form_Unload(0)
End Sub

Private Sub btnSearch_Click()
  Call QueryLog
End Sub

'
' ***** TEXT BOXES *****
'
Private Sub tbDateFrom_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> "/") Then KeyAscii = 0
End Sub

Private Sub tbDateFrom_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbDateTo_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> "/") Then KeyAscii = 0
End Sub

Private Sub tbDateTo_GotFocus()
  If (tbDateFrom.Text <> "") And (tbDateTo.Text = "") Then tbDateTo.Text = tbDateFrom.Text
  SendKeys "{home}+{end}"
End Sub

Private Sub tbTimeFrom_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> ":") Then KeyAscii = 0
End Sub

Private Sub tbTimeFrom_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbTimeTo_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> ":") Then KeyAscii = 0
End Sub

Private Sub tbTimeTo_GotFocus()
  If (tbDateFrom.Text <> "") And (tbDateTo.Text = "") Then tbTimeTo.Text = tbTimeFrom.Text
  SendKeys "{home}+{end}"
End Sub

Private Sub tbSocket_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Sub

Private Sub tbSocket_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbData_GotFocus()
  SendKeys "{home}+{end}"
End Sub


