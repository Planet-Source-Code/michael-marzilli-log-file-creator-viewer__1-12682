Attribute VB_Name = "MainMod"
' LOG FILE CREATOR AND VIEWER
' WITH SEARCH CAPABILITIES
' by     : Michael Marzilli
' date   : November 10, 2000
' site   : www.anthro-morphs.com
' license: You are free to use this code in your own projects,
'          so long as you do not try to pass it off as your own.
'          Send comments to programmer@anthro-morphs.com.


'GLOBAL VARIABLES NEEDED FOR LOGGING - REQUIRED FIELDS
 Global Const MAX_TYPES = 3                'This is the total number of TYPES
 Global Const MAX_LEN = 6                  'The Max Length of each TYPE name
 Global Const LOG_FILE = "log.txt"         'The Name of the Log File
 Global logType(1 To 10) As String         'The Name of the TYPES

'FORM VARIABLES
 Public fMain As frmMain
 Public fLog As frmLog

Public Sub main()
  Dim i As Integer
  
 'Create the Type Names in an Array of Strings - This MUST be done.
  logType(1) = "TYPE#1"
  logType(2) = "TYPE#2"
  logType(3) = "TYPE#3"
   
 'Make sure that the Type Name (logType(i)) is exactly MAX_LEN characters long
 'This MUST be done.
  For i = 1 To MAX_TYPES
    If (Len(logType(i)) < MAX_LEN) Then logType(i) = logType(i) & Left(Space(MAX_LEN), MAX_LEN - Len(logType(i)))
  Next i


  Set fMain = New frmMain
  Call fMain.Form_Load
  fMain.Show
End Sub
