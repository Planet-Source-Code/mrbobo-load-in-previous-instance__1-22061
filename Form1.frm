VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load In Previous Instance Example"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "The Demo"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   6735
      Begin VB.CommandButton cmdShowCode 
         Height          =   495
         Left            =   240
         Picture         =   "Form1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "View Re-Usable Code sample"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdHow 
         Height          =   495
         Left            =   240
         Picture         =   "Form1.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "How does it work ?"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   240
         Picture         =   "Form1.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save as a  ""*.bbt"" file."
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmdNew 
         Height          =   495
         Left            =   240
         Picture         =   "Form1.frx":0950
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   2295
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   5535
      End
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Description - When all else fails, READ THE INSTRUCTIONS !!!!!!"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmdDisAssociate 
         Caption         =   "Remove association"
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton cmdAssociate 
         Caption         =   "Create association"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtInstructions 
         Height          =   1575
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Form1.frx":0A52
         Top             =   360
         Width           =   6255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Used to send a string to another window
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETTEXT = &HC

Private Sub Form_Load()
Dim mycommand As String
mycommand = Command()
If mycommand <> "" Then 'If shelled through Explorer
    If App.PrevInstance Then
        'If there's a previous instance of
        'your app send the file to that instance
        LoadinPrevInst mycommand
        End
    Else
        'otherwise load it
        Text1.Text = OneGulp(mycommand)
    End If
End If
'Put the hwnd of txtCommand in registry
SaveSetting App.Title, "ActiveWindow", "Handle", Str(txtCommand.hwnd)
End Sub
Private Sub LoadinPrevInst(mfile As String)
'If we ended up here there must be a previous instance
'Read registry - get the handle of the previous instances'
'txtCommand, then add to it the name of the file
'the previous instance should load.
'The txtCommand_Change event in the previous instance
'will load the shelled file into the previous instance
Dim temp As String, mhw As Long
    temp = GetSetting(App.Title, "ActiveWindow", "Handle")
    mhw = CLng(Val(temp))
    SendMessage mhw, WM_SETTEXT, 0, ByVal CStr(mfile)
End Sub

Private Sub txtCommand_Change()
'Load a shelled file
If FileExists(txtCommand.Text) Then Text1.Text = OneGulp(txtCommand.Text)
End Sub
'********************************************************
'All above is the Load in Previous instance code
'The following is just for this example
'******************************************************

Private Sub cmdAssociate_Click()
Associate
End Sub

Private Sub cmdDisAssociate_Click()
DisAssociate
End Sub

Private Sub cmdHow_Click()
    Text1.Text = txtHow
End Sub

Private Sub cmdNew_Click()
    Text1.Text = ""
End Sub

Private Sub cmdSave_Click()
Dim sfile As String
sfile = ShowSave
If sfile <> "" Then FileSave Text1.Text, ChangeExt(sfile, "bbt")
End Sub

Private Sub cmdShowCode_Click()
Text1.Text = txtCode
End Sub

