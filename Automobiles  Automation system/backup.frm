VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Backup "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If MsgBox("If the file already exists" & vbNewLine & _
"to the selected lacation" & vbNewLine & _
"then it will be replaced by new one..." & vbNewLine & _
"Proceed ?", vbQuestion + vbYesNo, "Message") = vbYes Then
FileCopy "e:\Laxmipati Automobiles new\Other files\Master_Database1.accdb", Dir1.Path & "\Master_Database1.accdb"
MsgBox "Backed Up successfully.", vbExclamation, "Message"
Unload Me
End If
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
