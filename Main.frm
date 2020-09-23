VERSION 5.00
Begin VB.Form Main 
   Caption         =   "ElFerTagRemover"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.FileListBox FileList 
      Height          =   2625
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   4455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandRemove_Click()
If FileList.FileName = "" Then MsgBox ("No file selected"): Exit Sub

CommandRemove.Enabled = False

List1.Clear

If IsID31Present(FileList.FileName) Then List1.AddItem "ID3 v1 Present"
If IsID32Present(FileList.FileName) Then List1.AddItem "ID3 v2 Present"
List1.AddItem "ID31Size: " + Str$(ID31Size(FileList.FileName))
List1.AddItem "ID32Size: " + Str$(ID32Size(FileList.FileName))
List1.AddItem "MP3Size: " + Str$(MP3Size(FileList.FileName))
List1.AddItem "FileLen: " + Str$(FileLen(FileList.FileName))
RemoveID31Tag (FileList.FileName)
RemoveID32Tag (FileList.FileName)

FileList.Refresh
End Sub

