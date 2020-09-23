VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Thumbnail Viewer Sample"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ctlThumbViewer ctlThumbViewer1 
      Height          =   4695
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   4695
      _extentx        =   8281
      _extenty        =   8281
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   120
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctlThumbViewer1_dblclick(picNum As Integer)
Form1.Caption = ctlThumbViewer1.getFileName(picNum)
Form2.Show
Form2.Picture1.Picture = LoadPicture(Form1.Caption)

End Sub

Private Sub ctlThumbViewer1_mousedown(picNum As Integer, Button As Integer)
Form1.Caption = "Pic #" & Str(picNum) & " Button:" & Str(Button)
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_PathChange()
' When the path changes on the file list,
' the loop loads the image list array in
' the control with all of the files
' contained within the file list (full path)

Dim i As Integer

If File1.ListCount > 0 Then
'reset the control display
    ctlThumbViewer1.clearImages

'load the array
    For i = 0 To File1.ListCount
    ctlThumbViewer1.addImage (Dir1.Path & "\" & File1.List(i))
    Next

'update the thumbnails
    ctlThumbViewer1.refreshThumbs
End If

End Sub

Private Sub Form_Load()
Dir1.Path = App.Path
End Sub

Private Sub Form_Resize()

If WindowState <> 1 Then
ctlThumbViewer1.Width = Form1.Width - ctlThumbViewer1.Left - 250
ctlThumbViewer1.Height = Form1.Height - ctlThumbViewer1.Top - 650
File1.Height = Form1.Height - File1.Top - 650
End If

End Sub
