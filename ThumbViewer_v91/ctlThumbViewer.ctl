VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlThumbViewer 
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ScaleHeight     =   6660
   ScaleWidth      =   6300
   Begin VB.PictureBox picBuffer 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   840
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox controlPic 
      ClipControls    =   0   'False
      Height          =   6075
      Left            =   0
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   1
      Top             =   0
      Width           =   5880
      Begin VB.PictureBox thumbHolder 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5565
         Left            =   75
         ScaleHeight     =   371
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   384
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   5760
         Begin VB.Image Img 
            Height          =   510
            Index           =   0
            Left            =   3090
            Top             =   4575
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Shape Pic2 
            BackColor       =   &H80000001&
            BorderWidth     =   5
            FillColor       =   &H00FFFFFF&
            Height          =   1230
            Index           =   0
            Left            =   -3000
            Top             =   0
            Visible         =   0   'False
            Width           =   1695
         End
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   6075
      LargeChange     =   85
      Left            =   5880
      TabIndex        =   0
      Top             =   0
      Width           =   285
   End
   Begin MSComctlLib.ImageList imgIconList 
      Left            =   120
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlThumbViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Thanks for the feedback.
' This version is greatly improved.
'
' The sample form shows how to initiate the control and
' load it with images.
' The code and commenting could be better but it works
' well and what do you expect for free!
'
' Steve D.
'
' Info:
'
' v.91
' Removed DisplayPics sub (merged into other subs)
' Fixed Thumbnail display sizing and display issues
' Fixed scroll bar change sizes
' Added selection and higlighting of pics
' Added scrolling
' Added events (mousedown, mouseup, click, dblclick)
' Added display file icons property (code by "the_cleaner")
' Added display pic property
' Improved display time (somewhat)
' Improved Sample App (yay)
' Improved function and variable management
'
' v.90
' This control is based on code obtained from a larger
' image viewing application written by Ramón A. Gimenez
' The original app was in Spanish and provided a number
' of image viewing functions. The most useful was the
' thumbnail viewer, and it was to well integrated into
' the host app to be useful in other programs.
' This control is the result of extracting the thumbnail
' viewer of the program in addition to translating the code
' from Spanish to English (I don't speak Spanish).
' It is now a self-contained user control that can be
' easily added to any project.
'
' The code is copyrighted by its respective owners, but
' can be used in any project you want.
'


Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Dim FileInfo As typSHFILEINFO

Public TopMargin              As Integer
Public DisplayPad             As Integer
Public PicHeight              As Integer
Public PicWidth               As Integer
Public PicGap                 As Integer
Public AutoColumn             As Boolean
Public ShowIcons              As Boolean
Public ShowPics               As Boolean
Public NumofCols              As Integer
Public NumofRows              As Integer

Private ImageSpace            As Integer
Private MarginSize            As Integer
Private NumofPics             As Integer
Dim imageFiles()              As String

Public currentPic             As Integer
Private lastPic               As Integer

Event click(picNum As Integer)
Event dblclick(picNum As Integer)
Event mousedown(picNum As Integer, Button As Integer)
Event mouseup(picNum As Integer, Button As Integer)


Public Sub refreshThumbs()
'call this sub to refresh the display
If NumofPics > 0 Then
    unloadPics 'get rid of previous controls
    thumbHolder.Visible = True 'show thumbnail holder
    ImageSpace = PicHeight + PicGap
    NumofCols = calcCols()  'figure out number of columns
    NumofRows = calcRows()  'figure out number of rows
    createPics  'load picture controls
    loadPics  'place picture frames
    loadImages  'load images from disk and display
End If
End Sub


Public Sub clearImages()
'this resets the array and clears everything
ReDim imageFiles(0)
imageFiles(0) = ""
imgIconList.ListImages.Clear
unloadPics
End Sub

Public Function getFileName(picNum As Integer)
'call this function to read the file name from the array by index
getFileName = imageFiles(picNum)
End Function

Public Function addImage(FileName As String)
' call this function to add to the file list array externally
Dim i As Integer
Dim r As Integer
    
' if the icon option is true then load the file icons into
' the icon image list
If ShowIcons = True Then
    r = ExtractIcon(FileName, imgIconList, picBuffer, 32)
End If

i = UBound(imageFiles)

ReDim Preserve imageFiles(i + 1)
imageFiles(UBound(imageFiles)) = FileName
NumofPics = UBound(imageFiles) + 1

End Function

Public Function calcRows() As Integer
If AutoColumn = True Then
    calcRows = Int(NumofPics / NumofCols) + 1
Else
    calcRows = NumofRows
End If
End Function

Public Function calcCols() As Integer
If AutoColumn = True Then
    calcCols = Int(thumbHolder.ScaleWidth / ((PicWidth + PicGap)))
Else
    calcCols = NumofCols
End If
End Function

Public Sub createPics()
On Error Resume Next

    Dim i As Integer
    
    For i = 0 To NumofPics - 1
        Load Pic2(i)
        Pic2(i).Visible = True
        Load Img(i)
    Next i
    
End Sub


Public Sub loadPics()
On Error Resume Next
Dim pH As Double

    Dim i As Integer, j As Integer, n As Integer
    
    DisplayPad = (controlPic.ScaleWidth) - (ImageSpace * (NumofCols))
    MarginSize = DisplayPad / 2
    
    pH = (NumofRows * (ImageSpace + TopMargin))
    thumbHolder.Height = (pH)
    thumbHolder.Refresh
    controlPic.Refresh
    
    VScroll.Max = controlPic.ScaleHeight - thumbHolder.ScaleHeight + 10
    VScroll.SmallChange = (PicHeight + PicGap) ' * 15
    VScroll.LargeChange = (controlPic.Height / ((PicHeight + 10))) * (PicHeight)
    
    n = 1
    For i = 0 To NumofRows - 1
        For j = 0 To NumofCols - 1
            If n >= NumofPics - 1 Then
                Exit Sub
            End If
            Pic2(n).Visible = False
            Pic2(n).Left = (j * (Pic2(n).Width + PicGap)) + (PicGap)
            Pic2(n).Top = (i * (((Pic2(n).Height) + TopMargin))) + (TopMargin)
            Pic2(n).Visible = True
        
Img(n).Visible = False
    If ShowIcons = True Then
        Img(n).Left = (Pic2(n).Left) + (Img(n).Width - ((Pic2(n).Width + PicGap)) / 2)
        Img(n).Top = (Pic2(n).Top) + (((Pic2(n).Height - (Img(n).Height)) / 2) + topborder)
        Img(n).Width = Pic2(n).Height
        Img(n).Height = Pic2(n).Height
        Img(n).Picture = imgIconList.ListImages(n).Picture
    End If
Img(n).Visible = True
            n = n + 1
        Next j
    Next i
        
End Sub


Public Sub loadImages()
On Error GoTo errorHandler
    Dim ratio As Double
    Dim i As Integer
    Dim p As PictureBox
    
    For i = 1 To NumofPics - 1
        Img(i).Stretch = True
        If ShowPics = True Then
        Img(i).Visible = False
        Img(i).Picture = LoadPicture(imageFiles(i))
        ratio = Img(i).Picture.Width / Img(i).Picture.Height
        
        If Img(i).Picture.Height >= Img(i).Picture.Width Then
            Img(i).Height = Pic2(i).Height - (topborder)
            Img(i).Width = Pic2(i).Width * ratio
            Else
            Img(i).Width = Pic2(i).Width
            Img(i).Height = Pic2(i).Height / ratio - (topborder)
        End If

        Img(i).Left = (Pic2(i).Left) + ((Pic2(i).Width - Img(i).Width) / 2)
        Img(i).Top = (Pic2(i).Top) + (((Pic2(i).Height - (Img(i).Height)) / 2) + topborder)
        
        Img(i).Visible = True
        Img(i).Refresh
        End If
    Next i
Exit Sub
errorHandler:
Resume Next

End Sub


Public Sub unloadPics()
On Error Resume Next

    Dim i As Integer
    
    For i = 1 To NumofPics - 1
        Unload Pic2(i)
        Unload Img(i)
    Next i
    
End Sub


Private Sub UserControl_Initialize()
    
    clearImages
    resizeThumbs
    Me.AutoColumn = True
    Me.ShowIcons = True
    Me.ShowPics = True
    Me.PicHeight = 75
    Me.PicWidth = 100
    Me.TopMargin = 10
    Me.PicGap = 10
    
Pic2(0).Height = PicHeight ' * 15
Pic2(0).Width = PicWidth ' * 15

End Sub



Private Sub UserControl_Resize()

    resizeThumbs

End Sub

Private Sub resizeThumbs()
    
    controlPic.Left = 0
    controlPic.Top = 0

    controlPic.Width = Width - VScroll.Width
    controlPic.Height = Height

    VScroll.Top = 0
    VScroll.Left = controlPic.Width
    VScroll.Height = controlPic.Height

    thumbHolder.Width = controlPic.ScaleWidth - (thumbHolder.Left * 2)
    
End Sub

Private Sub VScroll_Change()
    
    thumbHolder.Top = VScroll.Value

End Sub

Private Sub handleClick(picNum As Integer)
lastPic = currentPic
Pic2(lastPic).BorderColor = "&H000000"
currentPic = picNum
Pic2(currentPic).BorderColor = "&H0000FF"

End Sub

Private Sub VScroll_Scroll()
    thumbHolder.Top = VScroll.Value

End Sub

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDC, 0, 0, ILD_TRANSPARENT)
        '.Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mousedown(Index, Button)
End Sub

Private Sub Img_Mouseup(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent mousedown(Index, Button)
End Sub

Private Sub Img_Click(Index As Integer)
handleClick (Index)
RaiseEvent click(Index)
End Sub

Private Sub Img_DblClick(Index As Integer)
handleClick (Index)
RaiseEvent dblclick(Index)
End Sub


