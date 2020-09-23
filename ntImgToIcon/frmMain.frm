VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ntImageToIcon"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4065
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDontSaveOutput 
      Caption         =   "Don't save output"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Used if you only want to copy to clipboard."
      Top             =   740
      Width           =   1575
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   40
      Width           =   255
   End
   Begin VB.CommandButton cmdCopyToClipboard 
      Caption         =   "Copy to Clipboard"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox chkFormOnTop 
      Caption         =   "Always on top"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Sets the app so its always on top."
      Top             =   480
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSComctlLib.ImageList imlConvert 
      Left            =   4200
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblInput 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgOutput 
      Height          =   855
      Left            =   360
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image imgInput 
      Height          =   855
      Left            =   360
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblSupportedImageTypes 
      Caption         =   "Converts BMP, GIF, JPG, WMF/EMF to ICO."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3345
   End
   Begin VB.Label lblUseInfo 
      Caption         =   "Drag and drop to convert an image."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strProgramName As String

'Created 3/19/00
'By Shannon Little
'codingman@yahoo.com
'http://go.to/neotrix

'It doesn't work too well with 256 color icons, I think the ExtractIcon
'method in the image list control doesn't work correctly because
'it imports 256 color images correctly but not when exporting them



Private Sub chkFormOnTop_Click()
    If chkFormOnTop.Value = vbChecked Then
        SetTopWindow hWnd, True
    Else
        SetTopWindow hWnd, False
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub cmdCopyToClipboard_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetData imgOutput
End Sub

Private Sub Form_Load()
    lblInput.Visible = False
    lblOutput.Visible = False
    imgInput.Visible = False
    imgOutput.Visible = False
    Height = 1500
    Width = 5000
    strProgramName = "ntImgToIcon"
    chkFormOnTop_Click
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Data.GetFormat(vbCFFiles) Then
        ConvertImages Data
    End If
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        'If the data is in the proper format, files
        'Show a good drop cursor, else show a no-drop
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub ConvertImages(Data As DataObject)
    On Error GoTo lError
    Dim tFile
    
    For Each tFile In Data.Files
        'Doesn't do icons
        If LCase(StrReturnRightFromEnd(tFile, ".")) <> "ico" Then
            Caption = "ntImageToIcon - " & StrReturnRightFromEnd(tFile, "\")
            imgInput.Picture = LoadPicture(tFile)
            imlConvert.ListImages.Clear
            imlConvert.ListImages.Add , , imgInput.Picture
            imgOutput.Picture = imlConvert.ListImages.Item(1).Picture
            
            'Now reload the output image if it was saved
            If chkDontSaveOutput.Value = vbUnchecked Then
                SavePicture imgOutput.Picture, Left(tFile, Len(tFile) - 3) & "ico"
                imgOutput.Picture = LoadPicture(Left(tFile, Len(tFile) - 3) & "ico")
            End If
            
            If imgInput.Picture = 0 Then   'Couldn't convert
                Caption = "ntImageToIcon - Couldn't convert image"
                cmdCopyToClipboard.Enabled = False
            Else
                cmdCopyToClipboard.Enabled = True
            End If
        Else
            'Its an icon, just quit
            Caption = "ntImageToIcon - Can't convert icons to icons"
            'If only 1 file was dragged and dropped then exit
            If Data.Files.Count = 1 Then Exit Sub
        End If
    Next tFile
    
    'imgOutput.Picture = LoadPicture(Left(tFile, Len(tFile) - 3) & "ico")
    'Resize form height
    lblInput.Visible = True
    lblOutput.Visible = True
    imgInput.Visible = True
    imgOutput.Visible = True
    Height = 2500 + 2 * imgInput.Height
    Width = 800 + imgInput.Width
    If Width < 5000 Then Width = 5000
    If Height < 1500 Then Height = 1500
    'Reposition img and lbl controls
    lblOutput.Top = imgInput.Height + imgInput.Top + 150
    imgOutput.Top = lblOutput.Top + 350
    
    Exit Sub
lError:
    Select Case Err.Number
        Case 7: TypError "Out of memory. Image is too large to be converted to an icon."
        Case 481:   'Not a valid picture
            TypError "Not a valid picture"
        Case Else: GenError
    End Select
End Sub

Private Sub Form_Resize()
    cmdAbout.Left = Width - cmdAbout.Width - 100
End Sub
