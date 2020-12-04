VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmVB6x 
   Caption         =   "Visual Basic 6.5"
   ClientHeight    =   2070
   ClientLeft      =   315
   ClientTop       =   1170
   ClientWidth     =   4185
   Icon            =   "frmVB6x.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4185
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   240
      Top             =   960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":12AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":1600
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":1E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVB6x.frx":21C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   960
      Top             =   15
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmVB6x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function apiSetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private once As Boolean

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim b As Boolean
    b = Not (vbEnv.ActiveVBProject Is Nothing)
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Remove Project").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Sa&ve Project").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Sav&e Project As...").enabled = b
    If vbEnv.SelectedVBComponent Is Nothing Then
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Save Code Module").enabled = False
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Save Code Module &As...").enabled = False
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Print Code Module...").enabled = False
        vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Remove Code Module").enabled = False
    Else
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Save Code Module").enabled = True
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Save Code Module &As...").enabled = True
        vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Print Code Module...").enabled = True
        vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Remove Code Module").enabled = True
    End If
    Dim vbComBarButton As Object
    Set vbComBarButton = vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Make Project &Group... ")
    If vbEnv.VBProjects.Count < 2 Then
       ' vbComBarButton.Picture = LoadPicture(app.Path & "\images\Blank.bmp")
        vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(1).Picture
        vbComBarButton.enabled = False
        'vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Make Project &Group... ").enabled = False
    Else
        vbComBarButton.enabled = True
        'vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Make Project &Group... ").enabled = True
        'vbComBarButton.Picture = LoadPicture(app.Path & "\images\ProjectGroup16.bmp")
        vbComBarButton.Picture = frmVB6x.ImageList1.ListImages.Item(2).Picture
    End If
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("Ma&ke .exe...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Import File").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&File").Controls("&Export File").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Add File...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("Refere&nces...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("C&omponents... ").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("&Additional Controls...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Project").Controls("Prop&erties...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("&Start").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("Start With &Full Compile").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("Brea&k ").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Run").Controls("&End").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("Add &Procedure...").enabled = b
    vbEnv.CommandBars("Menu Bar").Controls("&Tools").Controls("&Options... ").enabled = b
End Sub

Private Sub Timer2_Timer()
LoadToolbox
'Timer2.enabled = False
End Sub
Private Sub LoadToolbox()
    On Error Resume Next
    Dim twnd As Long
    twnd = apiFindWindow(vbNullString, "Toolbox")
    If twnd = 0 Then Exit Sub
    Dim vwnd As Long
    vwnd = apiFindWindow(vbNullString, vbEnv.MainWindow.caption)
    If vwnd = 0 Then Exit Sub
    Dim rt  As RECT
    Dim rv  As RECT
    Dim rvc As RECT
    If apiGetWindowRect(twnd, rt) <> 0 Then
        If apiGetWindowRect(vwnd, rv) <> 0 Then
            Dim w As Long
            Dim h As Long
            w = (rt.Right - rt.Left)
            h = (rt.Bottom - rt.Top)
            apiMoveWindow twnd, 6, 64, w, h, 1
        End If
    End If
    apiSetParent twnd, vwnd
End Sub
Private Sub tmrMain_Timer()
    tmrMain.enabled = False
    apc.FormatCodeModule vbEnv.SelectedVBComponent.CodeModule, True
End Sub
