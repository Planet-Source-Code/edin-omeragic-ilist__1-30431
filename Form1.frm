VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1230
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":154C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2228
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":49DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.IList IList1 
      Height          =   6720
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5010
      _extentx        =   8837
      _extenty        =   11853
      itemheight      =   25
      font            =   "Form1.frx":6C90
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'coded by edin omeragic

Dim X As Long

Private Sub Form_Load()
    
    Dim I As Long
    
    Set IList1.ImageList = ImageList1
    IList1.ItemHeight = 40
    IList1.SetPos 40, 4, 40, 20, 4, 4
    
    Dim T As Single
    Dim elt As Single
    
    T = Timer
    Dim NumItms
    NumItms = 1000
    For I = 1 To NumItms
        IList1.AddItem "Caption" + CStr(I), "Description" + CStr(I), , CLng(Rnd * 7) + 1
    Next
    
    Caption = CStr(NumItms) + " items for " + FormatNumber(Timer - T, 3) + " sec added"
    
End Sub

Private Sub IList1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then _
        PopupMenu mnuEdit
End Sub

Private Sub IList1_OnSelect()
    On Error Resume Next
    If X = 0 Then
        X = 1
        Exit Sub
    End If
    Dim Itm As CItem
    Set Itm = IList1.Item(IList1.Selected)
    
    Set Me.Icon = ImageList1.ListImages(Itm.Icon).ExtractIcon
    Me.Caption = Itm.Caption + "/" + Itm.Text
End Sub

Private Sub mnuRemove_Click()
    IList1.Remove IList1.Selected
End Sub

Private Sub mnuRename_Click()
    Dim X As String
    
    
    On Error Resume Next
    
    Dim Itm As CItem
    Set Itm = IList1.Item(IList1.Selected)
    X = InputBox("Enter new caption", , Itm.Caption)
    If X <> "" Then
        IList1.SetCaption IList1.Selected, X
    End If
    
    
    
    
End Sub
