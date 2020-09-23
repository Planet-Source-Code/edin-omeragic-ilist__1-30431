VERSION 5.00
Begin VB.UserControl IList 
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3525
      Left            =   90
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   282
      TabIndex        =   0
      Top             =   45
      Width           =   4290
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3
         Left            =   1035
         Top             =   1890
      End
      Begin VB.VScrollBar Scroll 
         Height          =   2805
         Left            =   4005
         Max             =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim CItems As New Collection
'Default Property Values:
Const m_def_Selected = 0
Const m_def_IconPosX = 0
Const m_def_IconPosY = 0
Const m_def_CaptionPosX = 0
Const m_def_CaptionPosY = 0
Const m_def_TextPosX = 0
Const m_def_TextPosY = 0
Const m_def_ItemHeight = 20
'Property Variables:

Dim m_Selected As Long
Dim m_IconPosX As Long
Dim m_IconPosY As Long
Dim m_CaptionPosX As Long
Dim m_CaptionPosY As Long
Dim m_TextPosX As Long
Dim m_TextPosY As Long
Dim m_ImageList As ImageList
Dim m_ItemHeight As Long

Dim Working As Boolean
Dim m_Scroll As Integer
'Event Declarations:
Event DblClick() 'MappingInfo=P,P,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=P,P,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=P,P,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=P,P,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=P,P,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=P,P,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=P,P,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Click() 'MappingInfo=P,P,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Public Event OnSelect()
'#######################################################################
'#######################################################################
Sub AddItem(Caption As String, Text As String, Optional Key As Variant, Optional Icon As Variant)
    Dim Item As New CItem
    Item.Caption = Caption
    Item.Text = Text
    If IsMissing(Icon) Then
        Item.Icon = 0
    Else
        Item.Icon = Icon
    End If
    If IsMissing(Key) Then
        CItems.Add Item
    Else
        CItems.Add Item
    End If
    SetScroll
End Sub
Sub Remove(Key As Variant)
    On Error Resume Next
    CItems.Remove Key
    Redraw
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "IList", Err.Description
    End If
End Sub
Sub Clear()
    Set CItems = Nothing
End Sub
Function Item(Key) As CItem
    On Error Resume Next
    Set Item = CItems.Item(Key)
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "IList", Err.Description
    End If
End Function
Function Count() As Long
    Count = CItems.Count
End Function
'#######################################################################
'#######################################################################

Private Sub P_Paint()
    Redraw
End Sub

Private Sub Scroll_Change()
    Redraw
End Sub

Private Sub Scroll_Scroll()
    Redraw
End Sub





Private Sub Timer1_Timer()


    Select Case m_Scroll
        Case Is = 1
            P_KeyDown vbKeyUp, 0
       Case Is = 2
            P_KeyDown vbKeyDown, 0
    End Select

End Sub

Private Sub UserControl_Paint()
    Redraw
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Redraw
    P.Move 0, 0, ScaleWidth, ScaleHeight
    Scroll.Move P.ScaleWidth - Scroll.Width, 0, Scroll.Width, P.ScaleHeight
    SetScroll
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,20
Public Property Get ItemHeight() As Long
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Long)
    m_ItemHeight = New_ItemHeight
    PropertyChanged "ItemHeight"
End Property
'#######################################################################
'#######################################################################
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ItemHeight = m_def_ItemHeight
    m_IconPosX = m_def_IconPosX
    m_IconPosY = m_def_IconPosY
    m_CaptionPosX = m_def_CaptionPosX
    m_CaptionPosY = m_def_CaptionPosY
    m_TextPosX = m_def_TextPosX
    m_TextPosY = m_def_TextPosY
    m_Selected = m_def_Selected
'    Set m_FontCaption = Ambient.Font
'    Set m_FontText = Ambient.Font
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_ItemHeight = PropBag.ReadProperty("ItemHeight", m_def_ItemHeight)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_IconPosX = PropBag.ReadProperty("IconPosX", m_def_IconPosX)
    m_IconPosY = PropBag.ReadProperty("IconPosY", m_def_IconPosY)
    m_CaptionPosX = PropBag.ReadProperty("CaptionPosX", m_def_CaptionPosX)
    m_CaptionPosY = PropBag.ReadProperty("CaptionPosY", m_def_CaptionPosY)
    m_TextPosX = PropBag.ReadProperty("TextPosX", m_def_TextPosX)
    m_TextPosY = PropBag.ReadProperty("TextPosY", m_def_TextPosY)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    P.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set P.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, m_def_ItemHeight)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("IconPosX", m_IconPosX, m_def_IconPosX)
    Call PropBag.WriteProperty("IconPosY", m_IconPosY, m_def_IconPosY)
    Call PropBag.WriteProperty("CaptionPosX", m_CaptionPosX, m_def_CaptionPosX)
    Call PropBag.WriteProperty("CaptionPosY", m_CaptionPosY, m_def_CaptionPosY)
    Call PropBag.WriteProperty("TextPosX", m_TextPosX, m_def_TextPosX)
    Call PropBag.WriteProperty("TextPosY", m_TextPosY, m_def_TextPosY)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", P.MousePointer, 0)
    Call PropBag.WriteProperty("Font", P.Font, Ambient.Font)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As ImageList
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As ImageList)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
End Property


Public Sub Redraw()
    
    Dim I As Long
 
    Dim Y As Long
    
    On Error Resume Next
    
    If Selected = 0 Then
        If Count > 0 Then
            Selected = 1
        End If
    End If
    SetScroll

    P.Cls
    Dim Ips As Long
    Ips = RoundEx(P.ScaleHeight / m_ItemHeight)
    
    For I = Scroll.Value + 1 To Count
        DrawItem I
        If I > Scroll.Value + Ips + 1 Then Exit For
    Next
    P.Refresh
    Working = False
End Sub


Sub DrawItem(Index)

On Error Resume Next

Dim Y As Long
Dim Itm As CItem

Set Itm = CItems(Index)
Y = (Index - Scroll.Value - 1) * m_ItemHeight

'Set forecolor and backcolor
If Selected = Index Then
   P.ForeColor = vbHighlightText
   Rectangle 0, Y, P.ScaleWidth, m_ItemHeight
Else
   P.ForeColor = vbButtonText
End If

'Print caption
P.FontBold = True
PrintAt m_CaptionPosX, Y + m_CaptionPosY, Itm.Caption
'Print text
P.FontBold = False
PrintAt m_TextPosX, Y + m_TextPosY, Itm.Text
'Draw picture
P.PaintPicture m_ImageList.ListImages(Itm.Icon).ExtractIcon, _
               m_IconPosX, m_IconPosY + Y
End Sub


'Api (heh)
Sub PrintAt(X As Long, Y As Long, Text As String)
    P.CurrentX = X
    P.CurrentY = Y
    P.Print Text
End Sub
Sub MoveTo(X, Y)
    P.CurrentX = X
    P.CurrentY = Y
End Sub

Sub LineTo(X, Y, Optional Color As Long = 0)
    P.Line -(X, Y), Color
End Sub
Sub TextOut(Text As String)
    P.Print Text
End Sub
Sub Rectangle(X As Long, Y As Long, Width As Long, Height As Long, _
              Optional Color As Long = vbHighlight)
    P.Line (X, Y)-Step(Width, Height), Color, BF
End Sub
Function RoundEx(X)
    If X > CLng(X) Then
        RoundEx = CLng(X) + 1
    Else
        RoundEx = CLng(X)
    End If
End Function




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get IconPosX() As Long
    IconPosX = m_IconPosX
End Property

Public Property Let IconPosX(ByVal New_IconPosX As Long)
    m_IconPosX = New_IconPosX
    PropertyChanged "IconPosX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get IconPosY() As Long
    IconPosY = m_IconPosY
End Property

Public Property Let IconPosY(ByVal New_IconPosY As Long)
    m_IconPosY = New_IconPosY
    PropertyChanged "IconPosY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionPosX() As Long
    CaptionPosX = m_CaptionPosX
End Property

Public Property Let CaptionPosX(ByVal New_CaptionPosX As Long)
    m_CaptionPosX = New_CaptionPosX
    PropertyChanged "CaptionPosX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionPosY() As Long
    CaptionPosY = m_CaptionPosY
End Property

Public Property Let CaptionPosY(ByVal New_CaptionPosY As Long)
    m_CaptionPosY = New_CaptionPosY
    PropertyChanged "CaptionPosY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TextPosX() As Long
    TextPosX = m_TextPosX
End Property

Public Property Let TextPosX(ByVal New_TextPosX As Long)
    m_TextPosX = New_TextPosX
    PropertyChanged "TextPosX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TextPosY() As Long
    TextPosY = m_TextPosY
End Property

Public Property Let TextPosY(ByVal New_TextPosY As Long)
    m_TextPosY = New_TextPosY
    PropertyChanged "TextPosY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Selected() As Long
    Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Long)

Dim Y As Long
Dim T As Long
If New_Selected > Count Then New_Selected = Count

If New_Selected <> m_Selected Then
        'Clear
        T = m_Selected
        m_Selected = New_Selected
        
        Y = (T - Scroll.Value - 1) * m_ItemHeight
        Rectangle 0, Y, P.ScaleWidth, m_ItemHeight, vbWhite
        DrawItem T
        DrawItem m_Selected
        
        RaiseEvent OnSelect
End If

    PropertyChanged "Selected"
End Property
Sub SetPos(CaptionX As Long, CaptionY As Long, _
           TextX As Long, TextY As Long, _
           IconX As Long, IconY As Long)
    m_CaptionPosX = CaptionX
    m_CaptionPosY = CaptionY
    m_TextPosX = TextX
    m_TextPosY = TextY
    m_IconPosX = IconX
    m_IconPosY = IconY
    Redraw
End Sub

Function IsVisible(Index As Long) As Boolean
    Dim Ips As Long
    Ips = (P.ScaleHeight \ m_ItemHeight)
    If Index > Scroll.Value And Index < Scroll.Value + Ips + 1 Then
        IsVisible = True
    End If
End Function

Sub ScrollTo(Index As Long)
    Dim Ips As Long
    Ips = (P.ScaleHeight \ m_ItemHeight)
    If Scroll.Visible = False Then Exit Sub
    If Count > Index + Ips Then
        Scroll.Value = Index - 1
    Else
        Scroll.Value = Count - Ips
    End If
End Sub


Private Sub SetScroll()
    
    Scroll.Max = Count - Int(P.ScaleHeight / m_ItemHeight)
    If Scroll.Max <= 0 Then
        Scroll.Max = 0
        Scroll.Visible = False
    Else
        Scroll.Visible = True
    End If
End Sub

Private Sub P_Click()
    RaiseEvent Click
  
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = P.hWnd
End Property

Private Sub P_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Selected = RoundEx(Y / m_ItemHeight) + Scroll.Value
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = P.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set P.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub P_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button = 1 Then
        If Y > 0 And Y < P.ScaleHeight Then
            Timer1.Enabled = False
            Selected = RoundEx(Y / m_ItemHeight) + Scroll.Value
        Else
        
            If Y < 0 Then
                m_Scroll = 1
            ElseIf Y > P.ScaleHeight Then
                m_Scroll = 2
            End If
            Timer1.Enabled = True
        End If
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = P.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    P.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub P_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub P_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = P.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set P.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=P,P,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = P.hDC
End Property

Private Sub P_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    On Error Resume Next
    If Working = True Then Exit Sub
    Select Case KeyCode
        Case Is = vbKeyUp
            If Selected > 1 Then Selected = Selected - 1
            If IsVisible(Selected) = False Then
                If Scroll.Value > 0 Then
                    'DoEvents
                    Working = True
                    Scroll.Value = Scroll.Value - 1
                End If
            End If
        Case Is = vbKeyDown
            If Selected < Count Then Selected = Selected + 1
            If IsVisible(Selected) = False Then
                If Scroll.Value < Scroll.Max Then
                    'DoEvents
                    Working = True
                    Scroll.Value = Scroll.Value + 1
                End If
            End If
    End Select
End Sub

Private Sub P_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub P_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Sub SetCaption(Index, Caption As String)
    CItems(Index).Caption = Caption
    DrawItem Index
End Sub


Public Sub SetText(Index, Text As String)
    CItems(Index).Text = Text
    DrawItem Index
End Sub
