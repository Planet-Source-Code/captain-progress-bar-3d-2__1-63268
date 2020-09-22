VERSION 5.00
Begin VB.UserControl Progress3D 
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6060
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   ToolboxBitmap   =   "Progress3D.ctx":0000
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   0
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Progress3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_TextColor = 0
Const m_def_ShowProgress = 1
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
'Property Variables:
Dim m_TextColor As OLE_COLOR
Dim m_ShowProgress As Boolean
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
'Event Declarations:
Event Click() 'MappingInfo=pic,pic,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=pic,pic,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=pic,pic,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=pic,pic,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=pic,pic,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=pic,pic,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."

Private Sub fill3D(obj As Object, X As Single, Y As Single, wid As Long, hgt As Long) ',clr as Long = obj.forecolor)
Dim l1 As Integer, l2 As Integer
Dim dr As Integer, dg As Integer, db As Integer
Dim r As Integer, g As Integer, b As Integer
Dim clr As Long
Dim dy As Integer
Dim nr As Single, ng As Single, nb As Single

clr = obj.ForeColor

r = clr And &HFF&
g = (clr And &HFF00&) \ &H100
b = (clr And &HFF0000) \ &H10000

l1 = hgt / 2
l2 = (hgt - l1)

dr = (255 - r) / l1
dg = (255 - g) / l1
db = (255 - b) / l1

For dy = 0 To l1
    nr = 255 - dy * dr
    If nr < 0 Then nr = 0
    ng = 255 - dy * dg
    If ng < 0 Then ng = 0
    nb = 255 - dy * db
    If nb < 0 Then nb = 0
    obj.Line (X, Y + dy)-Step(wid, 0), RGB(nr, ng, nb)
Next

dr = r / l2
dg = g / l2
db = b / l2

For dy = 1 To l2
    nr = r - dr * dy
    If nr < 0 Then nr = 0
    ng = g - dg * dy
    If ng < 0 Then ng = 0
    nb = b - db * dy
    If nb < 0 Then nb = 0
    obj.Line (X, Y + l1 + dy)-Step(wid, 0), RGB(nr, ng, nb)
Next
End Sub

Private Sub refreshAll()
    If m_Value < m_Min Then
        m_Value = m_Min
        PropertyChanged "Value"
    ElseIf m_Value > m_Max Then
        m_Value = m_Max
        PropertyChanged "Value"
    End If
    pic.Cls
    fill3D pic, 0, 0, pic.ScaleWidth * ((m_Value - m_Min) / (m_Max - m_Min)), pic.ScaleHeight
    DisplayProgress
End Sub

Private Sub pic_LostFocus()
    refreshAll
End Sub

Private Sub UserControl_GotFocus()
    refreshAll
End Sub

Private Sub UserControl_LostFocus()
    refreshAll
End Sub

Private Sub UserControl_Paint()
    refreshAll
End Sub

Private Sub UserControl_Resize()
    pic.Height = UserControl.ScaleHeight
    pic.Width = UserControl.ScaleWidth
    refreshAll
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = pic.BackColor
    refreshAll
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    pic.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = pic.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    pic.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    refreshAll
End Property

Private Sub pic_Click()
    RaiseEvent Click
End Sub

Private Sub pic_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub pic_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub pic_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    pic.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_Description = "Sets/Returns the maximum value."
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    If New_Min > m_Max Then New_Min = m_Max
    m_Min = New_Min
    PropertyChanged "Min"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get Max() As Long
Attribute Max.VB_Description = "Sets/Returns the maximum value."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    If New_Max < m_Min Then New_Max = m_Min
    m_Max = New_Max
    PropertyChanged "Max"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Sets/Returns the current value."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    m_Value = New_Value
    PropertyChanged "Value"
    refreshAll
End Property

Public Sub DisplayProgress()
If m_ShowProgress = True Then
    Dim prog As String
    prog = CInt(((m_Value - m_Min) / (m_Max - m_Min)) * 100) & "%"
    pic.CurrentX = (pic.ScaleWidth - TextWidth(prog)) / 2
    pic.CurrentY = (pic.ScaleHeight - TextHeight(prog)) / 2
    Dim a As Long
    a = pic.ForeColor
    pic.ForeColor = m_TextColor
    pic.Print prog
    pic.ForeColor = a
End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_ShowProgress = m_def_ShowProgress
    m_TextColor = m_def_TextColor
    refreshAll
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pic.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    pic.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    pic.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_ShowProgress = PropBag.ReadProperty("ShowProgress", m_def_ShowProgress)
    Set pic.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", pic.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", pic.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("BorderStyle", pic.BorderStyle, 1)
    Call PropBag.WriteProperty("ShowProgress", m_ShowProgress, m_def_ShowProgress)
    Call PropBag.WriteProperty("Font", pic.Font, Ambient.Font)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = pic.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    pic.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ShowProgress() As Boolean
Attribute ShowProgress.VB_Description = "Sets/Returns the value to show the progress percentage"
    ShowProgress = m_ShowProgress
End Property

Public Property Let ShowProgress(ByVal New_ShowProgress As Boolean)
    m_ShowProgress = New_ShowProgress
    PropertyChanged "ShowProgress"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=pic,pic,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = pic.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set pic.Font = New_Font
    PropertyChanged "Font"
    refreshAll
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Sets/Returns Color of the text"
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    refreshAll
End Property
