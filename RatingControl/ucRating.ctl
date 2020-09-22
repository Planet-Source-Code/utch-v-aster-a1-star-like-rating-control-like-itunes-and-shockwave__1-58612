VERSION 5.00
Begin VB.UserControl ucRating 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ScaleHeight     =   405
   ScaleWidth      =   8625
   Begin VB.PictureBox picZero 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   30
      ScaleHeight     =   375
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   30
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3750
      Top             =   0
   End
   Begin VB.PictureBox picStar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   210
      Picture         =   "ucRating.ctx":0000
      ScaleHeight     =   375
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   30
      Width           =   180
   End
   Begin VB.PictureBox picStarOff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   900
      Picture         =   "ucRating.ctx":0275
      ScaleHeight     =   375
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picStarOff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1110
      Picture         =   "ucRating.ctx":0684
      ScaleHeight     =   375
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picStarOn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1320
      Picture         =   "ucRating.ctx":0A81
      ScaleHeight     =   375
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picStarOn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1530
      Picture         =   "ucRating.ctx":0E8E
      ScaleHeight     =   375
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Shape border 
      Height          =   165
      Left            =   0
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "ucRating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI
X As Long
Y As Long
End Type

Dim mWidth As Long
Dim mHeight As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public WatchHWND As Long


Const m_def_Value = 0
Const m_def_MaxStars = 5
Const m_def_Enabled = 0
Const m_def_ToolTipText = ""
Const m_def_OverValue = 0
Const m_def_BorderColor = vbBlack

Dim mArray As String

Dim m_Value As Single
Dim m_MaxStars As Integer
Dim m_Enabled As Boolean
Dim m_ToolTipText As String
Dim m_OverValue As Variant
Dim m_BorderColor As Long

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseLeave()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
   
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
  ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
  m_ToolTipText = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property

Private Sub picStar_Click(Index As Integer)
  m_Value = (Index + 1) * 0.5
  PropertyChanged "Value"
  DrawStars
  RaiseEvent Click
End Sub

Private Sub picStar_DblClick(Index As Integer)
  RaiseEvent DblClick
End Sub

Private Sub picStar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
  WatchHWND = UserControl.hWnd
  Timer1.Enabled = True
  
  m_OverValue = (Index + 1) * 0.5
  PropertyChanged "OverValue"
  
  Dim i As Integer
  For i = 0 To picStar.UBound
    If i <= Index Then
      If i / 2 = i \ 2 Then
        picStar(i).Picture = picStarOn(0).Image
      Else
        picStar(i).Picture = picStarOn(1).Image
      End If
    Else
      If i / 2 = i \ 2 Then
        picStar(i).Picture = picStarOff(0).Image
      Else
        picStar(i).Picture = picStarOff(1).Image
      End If
    End If
  Next
End Sub

Private Sub picZero_Click()
  m_Value = 0
  RaiseEvent Click
  PropertyChanged "Value"
  DrawStars
End Sub

Private Sub picZero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  
  m_OverValue = 0
  PropertyChanged "OverValue"
  
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
  For i = 0 To picStar.UBound
    If i / 2 = i \ 2 Then
      picStar(i).Picture = picStarOff(0).Image
    Else
      picStar(i).Picture = picStarOff(1).Image
    End If
  Next
End Sub

Private Sub Timer1_Timer()
  Dim pCursor As POINTAPI
  Dim hWindow As Long
  Dim nRet As Long
  Dim szText As String
  Dim X
  
  GetCursorPos pCursor
  X = WindowFromPointXY(pCursor.X, pCursor.Y)
  If InStr(1, mArray, Trim(Str(X))) = 0 Then
    RaiseEvent MouseLeave
    DrawStars
    Timer1.Enabled = False
  End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Enabled = m_def_Enabled
  m_ToolTipText = m_def_ToolTipText
  m_MaxStars = m_def_MaxStars
  m_Value = m_def_Value
  m_OverValue = m_def_OverValue
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  WatchHWND = UserControl.hWnd
  Timer1.Enabled = True
  DrawStars
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
  m_MaxStars = PropBag.ReadProperty("MaxStars", m_def_MaxStars)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  
  DrawStars
  m_OverValue = PropBag.ReadProperty("OverValue", m_def_OverValue)
  border.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
End Sub

Private Sub UserControl_Resize()
  Width = mWidth
  Height = mHeight
  border.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
  Call PropBag.WriteProperty("MaxStars", m_MaxStars, m_def_MaxStars)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("OverValue", m_OverValue, m_def_OverValue)
  Call PropBag.WriteProperty("BorderColor", border.BorderColor, -2147483640)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MaxStars() As Integer
  MaxStars = m_MaxStars
End Property

Public Property Let MaxStars(ByVal New_MaxStars As Integer)
  m_MaxStars = New_MaxStars
  PropertyChanged "MaxStars"
  If m_Value > m_MaxStars Then
    m_Value = m_MaxStars
    PropertyChanged "Value"
  End If
  DrawStars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Value() As Variant
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
  m_Value = New_Value
  PropertyChanged "Value"
  DrawStars
End Property

Sub DrawStars()
  Dim X As Single
  Dim i As Integer
  For X = 0 To picStar.UBound
    If X <> 0 Then Unload picStar(X)
  Next
  
  mArray = "," & UserControl.hWnd & "," & picStar(0).hWnd & ","
  
  For X = 0 To m_MaxStars - 0.5 Step 0.5
    i = (X / 0.5)
    If i > 0 Then
      Load picStar(i)
      mArray = mArray & picStar(i).hWnd & ","
      picStar(i).Top = picStar(i - 1).Top
      picStar(i).Left = picStar(i - 1).Left + picStar(i - 1).Width
      picStar(i).Visible = True
      If i / 2 = i \ 2 Then
        If X >= m_Value Then
          picStar(i).Picture = picStarOff(0).Image
        Else
          picStar(i).Picture = picStarOn(0).Image
        End If
      Else
        If X >= m_Value Then
          picStar(i).Picture = picStarOff(1).Image
        Else
          picStar(i).Picture = picStarOn(1).Image
        End If
      End If
    Else
      If m_Value = 0 Then
        picStar(0).Picture = picStarOff(0).Image
      Else
        picStar(0).Picture = picStarOn(0).Image
      End If
    End If
  Next
  
  mWidth = (picStar(i).Width + picStar(i).Left) + 60
  mHeight = picStar(i).Height + 60
  UserControl_Resize
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get OverValue() As Variant
  OverValue = m_OverValue
End Property

Public Property Let OverValue(ByVal New_OverValue As Variant)
  m_OverValue = New_OverValue
  PropertyChanged "OverValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=border,border,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
  BorderColor = border.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
  border.BorderColor() = New_BorderColor
  PropertyChanged "BorderColor"
End Property

