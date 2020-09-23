VERSION 5.00
Begin VB.UserControl PDial 
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ToolboxBitmap   =   "PDial.ctx":0000
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1170
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   1
      Top             =   1305
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2520
      Picture         =   "PDial.ctx":0312
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   0
      Top             =   855
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "PDial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------
'
' This Code is done by Jörg Pfeffer (Peppa)
' Feel free to use this Code/Control for your Projects
' Sorry for not commenting
'
'-------------------------------------------------------------

Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal fuFlags As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


'DrawSateTypes
Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4
Const DSS_NORMAL = &H0
Const DSS_UNION = &H10 ' Dither
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80 ' Draw in colour of brush specified in hBrush
Const DSS_RIGHT = &H8000


Const Pi As Double = 3.14159265358979
Const TwoPI As Double = 2 * Pi

Dim m_Min As Integer     'Minimum value
Dim m_Max As Integer     'Maximum value
Dim m_Value As Integer   'Current dial value
Dim m_NullGrad As Integer
Dim Winkel As Single

Dim m_ToolTipText As String
Dim m_LColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_KnobColor As OLE_COLOR
Dim m_DrehColor As OLE_COLOR
Dim m_DrehColOff As OLE_COLOR
Dim m_DrehShow As Byte
Dim m_KnobImage As StdPicture
Dim m_Transparent As Boolean
Dim m_LPoint As Boolean
Dim m_TextShow As Byte
Dim m_TextColor As OLE_COLOR
Dim m_TicksColor As OLE_COLOR
Dim m_Text As String
Dim m_AutoSize As Boolean
Dim m_Abstand As Integer


Dim m_LRadius As Integer
Dim m_KnobRadius As Integer
Dim m_DrehRadius As Integer

Dim WasMax As Boolean
Dim WasMin As Boolean

Dim OldX As Integer
Dim OldY As Integer

Dim imgW As Integer
Dim imgH As Integer

Dim uCHi As Integer
Dim uCWi As Integer

Event Changing(iValue As Integer) 'Fires when angle changes during movement
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick


Private Type POINT
  X As Integer
  Y As Integer
End Type



Private Function GetAngle(Focus As POINT) As Integer
Dim Radians As Double
Dim offset As Integer
Dim os As POINT

os.X = Fix(uCWi / 2)
os.Y = Fix(uCHi / 2)

'''Debug.Print Focus.X & ":" & os.X & "  /  " & Focus.Y & ":" & os.Y

  ' Get angle value in radians
  If (Focus.X - os.X) <> 0 Then
    Radians = Atn((Abs(Focus.Y - os.Y)) / ((Abs(Focus.X - os.X))))
  Else
    ' Vertical. Determine if azimuth is north or south
    If Focus.Y > os.Y Then
      GetAngle = 0#
      Exit Function
    Else
      GetAngle = 180#
      Exit Function
    End If
  End If


  ' Determine Offset
  If (Focus.X > os.X) And (Focus.Y + 0.01 < os.Y) Then
  ' upper left quadrant
    offset = -90

  ElseIf (Focus.X > os.X) And (Focus.Y + 0.01 > os.Y) Then
  ' upper right quadrant
    offset = 90

  ElseIf (Focus.X < os.X) And (Focus.Y + 0.01 > os.Y) Then
  ' Lower right quadrant
    offset = -270

  ElseIf (Focus.X < os.X) And (Focus.Y + 0.01 < os.Y) Then
  ' Lower left quadrant
    offset = 270

  End If

  GetAngle = Abs(offset + (Radians * (180 / Pi)))
  GetAngle = GetAngle + 180
  If GetAngle > 360 Then GetAngle = GetAngle - 360

End Function



Public Property Get Abstand() As Integer
    Abstand = m_Abstand
End Property

Public Property Let Abstand(ByVal New_Abstand As Integer)
    m_Abstand = New_Abstand
    PropertyChanged "Abstand"
    Call UserControl_Resize
End Property




Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize
End Property




Public Property Get BackColor() As OLE_COLOR
 BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 m_BackColor = New_BackColor
 PropertyChanged "BackColor"
 ResetDial
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
'    ResetDial
End Property


Public Property Get NullGrad() As Integer
    NullGrad = m_NullGrad
End Property

Public Property Let NullGrad(ByVal New_NullGrad As Integer)
    m_NullGrad = New_NullGrad
    PropertyChanged "NullGrad"
    ResetDial
End Property



Public Property Get min() As Long
  min = m_Min
End Property

Public Property Let min(ByVal New_Min As Long)
 If New_Min < m_Max Then
  m_Min = New_Min
  If m_Min > m_Value Then m_Value = m_Min
  ResetDial
  PropertyChanged "Min"
 End If
End Property

Public Property Get max() As Long
  max = m_Max
End Property

Public Property Let max(ByVal New_Max As Long)
 If New_Max > m_Min Then
  m_Max = New_Max
  If m_Max < m_Value Then m_Value = m_Max
  ResetDial
  PropertyChanged "Max"
 End If
End Property


Public Property Get Value() As Long
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
 If New_Value >= m_Max Then New_Value = m_Max
 If New_Value <= m_Min Then New_Value = m_Min
 m_Value = Fix(New_Value)
 PropertyChanged "Value"
 ResetDial
 RaiseEvent Changing(Value)

End Property

Public Property Get KnobColor() As OLE_COLOR
 KnobColor = m_KnobColor
End Property

Public Property Let KnobColor(ByVal New_KnobColor As OLE_COLOR)
 m_KnobColor = New_KnobColor
 PropertyChanged "KnobColor"
ResetDial
End Property


Public Property Get KnobRadius() As Integer
    KnobRadius = m_KnobRadius
End Property

Public Property Let KnobRadius(ByVal New_KnobRadius As Integer)
    m_KnobRadius = New_KnobRadius
    PropertyChanged "KnobRadius"
    ResetDial
End Property

Public Property Get KnobImage() As StdPicture
    Set KnobImage = m_KnobImage
End Property

Public Property Set KnobImage(ByVal New_KnobImage As StdPicture)
    Set m_KnobImage = New_KnobImage
    PropertyChanged "KnobImage"
    Call BeVor
End Property


Public Property Get LColor() As OLE_COLOR
 LColor = m_LColor
End Property

Public Property Let LColor(ByVal New_LColor As OLE_COLOR)
 m_LColor = New_LColor
 PropertyChanged "LColor"
ResetDial
End Property

Public Property Get LRadius() As Integer
    LRadius = m_LRadius
End Property

Public Property Let LRadius(ByVal New_LRadius As Integer)
    m_LRadius = New_LRadius
    PropertyChanged "LRadius"
    ResetDial
End Property



Public Property Get DrehColor() As OLE_COLOR
 DrehColor = m_DrehColor
End Property

Public Property Let DrehColor(ByVal New_DrehColor As OLE_COLOR)
 m_DrehColor = New_DrehColor
 PropertyChanged "DrehColor"
ResetDial
End Property

Public Property Get DrehColOff() As OLE_COLOR
 DrehColOff = m_DrehColOff
End Property

Public Property Let DrehColOff(ByVal New_DrehColOff As OLE_COLOR)
 m_DrehColOff = New_DrehColOff
 PropertyChanged "DrehColOff"
ResetDial
End Property

Public Property Get DrehRadius() As Integer
    DrehRadius = m_DrehRadius
End Property

Public Property Let DrehRadius(ByVal New_DrehRadius As Integer)
    m_DrehRadius = New_DrehRadius
    PropertyChanged "DrehRadius"
    ResetDial
End Property


Public Property Get DrehShow() As Byte
 DrehShow = m_DrehShow
End Property

Public Property Let DrehShow(ByVal New_DrehShow As Byte)
 m_DrehShow = New_DrehShow
 PropertyChanged "DrehShow"
ResetDial
End Property



Public Property Get TRANSPARENT() As Boolean
    TRANSPARENT = m_Transparent
End Property

Public Property Let TRANSPARENT(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    ResetDial
End Property

Public Property Get LPoint() As Boolean
    LPoint = m_LPoint
End Property

Public Property Let LPoint(ByVal New_LPoint As Boolean)
    m_LPoint = New_LPoint
    PropertyChanged "LPoint"
    ResetDial
End Property

Public Property Get TextColor() As OLE_COLOR
 TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
 m_TextColor = New_TextColor
 PropertyChanged "TextColor"
ResetDial
End Property


Public Property Get TicksColor() As OLE_COLOR
 TicksColor = m_TicksColor
End Property

Public Property Let TicksColor(ByVal New_TicksColor As OLE_COLOR)
 m_TicksColor = New_TicksColor
 PropertyChanged "TicksColor"
ResetDial
End Property


Public Property Get TextShow() As Byte
    TextShow = m_TextShow
End Property

Public Property Let TextShow(ByVal New_TextShow As Byte)
    m_TextShow = New_TextShow
    PropertyChanged "TextShow"
    ResetDial
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    ResetDial
End Property



Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get HasDC() As Boolean
    HasDC = UserControl.HasDC
End Property




'-----------------------------



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
   UserControl_MouseMove Button, Shift, X, Y
 End If
RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim fc As POINT
Dim iWert As Integer
Dim isAngle As Integer

If OldX = X And OldY = Y Then Exit Sub

If Button <> 1 Then GoTo XOut:

fc.X = X
fc.Y = Y

OldX = X
OldY = Y

isAngle = GetAngle(fc)
iWert = m_Max / (340) * (isAngle - 20)

If WasMax = True And isAngle <= 180 Then Exit Sub
If WasMin = True And isAngle >= 180 Then Exit Sub

WasMax = False: WasMin = False
If isAngle <= 20 Then iWert = 0: WasMin = True:
If isAngle >= 340 Then iWert = m_Max: WasMax = True

If iWert = m_Value Then DoEvents: Exit Sub

m_Value = iWert
DoEvents
ResetDial

 RaiseEvent Changing(Value)
 
XOut:
 RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WasMax = False
WasMin = False
RaiseEvent MouseUp(Button, Shift, X, Y)
RaiseEvent Changing(Value)
DoEvents
End Sub

Private Sub UserControl_Resize()

 If m_AutoSize = True Then
   UserControl.Height = (Image1.Height + m_Abstand + 2) * Screen.TwipsPerPixelY
   UserControl.Width = (Image1.Width + m_Abstand) * Screen.TwipsPerPixelX
 Else
   UserControl.Height = UserControl.Width + 1
 End If
 
 uCWi = UserControl.ScaleWidth
 uCHi = UserControl.ScaleHeight - 2
 
 Call BeVor
End Sub

Private Sub UserControl_DblClick()
 RaiseEvent DblClick
End Sub




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Abstand = PropBag.ReadProperty("Abstand", 2)
    m_AutoSize = PropBag.ReadProperty("AutoSize", True)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    
    m_NullGrad = PropBag.ReadProperty("NullGrad", 12)
    m_Min = PropBag.ReadProperty("Min", 0)
    m_Max = PropBag.ReadProperty("Max", 100)
    m_Value = PropBag.ReadProperty("Value", 0)
    
    m_LColor = PropBag.ReadProperty("LColor", &HFF&)
    m_LPoint = PropBag.ReadProperty("LPoint", False)
    m_LRadius = PropBag.ReadProperty("LRadius", 7)
    
    m_DrehColor = PropBag.ReadProperty("DrehColor", vbWhite)
    m_DrehColOff = PropBag.ReadProperty("DrehColOff", vbBlue)
    m_DrehRadius = PropBag.ReadProperty("DrehRadius", 16)
    m_DrehShow = PropBag.ReadProperty("DrehShow", 1)
    
    m_KnobColor = PropBag.ReadProperty("KnobColor", vbWhite)
    Set m_KnobImage = PropBag.ReadProperty("KnobImage", Nothing)
    m_KnobRadius = PropBag.ReadProperty("KnobRadius", 9)
    
    m_TextColor = PropBag.ReadProperty("TextColor", vbBlack)
    m_TicksColor = PropBag.ReadProperty("TicksColor", m_TextColor)
    
    m_TextShow = PropBag.ReadProperty("TextShow", 2)
    m_Text = PropBag.ReadProperty("Text", "Lo Hi")

    m_Transparent = PropBag.ReadProperty("Transparent", True)

Call BeVor
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Abstand", m_Abstand, 2)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, True)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, "")
    Call PropBag.WriteProperty("BackColor", m_BackColor, vbButtonFace)
    
    Call PropBag.WriteProperty("NullGrad", m_NullGrad, 12)
    
    Call PropBag.WriteProperty("Min", m_Min, 0)
    Call PropBag.WriteProperty("Max", m_Max, 100)
    Call PropBag.WriteProperty("Value", m_Value, 0)
    
    Call PropBag.WriteProperty("LColor", m_LColor, &HFF&)
    Call PropBag.WriteProperty("LPoint", m_LPoint, False)
    Call PropBag.WriteProperty("LRadius", m_LRadius, 7)
    
    Call PropBag.WriteProperty("KnobColor", m_KnobColor, vbWhite)
    Call PropBag.WriteProperty("KnobImage", m_KnobImage, Nothing)
    Call PropBag.WriteProperty("KnobRadius", m_KnobRadius, 9)
    
    Call PropBag.WriteProperty("DrehColor", m_DrehColor, vbWhite)
    Call PropBag.WriteProperty("DrehColOff", m_DrehColOff, vbBlue)
    Call PropBag.WriteProperty("DrehShow", m_DrehShow, 1)
    Call PropBag.WriteProperty("DrehRadius", m_DrehRadius, 16)
    
    Call PropBag.WriteProperty("TextColor", m_TextColor, vbBlack)
    Call PropBag.WriteProperty("TicksColor", m_TicksColor, m_TextColor)
    Call PropBag.WriteProperty("TextShow", m_TextShow, 2)
    Call PropBag.WriteProperty("Text", m_Text, "Lo Hi")

    Call PropBag.WriteProperty("Transparent", m_Transparent, True)

'ResetDial
End Sub

Private Sub UserControl_Initialize()
 '
 
End Sub


Private Sub UserControl_InitProperties()
 m_AutoSize = True
 m_Abstand = 2
 m_NullGrad = 12
 m_Min = 0
 m_Max = 100
 m_Value = 0
 m_LColor = vbRed
 m_LRadius = 7
 m_LPoint = False
 m_BackColor = vbButtonFace
 m_DrehColor = vbWhite
 m_DrehColOff = vbBlue
 m_DrehShow = 1
 m_DrehRadius = 16
 Set m_KnobImage = Picture1
 m_KnobRadius = 9
 m_KnobColor = RGB(255, 255, 255)
 m_Transparent = True
 m_TextShow = 2
 m_Text = "Lo Hi"
 m_TextColor = vbBlack
 m_TicksColor = vbBlack
 Call UserControl_Resize

End Sub

Private Sub BeVor()
Dim eDraw As Long

 UserControl.AutoRedraw = True
 Set UserControl.Picture = Nothing
 UserControl.Cls
 
 eDraw = DST_BITMAP
 UserControl.BackColor = m_BackColor
 Set Image1 = m_KnobImage
 If Image1 = 0 Then
  Set Image1 = Picture1
 End If
 
 imgW = Image1.Width
 imgH = Image1.Height
 
 If Image1 <> 0 Then
  If Image1.Picture.Type = vbPicTypeIcon Then eDraw = DST_ICON
  eDraw = eDraw Or DSS_NORMAL Or DSS_RIGHT
  DrawState UserControl.hdc, 0, 0, Image1.Picture.Handle, 0, Fix((uCWi - Image1.Width) / 2), Fix((uCHi - Image1.Height) / 2), Image1.Width, Image1.Height, eDraw
 End If
 
 UserControl.DrawWidth = 1
 Set UserControl.Picture = UserControl.Image
 UserControl.AutoRedraw = False
 
' Set Image1 = Nothing
 Call ResetDial
 
End Sub

Private Sub ResetDial()
Dim Cn As Integer
Dim mAB As Integer
Dim RCn As Integer
Dim R As Integer
Dim RS As Integer
Dim RS2 As Integer
Dim aMax As Integer
Dim XPos As Integer
Dim YPos As Integer
Dim XPos2 As Integer
Dim YPos2 As Integer
  Dim NulPos As Integer
  Dim St As Double, En As Single
  Dim Dg As Single
  Dim Dg2 As Single
  Dim aValue As Single
  Dim Col As OLE_COLOR
  Dim dVal As Integer
  Dim dMax As Integer
  Dim T As Integer
  Dim I As Integer
  Dim TX As String
    
    
UserControl.FillColor = m_LColor
aMax = m_Max '- m_Min
Cn = Fix(uCWi / 2)
R = m_LRadius
m_Min = 0
RS = R
RCn = Cn

UserControl.AutoRedraw = True
 UserControl.Cls
' Call BeVor
 
'Innen Knopf Farbe
 If m_KnobColor <> vbWhite Then
   UserControl.FillStyle = 0
   UserControl.DrawWidth = 1
   UserControl.DrawMode = 13
   UserControl.FillColor = m_KnobColor
   UserControl.Circle (Cn, Cn), m_KnobRadius, m_KnobColor
   UserControl.DrawMode = 13
 End If


  UserControl.FillStyle = 0


 NulPos = m_NullGrad
 Winkel = Fix((360 / (360 - NulPos) * ((360 / m_Max * m_Value))))
  If Winkel <= NulPos Then Winkel = NulPos
  If Winkel >= 360 - NulPos Then Winkel = 360 - NulPos
 
 If m_LPoint = False Then
'------------------ Zeiger
    Dg = 360 / (360 + NulPos + NulPos) * (Winkel - 270 + 3 - (NulPos / 2))
    If Dg >= 0 Then Dg = Dg - 360
 
    Dg2 = 360 / (360 + NulPos + NulPos) * (Winkel - 270 - 3 - (NulPos / 2))
    If Dg2 >= 0 Then Dg2 = Dg2 - 360
    
    If Abs(Dg) > Abs(Dg2) Then Dg = 0
      
      St = (Dg * Pi / 180) 'Umwandlung von Grad in Bogenmaß
      En = (Dg2 * Pi / 180)
      UserControl.Circle (RCn, RCn), m_LRadius, m_LColor, St, En
 Else
'----------------- Punkt
   UserControl.DrawWidth = 3
   mAB = 1
   R = R
   aValue = Winkel: GoSub CalcAusen
   Col = m_LColor
   UserControl.Line (XPos, YPos)-(XPos2, YPos2), Col
   UserControl.DrawWidth = 1
 
 End If


'Außen Laufbahn
 If m_DrehShow >= 1 Then
    mAB = 1
    
    dVal = Fix((360 / (360 - NulPos) * ((360 / m_Max * m_Value))))
    For T = 0 To dVal
      R = m_DrehRadius
      aValue = T: GoSub CalcAusen
      UserControl.Line (XPos, YPos)-(XPos2, YPos2), m_DrehColor
    
      R = m_DrehRadius + 1: GoSub CalcAusen
      UserControl.Line (XPos, YPos)-(XPos2, YPos2), m_DrehColor
    Next T
 
     If m_DrehShow = 2 Then
       dMax = Fix((360 / (360 - NulPos) * 360))
       For T = dVal To dMax
         R = m_DrehRadius
         aValue = T: GoSub CalcAusen
         UserControl.Line (XPos, YPos)-(XPos2, YPos2), m_DrehColOff
       Next T
      
     End If

 End If



 If m_TextShow >= 1 Then
'----- Rasterung
    R = m_DrehRadius
    mAB = m_Abstand
    NulPos = 0
   
    For T = 1 To 7
      aValue = (45 * T): GoSub CalcAusen
      UserControl.Line (XPos, YPos)-(XPos2, YPos2), m_TicksColor
    Next T
 End If


'Beschriftung
  UserControl.FillStyle = 1
  UserControl.DrawWidth = 1
  UserControl.DrawMode = 13

  If m_TextShow >= 2 Then
    I = InStr(1, m_Text, " ", vbTextCompare)
    TX = Left(m_Text, I - 1)
    SetTextColor UserControl.hdc, m_TextColor
    TextOut UserControl.hdc, 0, UserControl.ScaleHeight - UserControl.TextHeight(TX) + 1, TX, Len(TX)

    TX = Mid(m_Text, I + 1, 100)
    SetTextColor UserControl.hdc, m_TextColor
    TextOut UserControl.hdc, uCWi - UserControl.TextWidth(TX) - 1, UserControl.ScaleHeight - UserControl.TextHeight(TX) + 1, TX, Len(TX)
  End If


If m_Transparent = True Then
 UserControl.MaskColor = UserControl.BackColor
 UserControl.BackStyle = 0
 Set UserControl.MaskPicture = UserControl.Image
Else
 UserControl.BackStyle = 1
End If
 
UserControl.AutoRedraw = False
Exit Sub





CalcAusen:
   
   If aValue <= NulPos Then aValue = NulPos
   If aValue >= 360 - NulPos Then aValue = 360 - NulPos
   
   Dg = 360 / (360 + NulPos + NulPos) * (aValue - 180)
   XPos = Cn + (R * (Sin((Dg * Pi) / 180)))
   YPos = Cn - (R * (Cos((Dg * Pi) / 180)))
   
   XPos2 = Cn + ((R + mAB) * (Sin((Dg * Pi) / 180)))
   YPos2 = Cn - ((R + mAB) * (Cos((Dg * Pi) / 180)))
Return



End Sub

