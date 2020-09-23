VERSION 5.00
Begin VB.Form FTest 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   StartUpPosition =   2  'CenterScreen
   Begin Test.PDial PDial1 
      Height          =   585
      Index           =   0
      Left            =   1035
      TabIndex        =   0
      Top             =   990
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1032
      KnobImage       =   "FTest.frx":000C
      TextColor       =   8454143
      TicksColor      =   14737632
   End
   Begin Test.PDial PDial1 
      Height          =   615
      Index           =   1
      Left            =   2655
      TabIndex        =   2
      Top             =   990
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   1085
      Abstand         =   4
      Max             =   24
      Value           =   8
      LColor          =   32768
      KnobImage       =   "FTest.frx":1042
      DrehColor       =   33023
      DrehColOff      =   12632256
      DrehShow        =   2
      TextColor       =   65280
      TicksColor      =   33023
      Text            =   "-12 +12"
   End
   Begin Test.PDial PDial1 
      Height          =   585
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   1035
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1032
      NullGrad        =   0
      Max             =   500
      Value           =   150
      LColor          =   16711680
      LPoint          =   -1  'True
      LRadius         =   5
      KnobImage       =   "FTest.frx":2078
      DrehColor       =   33023
      DrehColOff      =   12632256
      DrehShow        =   0
      TextColor       =   65280
      TicksColor      =   12632256
      TextShow        =   1
      Text            =   ""
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   3870
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2610
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   945
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

 PDial1(0).Value = 1
 PDial1(1).Value = 7
 PDial1(2).Value = 185

End Sub

Private Sub PDial1_Changing(Index As Integer, iValue As Integer)
  Label1(Index).Caption = PDial1(Index).Value
End Sub
