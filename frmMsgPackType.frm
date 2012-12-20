VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmMsgPackType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Type Selection"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4995
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton OPTLUMPSUM 
         Caption         =   "LUMPSUM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton OPTCARTON 
         Caption         =   "CARTOON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   0
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin WelchButton.lvButtons_H cmdOk 
         Default         =   -1  'True
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Caption         =   "&O.k"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    Select Packing Type     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmMsgPackType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If OPTCARTON.Value = True Then
   Me.Tag = "C"
Else
   Me.Tag = "L"
End If
Me.Visible = False
End Sub


Private Sub Form_Load()
  Call ColorComponent(Me)
End Sub
