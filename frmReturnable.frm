VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmReturnable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returnable Issue / Received for Metallic Cops and Pallet"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9135
   Begin VB.PictureBox picToolTip 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   29
      Top             =   7800
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Timer tmrTool 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1920
         Top             =   0
      End
      Begin VB.Label lblToolTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   120
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   7035
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12409
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   8438015
      BackColor       =   16777215
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      Begin VB.TextBox TXTPallets 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   14
         ToolTipText     =   "Enter the Description of Item."
         Top             =   4920
         Width           =   795
      End
      Begin VB.TextBox TXTADDRESS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2760
         Width           =   6315
      End
      Begin VB.TextBox TXTDCOD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2400
         Width           =   6315
      End
      Begin VB.TextBox TXTNOC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   7440
         MaxLength       =   6
         TabIndex        =   17
         ToolTipText     =   "Enter the Description of Item."
         Top             =   5280
         Width           =   1155
      End
      Begin VB.TextBox TXTRMRK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         MaxLength       =   49
         TabIndex        =   18
         ToolTipText     =   "Enter the Description of Item."
         Top             =   5640
         Width           =   6315
      End
      Begin VB.TextBox TXTNOT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "Enter the Description of Item."
         Top             =   5280
         Width           =   795
      End
      Begin VB.TextBox TXTNOB 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   16
         ToolTipText     =   "Enter the Description of Item."
         Top             =   5280
         Width           =   795
      End
      Begin VB.TextBox TXTBRNM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Enter the Description of Item."
         Top             =   2040
         Width           =   6315
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1680
         Width           =   6315
      End
      Begin VB.OptionButton optIssue 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Issue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         Top             =   3360
         Width           =   975
      End
      Begin VB.OptionButton optRecieved 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   3360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&dit"
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
         Image           =   "frmReturnable.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   5880
         TabIndex        =   4
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frmReturnable.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Save"
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
         Image           =   "frmReturnable.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3240
         TabIndex        =   2
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmReturnable.frx":14BE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7200
         TabIndex        =   5
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&xit"
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
         Image           =   "frmReturnable.frx":1910
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   1200
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56688641
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   600
         TabIndex        =   0
         Top             =   6240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Add"
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
         Image           =   "frmReturnable.frx":1D62
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid FLEXPLY 
         Height          =   885
         Left            =   2280
         TabIndex        =   13
         Top             =   3720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1561
         _Version        =   393216
         Cols            =   50
         BackColor       =   -2147483628
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Pallets"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   4920
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del. Address "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   2760
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Cops"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   5280
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "for Metallic Cops and Pallet."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label LBLDIVNM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Division : POY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   120
         Width           =   5895
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         Height          =   420
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "Slip Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblSlip 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Slip No.   :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   5640
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Tops"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   5280
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Bottoms"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   5280
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8895
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   240
         X2              =   8880
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   9000
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Returnable Issue / Received"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   720
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmReturnable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim PCOD As String
Dim COND As String
Dim INDEX As Long
Dim M_BRCD As String
Dim DIVCODE As String
Dim DIVNAME As String
Dim OPER As String
Dim DRCR As String
Public CHALLAN As String

Private Sub cmdAdd_Click()
    Call ClsData(Me)
    Call btn_sts(False)
    TXTVBDT.SetFocus
    TXTVBDT = Now
    LBLSLIP = GenVNO("RET", "000001")
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    Call ClsData(Me)
    LBLSLIP = GenVNO("RET", "000001")
    TXTVBDT = Now
    optRecieved.Value = True
    
    Dim i As Long
    For i = 1 To FLEXPLY.Cols - 1
      FLEXPLY.TextMatrix(1, i) = ""
      If FLEXPLY.Cols > 1 Then FLEXPLY.COL = 1
      FLEXPLY.CellBackColor = vbWhite
    Next
       
    cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim ANS As String, SQL As String, TEMPRS As New ADODB.Recordset
   
   Set TEMPRS = New ADODB.Recordset
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT * FROM ACCMST WHERE NAME ='" & TXTNAME & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      PCOD = Trim(TEMPRS!CODE)
   End If
   TEMPRS.Close
    
    If LBLSLIP.Caption = "" Then Exit Sub
    If SAVEFLAG = True Then Exit Sub
    
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.TITLE)
    If ANS = vbYes Then
       
       SQL = "UPDATE PKGSTK SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
       "' AND DVCD='" & DIVCODE & "' AND DBCD='000001' AND VTYP='RET' AND CHLN='" & Trim(LBLSLIP.Caption) & "'"
       
       CN.BeginTrans
       CN.Execute SQL
       CN.CommitTrans
       MsgBox "Data are Successfully Deleted."
       
    End If
                
    Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
    SAVEFLAG = False
    frmReturnableList.DIVCODE = DIVCODE
    frmReturnableList.M_DBCD = "000001"
    CHALLAN = Empty
    frmReturnableList.Show 1
    
    If CHALLAN = Empty Or CHALLAN = "" Then
       btn_sts (True)
       cmdAdd.Enabled = True
       SAVEFLAG = True
       cmdAdd.SetFocus
    Else
       btn_sts (False)
       TXTNAME.Enabled = True
       TXTNAME.SetFocus
    End If
End Sub

Private Sub CMDEXIT_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errPRIMARYKEY
    Dim SQL As String
    Dim BRCD As String, SCONSINEE As String, SADD As String
    
    Dim TEMPRS As New ADODB.Recordset
    Set TEMPRS = New ADODB.Recordset
    
    Dim Ctrl As Control
    
    TXTNAME.Text = Trim(TXTNAME.Text)
    
    For Each Ctrl In Me
        If TypeOf Ctrl Is TextBox Then
            Ctrl = Replace(Ctrl, "'", "", 1)
        End If
    Next
           
    If Trim(TXTNAME.Text) = "" Then
       MsgBox "Please Enter Party Name.", vbInformation, App.TITLE
       TXTNAME = Trim(TXTNAME)
       TXTNAME.SetFocus
       Exit Sub
    End If
    
    If Trim(TXTBRNM.Text) = "" Then
       MsgBox "Please Enter Broker Name.", vbInformation, App.TITLE
       TXTBRNM = Trim(TXTBRNM)
       TXTBRNM.SetFocus
       Exit Sub
    End If
    
    If Val(TXTNOB) = 0 And Val(TXTNOT) = 0 And Val(TXTNOC) = 0 Then
        MsgBox "Please Enter Valid Entry", vbInformation, App.TITLE
        TXTNOB.Enabled = True
        TXTNOB.SetFocus
        Exit Sub
    End If
           
   DRCR = "CR" 'CREDITOR
           
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT CODE,DRCR FROM ACCMST WHERE NAME ='" & TXTNAME & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      PCOD = Trim(TEMPRS!CODE)
      If Trim(TEMPRS!DRCR) = "D" Then DRCR = "DR" 'DEBTOR
   End If
   TEMPRS.Close
   
   If (Trim(txtDCOD) = "" Or Trim(TXTADDRESS) = "") And DRCR = "DR" Then
       MsgBox "Please Enter Consignee Name & Address.", vbInformation, App.TITLE
       txtDCOD = Trim(txtDCOD)
       txtDCOD.SetFocus
       Exit Sub
    End If
   
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT CODE FROM REFMST WHERE NAME ='" & TXTBRNM & "' AND CATA='B'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      BRCD = Trim(TEMPRS!CODE)
   End If
   TEMPRS.Close
   
   'CONSIGNEE
   If txtDCOD <> Empty Then
      SCONSINEE = GetCode("PADDMST", txtDCOD, "NAME", "CODE")
      If TEMPRS.State = 1 Then TEMPRS.Close
      TEMPRS.Open "SELECT SRNO FROM PADDMST WHERE CODE='" & SCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
      If Not TEMPRS.EOF Then
         SADD = TEMPRS!SRNO
      End If
      TEMPRS.Close
   End If
   '-------------------------------------------------------------
      
   If optIssue.Value = True Then
      OPER = "-"
   Else
      OPER = "+"
   End If
   
   On Error GoTo errPRIMARYKEY
   
   If SAVEFLAG = True Then
      LBLSLIP = GenVNO("RET", "000001")
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
       "' AND DVCD='" & DIVCODE & "' AND DBCD='000001' AND VTYP='RET' AND CHLN='" & Trim(LBLSLIP.Caption) & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
         MsgBox "Slip No. is Already Exist.", vbCritical
         Exit Sub
      End If
      RS.Close
   End If
         
      CN.BeginTrans
      
   If SAVEFLAG = True Then
      SQL = "INSERT INTO PKGSTK(COMP,UNIT,DVCD,DBCD,VTYP,CHLN,DATE,PCOD,BRCD,DCOD,ADDRESS,OPER,PALLETS,TOPPLY,BOTTOMPLY,QNTY,BRMK,RECSTAT) " & _
            "VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & "','000001', 'RET','" & LBLSLIP & _
            "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & PCOD & "','" & BRCD & "','" & SCONSINEE & _
            "','" & SADD & "','" & OPER & "','" & Val(TXTPallets) & "','" & Val(TXTNOT) & _
            "','" & Val(TXTNOB) & "','" & Val(TXTNOC) & "','" & TXTRMRK & "','A')"

      CN.Execute SQL
                  
      SQL = "UPDATE SERIALMASTER SET [SRNO]='" & LBLSLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
            "' AND VTYP='RET' AND CODE='000001' AND FYCD='" & FYCD & "'"
   
      CN.Execute SQL
      '-----------
      
     'SQL = "INSERT INTO STORETRAN(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,DATE,PCOD,ICOD,QNTY,[USER],[SYSR],OPER,COPS,RECSTAT)" & _
     '      "VALUES('" & compPth & "','" & UNCD & "','000001','000001', 'RET','" & LBLSLIP & _
     '      "','" & Format(txtvbdt, "MM/DD/YYYY") & "','" & PCOD & "','XXXXXXXXXX','" & Val(TXTNOC) & _
     '      "','" & cUName & "','N','+','" & Val(TXTNOC) & "','A')"
      
     'CN.Execute SQL
      
     CN.CommitTrans
     
   Else
   
    SQL = "UPDATE PKGSTK SET PALLETS='" & Val(TXTPallets) & "',TOPPLY='" & Val(TXTNOT) & "',BOTTOMPLY='" & Val(TXTNOB) & "',DCOD='" & SCONSINEE & _
    "',ADDRESS='" & SADD & "',PCOD='" & PCOD & "',BRCD='" & BRCD & "',OPER='" & OPER & _
    "',QNTY='" & Val(TXTNOC) & "',DATE ='" & Format(TXTVBDT, "MM/DD/YYYY") & "',BRMK='" & TXTRMRK & _
    "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    "' AND DVCD='" & DIVCODE & "' AND DBCD='000001' AND VTYP='RET' AND RECSTAT='A' AND CHLN='" & Trim(LBLSLIP) & "'"
    
    CN.Execute SQL
    
    'SQL = "UPDATE STORETRAN SET QNTY='" & Val(TXTNOC) & "',COPS='" & Val(TXTNOC) & _
    '      "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    '      "' AND DVCD='000001' AND DBCD='000001' AND VTYP='RET' AND RECSTAT='A' AND VBNO='" & Trim(LBLSLIP) & "'"
    '
    'CN.Execute SQL
    
    CN.CommitTrans
    
   End If
   
'PLY UPDATION COMMON FOR BOTH SAVE AND EDIT
Dim RSTMP As ADODB.Recordset: Set RSTMP = New ADODB.Recordset
Dim NOOFPLY As Double, i As Long, J As Long

If RSTMP.State = 1 Then RSTMP.Close
RSTMP.Open "SELECT * FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND DBCD='000001' AND VTYP='RET' AND CHLN='" & Trim(LBLSLIP) & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not RSTMP.EOF Then

i = 0
  For i = 1 To FLEXPLY.Cols - 1
    J = 0
    For J = 0 To RSTMP.Fields.COUNT - 1
      If Trim(RSTMP.Fields(J).NAME) = Trim(FLEXPLY.TextMatrix(0, i)) Then
         RSTMP.Fields(J).Value = Val(FLEXPLY.TextMatrix(1, i))
      End If
    Next
  Next
RSTMP.Update
End If

If RSTMP.State = 1 Then RSTMP.Close
'-------------------------------------------------
   If SAVEFLAG = True Then
    '  Call DAILYSTATUS("RET", PCOD, "000001", Val(TXTNOC), LBLSLIP.Caption, 0, cUName, "S", Now)
      MsgBox "Your Slip No. is: " & LBLSLIP
   Else
     ' Call DAILYSTATUS("RET", PCOD, "000001", Val(TXTNOC), LBLSLIP.Caption, 0, cUName, "U", Now)
      MsgBox "Slip No.: " & LBLSLIP & " Successfully Updated."
   End If
        
    Call btn_sts(True)
    Call ClsData(Me)
    Call cmdCancel_Click
    TXTVBDT = Now
    LBLSLIP = GenVNO("RET", "000001")
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
Resume
    CN.RollbackTrans
    If ERR.Number = -2147217873 Or -2147217900 Then
        TXTNAME.SetFocus
        MsgBox "This Name Already Registered With Other Category!!!", vbInformation, "Already Registered"
    Else
        ErrNumber = ERR.Number
        ErrMessage = ERR.Description
        frm_ErrorHandler.Show vbModal
    End If
    ERR.Clear
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or DIVNAME = Empty Then
  Unload Me
End If

    Call ColorComponent(Me)
    Me.BackColor = RGB(RED, GREEN, BLUE)
    If key_PressNew Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = "FLEXPLY" Then Exit Sub
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad

M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
DIVCODE = Empty: DIVNAME = Empty
  
If DIVCODE = Empty Then
   DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
   DIVCODE = Key
End If
  LBLDIVNM = "Division Name : " & UCase(DIVNAME)
  
  Call btn_sts(True)
  Call setHeading
  TXTVBDT = Now
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  cmdExit.Cancel = True
  Me.Show
  Exit Sub
  
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    TXTNAME.Enabled = Not bool
    TXTNOB.Enabled = Not bool
    
    
    TXTNOT.Enabled = Not bool
    TXTNAME.Enabled = Not bool
    
    TXTRMRK.Enabled = Not bool
End Sub

Private Sub TXTADDRESS_GotFocus()
  TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   If DRCR = "D" Then
      If txtDCOD = Empty And txtDCOD.Enabled Then txtDCOD.SetFocus: Exit Sub
   End If
   
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtDCOD & "'", 0, Empty, "Select Consignee Address")
   ElseIf TXTADDRESS = Empty Then
     If DRCR = "D" Then
        M_DESC = Empty:   NEW_VISIBLE = False
        TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtDCOD & "'", 0, Empty, "Select Consignee Address")
     End If
   End If
   
End Sub

Private Sub TXTADDRESS_LostFocus()
  TXTADDRESS.BackColor = vbWhite
End Sub

Private Sub TXTBRNM_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      TXTBRNM = Empty
    ElseIf KeyCode = vbKeyF2 Or (Trim(TXTBRNM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTBRNM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, TXTBRNM.Text, "SELECT AGENT FROM LIST")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "B"
            TXTBRNM.Text = ""
            Frm_Ref_FAS.Show
        Else
            M_BRCD = Key
        End If
    End If
  Me.KeyPreview = True
End Sub

Private Sub TXTDCOD_GotFocus()
 txtDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consinee Name ")
  ElseIf txtDCOD = Empty Then
     If DRCR = "D" Then
        M_DESC = Empty:   NEW_VISIBLE = False
        txtDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consinee Name ")
     End If
  End If
 Me.KeyPreview = True
End Sub

Private Sub TXTDCOD_LostFocus()
  txtDCOD.BackColor = vbWhite
End Sub

Private Sub txtName_GotFocus()
  TXTNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
  
  If TXTNAME = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For A/C Party Master Help", "", TXTNAME.Left, TXTNAME.Top + TXTNAME.Height + 100
  Else
      ToolTip Me, "Press {F2} For A/C Party Master Help", "", TXTNAME.Left, TXTNAME.Top + TXTNAME.Height + 100
  End If
End Sub

Private Sub TXTNAME_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      TXTNAME = Empty
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And TXTNAME = Empty) Then
     M_DESC = Empty:   NEW_VISIBLE = True
     TXTNAME = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
     If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTNAME.Text = ""
            frm_Acc.Show
     Else
            TXTNAME.Tag = Key
     End If
    
  End If
  
  If TXTNAME <> Empty Then
     DRCR = GetCode("ACCMST", TXTNAME, "NAME", "DRCR")
  End If
    
  Me.KeyPreview = True
End Sub

Private Sub txtName_LostFocus()
  TXTNAME.BackColor = vbWhite
  picToolTip.Visible = False
End Sub

Private Sub TXTNOC_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTNOC, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTNOB_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTNOB, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTNOT_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTNOT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTNOC_LostFocus()
TXTNOC.BackColor = vbWhite
End Sub

Private Sub TXTNOC_GotFocus()
 TXTNOC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTNOB_LostFocus()
TXTNOB.BackColor = vbWhite
End Sub

Private Sub TXTNOB_GotFocus()
 TXTNOB.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtNOT_LostFocus()
TXTNOT.BackColor = vbWhite
End Sub

Private Sub txtNOT_GotFocus()
 TXTNOT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTBRNM_GotFocus()
 TXTBRNM.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
  
  If TXTBRNM = Empty Then
      ToolTip Me, "Press {F2} / {Enter} For Broker Master Help", "", TXTBRNM.Left, TXTBRNM.Top + TXTBRNM.Height + 100
  Else
      ToolTip Me, "Press {F2} For Broker Master Help", "", TXTBRNM.Left, TXTBRNM.Top + TXTBRNM.Height + 100
  End If
End Sub

Private Sub TXTBRNM_LostFocus()
TXTBRNM.BackColor = vbWhite
  picToolTip.Visible = False
End Sub

Private Sub TXTRMRK_GotFocus()
 TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRMRK_LostFocus()
TXTRMRK.BackColor = vbWhite
End Sub

Private Sub FLEXPLY_EnterCell()
  If FLEXPLY.ROW = 0 Then Exit Sub
  FLEXPLY.CellBackColor = RGB(BRED, BGREEN, BBLUE)
  FLEXPLY.ColWidth(FLEXPLY.COL - 1) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL - 1))
  FLEXPLY.ColWidth(FLEXPLY.COL) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL))
  If FLEXPLY.COL + 1 < FLEXPLY.Cols Then
    FLEXPLY.ColWidth(FLEXPLY.COL + 1) = 155 * Len(FLEXPLY.TextMatrix(FLEXPLY.ROW - 1, FLEXPLY.COL + 1))
  End If
End Sub

Private Sub FLEXPLY_KeyPress(KeyAscii As Integer)
  If FLEXPLY.COL = 0 And (FLEXPLY.ROW = 1 Or FLEXPLY.ROW = 2) Then Exit Sub
  
  On Error GoTo LAST
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  
  Set MSTDAT = New ADODB.Recordset
  
  FWD_COL = False
  ALLOW_KEY = False
  
  If FLEXPLY.ROW > 0 Then
    If InStr(1, FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL), ".") > 0 And KeyAscii = 46 Then
        KeyAscii = 0
        Exit Sub
    End If
  End If
  
  If FLEXPLY.ROW > 0 Then    'START ------------>>>
  
  'CASE1
  If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
  Else
      ALLOW_KEY = False
  End If
  
  'CASE2
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  
  'CASE3
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL))
    If lnth > 0 Then
      FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) = Mid(FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL), 1, lnth - 1)
      Exit Sub
    End If
  End If
  
  'CASE4
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
       'If FLEXPLY.COL > 2 Then FLEXPLY.ColWidth(FLEXPLY.COL - 2) = 0
    Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  'ACTION 1
  If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
      FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) = FLEXPLY.TextMatrix(FLEXPLY.ROW, FLEXPLY.COL) + Chr(KeyAscii)
    End If
  End If
  
  'ACTION 2
  If ENTER_PRESS = True And FLEXPLY.COL + 1 < FLEXPLY.Cols Then
     FLEXPLY.COL = FLEXPLY.COL + 1
  ElseIf ENTER_PRESS = True Then
     FLEXPLY.CellBackColor = vbWhite
     SendKeys "{TAB}"
  End If
  
  End If   'END<<-----------------------
      
  Exit Sub
LAST:
  MsgBox ERR.Description
  FLEXPLY.SetFocus
  Exit Sub
End Sub

Private Sub FLEXPLY_LeaveCell()
  FLEXPLY.CellBackColor = vbWhite
End Sub


Private Sub setHeading()
Dim i As Long, J As Long
With FLEXPLY
    .TextMatrix(0, 0) = "PlyName"
    .TextMatrix(1, 0) = "No.ofPly"
    .ColWidth(0) = 850
    .ColWidth(0) = 850
End With

Dim COUNT As Long
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If GETRS.EOF Then FLEXPLY.Enabled = False
Do While Not GETRS.EOF
    COUNT = COUNT + 1
    FLEXPLY.TextMatrix(0, COUNT) = Trim(GETRS!NAME & "")
    FLEXPLY.ColWidth(COUNT) = 155 * Len(Trim(GETRS!NAME & ""))
GETRS.MoveNext
Loop
GETRS.Close

FLEXPLY.Cols = COUNT + 1

End Sub


Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{TAB}"
End If
End Sub
