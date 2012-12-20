VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmGRPacking 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goods Return Entry / Sale Return "
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9615
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   4755
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8387
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632319
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
      Begin VB.TextBox txtRMK 
         Height          =   285
         Left            =   1080
         MaxLength       =   249
         TabIndex        =   13
         Top             =   3360
         Width           =   8295
      End
      Begin VB.TextBox TXTTWIST 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "S"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtSubGRD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox TXTBOXES 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox TXTLTNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TXTDENI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox TXTGRAD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox TXTCOP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TXTTRWT 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   7320
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "0"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox TXTGRWT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   9
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox TXTNTWT 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "0"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox TXTPCOD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   285
         Left            =   7320
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   24313857
         CurrentDate     =   39347
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   4920
         TabIndex        =   17
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":0000
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6120
         TabIndex        =   18
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":039A
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2520
         TabIndex        =   15
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":0934
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3720
         TabIndex        =   16
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":0ECE
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   7320
         TabIndex        =   20
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":1320
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   3960
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
         Image           =   "frmGRPacking.frx":18BA
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   240
         TabIndex        =   37
         Top             =   3360
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   120
         X2              =   9480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label LBLSZO 
         BackStyle       =   0  'Transparent
         Caption         =   "{S/Z/0}"
         Enabled         =   0   'False
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
         Left            =   7920
         TabIndex        =   36
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Grade"
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
         Left            =   6120
         TabIndex        =   35
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cops"
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
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GR No."
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
         Left            =   6120
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   5  'Not Copy Pen
         X1              =   120
         X2              =   9480
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3600
         TabIndex        =   32
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name :"
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
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   9375
      End
      Begin VB.Label LBLDESC1 
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
         Left            =   1920
         TabIndex        =   30
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name  "
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
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Da&te "
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
         Left            =   6120
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LBLCFG 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
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
         Left            =   6120
         TabIndex        =   27
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label LBLLOT 
         BackStyle       =   0  'Transparent
         Caption         =   "LotNo"
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
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label LBLNOCOPS 
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Wgt."
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
         Left            =   6120
         TabIndex        =   24
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Wgt. "
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
         Left            =   6120
         TabIndex        =   23
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Tare Wgt."
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
         Left            =   6120
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         BorderStyle     =   5  'Dash-Dot-Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H00000080&
         Height          =   2715
         Left            =   120
         Top             =   1920
         Width           =   9375
      End
      Begin VB.Label LBLPCOD 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name "
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
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmGRPacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsShadeReq As Boolean
Public TWSTREQ As String
Dim ERROROCCUR As Boolean
Public DIVCODE As String
Public LSPKGCOD As String
Dim M_DBCD As String
Dim PKGNGCD As String
Public GRADE As String
Public CHALLAN As String
Dim SUBPKG As String
Dim SUBPKGCODE As String
Dim INFORS As New ADODB.Recordset
Dim COMPORTX As Integer
Dim SAVEFLAG As Boolean
Dim SQL As String
Dim M_PCOD As String
Dim FINITMCOD As String
Dim LOCCOD  As String
Dim DIVCFG As String

Public Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = bool
    cmdCancel.Enabled = bool
    cmdAdd.Enabled = Not bool
    cmdEdit.Enabled = Not bool
    cmdDelete.Enabled = Not bool
    
    TXTPCOD.Enabled = bool
    TXTLTNO.Enabled = bool
    TXTDENI.Enabled = bool
    TXTVBNO.Enabled = bool
    TXTVBDT.Enabled = bool
    TXTGRAD.Enabled = bool
    txtSubGRD.Enabled = bool
    TXTTWIST.Enabled = bool
    TXTBOXES.Enabled = bool
    TXTCOP.Enabled = bool
    TXTGRWT.Enabled = bool
    TXTTRWT.Enabled = bool
    TXTNTWT.Enabled = bool
    txtRMK.Enabled = bool
End Sub

Private Sub cmdAdd_Click()

Me.Caption = "HELLO"
Me.Caption = "HELLO121332212223"


    Call btn_sts(True)
    TXTPCOD.Enabled = True
    TXTPCOD.SetFocus
    SAVEFLAG = True
    cmdCancel.Cancel = True
    CHALLAN = GenGRPACKINGVNO("FGR")
    TXTVBNO = CHALLAN
End Sub

Private Sub cmdCancel_Click()
   cmdExit.Cancel = True
   Call btn_sts(False)
   SAVEFLAG = True
   Call ClsData(Me)
   cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    SAVEFLAG = False
    Call btn_sts(True)
    frmGRPackingList.Show 1
    
    If TXTVBNO = Empty Or TXTVBNO = "" Then
        Call cmdCancel_Click
        Exit Sub
    End If
    
    Dim ANS
    
    ANS = MsgBox("Do You want to Delete this record?", vbYesNo + vbQuestion, App.TITLE)
    
    If ANS = vbYes Then
        CN.BeginTrans
            CN.Execute "UPDATE GRPACKING SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VBNO='" & TXTVBNO & "'"
        CN.CommitTrans
    End If
    
    Call cmdCancel_Click
    
End Sub

Private Sub cmdEdit_Click()
    SAVEFLAG = False
    Call btn_sts(True)
    frmGRPackingList.Show 1
    
    If TXTVBNO = Empty Or TXTVBNO = "" Then
        Call cmdCancel_Click
        Exit Sub
    End If
    
    TXTPCOD.SetFocus
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

If CheckData Then Exit Sub

    If SAVEFLAG Then
    
       CHALLAN = GenGRPACKINGVNO("FGR")
       TXTVBNO = CHALLAN
        
       If RS.State = 1 Then RS.Close
       RS.Open "SELECT VBNO FROM GRPACKING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND VBNO='" & CHALLAN & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
       If Not RS.EOF Then
          MsgBox "GR Slip No. is Already Exist.", vbCritical
          Exit Sub
       End If
       RS.Close
             
    End If
    
    CN.BeginTrans
    
    M_DBCD = "000004"
    
    CN.Execute "DELETE FROM GRPACKING WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                  "' AND VBNO='" & TXTVBNO & "'"
       
    CN.Execute "INSERT INTO GRPACKING([COMP],[UNIT],[DVCD],[PKG_STCOD],[VBNO],SRCH,[VBDT],[PCOD],[LOTNO]," & _
           "[ICOD],[GRAD],[SUBGRD],[BOXES],[COPS],[GRSWGT],[TRWGT],[NETWGT],[RECSTAT],RMK) VALUES('" & compPth & _
           "','" & UNCD & "','" & DIVCODE & "','000001','" & CHALLAN & _
           "','1','" & Format(TXTVBDT, "MM/DD/YYYY") & _
           "','" & M_PCOD & "','" & TXTLTNO & _
           "','" & FindFinItemCode & "','" & GRADE & _
           "','" & FindSubGradeCode & "','" & Val(TXTBOXES) & "','" & Val(TXTCOP) & "','" & Val(TXTGRWT) & _
           "','" & Val(TXTTRWT) & "','" & Val(TXTNTWT) & "','A','" & Trim(txtRMK) & "') "
           
    If SAVEFLAG Then
       Dim UPSQL As String
       UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & CHALLAN & "' WHERE COMP='" & compPth & _
            "' AND UNIT = '" & UNCD & "' AND VTYP='FGR' AND FYCD='" & FYCD & "' "
            
       CN.Execute UPSQL
    End If
           
    CN.CommitTrans
    
    MsgBox "GR Packing Save Successfully", vbInformation
    Call cmdCancel_Click
                  
Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
End Sub

Private Sub Form_Activate()
  If DIVCODE = Empty Or Trim(LBLDESC1.Caption) = "XXXXXXXXXX" Then
     MsgBox "Select Division For Packing."
     Unload Me
  End If
  
 'For GR Slip
  If CHALLAN = Empty Then
     Unload Me
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

   If UCase(ActiveControl.NAME) = "TXTGRWT" Then
      Exit Sub
   End If
   
   If UCase(ActiveControl.NAME) = "TXTCOP" Then
      Exit Sub
   End If
   
   If UCase(ActiveControl.NAME) = "TXTTRWT" Then
      Exit Sub
   End If
   
   If UCase(ActiveControl.NAME) = "TXTNTWT" Then
      Exit Sub
   End If
   
   If UCase(ActiveControl.NAME) = "TXTBOXES" Then
      Exit Sub
   End If

   If KeyAscii = 13 Then SendKeys "{TAB}"
   
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  SAVEFLAG = True
  
  '-------DIVISION NAME
  M_DESC = Empty: Key = Empty:  NEW_VISIBLE = False:  DIVCODE = Empty
  LBLDESC1.Caption = Empty
  If DIVCODE = Empty Then
    LBLDESC1 = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    If LBLDESC1 <> Empty Then DIVCODE = Key Else LBLDESC1 = "???????": Unload Me
  End If
       
  LBLCFG.Caption = LabelDisplay(DIVCODE & "", UNCD)
  
  Dim TEMPRS As New ADODB.Recordset
  If TEMPRS.State = 1 Then TEMPRS.Close
  TEMPRS.Open "SELECT *FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND RECSTAT='A' AND CODE='" & DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
  If Not TEMPRS.EOF Then
      DIVCFG = Trim(TEMPRS!CFGTYP & "")
  End If
  
  IsShadeReq = False
  
  If IsTwistReq(DIVCODE) = "Y" Then
     TWSTREQ = "Y"
     LBLSZO.Enabled = True: TXTTWIST.Enabled = True
     Label5.Enabled = True
     LBLSZO.Visible = True
     Label5.Caption = "Twist"
     txtSubGRD.Visible = False
  ElseIf SetIsShadeReq(DIVCODE) = "Y" Then
     IsShadeReq = True
     Label5.Caption = "Shade"
     Label5.Enabled = True
     LBLSZO.Visible = False
     TXTTWIST.Enabled = False
     TXTTWIST.Visible = False
     txtSubGRD.Enabled = True
     txtSubGRD.Visible = True
  ElseIf DIVCFG = "GD" Or DIVCFG = "SD" Then
     txtSubGRD.Visible = False
     Label5.Visible = False
     LBLSZO.Visible = False
     TXTTWIST.Enabled = False
     TXTTWIST.Visible = False
  Else
     txtSubGRD.Visible = True
     Label5.Visible = True
     LBLSZO.Visible = False
     TXTTWIST.Enabled = False
     TXTTWIST.Visible = False
  End If
  
  'For Raw Material Consumption Slip
  CHALLAN = GenGRPACKINGVNO("FGR")
  
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  TXTVBDT.Value = Now

  Call btn_sts(False)
  
JUMP:
End Sub

Private Sub txtCop_GotFocus()
  TXTCOP.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCop_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(TXTCOP) > 0 Then
       SendKeys "{TAB}"
       Exit Sub
    End If
  
    If KeyAscii < 48 Or KeyAscii > 57 Then             ' 0- 9
       If KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtBOXES_LostFocus(): TXTBOXES.BackColor = vbWhite: End Sub

Private Sub TXTBOXES_GotFocus()
  TXTBOXES.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTBOXES_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(TXTBOXES) > 0 Then
       SendKeys "{TAB}"
       Exit Sub
    End If
  
    If KeyAscii < 48 Or KeyAscii > 57 Then             ' 0- 9
       If KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtCop_LostFocus(): TXTCOP.BackColor = vbWhite: End Sub

Private Sub txtDENI_GotFocus()
  TXTDENI.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtDENI_LostFocus(): TXTDENI.BackColor = vbWhite: End Sub

Private Sub TXTGRAD_GotFocus()
    TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If Trim(TXTGRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False: Key = Empty
    TXTGRAD.Text = SearchList1("SELECT TOP 20 CODE,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT " & LBLCFG.Caption)
    TXTGRAD.Tag = Key
    GRADE = Key
  End If
End Sub

Private Sub TXTGRAD_LostFocus(): TXTGRAD.BackColor = vbWhite: End Sub

Private Sub TXTGRWT_Change()
   If Val(TXTTRWT) > 0 Then
      TXTNTWT = Val(TXTGRWT) - Val(TXTTRWT)
   ElseIf Val(TXTGRWT) > 0 Then
      TXTNTWT = Val(TXTGRWT)
   End If
End Sub

Private Sub txtGRWT_GotFocus()
  TXTGRWT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtGRWT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Val(TXTGRWT) > 0 Then
     SendKeys "{TAB}"
     Exit Sub
  End If
  If CheckNumericKey(KeyAscii, TXTGRWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtGRWT_LostFocus()
  TXTGRWT.BackColor = vbWhite
End Sub

Private Sub txtRMK_GotFocus()
    txtRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRMK_LostFocus()
    txtRMK.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_GotFocus()
    txtSubGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Trim(txtSubGRD.Text) = Empty And KeyCode = 13) Or KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: Key = Empty
        txtSubGRD.Text = SearchList1("SELECT DISTINCT SUBGRD,NAME FROM SUBGRDMST WHERE COMP='" & _
                    compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND GRAD='" & _
                    GRADE & "'", 0, txtSubGRD, "SELECT SHADE")
        txtSubGRD.Tag = Key
    End If
End Sub

Private Sub TXTSUBGRD_LostFocus()
    txtSubGRD.BackColor = vbWhite
End Sub

Private Sub TXTTRWT_Change()
   If Val(TXTGRWT) > 0 And Val(TXTTRWT) > 0 Then
      TXTNTWT = Val(TXTGRWT) - Val(TXTTRWT)
   End If
End Sub

Private Sub TXTTRWT_GotFocus()
  TXTTRWT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtTRWT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Val(TXTTRWT) > 0 Then
     SendKeys "{TAB}"
     Exit Sub
  End If
  If CheckNumericKey(KeyAscii, TXTTRWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTTRWT_LostFocus()
  TXTTRWT.BackColor = vbWhite
End Sub

Private Sub TXTNTWT_GotFocus()
  TXTNTWT.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtNTWT_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Val(TXTNTWT) > 0 Then
     SendKeys "{TAB}"
     Exit Sub
  End If
  If CheckNumericKey(KeyAscii, TXTNTWT, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTNTWT_LostFocus()
  TXTNTWT.BackColor = vbWhite
End Sub

Private Sub txtLTNO_Change()
    If TXTLTNO <> Empty Then FindFinishItem
End Sub

Private Sub txtltno_GotFocus()
  TXTLTNO.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)

Dim SQL As String: Me.KeyPreview = False
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTLTNO = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTLTNO = Empty) Then
   M_DESC = Empty:   NEW_VISIBLE = False: Key = Empty
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND ACTIVE = 'Y' "
   TXTLTNO = SearchList(SQL)
End If

If TXTLTNO <> Empty Then FindFinishItem

If SAVEFLAG Then
   TXTLTNO.Tag = TXTLTNO
End If

Me.KeyPreview = True
End Sub

Private Sub txtltno_LostFocus(): TXTLTNO.BackColor = vbWhite: End Sub

Private Sub txtPCOD_GotFocus()
  TXTPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or (Trim(TXTPCOD.Text) = Empty And KeyCode = 13) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTPCOD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM ACCMST", 0, TXTPCOD, "List of Party A/c")
   ElseIf KeyCode = vbKeyDelete Then
        TXTPCOD = Empty
    End If
 Me.KeyPreview = True
End Sub

Private Sub txtPCOD_LostFocus(): TXTPCOD.BackColor = vbWhite: End Sub

Private Sub FindFinishItem()
Dim RSITM As ADODB.Recordset: Set RSITM = New ADODB.Recordset
Dim FICD As String

If RSITM.State = 1 Then RSITM.Close
RSITM.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND LTNO='" & TXTLTNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSITM.EOF Then FICD = RSITM!FICD
RSITM.Close

If FICD <> Empty Then
  If RSITM.State = 1 Then RSITM.Close
  RSITM.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RSITM.EOF Then
     TXTDENI = RSITM!NAME
  Else
     TXTDENI = Empty
  End If
  RSITM.Close
End If

End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{TAB}"
   End If
End Sub

Private Function CheckData() As Boolean
 CheckData = False

 Dim DBCDRS As ADODB.Recordset
 Set DBCDRS = New ADODB.Recordset

 If Val(TXTCOP) <= 0 And TXTCOP.Enabled Then MsgBox "Please Enter No.of Cops !!", vbInformation: TXTCOP.SetFocus: CheckData = True: Exit Function
 If Val(TXTGRWT) <= 0 And TXTGRWT.Enabled Then MsgBox "Please Enter Gross Weight !!", vbInformation, "Weight Missing !!": TXTGRWT.SetFocus: CheckData = True: Exit Function
 If Val(TXTNTWT) <= 0 And TXTNTWT.Enabled Then MsgBox "Net Weight is not proper !!", vbInformation, "Weight Missing !!": TXTNTWT.SetFocus: CheckData = True: Exit Function
  
 If TXTDENI = Empty Then MsgBox "Please Select Proper Item !!", vbInformation, "Item Missing !!": TXTLTNO.SetFocus: CheckData = True: Exit Function
 If TXTLTNO = Empty Then MsgBox "Please Select Proper Lot !!", vbInformation, "Lot Missing !!": TXTLTNO.SetFocus: CheckData = True: Exit Function
 If TXTGRAD = Empty Then MsgBox "Please Select Grade !!", vbInformation, "Grade Missing !!": TXTGRAD.SetFocus: CheckData = True: Exit Function
   
    If DBCDRS.State = 1 Then DBCDRS.Close
    DBCDRS.Open "SELECT CODE FROM ACCMST WHERE NAME='" & TXTPCOD & "' ", CN, adOpenKeyset, adLockPessimistic
    If DBCDRS.EOF Then
       CheckData = True
       MsgBox "Party Required in case of JobWork/GR PACKING !!", vbCritical
       TXTPCOD.SetFocus
       Exit Function
    Else
       M_PCOD = DBCDRS!CODE
    End If
        
    If DBCDRS.State = 1 Then DBCDRS.Close
    DBCDRS.Open "SELECT CODE FROM GRDMST WHERE GRAD='" & TXTGRAD.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not DBCDRS.EOF Then
       GRADE = Trim(DBCDRS!CODE & "")
    Else
       CheckData = True
       GRADE = Empty
       MsgBox "Grade Required !!", vbCritical
       TXTGRAD.SetFocus
       Exit Function
    End If
    DBCDRS.Close
    
End Function

Private Function FindFinItemCode() As String
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND NAME ='" & TXTDENI & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   FindFinItemCode = GRRS!CODE
Else
   FindFinItemCode = Empty
End If
GRRS.Close
End Function

Public Function GenGRPACKINGVNO(VTYP As String) As String

GenGRPACKINGVNO = "XXXXXXXXXX"
If Trim(VTYP) = Empty Then Exit Function

Dim NO As Double: NO = 0
Dim SRNO As String, STFY As String, ENFY As String
Dim GENRS As ADODB.Recordset
Set GENRS = New ADODB.Recordset

Dim NSQL As String

NSQL = "SELECT LEFT(SRNO,6) AS SRNO FROM SERIALMASTER WHERE COMP='" & compPth & _
        "' AND UNIT = '" & UNCD & "' AND VTYP='" & VTYP & "' AND FYCD='" & FYCD & "'"
   
START:
If GENRS.State = 1 Then GENRS.Close
GENRS.Open NSQL, CN, adOpenDynamic, adLockOptimistic

If GENRS.EOF Then
       FYCD = Mid(Year(FSDT), 3, 2) + Mid(Year(FEDT), 3, 2)
       STFY = Format(FSDT, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
       ENFY = Format(FEDT, "YYYY/MM/DD")      'Mid(Year(FSDT), 1, 4)
       SRNO = "000000" & FYCD

   CN.Execute "INSERT INTO SERIALMASTER(COMP,UNIT,DVCD,VTYP,CODE,NAME,SRNO,FYCD,STFY,ENFY) " & _
              "VALUES('" & compPth & "','" & UNCD & "','','" & VTYP & _
              "','','GR PACKING','" & SRNO & "','" & FYCD & "','" & STFY & "','" & ENFY & "')"
         
   GoTo START
End If

If Not GENRS.EOF Then
   NO = Val(GENRS!SRNO)
   NO = NO + 1
End If

GENRS.Close
   
   If NO < 10 Then
     GenGRPACKINGVNO = "00000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100 Then
     GenGRPACKINGVNO = "0000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 1000 Then
     GenGRPACKINGVNO = "000" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 10000 Then
     GenGRPACKINGVNO = "00" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 100000 Then
     GenGRPACKINGVNO = "0" + Trim(nstr(NO, 1, 0))
   ElseIf NO < 1000000 Then
     GenGRPACKINGVNO = Trim(nstr(NO, 1, 0))
   End If
      
   GenGRPACKINGVNO = GenGRPACKINGVNO & FYCD
   
End Function

Private Function FindSubGradeCode() As String
'SubGradename = ""

Dim LOTRS As ADODB.Recordset
Set LOTRS = New ADODB.Recordset
Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset
Dim COPSWGT As Double

'If IsShadeReq Then
   If GRRS.State = 1 Then GRRS.Close
   GRRS.Open "SELECT SUBGRD FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND NAME='" & txtSubGRD & "' AND GRAD=" & GRADE & " AND DVCD='" & _
             DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
   If Not GRRS.EOF Then
      FindSubGradeCode = Trim(GRRS!SUBGRD & "")
      Exit Function
   End If
   GRRS.Close
'End If

If TWSTREQ = "Y" Then
   FindSubGradeCode = Trim(TXTTWIST)
   Exit Function
End If

End Function


