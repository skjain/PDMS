VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmWastageDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wastage Dispatch"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   11355
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   4995
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8811
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
      Begin VB.ComboBox cmbDispatchType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmWastageDispatch.frx":0000
         Left            =   2160
         List            =   "frmWastageDispatch.frx":0002
         TabIndex        =   4
         Tag             =   "0"
         Text            =   "Select Type of Dispatch"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TXTDVNM 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   525
         Left            =   7560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txtConsinee 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox BRMK 
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3600
         Width           =   10935
      End
      Begin VB.TextBox TXTRATE 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7800
         MaxLength       =   200
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtDCOD 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox TXTITM 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox TXTAMNT 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox TXTSTKQTY 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   9360
         TabIndex        =   5
         Top             =   960
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
         Format          =   24313857
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1440
         TabIndex        =   1
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":1128
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6360
         TabIndex        =   3
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":157A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   7560
         TabIndex        =   30
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Save/&Print"
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
         Image           =   "frmWastageDispatch.frx":19CC
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3840
         TabIndex        =   34
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":1F66
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   5040
         TabIndex        =   35
         Top             =   4320
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
         Image           =   "frmWastageDispatch.frx":2300
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Dispatch :"
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
         TabIndex        =   33
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11400
         TabIndex        =   32
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dispatch Entry of Wastage Can't Be Edit."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   11400
         TabIndex        =   31
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   9240
         X2              =   9240
         Y1              =   2400
         Y2              =   3240
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "Wastage Challan Date :"
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
         Left            =   6960
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   11055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name"
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
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblBill 
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
         Left            =   9120
         TabIndex        =   27
         Top             =   480
         Width           =   1575
      End
      Begin VB.Shape BORDER 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Wastage Challan No.   :"
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
         Left            =   6960
         TabIndex        =   26
         Top             =   480
         Width           =   2055
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7440
         X2              =   7440
         Y1              =   1440
         Y2              =   3240
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Name"
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
         Height          =   255
         Left            =   4680
         TabIndex        =   25
         Tag             =   "S"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2655
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1440
         Width           =   11175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Address"
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
         Height          =   255
         Left            =   7920
         TabIndex        =   24
         Tag             =   "S"
         Top             =   1515
         Width           =   2175
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Party"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Tag             =   "S"
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   1440
         Y2              =   3240
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Height          =   255
         Left            =   8040
         TabIndex        =   22
         Tag             =   "S"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qnty."
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
         Height          =   255
         Left            =   6000
         TabIndex        =   21
         Tag             =   "S"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Tag             =   "S"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   11280
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks (If Any) :"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Tag             =   "S"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   255
         Left            =   9720
         TabIndex        =   18
         Tag             =   "S"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00000080&
         X1              =   5640
         X2              =   5640
         Y1              =   2400
         Y2              =   3240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Quantity"
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
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Tag             =   "S"
         Top             =   2520
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmWastageDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOAD As String
Dim DIVCODE As String
Dim DIVNAME As String
Dim SQL As String
Dim SAVEFLAG As Boolean
Public M_DBCD As String
Dim SPARTY As String
Dim SCONSINEE As String
Dim SITEM As String
Dim SADD As String
Dim BOX_PKG_REQ As String
Dim ALLOWEDITDEL As Boolean
Public CHALLAN As String

Private Sub cmbDispatchType_Click()
  Call SetDBCD
  lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End Sub

Private Sub cmdCancel_Click()
  TXTDVNM.Tag = TXTDVNM
  Dim DPFLISTINDEX As Long
  DPFLISTINDEX = cmbDispatchType.ListIndex
    
  Call ClsData(Me)
  txtQty.Tag = 0
  TXTDVNM = TXTDVNM.Tag
  SAVEFLAG = True
    Call btn_sts(True)
    If zoomflag = True Then
       Call CMDEXIT_Click
       Exit Sub
    End If
    
  If cmbDispatchType.ListCount > 1 Then cmbDispatchType.ListIndex = DPFLISTINDEX
  Call SetDBCD
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  
  CHALLAN = Empty
    
  frmWastageDispatchList.DIVCODE = DIVCODE
  frmWastageDispatchList.DIVNAME = DIVNAME
  frmWastageDispatchList.VTCD = M_DBCD
  
  frmWastageDispatchList.Show 1
     
  If IsSaleExist Then
     MsgBox "Sale Bill Exist Against this Challan."
     txtQty = 0: TXTRATE = 0: TXTAMNT = 0: TXTSTKQTY = 0
     Call cmdCancel_Click
     Exit Sub
  End If
  
  If CHALLAN <> Empty Then
     Dim AYS
     AYS = MsgBox("Are you sure to delete this Challan ", vbYesNo)
     If AYS = vbYes Then
        CN.BeginTrans
        
        Call SetFIFOReverseForJobworkStock
        
        CN.Execute "DELETE FROM SPTRAN WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND VTYP = 'DPF' AND DBCD='" & M_DBCD & _
           "' AND VBNO='" & lblBill & "' AND RECSTAT<>'D'"
       
        Call DAILYSTATUS("WST", GetCode("ACCMST", txtCONSINEE, "NAME", "CODE"), M_DBCD, Val(txtQty), lblBill, Val(TXTAMNT), cUName, "D", Now, TXTVBDT)
        CN.CommitTrans
      End If
      
      MsgBox "Your Challan No. " + lblBill + " is Successfully Deleted."
      
  End If
  
 Call cmdCancel_Click
 lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
     If ReadConfigMaster("0017", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  
  SAVEFLAG = False
  
  CHALLAN = Empty
    
  frmWastageDispatchList.DIVCODE = DIVCODE
  frmWastageDispatchList.DIVNAME = DIVNAME
  frmWastageDispatchList.VTCD = M_DBCD
  
  frmWastageDispatchList.Show 1
     
  If IsSaleExist Then
     MsgBox "Sale Bill Exist Against this Challan."
     txtQty = 0: TXTRATE = 0: TXTAMNT = 0: TXTSTKQTY = 0
     Call cmdCancel_Click
     Exit Sub
  End If
      
        
  If CHALLAN <> Empty Then
     lblBill = CHALLAN
     btn_sts (False)
     cmbDispatchType.Enabled = False
     If txtCONSINEE.Enabled Then txtCONSINEE.SetFocus
     TXTVBDT.Enabled = False
  Else
     btn_sts (True)
     cmdAdd.SetFocus
  End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim INDEX As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double

If INVALIDDATA Then Exit Sub

Call SetInternal

If SAVEFLAG Then
   lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End If

If Val(txtQty) > TXTSTKQTY Then
   MsgBox "Challan Quantity Exceed From Stock Quantity."
   txtQty.Enabled = True: txtQty.SetFocus: Exit Sub
End If

If SAVEFLAG Then
   Dim NSQL As String
   Dim MSGS As String: MSGS = "Unit"
   lblBill = GenDPFVNO("DPF", M_DBCD, DIVCODE)
   
   NSQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & _
           "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & lblBill & "' "
   
   If UNT_DIVSERIES_REQ = "Y" Then
      NSQL = NSQL & " AND DVCD='" & DIVCODE & "' "
      MSGS = "Division"
   End If
   
   If RS.State Then RS.Close
   RS.Open NSQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      MsgBox "Wastage Challan No. " & SLIP & " Already Exist. Check Last No. In " & MSGS & " Configuration", vbCritical
      Exit Sub
   End If
   RS.Close
End If

CN.BeginTrans

If SAVEFLAG = True Then
   lblBill = GenDPFVNO("DPF", M_DBCD, DIVCODE)
End If

If BOX_PKG_REQ = "Y" Then
   Call SetFIFOReverseForJobworkStock
   Call SETBOXREGISTER
Else
   Call SETPKGMAN
End If

Dim AMT As Double
AMT = 0
If RoundOffReq = True Then
  AMT = Round(txtQty * Val(TXTRATE), 0)
 Else
  AMT = txtQty * Val(TXTRATE)
End If

Dim NEWSQL As String
NEWSQL = "DELETE FROM SPTRAN WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND VTYP = 'DPF' AND DBCD='" & M_DBCD & _
           "' AND VBNO='" & lblBill & "' AND RECSTAT<>'D' "
           
If UNT_DIVSERIES_REQ = "Y" Then
   NEWSQL = NEWSQL & " AND DVCD='" & DIVCODE & "' "
End If

CN.Execute NEWSQL

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "DCOD,ADDRESS,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA4)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','DPF','" & M_DBCD & "','" & lblBill & "','" & lblBill & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & "','" & SCONSINEE & _
"','" & SADD & "','WASTE','" & SITEM & "','0','','0','" & txtQty & "'," & TXTRATE & "," & AMT & _
",'Q','N','" & cUName & "','-','A','0','" & Trim(BRMK) & "')"

CN.Execute SQL

If SAVEFLAG Then
    Dim UPSQL As String
    UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & Trim(lblBill) & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
             "' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "' "
    
    If UNT_DIVSERIES_REQ = "Y" Then
       UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
    End If
     
    CN.Execute UPSQL
End If
'-------------------------
'DAILYSTATUS ENTRY
 If SAVEFLAG = True Then
  Call DAILYSTATUS("DPF", GetCode("ACCMST", txtCONSINEE, "NAME", "CODE"), M_DBCD, Val(txtQty), lblBill, Val(TXTAMNT), cUName, "N", Now, TXTVBDT)
 Else
  Call DAILYSTATUS("DPF", GetCode("ACCMST", txtCONSINEE, "NAME", "CODE"), M_DBCD, Val(txtQty), lblBill, Val(TXTAMNT), cUName, "M", Now, TXTVBDT)
 End If
'------------------------
CN.CommitTrans
lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
If SAVEFLAG Then
  MsgBox "Your Challan No. is : " & lblBill
Else
   MsgBox "Challan No. : " & SLIP & " is Successfully Edited."
End If

Call cmdCancel_Click

Exit Sub
LAST:
MsgBox ERR.Description
Resume
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
Call cmdSave_Click
End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or DIVNAME = Empty Then
   Unload Me
   Exit Sub
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call btn_sts(True)
   
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty: DIVNAME = Empty
  
  If DIVCODE = Empty Then
    DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
  TXTDVNM = DIVNAME
  TXTVBDT = Date
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  
  
  Call SetDispatchType
  
  'M_DBCD = "000005"
  
    If zoomflag = True Then
        btn_sts (False)
        SAVEFLAG = False
    Else
        btn_sts (True)
    End If
 BOX_PKG_REQ = "Y"
 SAVEFLAG = True
 Call SetInternal
 
JUMP:
End Sub

Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
    If cmbDispatchType.Enabled Then cmbDispatchType.SetFocus
    SAVEFLAG = True
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtDCOD & "'", 0, Empty, "Select A/c Party Filtered by Party Group")
   End If
End Sub

Private Sub txtCONSINEE_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtCONSINEE = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtCONSINEE = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
  End If
      
  Me.KeyPreview = True
  
End Sub

Private Sub txtConsinee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCONSINEE = Empty Then
        Call txtCONSINEE_KeyDown(vbKeyF2, 0)
    End If
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtDCOD = Empty
  ElseIf KeyCode = vbKeyF2 Or txtDCOD = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtDCOD = SearchList1("Select DISTINCT CODE,NAME From PADDMST", 0, Empty, "Select Consinee Name ")
  End If
  
 Me.KeyPreview = True
End Sub

Private Sub TXTITM_Change()

 If BOX_PKG_REQ = "Y" Then
 Call FindStock_BOXREGISTER
Else
 Call FindStock_PKGMAN
End If

End Sub

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or (Trim(txtitm) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtitm.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, txtitm.Text, "SELECT FINISH ITEM FROM LIST")
    End If
    Me.KeyPreview = True
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, txtQty, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTRATE, Me) = 0 Then KeyAscii = 0
 If RoundOffReq = True Then
   TXTAMNT = Round(Val(txtQty) * Val(TXTRATE), 0)
  Else
   TXTAMNT = Val(txtQty) * Val(TXTRATE)
 End If
 TXTAMNT = nstr(TXTAMNT, 12, 2)
End Sub

Private Sub TXTAMNT_LostFocus()
TXTAMNT.BackColor = vbWhite
End Sub

Private Sub TXTQTY_LostFocus()
txtQty.BackColor = vbWhite
End Sub

Private Sub txtRate_LostFocus()
TXTRATE.BackColor = vbWhite
End Sub

Private Sub BRMK_GotFocus()
BRMK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_GotFocus()
   TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTADDRESS_LostFocus()
   TXTADDRESS.BackColor = vbWhite
End Sub

Private Sub txtAmnt_GotFocus()
   TXTAMNT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtConsinee_GotFocus()
  txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtConsinee_LostFocus()
 txtCONSINEE.BackColor = vbWhite
End Sub

Private Sub TXTDCOD_GotFocus()
 txtDCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTDCOD_LostFocus()
txtDCOD.BackColor = vbWhite
End Sub

Private Sub TXTITM_GotFocus()
  txtitm.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTITM_LostFocus()
txtitm.BackColor = vbWhite
End Sub

Private Sub TXTQTY_GotFocus()
 txtQty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtRate_GotFocus()
  TXTRATE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub BRMK_LostFocus()
  BRMK.BackColor = vbWhite
End Sub

Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    txtCONSINEE.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    txtDCOD.Enabled = Not Yes
    TXTADDRESS.Enabled = Not Yes
    TXTSTKQTY.Enabled = Not Yes
    txtitm.Enabled = Not Yes
    txtQty.Enabled = Not Yes
    TXTRATE.Enabled = Not Yes
    TXTAMNT.Enabled = Not Yes
    BRMK.Enabled = Not Yes
End Sub

Private Sub TimerBillNo_Timer()
    Static ctr As Integer
    
    If ctr Mod 45 = 0 And ctr <= 45 Then
        lblAlert.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        BORDER.BorderColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
        lblBill.ForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE)
    ElseIf ctr Mod 75 = 0 And ctr <= 75 Then
        lblAlert.ForeColor = vbRed
        BORDER.BorderColor = vbRed
        lblBill.ForeColor = vbRed
    ElseIf ctr Mod 105 = 0 And ctr <= 105 Then
        lblAlert.ForeColor = vbBlue
        BORDER.BorderColor = vbBlue
        lblBill.ForeColor = vbBlue
        ctr = 0
    End If
    
    ctr = ctr + 15
End Sub

Private Sub SetInternal()

Dim GRRS As ADODB.Recordset
Set GRRS = New ADODB.Recordset

Call SetDBCD

If txtitm = Empty Then Exit Sub

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & _
"' AND DVCD = '" & DIVCODE & "' AND NAME = '" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
   SITEM = Trim(GRRS!CODE & "")
End If
GRRS.Close

If txtCONSINEE = Empty Or txtDCOD = Empty Or TXTADDRESS = Empty Then Exit Sub

SPARTY = GetCode("ACCMST", txtCONSINEE, "NAME", "CODE")
SCONSINEE = GetCode("PADDMST", txtDCOD, "NAME", "CODE")

If GRRS.State = 1 Then GRRS.Close
GRRS.Open "SELECT SRNO FROM PADDMST WHERE CODE='" & SCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
If Not GRRS.EOF Then
  SADD = GRRS!SRNO
Else
  SADD = Empty
End If
GRRS.Close

End Sub

Private Sub UPDATEDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "DPF"
  DLYSTA!PCOD = txtCONSINEE
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = Val(txtQty)
  DLYSTA!VBNO = lblBill & ""
  DLYSTA!AMNT = Val(TXTAMNT)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub FindStock_PKGMAN()

If txtitm = Empty Then Exit Sub

Call SetInternal

Dim PACKEDQTY As Double: PACKEDQTY = 0
Dim DISPATCHEDQTY As Double: DISPATCHEDQTY = 0

SQL = "SELECT SUM(ISNULL(QNTY,0)) AS PACKED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='PPF' AND FINITMCOD='" & SITEM & "' AND OPER='+' AND RECSTAT='A' "
SQL = SQL & "AND DBCD ='000006'"

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
  PACKEDQTY = Val(Trim(CHKRS!PACKED & ""))
End If
CHKRS.Close

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS DISPACHED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND FINITMCOD='" & SITEM & _
"' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not CHKRS.EOF Then
 DISPATCHEDQTY = Val(Trim(CHKRS!DISPACHED & ""))
End If
CHKRS.Close

TXTSTKQTY = PACKEDQTY - DISPATCHEDQTY

If Trim(txtitm) = Trim(txtitm.Tag) And Not SAVEFLAG Then
   TXTSTKQTY = TXTSTKQTY + Val(txtQty.Tag)
End If

TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
TXTSTKQTY = Trim(TXTSTKQTY)
End Sub

Private Sub FindStock_BOXREGISTER()

If txtitm = Empty Then Exit Sub

Call SetInternal

Dim PACKEDQTY As Double: PACKEDQTY = 0

SQL = "SELECT SUM(ISNULL(GRSWGT - DSPWGT,0)) AS PACKED FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='000006' AND ICOD='" & SITEM & "' AND DSPWGT <> GRSWGT AND RECSTAT='A'"

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
  PACKEDQTY = Val(Trim(CHKRS!PACKED & ""))
End If
CHKRS.Close

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM SPTRAN WHERE COMP='" & compPth & _
            "' AND UNIT='" & UNCD & "' AND VTYP = 'DPF' AND DBCD='" & M_DBCD & _
            "' AND VBNO='" & lblBill & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
   PACKEDQTY = PACKEDQTY + Val(CHKRS!QNTY)
End If
CHKRS.Close

TXTSTKQTY = PACKEDQTY
TXTSTKQTY = nstr(TXTSTKQTY, 12, 3)
TXTSTKQTY = Trim(TXTSTKQTY)

End Sub

Private Function INVALIDDATA() As Boolean

If txtCONSINEE = Empty Then
  If txtCONSINEE.Enabled Then txtCONSINEE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtDCOD = Empty Then
  If txtDCOD.Enabled Then txtDCOD.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTADDRESS = Empty Then
  If TXTADDRESS.Enabled Then TXTADDRESS.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtitm = Empty Then
  If txtitm.Enabled Then txtitm.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If txtQty = Empty Or Val(txtQty) = 0 Then
  If txtQty.Enabled Then txtQty.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTRATE = Empty Then
  If TXTRATE.Enabled Then TXTRATE.SetFocus
  INVALIDDATA = True
  Exit Function
End If

If TXTAMNT = Empty Then
  If TXTAMNT.Enabled Then TXTAMNT.SetFocus
  INVALIDDATA = True
  Exit Function
End If
End Function

Private Sub SETPKGMAN()
On Error GoTo LAST:

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,SRNO,SRCH,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','DPF',"
SQL = SQL & "'1','1','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & lblBill & "','000000','','" & SITEM & _
"','0','0','" & txtQty & "','N','" & cUName & "','-','A')"

CN.Execute SQL

Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub SETBOXREGISTER()
On Error GoTo LAST:
Dim TMPQTY As Double
Dim DSPQTY As Double: DSPQTY = Val(txtQty)
Dim RSFIRST As ADODB.Recordset
Set RSFIRST = New ADODB.Recordset
Dim RSSECOND As ADODB.Recordset
Set RSSECOND = New ADODB.Recordset

If RSFIRST.State = 1 Then RSFIRST.Close
RSFIRST.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
"' AND DBCD='000006' AND ICOD='" & SITEM & "' AND DSPWGT <> GRSWGT AND RECSTAT='A' ORDER BY VBNO", CN, adOpenDynamic, adLockOptimistic
Do While Not RSFIRST.EOF
   If RSSECOND.State = 1 Then RSSECOND.Close
   RSSECOND.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND DBCD='000006' AND ICOD='" & SITEM & "' AND DSPWGT <> GRSWGT AND RECSTAT='A' AND VBNO='" & Trim(RSFIRST!VBNO) & "'", CN, adOpenDynamic, adLockOptimistic
   If Not RSSECOND.EOF Then TMPQTY = Val(RSSECOND!GRSWGT) - Val(RSSECOND!DSPWGT)
        
   If DSPQTY >= TMPQTY Then
    SQL = "UPDATE BOXREGISTER SET DSPWGT = DSPWGT + " & TMPQTY & ",VTYP='DPF',RVBNO='" & Trim(lblBill) & "',RVBDT= '" & Format(TXTVBDT, "MM/DD/YYYY") & _
    "',RDBC = '" & M_DBCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
    "' AND DBCD='000006' AND ICOD='" & SITEM & "' AND DSPWGT <> GRSWGT AND RECSTAT='A' AND VBNO='" & Trim(RSFIRST!VBNO) & "'"
    CN.Execute SQL
    DSPQTY = DSPQTY - TMPQTY
   ElseIf DSPQTY > 0 Then
     SQL = "UPDATE BOXREGISTER SET DSPWGT = DSPWGT + " & DSPQTY & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
     "' AND DVCD='" & DIVCODE & "' AND DBCD='000006' AND ICOD='" & SITEM & _
     "' AND DSPWGT <> GRSWGT AND RECSTAT='A' AND VBNO='" & Trim(RSFIRST!VBNO) & "'"
    CN.Execute SQL
    DSPQTY = 0
   End If
   
   Me.Caption = "Series : " & Trim(RSFIRST!VBNO)
   
RSFIRST.MoveNext
Loop
RSFIRST.Close

Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Function RoundOffReq() As Boolean
RoundOffReq = False

Dim RFRS As ADODB.Recordset
Set RFRS = New ADODB.Recordset
        
If RFRS.State = 1 Then RFRS.Close
RFRS.Open "SELECT ITEMRO FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ITEMRO ='Y'", CN, adOpenDynamic, adLockOptimistic
      If Not RFRS.EOF Then
         RoundOffReq = True
         Exit Function
      End If
      RFRS.Close
End Function

Private Sub SetDispatchType()
Dim DEFAULTDBCD As String, Defaultindex As Long, dctr As Long, DEFAULTTYPE As String
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
Defaultindex = 1
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT TOP 1 DBCD FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND VTYP='DPF' AND RECSTAT<>'D' AND LTNO='WASTE' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic
If Not PKTYPRS.EOF Then
   DEFAULTDBCD = Trim(PKTYPRS!dbcd & "")
End If
PKTYPRS.Close

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND (CODE IN ('000001','000002') OR NAME LIKE '%WASTAGE%') AND FYCD='" & FYCD & "' AND NAME<>'' order by name desc", CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then M_DBCD = Trim(PKTYPRS!CODE)

Do While Not PKTYPRS.EOF
  dctr = dctr + 1
  If DEFAULTDBCD = Trim(PKTYPRS!CODE & "") Then DEFAULTTYPE = Trim(PKTYPRS!NAME & ""): M_DBCD = Trim(PKTYPRS!CODE & "")
  cmbDispatchType.AddItem Trim(PKTYPRS!NAME & "")
  PKTYPRS.MoveNext
Loop

lblBill.Caption = GenDPFVNO("DPF", M_DBCD, DIVCODE)
If cmbDispatchType.ListCount > 1 Then cmbDispatchType.Text = DEFAULTTYPE
End Sub

Private Sub SetDBCD()
  Dim DBCDRS As ADODB.Recordset
  Set DBCDRS = New ADODB.Recordset
  If DBCDRS.State = 1 Then DBCDRS.Close
  DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND VTYP='DPF' AND NAME = '" & cmbDispatchType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not DBCDRS.EOF Then
     M_DBCD = Trim(DBCDRS!CODE & "")
  Else
     M_DBCD = Empty
  End If
  DBCDRS.Close
End Sub


Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
  SendKeys "{TAB}"
End If

End Sub

'FIFO----------------------
Private Sub SetFIFOReverseForJobworkStock()
On Error GoTo FIFOERR

'VARIABLE DECLARATION
Dim INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double, NETWGT As Double
Dim ITMCODE As String

Dim TEMPRS As ADODB.Recordset: Set TEMPRS = New ADODB.Recordset
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT ISNULL(SUM(QNTY),0) AS QNTY FROM SPTRAN WHERE COMP='" & compPth & _
            "' AND UNIT='" & UNCD & "' AND VTYP = 'DPF' AND DBCD='" & M_DBCD & _
            "' AND VBNO='" & lblBill & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not TEMPRS.EOF Then
   NETWGT = Val(TEMPRS!QNTY)
   BALQNTY = Val(TEMPRS!QNTY)
End If
TEMPRS.Close

If NETWGT = 0 Then Exit Sub
   
'FIND JOBCARD RECEIPE
If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT NTWGT,DSPWGT,* FROM BOXREGISTER " & _
            "INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND " & _
            "FINITMMST.UNIT=BOXREGISTER.UNIT AND FINITMMST.DVCD=BOXREGISTER.DVCD AND " & _
            "FINITMMST.CODE=BOXREGISTER.ICOD AND FINITMMST.NAME='" & txtitm & "' " & _
            "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
            "' AND BOXREGISTER.DBCD='000006' AND BOXREGISTER.LOTNO='WASTE' AND DSPWGT > 0 ORDER BY BOXREGISTER.VBDT DESC", CN, adOpenDynamic, adLockOptimistic
            
    Do While Not FIFORS.EOF
   
        TMPQTY = Val(FIFORS!DSPWGT)
            
        If BALQNTY > TMPQTY Then
           FIFORS!DSPWGT = 0
           BALQNTY = BALQNTY - TMPQTY
           FIFORS.Update
        ElseIf BALQNTY > 0 Or BALQNTY = TMPQTY Then
           FIFORS!DSPWGT = Val(FIFORS!DSPWGT) - BALQNTY
           FIFORS.Update
           BALQNTY = 0
           Exit Do
        End If
                
    FIFORS.MoveNext
    Loop
    FIFORS.Close
    

   
Exit Sub
FIFOERR:
MsgBox ERR.Description
Resume
End Sub

Private Function IsSaleExist() As Boolean
'default
IsSaleExist = False
'-----------------------------------
Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
Dim SQL As String

  'CODE TO CHECK SALE BILL EXIST
  SQL = "SELECT TOP 1 VBNO FROM SPTRAN "
  SQL = SQL & "WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND DVCD ='" & DIVCODE & "' AND DBCD ='" & M_DBCD & _
  "'  AND VBNO ='" & CHALLAN & "' AND VTYP='DPF' AND RECSTAT='A' AND SVBN IS NOT NULL"
   
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If Not CHKRS.EOF Then
     IsSaleExist = True
  End If
  '---------------------------------
End Function
