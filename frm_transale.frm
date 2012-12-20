VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_transale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Purchase Trancasation"
   ClientHeight    =   6900
   ClientLeft      =   540
   ClientTop       =   1170
   ClientWidth     =   11025
   LinkTopic       =   "FORM1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11025
   Begin VB.Frame FRMBTRM 
      Height          =   2415
      Left            =   7200
      TabIndex        =   27
      Top             =   4440
      Width           =   3855
      Begin VB.TextBox TXTADLS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtBEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox TXTBNET 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   1965
         Width           =   1905
      End
      Begin MSFlexGridLib.MSFlexGrid flexBTRM 
         Height          =   1635
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   0
         Cols            =   3
         FixedRows       =   0
         Appearance      =   0
      End
      Begin VB.Label LBLNET 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   33
         Top             =   2040
         Width           =   1305
      End
   End
   Begin VB.Frame FRMLRDTL 
      Height          =   1695
      Left            =   0
      TabIndex        =   26
      Top             =   4440
      Width           =   7095
      Begin VB.TextBox TXTLRNO 
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TXTTRNM 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox TXTVHCL 
         Height          =   285
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   41
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox TXTCRDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   47
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXTPRTM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   49
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TXTRMTM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         MaxLength       =   5
         TabIndex        =   51
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   840
         MaxLength       =   250
         TabIndex        =   53
         Top             =   1320
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker TXTLRDT 
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   80478209
         CurrentDate     =   39347
      End
      Begin VB.Label LBLLRNO 
         Caption         =   "L.R.No."
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LBLLRDT 
         Caption         =   "L.R.Dt."
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LBLTRNM 
         Caption         =   "Name of Transport"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LBLVHCL 
         Caption         =   "Vechicle No."
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LBLCRDS 
         Caption         =   "Cr.Days"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   615
      End
      Begin VB.Label LBLPRTM 
         Caption         =   "Prepration Time"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LBLRMTM 
         Alignment       =   2  'Center
         Caption         =   "Removal Time :"
         Height          =   255
         Left            =   4800
         TabIndex        =   50
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LBLRMRK 
         Caption         =   "Remark "
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10920
      TabIndex        =   59
      ToolTipText     =   "Click to Add New Item"
      Top             =   6720
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3600
      TabIndex        =   58
      Top             =   6360
      Width           =   795
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Cancel Invoice"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   57
      Top             =   6360
      Width           =   795
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      TabIndex        =   56
      Top             =   6360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1935
      TabIndex        =   55
      Top             =   6360
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      TabIndex        =   54
      Top             =   6360
      Width           =   795
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   2
      Top             =   6360
      Width           =   795
   End
   Begin VB.Frame frm_head 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   11175
      Begin VB.TextBox TXTRTORTAX 
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox TXTDLPTY 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   255
         Left            =   9720
         TabIndex        =   19
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   80478209
         CurrentDate     =   39347
      End
      Begin VB.TextBox TXTCOMINV 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TXTVBNO 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TXTTAXNAM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox TXTBRNM 
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TXTDBAC 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   7095
      End
      Begin VB.TextBox TXTCRAC 
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   7095
      End
      Begin VB.Shape Shape2 
         Height          =   1815
         Left            =   0
         Top             =   0
         Width           =   11055
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   8520
         Y1              =   0
         Y2              =   1800
      End
      Begin VB.Label LBLRTORTX 
         Caption         =   "Retail/Tax Inv."
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LBLDLPTY 
         Caption         =   "Delivery Party"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBLBILLDATE 
         Caption         =   "Bill Date"
         Height          =   255
         Left            =   8640
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Label LBLCOMMINV 
         Caption         =   "Comm InvNo"
         Height          =   255
         Left            =   8640
         TabIndex        =   20
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label LBLBILLNO 
         Caption         =   "Bill No. "
         Height          =   255
         Left            =   8640
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LBLTAXNAM 
         Caption         =   "Tax Reference"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LBLBRNM 
         Caption         =   "Agent Name"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label LBLDRAC 
         Caption         =   "Db A/c Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LBLCRAC 
         Caption         =   "Cr A/c Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame ITMFRM 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      TabIndex        =   22
      Top             =   1800
      Width           =   11055
      Begin VB.TextBox TXTITOT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox TXTTQTY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   39
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TXTTPCS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1815
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   20
         FixedCols       =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Total Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   44
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Total Carton"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   35
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label LBLGRS 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9360
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Gross Amount "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   24
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         Height          =   2055
         Left            =   0
         Top             =   480
         Width           =   11055
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "D.O. Quantity"
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
      Left            =   4440
      TabIndex        =   61
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label LBLDO 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   5760
      TabIndex        =   60
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   0
      Top             =   6240
      Width           =   7095
   End
   Begin VB.Label LBLDAYBOK 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "DAY BOOK : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label LBLDIV 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "DIVISION : "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frm_transale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FIL_ITM_COD As String
Public FIL_GRADE As String
Public FIL_PKGCOD As String
Public SEL_DOS_TYP As String
Public SEL_DOS_SRN As String
Public SEL_SCOD As String
Public SEL_RATE As Double
Public SEL_ORDN As String
Public SEL_OSRC As String
Public FRT_RAT As Double
Public SEL_TRCD As String
Public DOS_CANC_CLICK As Boolean
Public SALBOK As String
Public LR_REQ As String
Public m_dbcd As String
Dim M_OPER(0 To 10) As String
Dim M_PERC(0 To 10) As Double
Dim M_POSTCOD(0 To 10) As String
Dim M_NICK(0 To 10) As String
Dim M_POSTYESNO(0 To 10) As String
Dim M_FMLA(0 To 10) As String
Dim M_RDOF(0 To 10) As String
Dim M_BILRDOF As String
Dim saveflag As Boolean
Public isretail As String
Public FILT_TXRT As Boolean
Public MIN_DAT As Date
Dim calbtm As Boolean
Dim chgFlag As Boolean
Dim FRMOPER As String
Public m_srno As String
Dim sav_srfx  As String
Public M_EXCISABLE As String
Dim CHK_FLX As Boolean
Dim Emptycell As Boolean
Dim FLXROW As Double
Dim FLXCOL As Double
Public UNTCOD As String
Public UNTNAM As String
Public EDITRAT As Boolean
Public Sub cmdadd_Click()
  FIL_ITM_COD = Empty
  FIL_GRADE = Empty
  FIL_PKGCOD = Empty
  SEL_DOS_TYP = Empty
  SEL_DOS_SRN = Empty
  SEL_SCOD = Empty
  SEL_RATE = 0
  SEL_ORDN = Empty
  SEL_OSRC = Empty
  frm_transale.DOS_CANC_CLICK = False
  zoomflag = False
  FRMOPER = "*"
  m_srno = Empty
  saveflag = True
  btn_sts (False)

  'Generate Bill No. from daybook or UNTCFG depending upon exciseable or not
  Set rs = New ADODB.Recordset
  If rs.State = 1 Then rs.Close
  rs.Open "select BSFX,gpnr,VBNO from daybok where comp='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' and vtyp='SAL' AND NAME='" & SALBOK & "'", CN, adOpenDynamic, adLockOptimistic
  If rs.EOF Then
    Unload Me
  End If
  Dim BOK_GEN As Boolean
  Dim m_sfx
  Dim m_pfx
  
  If Not IsNull(rs!gpnr) Then
      If rs!gpnr = "Y" Then
          BOK_GEN = False
      Else
          BOK_GEN = True
      End If
  Else
      BOK_GEN = True
  End If
  m_sfx = Trim(rs!bsfx & "")
  sav_srfx = m_sfx
  m_pfx = ""
  FILT_TXRT = False
  M_EXCISABLE = "N"
  isretail = "N"
  If BOK_GEN = True Then
     If rs.State = 1 Then rs.Close
     rs.Open "select BSFX,gpnr,VBNO,TXBI from daybok where comp='" & compPth & "' AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' and vtyp='SAL' AND NAME='" & SALBOK & "'"
     m_sfx = Trim(rs!bsfx & "")
     If Mid(rs!txbi, 1, 1) = "Y" Then
       isretail = "Y"
      Else
       isretail = "N"
     End If
     TXTVBNO = GENVBNO("Select * from DAYBOK where COMP='" & compPth & "'  AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' AND [NAME]='" & SALBOK & "'", m_sfx, m_pfx)
     
    Else
     If rs.State = 1 Then rs.Close
     rs.Open "SELECT * FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
     If Not rs.EOF Then
       TXTVBNO = genVBNoEXC("Select * from UNTCFG where COMP='" & compPth & "' AND UNIT='" & UNCD & "'", "", "")
       M_EXCISABLE = "Y"
       FILT_TXRT = True
     End If
  End If
  TXTCOMINV = GENVBNO("Select * from DAYBOK where COMP='" & compPth & "'  AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' AND [NAME]='" & SALBOK & "'", m_sfx, m_pfx)
  
  Load frm_DOSList
  frm_DOSList.Show 1
  If frm_transale.DOS_CANC_CLICK = True Then
    Unload Me
    Exit Sub
  End If
  Load FRM_BATCHSELECTION
  FRM_BATCHSELECTION.Show 1
  If frm_transale.DOS_CANC_CLICK = True Then
    Unload Me
    Exit Sub
  End If
  'Load frm_EGPSAL
  'frm_EGPSAL.Show 1
  Call calADLS
  If TXTCRAC.Enabled = True Then
    TXTCRAC.SetFocus
  End If
End Sub
Private Sub cmdCancel_Click()
  ClsData (frm_transale)
  FIL_ITM_COD = Empty
  FIL_GRADE = Empty
  FIL_PKGCOD = Empty
  SEL_DOS_TYP = Empty
  SEL_DOS_SRN = Empty
  SEL_SCOD = Empty
  SEL_RATE = 0
  SEL_ORDN = Empty
  SEL_OSRC = Empty
  LBLDO.Caption = "0.000"
  frm_transale.DOS_CANC_CLICK = False
  Flex.Clear
  Flex.Rows = 2
  btn_sts (True)
  Call setflexhead
  If zoomflag = True Then
    Call cmdexit_Click
    Exit Sub
  End If
  
  cmdAdd.SetFocus
  m_srno = Empty
  Dim I As Integer
  For I = 0 To flexBTRM.Rows - 1
    flexBTRM.TextMatrix(I, 2) = "0.00"
  Next
  txtBNET.Text = "0.00"
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("0020", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  saveflag = False
  Dim SAVDAT As New ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  m_srno = Empty
  btn_sts (False)
  frm_SPLIST.Show 1
  'Check for Receipt and Payment Entires
  
  If Not m_srno = Empty Then
    If rs.State = 1 Then rs.Close
    rs.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
      If rs!AUST & "" = "C" Then
         MsgBox "These Entry Can Not Modify / Delete !! Entry Status Clear", vbInformation, "AUDITED"
         m_srno = Empty
         cmdCancel_Click
         Exit Sub
      End If
    End If
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM RPTRAN WHERE COMP='" & compPth & "' AND BSR1='SAL' AND BSR2='" & m_srno & "'", CN, adOpenForwardOnly, adLockPessimistic
    If Not SAVDAT.EOF Then
        MsgBox "Further transaction Exist. Can not Delete it"
        Call cmdCancel_Click
        Exit Sub
    End If
    Dim ays
    ays = MsgBox("Are you sure to delete the invoice ", vbYesNo)
    If ays = vbYes Then
      CN.BeginTrans
      
      Dim m_rtyp As String
      Dim m_rsrn As String
      
      If SAVDAT.State = 1 Then SAVDAT.Close
      CN.Execute "UPDATE SPTRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
      CN.Execute "UPDATE BILLMAIN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
      CN.Execute "UPDATE TRNMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
      'CN.Execute "UPDATE EGPMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
      CN.Execute "DELETE FROM DOTRAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
      Call UPDELSTATUS
      CN.CommitTrans
    End If
  End If
  Call cmdCancel_Click
  If zoomflag = True Then
    Call cmdexit_Click
    Exit Sub
  End If
End Sub

Private Sub cmdedit_Click()
  saveflag = False
End Sub

Private Sub cmdexit_Click()
  Unload Me
End Sub

Private Sub FLEX_Click()
  'cmddelitm.Enabled = True
End Sub
Private Sub FLEX_EnterCell()
  Flex.CellBackColor = RGB(247, 251, 217)
  Emptycell = True
End Sub
Private Sub flex_GotFocus()
  'FLEX.Col = 1
  'FLEX.ROW = 1
  Me.KeyPreview = False
  Flex.TextMatrix(Flex.ROW, 0) = Flex.ROW
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
  
  Flex.TextMatrix(Flex.ROW, 0) = Flex.ROW
  Dim ALLOW_KEY As Boolean
  Dim FWD_COL As Boolean
  Dim ENTER_PRESS As Boolean
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  FWD_COL = False
  ALLOW_KEY = False
  'Rate can Not be Modified
 
  If Flex.Col = 6 Or Flex.Col = 7 Or Flex.Col = 8 Or Flex.Col = 9 Then
    If InStr(1, Flex.TextMatrix(Flex.ROW, Flex.Col), ".") > 0 And KeyAscii = 46 Then
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If Emptycell = True And (Not KeyAscii = 13) Then
    If Flex.Col = 9 Then
      Emptycell = False
     Else
      Flex.TextMatrix(Flex.ROW, Flex.Col) = Empty
      Emptycell = False
    End If
  End If
  Select Case Flex.Col
   Case 1
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
      ALLOW_KEY = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 2
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 3
    NEW_VISIBLE = True
    M_DESC = Empty
    Key = Empty
    Flex.TextMatrix(Flex.ROW, 3) = SearchList1("select code,name from itmmst", 0, Flex.TextMatrix(Flex.ROW, 3), "SELECT ITEM FROM LIST")
    If key_PressNew = True Then
       M_DESC = ""
       Key = ""
       Flex.TextMatrix(Flex.ROW, 3) = ""
       frm_Item.Show
    End If
    ALLOW_KEY = True
   Case 4
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
      ALLOW_KEY = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 5
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii >= 65 And KeyAscii <= 90 Then         ' A-Z
      ALLOW_KEY = True
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then         'a-z
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    ElseIf KeyAscii = 47 Then                              '/
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 6
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 7
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 8
    If KeyAscii >= 48 And KeyAscii <= 57 Then             ' 0- 9
      ALLOW_KEY = True
    ElseIf KeyAscii = 46 Then                              '.
      ALLOW_KEY = True
    Else
      ALLOW_KEY = False
    End If
   Case 9
    ALLOW_KEY = True
   Case 10
    ALLOW_KEY = False
   Case 16
    If Chr(KeyAscii) = "S" Or Chr(KeyAscii) = "Z" Or Chr(KeyAscii) = "0" Or Chr(KeyAscii) = " " Then
      ALLOW_KEY = True
     Else
      ALLOW_KEY = False
    End If
  End Select
  If KeyAscii = vbKeyReturn Then
    ENTER_PRESS = True
   Else
    ENTER_PRESS = False
  End If
  If KeyAscii = 8 Then
    Dim lnth As Double
    lnth = Len(Flex.TextMatrix(Flex.ROW, Flex.Col))
    If lnth > 0 Then
      Flex.TextMatrix(Flex.ROW, Flex.Col) = Mid(Flex.TextMatrix(Flex.ROW, Flex.Col), 1, lnth - 1)
      Exit Sub
    End If
  End If
  If ALLOW_KEY = False Then
    If ENTER_PRESS = True Then
     Else
      KeyAscii = 0
      Exit Sub
    End If
  End If
  
  If ALLOW_KEY = True Then
    If ENTER_PRESS = False Then
      If Flex.Col = 9 Then
       Else
        Flex.TextMatrix(Flex.ROW, Flex.Col) = Trim(Flex.TextMatrix(Flex.ROW, Flex.Col)) + Chr(KeyAscii)
      End If
    End If
  End If
  FWD_COL = False
  If ENTER_PRESS = True Then
    Select Case Flex.Col
     Case 1
      FWD_COL = True
     Case 2
      If Len(Flex.TextMatrix(Flex.ROW, Flex.Col)) = 10 Then
        If IsDate(CDate(Flex.TextMatrix(Flex.ROW, Flex.Col))) Then
          FWD_COL = True
         Else
          FWD_COL = False
        End If
       Else
        FWD_COL = False
      End If
     Case 3
      If MSTDAT.State = 1 Then MSTDAT.Close
      MSTDAT.Open "select * from itmmst where name='" & Flex.TextMatrix(Flex.ROW, Flex.Col) & "'", CN, adOpenDynamic, adLockOptimistic
      If MSTDAT.EOF Then
        FWD_COL = False
       Else
        Flex.TextMatrix(Flex.ROW, 11) = MSTDAT!CODE
        FWD_COL = True
      End If
     Case 4
      FWD_COL = True
     Case 5
      FWD_COL = True
     Case 6
      If IsNumeric(Flex.TextMatrix(Flex.ROW, Flex.Col)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 7
      If IsNumeric(Flex.TextMatrix(Flex.ROW, Flex.Col)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 8
      If IsNumeric(Flex.TextMatrix(Flex.ROW, Flex.Col)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 9
      If IsNumeric(Flex.TextMatrix(Flex.ROW, Flex.Col)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
     Case 10
      If IsNumeric(Flex.TextMatrix(Flex.ROW, Flex.Col)) Then
        FWD_COL = True
       Else
        FWD_COL = False
      End If
    End Select
    
    If FWD_COL = True Then
      If Flex.Col = 9 Then
        Flex.TextMatrix(Flex.ROW, 8) = Format(Val(Flex.TextMatrix(Flex.ROW, 6)) * Val(Flex.TextMatrix(Flex.ROW, 7)), "########.000")
        Flex.TextMatrix(Flex.ROW, 10) = Round(Format(Val(Flex.TextMatrix(Flex.ROW, 9)) * Val(Flex.TextMatrix(Flex.ROW, 8)), "#########.00"), 0)
        Flex.Col = 9
        'Allowed to add row with msgbox
        'Check all the cell are filled
        Call CHKFLEX
        If Not CHK_FLX Then
          MsgBox "Invalid Data in item details "
          Flex.ROW = FLXROW
          Flex.Col = FLXCOL
          Flex.SetFocus
          Exit Sub
        End If
        Dim ays
        'AYS = MsgBox("Want to Add More Item ", vbYesNo)
        'If AYS = vbYes Then
        If Flex.ROW <> Flex.Rows - 1 Then
          Flex.ROW = Flex.ROW + 1
          Flex.Col = 7
         Else
          If flexBTRM.Enabled = True Then
            flexBTRM.SetFocus
            
           Else
            Call calADLS
            If TXTLRNO.Enabled = True Then
              TXTLRNO.SetFocus
             Else
              TXTCRDS.SetFocus
            End If
          End If
          Exit Sub
        End If
       Else
        Flex.Col = Flex.Col + 1
      End If
      Emptycell = True
    End If
  End If
End Sub
Private Sub flex_LeaveCell()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim I As Double
  If Flex.Col = 7 Then
    If Val(Flex.TextMatrix(Flex.ROW, 16)) < Val(Flex.TextMatrix(Flex.ROW, 7)) Then
      MsgBox "Stock Is only : " + nstr(Val(Flex.TextMatrix(Flex.ROW, 16)), 8, 0)
      Flex.Col = 7
      Flex.SetFocus
      Exit Sub
    End If
  End If
  Flex.CellBackColor = vbWhite
  Flex.TextMatrix(Flex.ROW, 8) = Format(Val(Flex.TextMatrix(Flex.ROW, 6)) * Val(Flex.TextMatrix(Flex.ROW, 7)), "########.000")
  Flex.TextMatrix(Flex.ROW, 10) = Round(Format(Val(Flex.TextMatrix(Flex.ROW, 9)) * Val(Flex.TextMatrix(Flex.ROW, 8)), "#########.00"), 0)
  FLEXROW = Flex.ROW
  FLEXCOL = Flex.Col
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For I = 1 To Flex.Rows - 1
    TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(Flex.TextMatrix(I, 7)), "######")
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(Flex.TextMatrix(I, 8)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(Flex.TextMatrix(I, 10)), "########.00")
  Next
  Flex.ROW = FLEXROW
  Flex.Col = FLEXCOL
  Flex.SetFocus
End Sub
Private Sub flex_LostFocus()
  Dim FLEXROW As Double
  Dim FLEXCOL As Double
  Dim I As Double
  Flex.CellBackColor = vbWhite
  Flex.TextMatrix(Flex.ROW, 8) = Format(Val(Flex.TextMatrix(Flex.ROW, 6)) * Val(Flex.TextMatrix(Flex.ROW, 7)), "########.000")
  Flex.TextMatrix(Flex.ROW, 10) = Round(Format(Val(Flex.TextMatrix(Flex.ROW, 9)) * Val(Flex.TextMatrix(Flex.ROW, 8)), "#########.00"), 0)
  FLEXROW = Flex.ROW
  FLEXCOL = Flex.Col
  TXTTPCS.Text = 0
  TXTTQTY.Text = 0
  TXTITOT.Text = 0
  For I = 1 To Flex.Rows - 1
    TXTTPCS.Text = Format(Val(TXTTPCS.Text) + Val(Flex.TextMatrix(I, 7)), "######")
    TXTTQTY.Text = Format(Val(TXTTQTY.Text) + Val(Flex.TextMatrix(I, 8)), "########.000")
    TXTITOT.Text = Format(Val(TXTITOT.Text) + Val(Flex.TextMatrix(I, 10)), "########.00")
  Next
  Flex.ROW = FLEXROW
  Flex.Col = FLEXCOL
End Sub

Private Sub Form_Activate()
  DIVCOD = Me.Tag
  Dim divisionmaster As New ADODB.Recordset
  Set divisionmaster = New ADODB.Recordset
  If divisionmaster.State = 1 Then divisionmaster.Close
  divisionmaster.Open "select * from DIVMST where code='" & DIVCOD & "' AND UNIT='" & UNCD & "' and comp='" & compPth & "'", CN
  If Not divisionmaster.EOF Then
    DIVNAM = divisionmaster!Name
    DIVCOD = divisionmaster!CODE
   Else
    DIVNAM = "??????"
  End If
  LBLDIV.Caption = "DIVSION : " + DIVNAM
  If DIVNAM = "??????" Then
    Unload Me
  End If
  SALBOK = Me.Caption
  FRMPARA = "SAL"
  If DIVNAM = "??????" Or SALBOK = "??????" Or Trim(SALBOK) = "" Then
    Unload Me
    Exit Sub
  End If
  If zoomflag = True Then
    saveflag = False
    btn_sts (False)
    CMDSAVE.Enabled = False
    cmdDelete.Enabled = False
    
  End If
  If LR_REQ = "N" Then
    TXTLRNO.Enabled = False
    TXTTRNM.Enabled = False
    TXTVHCL.Enabled = False
    TXTLRDT.Enabled = False
   Else
    TXTLRNO.Enabled = True
    TXTTRNM.Enabled = True
    TXTVHCL.Enabled = True
    TXTLRDT.Enabled = True
  End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
  If (ActiveControl.Name = "TXTCRAC" Or ActiveControl.Name = "TXTDBAC" Or ActiveControl.Name = "TXTDLPTY" Or ActiveControl.Name = "TXTBRNM" Or ActiveControl.Name = "TXTTAXNAM") Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub Form_Load()
  TXTVBDT.MinDate = FSDT
  TXTVBDT.MaxDate = FEDT
  LR_REQ = "N"
  Emptycell = True
  flexBTRM.ColWidth(0) = 1500
  flexBTRM.ColWidth(1) = 800
  flexBTRM.ColWidth(2) = 1200
  TXTVBDT = Date
  TXTLRDT = Date
  m_dbcd = Empty
  Set rs = New ADODB.Recordset
  FRMPARA = "SAL"
  Call CenterChild(frm_Main, Me)
  Call setflexhead
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  If Not zoomflag = True Then
    m_srno = Empty
    If DIVCOD = Empty Then
           
       DIVNAM = SearchList1("SELECT CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT DIVISION MASTER")
    End If
  End If
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND NAME='" & DIVNAM & "' AND UNIT='" & UNCD & "'", CN, adOpenKeyset, adLockPessimistic
  If Not rs.EOF Then
     DIVCOD = rs!CODE
     DIVNAM = rs!Name
     Me.Tag = DIVCOD
     LBLDIV.Caption = "DIVSION : " + DIVNAM
    Else
     LBLDIV.Caption = "DIVISION : " + "??????"
  End If
  'For Day Book
  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  If Not zoomflag = True Then
    If m_dbcd = Empty Then
     SALBOK = SearchList1("SELECT DBCD,NAME FROM DAYBOK WHERE COMP='" & compPth & "' AND VTYP='" & FRMPARA & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' ", 0, "SALES ACCOUNT (GENERAL)", "SELECT DAYBOK MASTER")
    End If
  End If
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT * FROM DAYBOK WHERE COMP='" & compPth & "' AND NAME='" & SALBOK & "' AND VTYP='" & FRMPARA & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' ", CN, adOpenKeyset, adLockPessimistic
  If Not rs.EOF Then
     m_dbcd = rs!dbcd
     SALBOK = rs!Name
     M_BILRDOF = rs!rdof & ""
     LR_REQ = rs!lrrq & ""
     LBLDAYBOK.Caption = "DAY BOOK : " + SALBOK
    Else
     LBLDAYBOK.Caption = "DAY BOOK : " + "???????"
     LR_REQ = "N"
  End If
  Me.Caption = SALBOK
  Call FIL_Billingterm
  Call btn_sts(True)
  Call TXTPRTM_GotFocus
  Call TXTRMTM_GotFocus
  Me.KeyPreview = True
End Sub
Private Sub setflexhead()
    Flex.TextMatrix(0, 0) = "Sr."
    Flex.TextMatrix(0, 1) = "Challan No."
    Flex.TextMatrix(0, 2) = "Challan Dt."
    Flex.TextMatrix(0, 3) = "Item Name"
    Flex.TextMatrix(0, 4) = "Batch No."
    Flex.TextMatrix(0, 5) = "Grade"
    Flex.TextMatrix(0, 6) = "Pkg-Type"
    Flex.TextMatrix(0, 7) = "Bags"
    Flex.TextMatrix(0, 8) = "Quanity"
    Flex.TextMatrix(0, 9) = "Rate"
    Flex.TextMatrix(0, 10) = "Amount"
    Flex.TextMatrix(0, 11) = "ICOD"
    Flex.TextMatrix(0, 12) = "RTYP"
    Flex.TextMatrix(0, 13) = "RSRN"
    Flex.TextMatrix(0, 14) = "ORDN"
    Flex.TextMatrix(0, 15) = "ORDRATE"
    Flex.TextMatrix(0, 16) = "Valid"
    Flex.TextMatrix(0, 17) = ""
    Flex.TextMatrix(0, 18) = ""
    Flex.TextMatrix(0, 19) = ""
    Flex.ColWidth(0) = 300
    Flex.ColWidth(1) = 900
    Flex.ColWidth(2) = 1000
    Flex.ColWidth(3) = 1500
    Flex.ColWidth(4) = 900
    Flex.ColWidth(5) = 1200
    Flex.ColWidth(6) = 900
    Flex.ColWidth(7) = 900
    Flex.ColWidth(8) = 1000
    Flex.ColWidth(9) = 900
    Flex.ColWidth(10) = 1300
    Flex.ColWidth(11) = 0
    Flex.ColWidth(12) = 0
    Flex.ColWidth(13) = 0
    Flex.ColWidth(14) = 0
    Flex.ColWidth(15) = 0
    Flex.ColWidth(16) = 0
    Flex.ColWidth(17) = 0
    Flex.ColWidth(18) = 0
    Flex.ColWidth(19) = 0
    Flex.ColAlignment(3) = 0
    Flex.ColAlignment(2) = 0
End Sub
Public Sub btn_sts(Yes As Boolean)
    CMDSAVE.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    cmdEdit.Enabled = Yes
    cmdDelete.Enabled = Yes
    frm_head.Enabled = Not Yes
    ITMFRM.Enabled = Not Yes
    FRMLRDTL.Enabled = Not Yes
    FRMBTRM.Enabled = Not Yes
End Sub
Public Sub FIL_Billingterm()
Dim cntr As Byte
  flexBTRM.Clear
  flexBTRM.Rows = 0
  M_OPER(0) = ""
  M_OPER(1) = ""
  M_OPER(2) = ""
  M_OPER(3) = ""
  M_OPER(4) = ""
  M_OPER(5) = ""
  M_OPER(6) = ""
  M_OPER(7) = ""
  M_OPER(8) = ""
  M_OPER(9) = ""
  M_PERC(0) = 0
  M_PERC(1) = 0
  M_PERC(2) = 0
  M_PERC(3) = 0
  M_PERC(4) = 0
  M_PERC(5) = 0
  M_PERC(6) = 0
  M_PERC(7) = 0
  M_PERC(8) = 0
  M_PERC(9) = 0
  M_POSTCOD(0) = ""
  M_POSTCOD(1) = ""
  M_POSTCOD(2) = ""
  M_POSTCOD(3) = ""
  M_POSTCOD(4) = ""
  M_POSTCOD(5) = ""
  M_POSTCOD(6) = ""
  M_POSTCOD(7) = ""
  M_POSTCOD(8) = ""
  M_POSTCOD(9) = ""
  M_NICK(0) = ""
  M_NICK(1) = ""
  M_NICK(2) = ""
  M_NICK(3) = ""
  M_NICK(4) = ""
  M_NICK(5) = ""
  M_NICK(6) = ""
  M_NICK(7) = ""
  M_NICK(8) = ""
  M_NICK(9) = ""
  M_POSTYESNO(0) = ""
  M_POSTYESNO(1) = ""
  M_POSTYESNO(2) = ""
  M_POSTYESNO(3) = ""
  M_POSTYESNO(4) = ""
  M_POSTYESNO(5) = ""
  M_POSTYESNO(6) = ""
  M_POSTYESNO(7) = ""
  M_POSTYESNO(8) = ""
  M_POSTYESNO(9) = ""
  M_FMLA(0) = ""
  M_FMLA(1) = ""
  M_FMLA(2) = ""
  M_FMLA(3) = ""
  M_FMLA(4) = ""
  M_FMLA(5) = ""
  M_FMLA(6) = ""
  M_FMLA(7) = ""
  M_FMLA(8) = ""
  M_FMLA(9) = ""
  M_RDOF(0) = ""
  M_RDOF(1) = ""
  M_RDOF(2) = ""
  M_RDOF(3) = ""
  M_RDOF(4) = ""
  M_RDOF(5) = ""
  M_RDOF(6) = ""
  M_RDOF(7) = ""
  M_RDOF(8) = ""
  M_RDOF(9) = ""
  Set rs = New ADODB.Recordset
  If rs.State = 1 Then rs.Close
  rs.Open "select * from config where comp='" & compPth & "' and vtyp='" & FRMPARA & "' AND DBCD='" & m_dbcd & "'  AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' order by srch", CN, adOpenKeyset, adLockPessimistic
  cntr = 0
  Do While Not rs.EOF
   flexBTRM.Rows = flexBTRM.Rows + 1
   flexBTRM.TextMatrix(cntr, 0) = rs!NICK & ""
   flexBTRM.TextMatrix(cntr, 1) = Format(rs!PERC, "#######.00")
   M_OPER(cntr) = Trim(rs!OPER)
   M_PERC(cntr) = rs!PERC
   M_POSTCOD(cntr) = Trim(rs!CODE)
   M_NICK(cntr) = Trim(rs!NICK)
   M_POSTYESNO(cntr) = Trim(rs!post)
   M_FMLA(cntr) = Trim(rs!FMLA)
   M_RDOF(cntr) = Trim(rs!rdof)
   rs.MoveNext
   cntr = cntr + 1
  Loop
  Dim TMP_FMLA(0 To 10) As String
  cntr = 0
  For cntr = 0 To 9
    
    M_FMLA(cntr) = Replace(M_FMLA(cntr), "GROSS TOTAL", "M_STOT ")
    M_FMLA(cntr) = Replace(M_FMLA(cntr), "TOTAL QUANTITY", "M_TQTY ")
    M_FMLA(cntr) = Replace(M_FMLA(cntr), "TOTAL PCS", "M_TPCS ")
    If M_NICK(0) <> "" Then
        If M_OPER(0) = "+" Then
          M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(0), "AMT_01 ")
         Else
          M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(0), " -AMT_01")
        End If
    End If
    If M_NICK(1) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(1), " +AMT_02")
    End If
    If M_NICK(2) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(2), " +AMT_03")
    End If
    If M_NICK(3) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(3), " +AMT_04")
    End If
    If M_NICK(4) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(4), " +AMT_05")
    End If
    If M_NICK(5) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(5), " +AMT_06")
    End If
    If M_NICK(6) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(6), " +AMT_07")
    End If
    If M_NICK(7) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(7), " +AMT_08")
    End If
    If M_NICK(8) <> "" Then
        M_FMLA(cntr) = Replace(M_FMLA(cntr), M_NICK(8), " +AMT_09")
    End If
  Next
  If flexBTRM.Rows > 0 Then
    'O.k
   Else
    flexBTRM.Enabled = False
  End If
End Sub

Private Sub LBLGRS_Change()
  TXTITOT = Format(LBLGRS, "#########.00")
End Sub
Private Sub TXTBRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTBRNM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTBRNM.Text = SearchList1("SELECT CODE,NAME FROM REFMST WHERE CATA='B'", 0, TXTBRNM, "List of Agent")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTBRNM.Text = ""
            Ref_Cat = "B"
            Frm_Ref_FAS.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTBRNM = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub


Private Sub TXTRMRK_KeyUp(KeyCode As Integer, Shift As Integer)

Dim rsNarration As Recordset

    Set rsNarration = New Recordset
    rsNarration.Open "Select * From NARRMAST Where KCod='" & KeyCode & "' And shift =" & Shift & " and Modul='SALES'", CN, adOpenDynamic, adLockOptimistic

    
    If rsNarration.EOF = False Then
        TXTRMRK = Trim(rsNarration!narr)
        TXTRMRK.SelStart = 1000
    End If
    
    rsNarration.Close
      
End Sub


Private Sub TXTCRAC_KeyDown(KeyCode As Integer, Shift As Integer)
    If TXTCRAC = Empty Then TXTCRAC = "SALES ACCOUNT (GENERAL)"
    Me.KeyPreview = False
    If KeyCode = vbKeyF2 Or Trim(TXTCRAC.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTCRAC.Text = SearchList1("SELECT CODE,NAME FROM ACCMST WHERE HCOD='000011'", 0, TXTCRAC, "List of Credit A/c")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTCRAC.Text = ""
            frm_Acc.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTCRAC = Empty
    End If
    Me.KeyPreview = True
    'If KeyCode = vbKeyReturn Then
    '  SendKeys "{TAB}"
    'End If
End Sub

Private Sub TXTCRAC_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And (Not Trim(TXTCRAC.Text) = Empty) Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub TXTDBAC_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn And (Not Trim(TXTDBAC.Text) = Empty) Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TXTCRDS_GotFocus()
  Me.KeyPreview = True
  TXTCRDS.SelStart = 0
  TXTCRDS.SelLength = Len(TXTCRDS)
End Sub

Private Sub TXTDBAC_Change()
  Dim SEL_TXCD As String
  Dim SEL_BRCD As String
  Dim rs As New ADODB.Recordset
  Set rs = New ADODB.Recordset
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockPessimistic
  If Not rs.EOF Then
    TXTRTORTAX.Text = rs!ttyp & ""
    TXTRTORTAX.Enabled = False
    TXTCRDS = Val(rs!CDAY)
    SEL_TXCD = rs!TXCD & ""
    SEL_BRCD = rs!BRCD & ""
  End If
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT * FROM REFMST WHERE CODE='" & SEL_TXCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    TXTTAXNAM = rs!Name
  End If
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT * FROM REFMST WHERE CODE='" & SEL_BRCD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not rs.EOF Then
    TXTBRNM = rs!Name
  End If
End Sub
Private Sub TXTDBAC_KeyDown(KeyCode As Integer, Shift As Integer)
   Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTDBAC.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDBAC.Text = SearchList1("SELECT CODE,NAME FROM ACCMST", 0, TXTCRAC, "List of Debit A/c")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            TXTDBAC.Text = ""
            frm_Acc.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDBAC = Empty
    End If
    Me.KeyPreview = True
    'If KeyCode = vbKeyReturn Then
    '  SendKeys "{TAB}"
    'End If
End Sub
Private Sub TXTDLPTY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTDLPTY.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTDLPTY.Text = SearchList1("SELECT CODE,NAME FROM REFMST WHERE CATA='Y'", 0, TXTDLPTY, "List of Delivery Party")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "Y"
            TXTDLPTY.Text = ""
            Frm_Ref_FAS.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTDLPTY = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTITOT_Change()
  If flexBTRM.Rows > 0 Then
    flexBTRM.Col = 0
    flexBTRM.ROW = 0
  End If
  
  calBTRM 0
  calADLS
End Sub

Private Sub txtLRDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub TXTLRDT_LostFocus()
  'If TXTLRDT < MIN_DAT Then
  '  MsgBox "L.R date should be grater than or equal to  challan date", vbInformation
  '  TXTLRDT.SetFocus
  '  Exit Sub
  'End If
End Sub

Private Sub TXTLRNO_GotFocus()
  Me.KeyPreview = True
  TXTLRNO.SelStart = 0
  TXTLRNO.SelLength = Len(TXTLRNO)
End Sub
Private Sub TXTPRTM_GotFocus()
  TXTPRTM.SelStart = 0
  TXTPRTM.SelLength = Len(TXTPRTM)
  If saveflag = True Then
    If Trim(TXTPRTM) = Empty Then
      TXTPRTM = Mid(Time(), 1, 5)
    End If
  End If
End Sub

Private Sub TXTRMRK_GotFocus()
  TXTRMRK.SelStart = 0
  TXTRMRK.SelLength = Len(TXTRMRK)
End Sub

Private Sub TXTRMTM_GotFocus()
  TXTRMTM.SelStart = 0
  TXTRMTM.SelLength = Len(TXTRMTM)
  If saveflag = True Then
    If Trim(TXTRMTM) = Empty Then
      TXTRMTM = Mid(Time(), 1, 5)
    End If
  End If
End Sub

Private Sub TXTTAXNAM_Change()
   Dim MSTDAT As New ADODB.Recordset
   Dim CSTPER As Double
   Dim I As Double
   If TXTTAXNAM.Text = "C-FORM" Then
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM REFMST WHERE NAME='C-FORM'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       CSTPER = MSTDAT!PERC
     End If
   
     I = 1
     For I = 1 To flexBTRM.Rows - 1
       If flexBTRM.TextMatrix(I, 0) = "VAT" Then
         flexBTRM.TextMatrix(I, 1) = 0
         flexBTRM.TextMatrix(I, 2) = 0
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "AVAT" Then
         flexBTRM.TextMatrix(I, 1) = 0
         flexBTRM.TextMatrix(I, 2) = 0
         Call calBTRM(I)
       End If
       
       If flexBTRM.TextMatrix(I, 0) = "CST" Then
         flexBTRM.TextMatrix(I, 1) = Format(CSTPER, "##.00")
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "FREIGHT" Then
         flexBTRM.TextMatrix(I, 1) = Format(FRT_RAT, "##.00")
         Call calBTRM(I)
       End If
     Next
   ElseIf TXTTAXNAM.Text = "WITHOUT FORM C" Then
    If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM REFMST WHERE NAME='WITHOUT FORM C'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       CSTPER = MSTDAT!PERC
     End If
   
     I = 1
     For I = 1 To flexBTRM.Rows - 1
       If flexBTRM.TextMatrix(I, 0) = "VAT" Then
         flexBTRM.TextMatrix(I, 1) = 0
         flexBTRM.TextMatrix(I, 2) = 0
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "AVAT" Then
         flexBTRM.TextMatrix(I, 1) = 0
         flexBTRM.TextMatrix(I, 2) = 0
         Call calBTRM(I)
       End If
       
       If flexBTRM.TextMatrix(I, 0) = "CST" Then
         flexBTRM.TextMatrix(I, 1) = Format(CSTPER, "##.00")
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "FREIGHT" Then
         flexBTRM.TextMatrix(I, 1) = Format(FRT_RAT, "##.00")
         Call calBTRM(I)
       End If
     Next
   Else
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTTAXNAM & "'", CN, adOpenDynamic, adLockOptimistic
     If Not MSTDAT.EOF Then
       CSTPER = MSTDAT!PERC
     End If
   
     I = 1
     For I = 1 To flexBTRM.Rows - 1
       If flexBTRM.TextMatrix(I, 0) = "VAT" Then
         flexBTRM.TextMatrix(I, 1) = Format(CSTPER, "##.00")
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "CST" Then
         flexBTRM.TextMatrix(I, 1) = 0
         flexBTRM.TextMatrix(I, 2) = 0
         Call calBTRM(I)
       End If
       If flexBTRM.TextMatrix(I, 0) = "FREIGHT" Then
         flexBTRM.TextMatrix(I, 1) = Format(FRT_RAT, "##.00")
         Call calBTRM(I)
       End If
     Next
   End If
End Sub

Private Sub TXTTAXNAM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTTAXNAM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTTAXNAM.Text = SearchList1("SELECT CODE,NAME FROM REFMST WHERE CATA='T'", 0, TXTTAXNAM, "List of Tax Catagoery")
        If key_PressNew = True Then
            M_DESC = ""
            Key = ""
            Ref_Cat = "T"
            TXTTAXNAM.Text = ""
            Frm_Ref_FAS.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTTAXNAM = Empty
    End If
    If KeyCode = vbKeyReturn Then
      SendKeys "{TAB}"
    End If
End Sub

Private Sub TXTTPCS_Change()
  calBTRM 0
End Sub

Private Sub TXTTQTY_Change()
  calBTRM 0
End Sub

Private Sub TXTTRNM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(TXTTRNM.Text) = Empty Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTTRNM.Text = SearchList1("SELECT CODE,NAME FROM REFMST WHERE CATA='R'", 0, TXTTRNM, "List of Transporter")
        If key_PressNew = True Then
            
            M_DESC = ""
            Key = ""
            TXTTRNM.Text = ""
            Ref_Cat = "R"
            Frm_Ref_FAS.Show
        End If
    ElseIf KeyCode = vbKeyDelete Then
        TXTTRNM = Empty
    End If
    'If KeyCode = vbKeyReturn Then
    '  SendKeys "{TAB}"
    'End If
End Sub

Private Sub TXTVBDT_Change()
  Dim I As Double
  I = 0
  For I = 1 To Flex.Rows - 1
    Flex.TextMatrix(I, 2) = Format(TXTVBDT.Value, "DD/MM/YYYY")
  Next
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
  If EDITRAT = True Then
    If KeyCode = vbKeyReturn Then
      If Flex.Rows > 0 Then
        Flex.Col = 7
        Flex.SetFocus
       Else
        SendKeys "{TAB}"
      End If
    End If
   Else
    If KeyCode = vbKeyReturn Then
      Flex.SetFocus
    End If
  End If
End Sub
Private Sub calBTRM(ByVal ICTR As Integer)

    Dim j As Integer, iFMLA(20) As Double, subTot As Double
    Dim c_FMLA(20) As String
    Dim l As Integer
    Dim m As Integer
    Dim B() As String
    subTot = 0
    Dim a() As String, K As Integer
    j = 0
    If flexBTRM.Rows = 0 Then Exit Sub
    For j = flexBTRM.ROW To flexBTRM.Rows - 1
        If Val(flexBTRM.TextMatrix(j, 1)) <> 0 Then flexBTRM.TextMatrix(j, 2) = 0
    Next j

    For j = flexBTRM.ROW To flexBTRM.Rows - 1
        c_FMLA(j) = Trim(M_FMLA(j))
        If Len(c_FMLA(j)) <= 6 Then
            Select Case c_FMLA(j)
                Case "M_STOT"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(TXTITOT.Text)) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "M_TQTY"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(TXTTQTY.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "M_TPCS"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(TXTTPCS.Text)), "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_01"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(0, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_02"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(1, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_03"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(2, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_04"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(3, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_05"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(4, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_06"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(5, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_07"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(6, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_08"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(7, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
                Case "AMT_09"
                    If saveflag Or (saveflag = False And Val(flexBTRM.TextMatrix(j, 1)) <> 0) Then
                        flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 1)) * Val(flexBTRM.TextMatrix(8, 2))) / 100, "##########.000")
                    Else
                        flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "#.000")
                    End If
            End Select
                
            If M_RDOF(j) = "Y" Then
                flexBTRM.TextMatrix(j, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(j, 2)), 0), "############.00")
            Else
                flexBTRM.TextMatrix(j, 2) = Format(flexBTRM.TextMatrix(j, 2), "############.00")
            End If
        Else
            c_FMLA(j) = Replace(c_FMLA(j), "M_STOT", Val(TXTITOT.Text))
            c_FMLA(j) = Replace(c_FMLA(j), "M_TQTY", Val(TXTTQTY.Text))
            c_FMLA(j) = Replace(c_FMLA(j), "M_TPCS", Val(TXTTPCS.Text))
            For K = 0 To j
                c_FMLA(j) = Replace(c_FMLA(j), "AMT_0" & K + 1, Format(flexBTRM.TextMatrix(K, 2), "##########.00"))
            Next K
            c_FMLA(j) = c_FMLA(j)
            a() = Split(c_FMLA(j), " ")
            
            Dim y As Double
            
            y = 0
            For K = 0 To UBound(a)
             y = y + Val(a(K))
            Next
                
            If Val(flexBTRM.TextMatrix(j, 1)) <> 0 Then
            
              flexBTRM.TextMatrix(j, 2) = Abs(y)
              
            End If
            
            If M_RDOF(j) = "N" Then
                If Val(flexBTRM.TextMatrix(j, 1)) <> 0 Then
                  flexBTRM.TextMatrix(j, 2) = Format((Val(flexBTRM.TextMatrix(j, 2)) * Val(flexBTRM.TextMatrix(j, 1))) / 100, "##########.00")
                  
                End If
            Else
                'If Val(flexBTRM.TextMatrix(J, 1)) <> 0 Then
                  flexBTRM.TextMatrix(j, 2) = Format(FormatNumber(Val(flexBTRM.TextMatrix(j, 2)) * Val(flexBTRM.TextMatrix(j, 1)) / 100, 0), "##########.00")
                'End If
            End If
            
        End If
MsubTot:
        If M_OPER(j) = "+" Then
            subTot = subTot + Val(flexBTRM.TextMatrix(j, 2))
        Else
            subTot = subTot - Val(flexBTRM.TextMatrix(j, 2))
        End If
        txtBNET.Text = Val(TXTITOT.Text) + subTot
    Next j

    
End Sub

Private Sub EditKeyCode(MSHFlexGrid As Control, Edt As Control, KeyCode As Integer, Shift As Integer)
   
    Dim ans As String
    chgFlag = True
    'Standard edit control processing.
   Select Case KeyCode
    
   Case 27   ' ESC: hide, return focus to MSHFlexGrid.
      Edt.Visible = False
      MSHFlexGrid.SetFocus
    
   Case 9    ' TAB return focus to mshflexgrid.
        If Flex.Col - 1 <> 7 And Flex.Col - 1 <> 0 Then Flex.TextMatrix(Flex.ROW, Flex.Col - 1) = 0
   Case 13    ' ENTER return focus to MSHFlexGrid.
         MSHFlexGrid.SetFocus
         If MSHFlexGrid.Col = 2 Then
            If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
               MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
               MSHFlexGrid.Col = 1
            End If
         Else
            MSHFlexGrid.Col = MSHFlexGrid.Col + 1
        End If
   Case 38      ' Up.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW > MSHFlexGrid.FixedRows Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW - 1
      End If

   Case 40      ' Down.
      MSHFlexGrid.SetFocus
      DoEvents
      If MSHFlexGrid.ROW < MSHFlexGrid.Rows - 1 Then
         MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
      End If
   End Select
   chgFlag = False
End Sub
Private Sub flexBTRM_DblClick()
    calbtm = True
    MSHFlexGridEdit flexBTRM, txtBEdit, 32 ' Simulate a space.
End Sub
Private Sub flexBTRM_GotFocus()
    Me.KeyPreview = False
    Msg "Billing Terms"
    flexBTRM.Col = 1
    flexBTRM.TopRow = 0
    flexBTRM.LeftCol = 1
End Sub
Private Sub flexBTRM_KeyPress(KeyAscii As Integer)
    If flexBTRM.Col = 2 And flexBTRM.ROW + 1 = flexBTRM.Rows Then calbtm = False Else calbtm = True
    MSHFlexGridEdit flexBTRM, txtBEdit, KeyAscii
    If KeyAscii = vbKeyReturn Then
      If flexBTRM.ROW Mod 4 = 0 And flexBTRM.Col = 2 And flexBTRM.ROW > 0 Then
         'SendKeys "{Down}"
         flexBTRM.TopRow = flexBTRM.ROW - 1
         'flexBTRM.Col = 1
      End If
    End If
End Sub
Private Sub MSHFlexGridEdit(MSHFlexGrid As Control, Edt As Control, KeyAscii As Integer)
    chgFlag = True
    ' Use the character that was typed.
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 12
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 14 To 26
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 13
      If MSHFlexGrid.Col = 2 Then
            If MSHFlexGrid.Rows <> MSHFlexGrid.ROW + 1 Then
                MSHFlexGrid.ROW = MSHFlexGrid.ROW + 1
            Else
                If TXTLRNO.Enabled = True Then
                  TXTLRNO.SetFocus
                 Else
                  TXTCRDS.SetFocus
                End If
            End If
            MSHFlexGrid.Col = 1
            Exit Sub
        Else
            MSHFlexGrid.Col = MSHFlexGrid.Col + 1
            Exit Sub
      End If
   Case 28 To 32
      Edt = MSHFlexGrid
      Edt.SelStart = 1000
   Case 27
        Edt.Text = Empty
        Exit Sub
   ' Anything else means replace the current text.
   Case Else
      Edt = Chr(KeyAscii)
      Edt.SelStart = 1
   End Select

   ' Show Edt at the right place.
   Edt.Move MSHFlexGrid.Left + MSHFlexGrid.CellLeft, _
      MSHFlexGrid.Top + MSHFlexGrid.CellTop, _
      MSHFlexGrid.CellWidth - 8, _
      MSHFlexGrid.CellHeight - 8
   Edt.Visible = True

   ' And make it work.
   Edt.SetFocus
   chgFlag = False
End Sub
Private Sub calADLS()
    Dim P As Integer
    TXTADLS.Text = Empty
    For P = 0 To flexBTRM.Rows - 1
        If M_OPER(P) = "-" Then
            TXTADLS.Text = Format(Val(TXTADLS.Text) - Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        Else
            TXTADLS.Text = Format(Val(TXTADLS.Text) + Val(flexBTRM.TextMatrix(P, 2)), "############.00")
        End If
    Next P
    If M_BILRDOF = "Y" Then
        txtBNET.Text = Format(FormatNumber(Val(TXTITOT.Text) + Val(TXTADLS.Text), 0), "##########.00")
    Else
        txtBNET.Text = Format(Val(TXTITOT.Text) + Val(TXTADLS.Text), "##########.00")
    End If
End Sub
Private Sub txtBEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode flexBTRM, txtBEdit, KeyCode, Shift
End Sub
Private Sub txtBEdit_KeyPress(KeyAscii As Integer)
   If KeyAscii = Asc(vbCr) Then KeyAscii = 0
End Sub
Private Sub flexBTRM_LeaveCell()
    If txtBEdit.Visible = False Then Exit Sub
    flexBTRM = txtBEdit
    txtBEdit.Visible = False
End Sub
Private Sub flexBTRM_RowColChange()
    If flexBTRM.Col = 1 Then
        'If flexBTRM.ROW = 0 Then txtDUTY.Text = Format(Val(txtDUTY.Text), "#########.00")
        If calbtm = True Then
            calBTRM Flex.ROW
        End If
    End If
    If flexBTRM.Rows > 7 Then
        If flexBTRM.ROW Mod 5 = 0 And flexBTRM.ROW <> 0 Then
            flexBTRM.TopRow = 5
        End If
    End If
    calADLS
End Sub
Private Function CHKSAVEDATA() As Boolean
  Dim CHKRS As New ADODB.Recordset
  Set CHKRS = New ADODB.Recordset
  'Credit A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from ACCMST WHERE NAME='" & TXTCRAC.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Credit A/c Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Debit A/c Code
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Debit A/c Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Delivery Party Name
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from REFMST WHERE NAME='" & TXTDLPTY.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Delivery Party Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Agent Name
  If CHKRS.State = 1 Then CHKRS.Close
  CHKRS.Open "SELECT * from REFMST WHERE NAME='" & TXTBRNM.Text & "'", CN, adOpenKeyset, adLockPessimistic
  If CHKRS.EOF Then
     MsgBox "Agent Name Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  'Sale Tax Catagoery
  If CHKRS.State = 1 Then CHKRS.Close
  If TXTTAXNAM.Enabled = True Then
    CHKRS.Open "SELECT * from REFMST WHERE NAME='" & TXTTAXNAM.Text & "'", CN, adOpenKeyset, adLockPessimistic
    If CHKRS.EOF Then
       MsgBox "Tax Catagoery Name Not Define ", vbCritical
       CHKSAVEDATA = False
       Exit Function
    End If
  End If
  'Transporter
  If TXTTRNM.Enabled = True Then
    If CHKRS.State = 1 Then CHKRS.Close
    CHKRS.Open "SELECT * from REFMST WHERE NAME='" & TXTTRNM.Text & "'", CN, adOpenKeyset, adLockPessimistic
    If CHKRS.EOF Then
       MsgBox "Transporter Name Not Define ", vbCritical
       CHKSAVEDATA = False
       Exit Function
    End If
  End If
  'Retail / Tax Catagoery
  If Trim(TXTRTORTAX) = Empty Then
     MsgBox "Retail/Tax Invoice Not Define ", vbCritical
     CHKSAVEDATA = False
     Exit Function
  End If
  Dim I As Double
  Dim sal_icod As String
  For I = 1 To Flex.Rows - 1
     sal_icod = Flex.TextMatrix(I, 11)
     If CHKRS.State = 1 Then CHKRS.Close
     CHKRS.Open "select * from itmmst where code='" & sal_icod & "'", CN, adOpenKeyset, adLockPessimistic
     If CHKRS.EOF Then
        MsgBox "Item Missing From Master !!! ", vbCritical
        CHKSAVEDATA = False
        Exit Function
     End If
     If Not IsNumeric(Flex.TextMatrix(I, 7)) Then
        MsgBox "Invalid No of Bags"
        CHKSAVEDATA = False
        Flex.SetFocus
        Exit Function
     End If
     
     If Not IsNumeric(Flex.TextMatrix(I, 8)) Then
        MsgBox "Invalid Quanity"
        CHKSAVEDATA = False
        Exit Function
     End If
     
     If Not IsNumeric(Flex.TextMatrix(I, 9)) Then
        MsgBox "Invalid Rate"
        CHKSAVEDATA = False
        Flex.SetFocus
        Exit Function
     End If
     
     If Not IsNumeric(Flex.TextMatrix(I, 10)) Then
        MsgBox "Invalid Amount"
        CHKSAVEDATA = False
        Flex.SetFocus
        Exit Function
     End If
     
     If Round(Val(Flex.TextMatrix(I, 10)), 0) = Round(Val(Flex.TextMatrix(I, 8)) * Val(Flex.TextMatrix(I, 9)), 0) Then
        'O.k
       Else
        MsgBox "Invalid Amount"
        CHKSAVEDATA = False
        Flex.SetFocus
        Exit Function
     End If
  Next
  CHKSAVEDATA = True
End Function
Private Sub CMDSAVE_Click()
  On Error GoTo LAST
  If Val(TXTTQTY) > Val(LBLDO) Then
    MsgBox "Despatch Quanity is more than the D.O. Quanity"
    Exit Sub
  End If
  If CHKSAVEDATA = False Then
    Exit Sub
  End If
  Dim I As Double
  
  For I = 1 To Flex.Rows - 1
   If (CDate(Flex.TextMatrix(I, 2))) > (TXTVBDT) Then
     MsgBox "Bill Date must be grater then challan date"
     TXTVBDT.SetFocus
     Exit Sub
   End If
  Next
  'If TXTVBDT > MIN_DAT Then
  '   MsgBox "Bill Date Should Grater then Challan Date "
  '   TXTVBDT.SetFocus
  '   Exit Sub
  'End If
  'If TXTLRDT > MIN_DAT Then
  '   MsgBox "L.R date should be grater than or equal to  challan date", vbInformation
  '   TXTLRDT.SetFocus
  '   Exit Sub
  'End If
  If m_srno = Empty Then
    'Genrate Sr. No.
    m_srno = pubGenSrNoBILL(TXTVBDT, "SAL")
  End If
  If saveflag = True Then
    Call GENVBNOONLINE
  End If
  Dim SAVDAT As ADODB.Recordset
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' AND VBNO='" & TXTVBNO & "' AND DBCD='" & m_dbcd & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    If SAVDAT!srno = m_srno Then
     Else
      MsgBox "Duplicate Bill No.Make Change in daybok for Invoice No."
      CMDSAVE.SetFocus
      Exit Sub
    End If
  End If
  Call SAVERECSAL
  Call UPDAILYSTATUS
  If saveflag = True Then
    MsgBox "Your Invoice No. is " + TXTVBNO.Text
  End If
  Call cmdCancel_Click
  If zoomflag = True Then
    Call cmdexit_Click
    Exit Sub
  End If
  Call TXTPRTM_GotFocus
  Call TXTRMTM_GotFocus
  Exit Sub
LAST:
  MsgBox Err.Description
  
End Sub
Private Sub GENVBNOONLINE()
    Dim BOK_GEN As Boolean
    Dim m_sfx
    Dim m_pfx
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select BSFX,gpnr,VBNO from daybok where comp='" & compPth & "' AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' and vtyp='SAL' AND NAME='" & SALBOK & "'", CN, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
      Unload Me
    End If
    If Not IsNull(rs!gpnr) Then
        If rs!gpnr = "Y" Then
            BOK_GEN = False
        Else
            BOK_GEN = True
        End If
    Else
        BOK_GEN = True
    End If
    m_sfx = Trim(rs!bsfx & "")
    m_pfx = ""
    M_EXCISABLE = "N"
    If BOK_GEN = True Then
       If rs.State = 1 Then rs.Close
       rs.Open "select BSFX,gpnr,VBNO from daybok where comp='" & compPth & "' AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' and vtyp='SAL' AND NAME='" & SALBOK & "'"
       m_sfx = Trim(rs!bsfx & "")
       TXTVBNO = GENVBNO("Select * from DAYBOK where COMP='" & compPth & "' AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' AND [NAME]='" & SALBOK & "'", m_sfx, m_pfx)
      Else
       If rs.State = 1 Then rs.Close
       rs.Open "SELECT * FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", CN, adOpenDynamic, adLockOptimistic
       If Not rs.EOF Then
         TXTVBNO = genVBNoEXC("Select * from UNTCFG where COMP='" & compPth & "' AND UNIT='" & UNCD & "'", "", "")
         M_EXCISABLE = "Y"
       End If
    End If
    TXTCOMINV = GENVBNO("Select * from DAYBOK where COMP='" & compPth & "' AND UNIT='" & UNCD & "'  AND DVCD='" & DIVCOD & "' AND [NAME]='" & SALBOK & "'", m_sfx, m_pfx)
End Sub
Private Sub SAVERECSAL()
  On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Dim M_CRAC As String
  Dim M_DRAC As String
  Dim M_PCOD As String
  Dim M_DCOD As String
  Dim M_CPCD As String
  Dim M_ARCD As String
  Dim M_TRCD As String
  Dim M_TXCD As String
  Dim M_BRCD As String
  Dim I As Double
  Dim j As Double
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  'Credit A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTCRAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  M_CRAC = SAVDAT!CODE & ""
  'Debit A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTDBAC.Text & "'", CN, adOpenDynamic, adLockOptimistic
  M_DRAC = SAVDAT!CODE & ""
  M_PCOD = SAVDAT!CODE & ""
  M_CPCD = SAVDAT!CPCD & ""
  M_ARCD = SAVDAT!ARCD & ""
  
  'Delivery Party
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTDLPTY.Text & "'", CN, adOpenDynamic, adLockOptimistic
  M_DCOD = SAVDAT!CODE & ""
  'Agent
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTBRNM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  M_BRCD = SAVDAT!CODE & ""
  'Transporter Code
  If TXTTRNM.Enabled = True Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTTRNM.Text & "'", CN, adOpenDynamic, adLockOptimistic
  
    M_TRCD = SAVDAT!CODE & ""
   Else
    M_TRCD = Empty
  End If
  'Tax Catagoery
  If Trim(TXTTAXNAM.Text) <> "" Then
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM REFMST WHERE NAME='" & TXTTAXNAM.Text & "'", CN, adOpenDynamic, adLockOptimistic
    M_TXCD = SAVDAT!CODE & ""
   Else
    M_TXCD = Empty
  End If
  CN.BeginTrans
  Call DELETESAL
  SQL = Empty
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
  If SAVDAT.EOF Then
      SAVDAT.AddNew
  End If
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "SAL"
  SAVDAT!srno = m_srno
  SAVDAT!SRCH = 1
  SAVDAT!dbcd = m_dbcd
  SAVDAT!Date = Format(TXTVBDT.Value, "YYYY/MM/DD")
  SAVDAT!VBNO = Trim(TXTVBNO.Text)
  SAVDAT!CVBN = Trim(TXTCOMINV.Text)
  SAVDAT!CRAC = M_CRAC
  SAVDAT!DRAC = M_DRAC
  SAVDAT!pcod = M_PCOD
  SAVDAT!DCOD = M_DCOD
  SAVDAT!BRCD = M_BRCD
  SAVDAT!CPCD = M_CPCD
  SAVDAT!ARCD = M_ARCD
  SAVDAT!TXCD = M_TXCD
  SAVDAT!TPCS = 0
  SAVDAT!tQty = 0
  SAVDAT!ITOT = Val(TXTITOT.Text)
  SAVDAT!BADJ = Val(txtBNET.Text) - Val(TXTITOT.Text)
  SAVDAT!BNET = Val(txtBNET.Text)
  SAVDAT!ttyp = Trim(TXTRTORTAX.Text)
  SAVDAT!CDAY = Val(TXTCRDS)
  SAVDAT!unit = UNCD
  If saveflag = True Then
    SAVDAT!SYSR = "N"
   Else
    SAVDAT!SYSR = "U"
  End If
  SAVDAT![User] = cUName & ""
  SAVDAT!DVCD = DIVCOD
  SAVDAT!unit = UNCD
  SAVDAT!TRCD = M_TRCD
  SAVDAT!LRNO = Trim(TXTLRNO.Text)
  SAVDAT!LRDT = Format(TXTLRDT.Value, "YYYY/MM/DD")
  SAVDAT!VHCL = Trim(TXTVHCL)
  SAVDAT!RECSTAT = "A"
  SAVDAT!PRTM = Trim(TXTPRTM.Text)
  SAVDAT!RMTM = Trim(TXTRMTM.Text)
  SAVDAT!BRMK = Trim(TXTRMRK.Text)
  I = 0
  For I = 0 To flexBTRM.Rows - 1
    j = 0
    For j = 0 To SAVDAT.Fields.Count - 1
      If Trim(SAVDAT.Fields(j).Name) = Trim(flexBTRM.TextMatrix(I, 0)) Then
        SAVDAT.Fields(j).Value = Val(flexBTRM.TextMatrix(I, 2))
      End If
      If Trim(SAVDAT.Fields(j).Name) = "PER" & Trim(flexBTRM.TextMatrix(I, 0)) Then
        SAVDAT.Fields(j).Value = Val(flexBTRM.TextMatrix(I, 1))
      End If
    Next
  Next
  Dim K As Double
  K = 1
  SAVDAT.Update
  If SAVDAT.State = 1 Then SAVDAT.Close
  I = 1
  For I = 1 To Flex.Rows - 1
    If SAVDAT.State = 1 Then SAVDAT.Close
    SAVDAT.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "' AND SRCH='" & I & "'", CN, adOpenDynamic, adLockOptimistic
    'Add Records for Sale Data
    If SAVDAT.EOF Then
      SAVDAT.AddNew
    End If
    Dim PKGCOD As String
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM REFMST WHERE NAME ='" & Flex.TextMatrix(I, 6) & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      PKGCOD = MSTDAT!CODE
     Else
      PKGCOD = Empty
    End If
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "SAL"
    SAVDAT!srno = m_srno
    SAVDAT!SRCH = I
    SAVDAT!unit = UNCD
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!CHLN = Flex.TextMatrix(I, 1)
    If IsDate(Flex.TextMatrix(I, 2)) = True Then
      SAVDAT!CHDT = Flex.TextMatrix(I, 2)
    End If
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = m_dbcd
    SAVDAT!CRAC = M_CRAC
    SAVDAT!DRAC = M_DRAC
    SAVDAT!pcod = M_PCOD
    SAVDAT!DCOD = M_DCOD
    SAVDAT!ICOD = Flex.TextMatrix(I, 11)
    SAVDAT!PCES = Val(Flex.TextMatrix(I, 7))
    SAVDAT!QNTY = Val(Flex.TextMatrix(I, 8))
    SAVDAT!GWGT = Val(Flex.TextMatrix(I, 8)) / 1000
    SAVDAT!twgt = 0
    SAVDAT!Rate = Val(Flex.TextMatrix(I, 9))
    SAVDAT!AMNT = Val(Flex.TextMatrix(I, 10))
    SAVDAT!QORP = "Q"
    SAVDAT!User = cUName
    SAVDAT!SYSR = "N"
    SAVDAT!OPER = "-"
    SAVDAT!DVCD = DIVCOD
    SAVDAT!GRAD = Flex.TextMatrix(I, 5)
    SAVDAT!ltno = Flex.TextMatrix(I, 4)
    SAVDAT!MRGN = ""
    SAVDAT!COPS = 0
    SAVDAT!TWST = "0"
    SAVDAT!RTYP = "SAL"
    SAVDAT!RSRN = m_srno
    SAVDAT!RSRC = I
    SAVDAT!RECSTAT = "A"
    SAVDAT!SDBC = m_dbcd
    SAVDAT!SVBN = TXTVBNO
    SAVDAT!SHCD = PKGCOD
    SAVDAT.Update
  Next
  If MSTDAT.State = 1 Then MSTDAT.Close
  MSTDAT.Open "SELECT ISNULL(SUM(PCES),0) AS TPCS,ISNULL(SUM(QNTY),0) AS TQTY FROM SPTRAN WHERE COMP='" & compPth & "' AND RTYP='SAL' AND RSRN='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
  If Not MSTDAT.EOF Then
    CN.Execute "UPDATE BILLMAIN SET TPCS='" & MSTDAT!TPCS & "', TQTY='" & MSTDAT!tQty & "' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  End If
  'UPDATE SPTRAN AND SPMAIN
  I = 1
  For I = 1 To Flex.Rows - 1
    CN.Execute "UPDATE SPMAIN SET LRNO='" & TXTLRNO & "', LRDT='" & Format(TXTLRDT, "MM/DD/YYYY") & "',TRCD='" & M_TRCD & "',VHCL='" & TXTVHCL & "'  WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  Next
  'Add Records In Tranman
  Dim BOK_TOT As Double
  BOK_TOT = 0
  BOK_TOT = Val(txtBNET)
  I = 0
  For I = 0 To flexBTRM.Rows - 1
    If M_POSTYESNO(I) = "Y" Then
      If M_OPER(I) = "+" Then
        BOK_TOT = BOK_TOT - Val(flexBTRM.TextMatrix(I, 2))
       Else
        BOK_TOT = BOK_TOT + Val(flexBTRM.TextMatrix(I, 2))
      End If
    End If
  Next
  'Post of day Book
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  If FRMPARA = "SAL" Then
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = FRMPARA
    SAVDAT!srno = m_srno
    SAVDAT!SRCH = 1
    SAVDAT![User] = cUName
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!ACOD = M_CRAC
    SAVDAT!RCOD = M_DRAC
    SAVDAT!damt = 0
    SAVDAT!camt = BOK_TOT
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
    SAVDAT!AMNT = BOK_TOT
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
   Else
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = FRMPARA
    SAVDAT!srno = m_srno
    SAVDAT!SRCH = 1
    SAVDAT![User] = cUName
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!ACOD = M_CRAC
    SAVDAT!RCOD = M_DRAC
    SAVDAT!damt = BOK_TOT
    SAVDAT!camt = 0
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
    SAVDAT!AMNT = BOK_TOT
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
  End If
  SAVDAT.Update
  'Post of Party A/c
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM TRNMAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  If FRMPARA = "SAL" Then
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = FRMPARA
    SAVDAT!srno = m_srno
    SAVDAT!SRCH = 2
    SAVDAT![User] = cUName
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!ACOD = M_DRAC
    SAVDAT!RCOD = M_CRAC
    SAVDAT!damt = Val(txtBNET)
    SAVDAT!camt = 0
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
    SAVDAT!AMNT = Val(txtBNET)
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
   Else
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = FRMPARA
    SAVDAT!srno = m_srno
    SAVDAT!SRCH = 2
    SAVDAT![User] = cUName
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!ACOD = M_DRAC
    SAVDAT!RCOD = M_CRAC
    SAVDAT!damt = 0
    SAVDAT!camt = Val(txtBNET)
    SAVDAT!VBNO = TXTVBNO
    SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
    SAVDAT!AMNT = Val(txtBNET)
    SAVDAT!DVCD = DIVCOD
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
  End If
  Dim TRNSRCH As Double
  TRNSRCH = 2
  SAVDAT.Update
  'Post of config if postyes="Y"
  I = 0
  For I = 0 To flexBTRM.Rows - 1
    If M_POSTYESNO(I) = "Y" Then

      TRNSRCH = TRNSRCH + 1
      If FRMPARA = "SAL" Then
         If M_OPER(I) = "+" Then
            If Val(flexBTRM.TextMatrix(I, 2)) <> 0 Then
              SAVDAT.AddNew
              SAVDAT!COMP = compPth
              SAVDAT!VTYP = FRMPARA
              SAVDAT!srno = m_srno
              SAVDAT!SRCH = TRNSRCH
              SAVDAT![User] = cUName
              SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
              SAVDAT!ACOD = M_POSTCOD(I)
              SAVDAT!RCOD = M_DRAC
              SAVDAT!damt = 0
              SAVDAT!camt = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!VBNO = TXTVBNO
              SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
              SAVDAT!AMNT = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!DVCD = DIVCOD
              SAVDAT!unit = UNCD
              SAVDAT!RECSTAT = "A"
              SAVDAT.Update
            End If
          Else
            If Val(flexBTRM.TextMatrix(I, 2)) <> 0 Then
              SAVDAT.AddNew
              SAVDAT!COMP = compPth
              SAVDAT!VTYP = FRMPARA
              SAVDAT!srno = m_srno
              SAVDAT!SRCH = TRNSRCH
              SAVDAT![User] = cUName
              SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
              SAVDAT!ACOD = M_POSTCOD(I)
              SAVDAT!RCOD = M_DRAC
              SAVDAT!damt = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!camt = 0
              SAVDAT!VBNO = TXTVBNO
              SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
              SAVDAT!AMNT = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!DVCD = DIVCOD
              SAVDAT!unit = UNCD
              SAVDAT!RECSTAT = "A"
              SAVDAT.Update
            End If
         End If
        Else
         If M_OPER(I) = "+" Then
            If Val(flexBTRM.TextMatrix(I, 2)) <> 0 Then
              SAVDAT.AddNew
              SAVDAT!COMP = compPth
              SAVDAT!VTYP = FRMPARA
              SAVDAT!srno = m_srno
              SAVDAT!SRCH = TRNSRCH
              SAVDAT![User] = cUName
              SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
              SAVDAT!ACOD = M_POSTCOD(I)
              SAVDAT!RCOD = M_DRAC
              SAVDAT!damt = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!camt = 0
              SAVDAT!VBNO = TXTVBNO
              SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
              SAVDAT!AMNT = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!DVCD = DIVCOD
              SAVDAT!unit = UNCD
              SAVDAT!RECSTAT = "A"
              SAVDAT.Update
            End If
          Else
            If Val(flexBTRM.TextMatrix(I, 2)) <> 0 Then
              SAVDAT.AddNew
              SAVDAT!COMP = compPth
              SAVDAT!VTYP = FRMPARA
              SAVDAT!srno = m_srno
              SAVDAT!SRCH = TRNSRCH
              SAVDAT![User] = cUName
              SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
              SAVDAT!ACOD = M_POSTCOD(I)
              SAVDAT!RCOD = M_DRAC
              SAVDAT!damt = 0
              SAVDAT!camt = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!VBNO = TXTVBNO
              SAVDAT!narr = "Sale Invoice No. : " + TXTVBNO
              SAVDAT!AMNT = Val(flexBTRM.TextMatrix(I, 2))
              SAVDAT!DVCD = DIVCOD
              SAVDAT!unit = UNCD
              SAVDAT!RECSTAT = "A"
              SAVDAT.Update
            End If
         End If
      End If
    End If
  Next
  'Add Records In DoTRAN
  I = 1
  Dim ORDN As String
  Dim BSCRAT As Double
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM DOTRAN WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
  SAVDAT.AddNew
  SAVDAT!COMP = compPth
  SAVDAT!VTYP = "SAL"
  SAVDAT!srno = m_srno
  SAVDAT!DVCD = DIVCOD
  SAVDAT!unit = UNCD
  SAVDAT!TXRT = TXTRTORTAX.Text
  SAVDAT!DONO = Trim(SEL_DOS_SRN)
  SAVDAT!DODT = Format(TXTVBDT, "MM/DD/YYYY")
  SAVDAT!pcod = M_PCOD
  SAVDAT!DCOD = M_DCOD
  SAVDAT!BRCD = M_BRCD
  SAVDAT!SCOD = SEL_SCOD
  SAVDAT!ICOD = FIL_ITM_COD
  SAVDAT!PKTP = FIL_PKGCOD
  SAVDAT!GRAD = FIL_GRADE
  SAVDAT!QNTY = Val(TXTTQTY)
  SAVDAT!GWGT = Val(TXTTQTY) / 1000
  SAVDAT!Rate = SEL_RATE
  SAVDAT!ARAT = SEL_RATE
  SAVDAT!BRMK = ""
  SAVDAT!PRDL = "N"
  SAVDAT!DFLG = "N"
  SAVDAT!ORDN = SEL_ORDN
  SAVDAT!OSRC = SEL_OSRC
  SAVDAT!CHLN = TXTVBNO
  SAVDAT!VBNO = TXTVBNO
  SAVDAT!RECSTAT = "A"
  SAVDAT.Update
  'update last bill no in DAYBOK
  If saveflag = True Then
    Dim x
    Dim sav_VBNO As String
    If Len(Trim(sav_srfx)) > 0 Then
      sav_VBNO = Mid(TXTCOMINV, Len(Trim(sav_srfx)) + 1, 5)
     Else
      sav_VBNO = TXTCOMINV
    End If
    CN.Execute "update DAYBOK set [VBNO]='" & sav_VBNO & "' where NAME ='" & SALBOK & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCOD & "' "
    If M_EXCISABLE = "Y" Then
      CN.Execute "update UNTCFG set exno='" & TXTVBNO & "' where comp='" & compPth & "'"
    End If
  End If
  Dim REC_AMT As Double
  Dim DBN_AMT As Double
  Dim CRN_AMT As Double
  Dim RET_AMT As Double
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT ISNULL(SUM(RAMT),0) AS RAMT, ISNULL(SUM(DBNA),0) AS DBNA, ISNULL(SUM(CRNA),0) AS CRNA, ISNULL(SUM(RETG),0) AS RETG FROM RPTRAN WHERE COMP='" & compPth & "' AND BSR1='SAL' AND BSR2='" & m_srno & "'", CN, adOpenDynamic, adLockOptimistic
  If Not SAVDAT.EOF Then
    CN.Execute "UPDATE BILLMAIN SET RAMT='" & SAVDAT!RAMT & "',DBNA='" & SAVDAT!DBNA & "',CRNA='" & SAVDAT!CRNA & "',RETG='" & SAVDAT!RETG & "' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  End If
  CN.CommitTrans
  Exit Sub
LAST:
 'Resume
 MsgBox Err.Description
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub
Private Sub DELETESAL()
  Dim SAVDAT As New ADODB.Recordset
  Dim m_rtyp As String
  Dim m_rsrn As String
  Set SAVDAT = New ADODB.Recordset
  If SAVDAT.State = 1 Then SAVDAT.Close
  CN.Execute "UPDATE SPTRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  CN.Execute "UPDATE BILLMAIN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  CN.Execute "UPDATE TRNMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  CN.Execute "UPDATE EGPMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  CN.Execute "DELETE FROM DOTRAN WHERE COMP='" & compPth & "' AND VTYP='SAL' AND SRNO='" & m_srno & "'"
  'For Reco. check here
End Sub
Private Sub TXTVBDT_LostFocus()
  'If TXTVBDT < MIN_DAT Then
  '  MsgBox "Bill date should be grater than challan date", vbInformation
  '  TXTVBDT.SetFocus
  '  Exit Sub
  'End If
End Sub
Private Sub TXTVBDT_Validate(cancel As Boolean)
  'If TXTVBDT < MIN_DAT Then
  '  Cancel = True
  'End If
End Sub
Private Sub TXTVHCL_GotFocus()
  TXTVHCL.SelStart = 0
  TXTVHCL.SelLength = Len(TXTVHCL)
End Sub
Private Sub CHKFLEX()
  CHK_FLX = True
  
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim CHKITM As String
  Dim FLXR As Double
  For FLXR = 1 To Flex.Rows - 1
    If Val(Flex.TextMatrix(Flex.ROW, 16)) < Val(Flex.TextMatrix(Flex.ROW, 7)) Then
      MsgBox "Stock Is only : " + nstr(Val(Flex.TextMatrix(Flex.ROW, 16)), 8, 0)
      FLXROW = FLXR
      FLXCOL = 7
      CHK_FLX = False
      Exit For
    End If
    If Trim(Flex.TextMatrix(FLXR, 7)) = Empty Then Exit For
    'If Not Trim(FLEX.TextMatrix(FLXR, 2)) = "" Then
    '  If Not IsDate(CDate(FLEX.TextMatrix(FLXR, 2))) Then
    '    CHK_FLX = False
    '    FLXROW = FLXR
    '    FLXCOL = 2
    '    Exit For
    '  End If
    ' Else
    '  CHK_FLX = False
    '  FLXROW = FLXR
    '  FLXCOL = 2
    '  Exit For
    'End If
    CHKITM = Flex.TextMatrix(FLXR, 11)
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ITMMST WHERE CODE='" & CHKITM & "'", CN, adOpenDynamic, adLockOptimistic
    If MSTDAT.EOF Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 3
       Exit For
    End If

    If Not IsNumeric(Flex.TextMatrix(FLXR, 7)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 7
       Exit For
    End If
    If Not IsNumeric(Flex.TextMatrix(FLXR, 8)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 8
       Exit For
    End If
    If Not IsNumeric(Flex.TextMatrix(FLXR, 9)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 9
       Exit For
    End If
    If Not IsNumeric(Flex.TextMatrix(FLXR, 10)) Then
       CHK_FLX = False
       FLXROW = FLXR
       FLXCOL = 10
       Exit For
    End If
    If Val(TXTTQTY) > Val(LBLDO) Then
      MsgBox "Sale Quantity is more than the D.O. Quanity"
      FLXROW = FLXR
      FLXCOL = 7
      Exit Sub
    End If
  Next
End Sub
Private Sub UPDELSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "SAL"
  DLYSTA!pcod = TXTDBAC
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(txtBNET)
  DLYSTA!CUSR = cUName
  DLYSTA!ACTN = "D"
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub
Private Sub UPDAILYSTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE CUSR='JNMOPAWCGBDSXZAS'", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "SAL"
  DLYSTA!pcod = TXTDBAC
  DLYSTA!dbcd = ""
  DLYSTA!QNTY = Val(TXTTQTY)
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = Val(txtBNET)
  DLYSTA!CUSR = cUName
  If saveflag = True Then
    DLYSTA!ACTN = "N"
   Else
    DLYSTA!ACTN = "E"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub
