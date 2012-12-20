VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmLumpSumDispatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Against Delivery Order"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleMode       =   0  'User
   ScaleWidth      =   11355
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12197
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
      Begin VB.ComboBox cmbPackingType 
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
         ItemData        =   "frmLumpSumDispatch.frx":0000
         Left            =   2040
         List            =   "frmLumpSumDispatch.frx":0002
         TabIndex        =   7
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Frame frmDTRNGE 
         Height          =   1200
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   11130
         Begin VB.ComboBox cmbSelection 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   1
            Tag             =   "0"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TXTSEARCH 
            Height          =   285
            Left            =   5160
            MaxLength       =   30
            TabIndex        =   4
            Top             =   720
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker txtToDate 
            Height          =   330
            Left            =   7680
            TabIndex        =   3
            Top             =   285
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
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
            Format          =   50266113
            CurrentDate     =   38429
         End
         Begin MSComCtl2.DTPicker txtFrDate 
            Height          =   330
            Left            =   5160
            TabIndex        =   2
            Top             =   285
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
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
            Format          =   50266113
            CurrentDate     =   38429
         End
         Begin WelchButton.lvButtons_H cmdGo 
            Height          =   615
            Left            =   9480
            TabIndex        =   5
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1085
            Caption         =   "&Filter"
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
            Image           =   "frmLumpSumDispatch.frx":0004
            cBack           =   -2147483633
         End
         Begin VB.Label LBLNAME 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   18
            Top             =   720
            Width           =   2715
         End
         Begin VB.Label lblToDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Date : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6720
            TabIndex        =   17
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblFrDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From Date : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   16
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label LBLSEARCH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Criteria :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   290
            Width           =   1395
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   10
         BackColor       =   16777215
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   7680
         TabIndex        =   9
         Top             =   6120
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
         Image           =   "frmLumpSumDispatch.frx":039E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   10320
         TabIndex        =   11
         Top             =   6120
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmLumpSumDispatch.frx":1128
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   330
         Left            =   6000
         TabIndex        =   8
         Top             =   6120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
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
         Format          =   50266113
         CurrentDate     =   38429
      End
      Begin WelchButton.lvButtons_H cmdSavePrint 
         Height          =   495
         Left            =   8880
         TabIndex        =   10
         Top             =   6120
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
         Image           =   "frmLumpSumDispatch.frx":157A
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Date :"
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
         Left            =   4680
         TabIndex        =   21
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of  Dispatch :"
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
         Left            =   120
         TabIndex        =   20
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label LBLDIV 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   19
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Lumpsum Dispatch Against Delivery Order"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   13
         Top             =   0
         Width           =   5055
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
         TabIndex        =   12
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLumpSumDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOAD As String
Public DIVCODE As String
Const strChecked = "þ"
Const strUnChecked = "q"
Dim RECSET As ADODB.Recordset
Public DIVNAME As String
Dim SIGNAL As Boolean
Dim M_DBCD As String
Dim BOX As Double
Dim COPS As Double

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
  If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim INDEX As Long
Dim FLAG As Boolean
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double

With MSFlexGrid1
For INDEX = 1 To MSFlexGrid1.Rows - 1
   If .TextMatrix(INDEX, 0) = strChecked And Val(.TextMatrix(INDEX, 9)) > 0 Then
      .ROW = INDEX
      FLAG = True
      Exit For
   End If
Next INDEX

If FLAG = False Then Exit Sub
   
Call SetGlobal

SLIP = GenDPFVNO("DPF", M_DBCD, DIVCODE)
 
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT * FROM ORDTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & .TextMatrix(.ROW, 18) & _
"' AND VTYP='DOS' AND DFLG<>'Y' AND RECSTAT='A' AND DOSTAT='Y' AND ORDN = '" & .TextMatrix(.ROW, 1) & _
"' AND DONO  = '" & .TextMatrix(.ROW, 2) & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

   'FOR DUPLICATE CHLLAN NO. STOP CHECKING
   Dim NSQL As String
   Dim MSGS As String: MSGS = "Unit"
   SLIP = GenDPFVNO("DPF", M_DBCD, DIVCODE)
   
   NSQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "'AND UNIT='" & UNCD & _
           "' AND VTYP='DPF' AND DBCD='" & M_DBCD & "' AND VBNO = '" & SLIP & "' "
   
   If UNT_DIVSERIES_REQ = "Y" Then
      NSQL = NSQL & " AND DVCD='" & DIVCODE & "' "
      MSGS = "Division"
   End If
   
   If RS.State Then RS.Close
   RS.Open NSQL, CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      MsgBox "Lumpsum Challan No. " & SLIP & " Already Exist. Check Last No. In " & MSGS & " Configuration", vbCritical
      Exit Sub
   End If
   RS.Close
   '===================================================================

CN.BeginTrans

SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD,QNTY,RATE,ARAT,"
SQL = SQL & "ORDN,OSRC,BRMK,PRDL,ICOD,TXRT,TXCD,RTCD,SLIP,SLIPDATE,RDBC,DELQNTY) VALUES ('" & compPth & "','" & UNCD & _
"','" & DIVCODE & "','DPF','" & M_DBCD & "','" & .TextMatrix(.ROW, 2) & "','" & Format(RECSET!DODt, "YYYY/MM/DD") & _
"','" & RECSET!PCOD & "','" & RECSET!DCOD & "','" & RECSET!SRCH & _
"','" & RECSET!BRCD & "','" & RECSET!ltno & "','" & RECSET!grad & "','" & RECSET!SUBGRD & "','" & .TextMatrix(.ROW, 9) & _
"'," & RECSET!RATE & "," & RECSET!ARAT & ",'" & RECSET!ORDN & _
"','1','','','" & RECSET!ICOD & "','" & RECSET!TXRT & "','" & RECSET!TXCD & "','" & RECSET!SUBGRD & _
"','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & M_DBCD & "','" & .TextMatrix(.ROW, 9) & "')"

CN.Execute SQL

Call FindDetails(Val(.TextMatrix(.ROW, 9)))

SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,"
SQL = SQL & "DCOD,SRCH,LTNO,ICOD,GRAD,PCES,QNTY,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
SQL = SQL & "RECSTAT,COPS,EXTRA1,EXTRA2,EXTRA3)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
"','DPF','" & M_DBCD & "','" & SLIP & "','" & SLIP & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
"','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & RECSET!PCOD & "','" & RECSET!BRCD & _
"','" & RECSET!DCOD & "','" & RECSET!SRCH & "','" & RECSET!ltno & "','" & RECSET!ICOD & "','" & RECSET!grad & _
"','" & BOX & "','" & .TextMatrix(.ROW, 9) & "'," & RECSET!RATE & "," & (Val(.TextMatrix(.ROW, 9)) * Val(RECSET!RATE)) & _
",'Q','N','" & cUName & "','-','A','" & (COPS * BOX) & "','" & RECSET!ORDN & _
"','" & .TextMatrix(.ROW, 2) & "','" & .TextMatrix(.ROW, 18) & "')"

CN.Execute SQL

SQL = "INSERT INTO PKGMAN (COMP,UNIT,DVCD,DBCD,VTYP,SRNO,SRCH,DATE,SLIPNO,PKG_STCOD,"
SQL = SQL & "LOTNO,FINITMCOD,GRAD,SUBGRAD,QNTY,SYSR,[USER],OPER,RECSTAT) VALUES "
SQL = SQL & "('" & compPth & "','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','DPF',"
SQL = SQL & "'1','1','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SLIP & "','000000',"
SQL = SQL & "'" & .TextMatrix(.ROW, 4) & "','" & .TextMatrix(.ROW, 15) & _
"','" & .TextMatrix(.ROW, 16) & "','" & .TextMatrix(.ROW, 17) & "','" & .TextMatrix(.ROW, 9) & _
"','N','" & cUName & "','-','A')"

CN.Execute SQL

Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & SLIP & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
         "' AND VTYP='DPF' AND CODE='" & M_DBCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
End If
 
CN.Execute UPSQL

SQL = "UPDATE ORDTRN SET DFLG ='Y' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND DBCD='" & .TextMatrix(.ROW, 18) & "' AND VTYP='DOS' AND "
SQL = SQL & "DFLG<>'Y' AND RECSTAT='A' AND DOSTAT='Y' AND ORDN = '" & .TextMatrix(.ROW, 1) & _
"' AND DONO  = '" & .TextMatrix(.ROW, 2) & "'"

CN.Execute SQL

SQL = "UPDATE ORDMAN SET LDSPDAT='" & Format(TXTVBDT, "MM/DD/YYYY") & "',DISPATCHQTY = DISPATCHQTY + " & Val(Trim(.TextMatrix(.ROW, 9))) & ",DOQTY = DOQTY - " & Val(Trim(.TextMatrix(.ROW, 8))) & " WHERE COMP='" & compPth & _
"' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & "' AND DBCD='" & .TextMatrix(.ROW, 18) & _
"' AND ORDN = '" & .TextMatrix(.ROW, 1) & "' AND ICOD = '" & RECSET!ICOD & "' AND TRCD='" & RECSET!grad & "'"

CN.Execute SQL

CN.CommitTrans

MsgBox "Your Challan No. is : " & SLIP
Call cmdGo_Click

End With

Exit Sub
LAST:
MsgBox ERR.Description
Exit Sub
End Sub

Private Sub cmdSavePrint_Click()
Call cmdSave_Click
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)

  M_DESC = Empty
  Key = Empty
  NEW_VISIBLE = False
  DIVCODE = Empty
  DIVNAME = Empty
  If DIVCODE = Empty Then
    DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A'  AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
    LBLDIV.Caption = "DIVISION : " + DIVNAME
    
  If PackingType(Key) = "C" Then MsgBox "Division Not Allowed Lumpsum Packing.Check Configuration": LOAD = "N": GoTo JUMP
  
  Call SetFlex
  
  cmbSelection.Clear
  cmbSelection.AddItem ("OrderNo.Wise")
  cmbSelection.AddItem ("DONO.Wise")
  cmbSelection.AddItem ("DateWise")
  cmbSelection.AddItem ("AgentWise")
  cmbSelection.AddItem ("A/c PartyWise")
  cmbSelection.AddItem ("ItemWise")
  cmbSelection.AddItem ("LotNo.Wise")
  cmbSelection.ListIndex = 0
  
  Call SetDispatchType
  TXTVBDT = Now
  txtFrDate.Value = GetMinDate
  txtToDate.Value = GetMaxDate
JUMP:
End Sub

Private Sub CMBSELECTION_Click()
SendKeys "{HOME}"
TXTSEARCH = Empty
Select Case cmbSelection.Text
Case "OrderNo.Wise"
      lblName.Caption = "OrderNo. : "
Case "DONO.Wise"
      lblName.Caption = "DONO : "
Case "DateWise"
      lblName.Caption = ""
Case "AgentWise"
      lblName.Caption = "Select Agent Name : "
Case "A/c PartyWise"
      lblName.Caption = "Select A/C Party Name : "
Case "ItemWise"
      lblName.Caption = "Select Item Name"
Case "LotNo.Wise"
      lblName.Caption = "Enter Lotno."
End Select
End Sub

Private Sub cmbSelection_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If cmbSelection.Text = "DateWise" Then
      SendKeys "{TAB}"
   Else
      TXTSEARCH.Enabled = True: TXTSEARCH.SetFocus
   End If
 End If
If KeyCode = vbKeyDelete Then
KeyCode = 0
End If
 
End Sub

Private Sub cmbSelection_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdGo_Click()

If cmbSelection.Text <> "DateWise" And TXTSEARCH = Empty Then FillList: Exit Sub

Select Case cmbSelection.Text
Case "OrderNo.Wise"
      Call FillList(" AND ORDTRN.ORDN ='" & Trim(TXTSEARCH) & "'")
Case "DONO.Wise"
      Call FillList(" AND ORDTRN.DONO ='" & Trim(TXTSEARCH) & "'")
Case "DateWise"
      Call FillList(" AND ORDTRN.DODT >='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND ORDTRN.DODT <='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'")
Case "AgentWise"
      Call FillList(" AND ORDTRN.BRCD ='" & Trim(TXTSEARCH.Tag) & "'")
Case "A/c PartyWise"
      Call FillList(" AND ORDTRN.PCOD ='" & Trim(TXTSEARCH.Tag) & "'")
Case "ItemWise"
      Call FillList(" AND ORDTRN.ICOD ='" & Trim(TXTSEARCH.Tag) & "'")
Case "LotNo.Wise"
      Call FillList(" AND ORDTRN.LTNO ='" & Trim(TXTSEARCH) & "'")
End Select
End Sub

Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.Rows <= 1 Then Exit Sub

With MSFlexGrid1
If .COL = 0 Then
If .TextMatrix(.ROW, .COL) = strChecked Then .TextMatrix(.ROW, 9) = Empty
End If

Dim STOCK As Double

If .COL = 0 Then STOCK = FindStock
If .COL = 0 Then
    If STOCK >= Val(.TextMatrix(.ROW, 8)) Then
       SIGNAL = False
       .TextMatrix(.ROW, 9) = nstr(.TextMatrix(.ROW, 8), 12, 3)
    Else
       SIGNAL = True
       .TextMatrix(.ROW, 9) = nstr(STOCK, 12, 3)
    End If
    
 Dim ROW As Long, COL As Long
 ROW = .ROW: COL = .COL
 Call TriggerCheckbox(.ROW, .COL)
    
 End If
 
 Dim FLAG As Boolean: FLAG = False
 Dim INDEX As Long
 For INDEX = 1 To .Rows - 1
    If .TextMatrix(INDEX, 0) = strChecked Then FLAG = True: Exit For
 Next INDEX
   
   
 If FLAG Then
 For INDEX = 1 To .Rows - 1
    .TextMatrix(INDEX, 0) = strUnChecked
    If INDEX <> ROW Then .TextMatrix(INDEX, 9) = Empty
    If INDEX <> ROW Then .ROW = INDEX: .COL = 9: .CellBackColor = vbWhite
 Next INDEX
    If ROW <> 0 Then .TextMatrix(ROW, 0) = strChecked
    .ROW = ROW: .COL = COL
 Else
     .COL = 0
 End If
 
.ROW = ROW: .COL = COL
If .COL = 9 And .TextMatrix(.ROW, 9) <> Empty Then
  
End If


End With




End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 Or KeyAscii = 32) And MSFlexGrid1.COL = 0 Then
     Call MSFlexGrid1_Click
     Exit Sub
End If
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub

Private Sub TXTSEARCH_GotFocus()
TXTSEARCH.BackColor = RGB(BRED, BGREEN, BBLUE)
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then TXTSEARCH = Empty: FillList: Exit Sub
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTSEARCH = Empty) Then
Select Case cmbSelection.Text
Case "OrderNo.Wise"
      lblName.Caption = "OrderNo. : "
Case "DONO.Wise"
      lblName.Caption = "DONO : "
Case "DateWise"
      lblName.Caption = ""
Case "AgentWise"
      lblName.Caption = "Select Agent Name : "
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM REFMST WHERE CATA='B'", 0, TXTSEARCH.Text, "SELECT AGENT FROM LIST")
      TXTSEARCH.Tag = Key
Case "A/c PartyWise"
      lblName.Caption = "Select A/C Party Name : "
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM ACCMST", 0, TXTSEARCH.Text, "SELECT A/C PARTY FROM LIST")
      TXTSEARCH.Tag = Key
Case "ItemWise"
      lblName.Caption = "Select Item Name"
      NEW_VISIBLE = False
      M_DESC = Empty
      Key = Empty
      TXTSEARCH = SearchList1("SELECT TOP 20 CODE, NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, TXTSEARCH, "SELECT FINISH ITEM FROM LIST")
      TXTSEARCH.Tag = Key
Case "LotNo.Wise"
      lblName.Caption = "Enter Lotno."
End Select
End If

If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub TXTSEARCH_LostFocus()
  TXTSEARCH.BackColor = vbWhite
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub

Private Sub SetFlex()
    With MSFlexGrid1
        .Cols = 19
        '.Rows = 40
        .FixedRows = 1
        .FixedCols = 0
                                    
        .TextMatrix(0, 1) = "OrderNo"
        .TextMatrix(0, 2) = "DONO"
        .TextMatrix(0, 3) = "DO Date"
        .TextMatrix(0, 10) = "Agent"
        .TextMatrix(0, 11) = "A/C Party"
        .TextMatrix(0, 12) = "Consinee"
        .TextMatrix(0, 4) = "Lotno"
        .TextMatrix(0, 5) = "Item Desc."
        .TextMatrix(0, 6) = "Grad"
        .TextMatrix(0, 7) = "Subgrad"
        .TextMatrix(0, 8) = "DO Qnty"
        .TextMatrix(0, 9) = "Chln Qnty"
        .TextMatrix(0, 13) = "Rate"
        .TextMatrix(0, 14) = "Remarks"
                      
        .TextMatrix(0, 15) = "ICOD"
        .TextMatrix(0, 16) = "GRAD"
        .TextMatrix(0, 17) = "SUBGRD"
                      
        .Rows = .Rows - 1
       
        .ColWidth(0) = 350
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1000
        .ColWidth(10) = 1200
        .ColWidth(11) = 1200
        .ColWidth(12) = 1140
        .ColWidth(4) = 1100
        .ColWidth(5) = 1200
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(13) = 1100
        .ColWidth(14) = 1140
        
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        
        'Call FillList
    End With
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With MSFlexGrid1
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
                 .ROW = iRow
                 .COL = 9
                 
                 If SIGNAL Then
                   .CellBackColor = &H8080FF
                 Else
                   .CellBackColor = &HC0FFC0
                 End If
                 
                 .CellFontBold = True
                'If (.Rows > 1 And .ROW <> .Rows - 1) Then
                '.COL = 9
                '.ROW = .ROW
                'End If
               
            Else
                .TextMatrix(iRow, 9) = Empty
                .TextMatrix(iRow, iCol) = strUnChecked
                 '.ROW = iRow
                 '.COL = 9
                 .CellBackColor = vbWhite
                 'If (.Rows > 1 And .ROW <> .Rows - 1) Then .ROW = .ROW + 1: .Col = 0
            End If
        
        If (.ROW > 13) Then .TopRow = .TopRow + 1
        End With
End Sub

Public Sub FillList(Optional FILTER As String)
Dim SQL As String
Dim M_ROW As Integer

MSFlexGrid1.Rows = 1
Screen.MousePointer = vbHourglass
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT ORDTRN.*,ORDMAN.PORD AS PORD,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,SUBGRDMST.NAME AS SUBGRADE "
SQL = SQL & "FROM ORDTRN INNER JOIN ORDMAN ON ORDTRN.ORDN=ORDMAN.ORDN INNER JOIN ACCMST "
SQL = SQL & "ON ACCMST.CODE=ORDTRN.PCOD INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND " & _
            "FINITMMST.UNIT=ORDTRN.UNIT AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD "
SQL = SQL & "INNER JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND ORDTRN.UNIT = SUBGRDMST.UNIT AND ORDTRN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "ORDTRN.GRAD = SUBGRDMST.GRAD AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "WHERE ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & UNCD & _
"' AND ORDTRN.DVCD='" & DIVCODE & "' AND ORDTRN.VTYP='DOS' AND ORDTRN.DFLG<>'Y' AND ORDTRN.RECSTAT='A' AND ORDTRN.DOSTAT='Y' "

If FILTER <> Empty Then SQL = SQL & FILTER

SQL = SQL & "ORDER BY ORDTRN.DONO,ORDTRN.DODT"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

Do While RECSET.EOF = False
   
     With MSFlexGrid1
           
            If M_ROW = Empty Or M_ROW = 0 Then
                .Rows = .Rows + 1
                M_ROW = .Rows - 1
            End If
            
            ''''Code For option Button
            .ROW = M_ROW
            .COL = 0
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .CellForeColor = RGB(LBLRED, LBLGREEN, LBLBLUE) 'vbRed
            .CellAlignment = flexAlignCenterCenter
            .Text = strUnChecked
            ''''''End of Option Button
'INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,
'GRAD,SUBGRD,QNTY,RATE,ARAT,"
'SQL & "ORDN,OSRC,BRMK,PRDL,ICOD,TXRT,TXCD)
   
            .TextMatrix(M_ROW, 1) = Trim(RECSET!ORDN)
            .TextMatrix(M_ROW, 2) = Trim(RECSET!DONO)
            .TextMatrix(M_ROW, 3) = Trim(RECSET!DODt)
            .TextMatrix(M_ROW, 10) = GetCode("REFMST", Trim(RECSET!BRCD), "CODE", "NAME")
            .TextMatrix(M_ROW, 11) = GetCode("ACCMST", Trim(RECSET!PCOD), "CODE", "NAME")
            .TextMatrix(M_ROW, 12) = GetCode("ACCMST", Trim(RECSET!PCOD), "CODE", "NAME")
            .TextMatrix(M_ROW, 4) = Trim(RECSET!ltno)
            .TextMatrix(M_ROW, 5) = Trim(RECSET!ITNM)
            .TextMatrix(M_ROW, 6) = GetCode("GRDMST", Trim(RECSET!grad), "CODE", "GRAD")
            .TextMatrix(M_ROW, 7) = Trim(RECSET!SUBGRADE)
            .TextMatrix(M_ROW, 8) = nstr(Trim(RECSET!QNTY), 12, 3)
            .TextMatrix(M_ROW, 13) = nstr(Trim(RECSET!RATE), 12, 2)
            .TextMatrix(M_ROW, 14) = Trim(RECSET!BRMK)
            
            .TextMatrix(M_ROW, 15) = Trim(RECSET!ICOD)
            .TextMatrix(M_ROW, 16) = Trim(RECSET!grad)
            .TextMatrix(M_ROW, 17) = Trim(RECSET!SUBGRD)
            .TextMatrix(M_ROW, 18) = Trim(RECSET!dbcd)
            M_ROW = Empty
  End With
  
  
  RECSET.MoveNext
  Loop
  RECSET.Close
  
  MSFlexGrid1.ROW = 1
  MSFlexGrid1.COL = 0
Screen.MousePointer = vbNormal
End Sub

Private Function FindStock() As Double
Dim PACKEDQTY As Double: PACKEDQTY = 0
Dim DISPATCHEDQTY As Double: DISPATCHEDQTY = 0

If MSFlexGrid1.Rows <= 1 Then Exit Function

With MSFlexGrid1

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS PACKED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='PPF' AND LOTNO='" & .TextMatrix(.ROW, 4) & _
"' AND FINITMCOD='" & .TextMatrix(.ROW, 15) & "' AND GRAD='" & .TextMatrix(.ROW, 16) & _
"' AND SUBGRAD='" & .TextMatrix(.ROW, 17) & _
"' AND DBCD NOT IN('000001','000005') AND OPER='+' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
 PACKEDQTY = Val(Trim(CHKRS!PACKED & ""))
End If
CHKRS.Close

If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT SUM(ISNULL(QNTY,0)) AS DISPACHED FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='DPF' AND LOTNO='" & .TextMatrix(.ROW, 4) & _
"' AND FINITMCOD='" & .TextMatrix(.ROW, 15) & "' AND GRAD='" & .TextMatrix(.ROW, 16) & _
"' AND SUBGRAD='" & .TextMatrix(.ROW, 17) & "' AND OPER='-' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic

If Not CHKRS.EOF Then
 DISPATCHEDQTY = Val(Trim(CHKRS!DISPACHED & ""))
End If
CHKRS.Close

FindStock = PACKEDQTY - DISPATCHEDQTY
End With
End Function

Private Sub SetDispatchType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT DISTINCT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND CODE NOT IN('000003','000004') AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'  AND NAME<>''", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
PKTYPRS.MoveNext
Loop

If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0

End Sub

Private Sub SetGlobal()
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   M_DBCD = Trim(DBCDRS!CODE & "")
Else
   M_DBCD = Empty
End If
DBCDRS.Close
End Sub

Private Sub FindDetails(QNTY As Double)
On Error GoTo LAST
With MSFlexGrid1
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset
If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open "SELECT ISNULL(NWPB,0) AS NET FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='PPF' AND LOTNO='" & .TextMatrix(.ROW, 4) & _
"' AND FINITMCOD = '" & .TextMatrix(.ROW, 15) & "' AND GRAD='" & .TextMatrix(.ROW, 16) & _
"' AND SUBGRAD='" & .TextMatrix(.ROW, 17) & "' ORDER BY DATE DESC", CN, adOpenDynamic, adLockOptimistic

BOX = 0

If Not FINDRS.EOF Then
  If Val(FINDRS!NET) > 0 Then BOX = QNTY / Val(FINDRS!NET)
End If
End With
Exit Sub

LAST:
MsgBox ERR.Description
End Sub
