VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_CopsInventory 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Returnable Cops Inventory"
   ClientHeight    =   2400
   ClientLeft      =   2895
   ClientTop       =   2790
   ClientWidth     =   5865
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   3000
      TabIndex        =   10
      Top             =   840
      Width           =   2820
      Begin MSComCtl2.DTPicker Opdt 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   240
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
         Format          =   53477377
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Run Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   5700
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   5
         Text            =   "100"
         Top             =   255
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2385
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Pre&view"
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
         Image           =   "frmRPT_CopsInventory.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   240
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
         Image           =   "frmRPT_CopsInventory.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Zoom %"
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
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5700
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   308
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmRPT_CopsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As ADODB.Recordset
Dim LUNIT As String

Private Sub Form_Activate()
   Call ColorComponent(Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Opdt = Date
  txtUNIT = UntNm
  txtUNIT.Tag = UNCD
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()

If txtUNIT.Text = Empty Then
   MsgBox "Please Select Unit !!", vbInformation, "Key Field Unit Is Missing"
   txtUNIT.Enabled = True
   txtUNIT.SetFocus
   Exit Sub
End If

If RS.State = 1 Then RS.Close
RS.Open "SELECT CODE FROM UNTMST WHERE COMP='" & compPth & "' AND NAME='" & txtUNIT & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
   LUNIT = Trim(RS!CODE & "")
Else
   MsgBox "Please Select Unit !!", vbInformation, "Key Field Unit Is Missing"
   txtUNIT.Enabled = True
   txtUNIT.SetFocus
   Exit Sub
End If

Call SetViewForCopsInventory

ReportName = Empty
ReportName = App.PATH & "\Reports\CopsInventory.rpt"
RPTN = "Cops Inventory Report "
                                
CRPT.Reset
crptConnect CRPT

PERIOD = "As on Date : " & Opdt

If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If
     
    CRPT.ReportFileName = ReportName
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT.Text & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
                           
         RPTN = RPTN + Space(5) + ReportName
                 
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
    Exit Sub
    
errPreview:
    ErrNumber = Err.Number
    ErrMessage = Err.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub Opdt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
End Sub

Private Sub TXTZOOM_GotFocus():: txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTZOOM_LostFocus(): txtZoom.BackColor = vbWhite: End Sub

Private Sub txtUNIT_GotFocus(): txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtUNIT_LostFocus():  txtUNIT.BackColor = vbWhite: End Sub

Private Sub SetViewForCopsInventory()
Dim QRY As String
On Error GoTo VWERR
'STORE : STR
QRY = "CREATE VIEW VW_COPS_INVENTORY AS " & _
      "SELECT COMP,UNIT,'STR' AS VTYP," & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DVCD='000001' AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' AND OPER='+' AND " & _
      "DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "') - " & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DVCD='000001' AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' AND OPER='-' AND " & _
      "DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "') AS QNTY FROM STORETRAN WHERE COMP='" & compPth & _
      "' AND UNIT='" & LUNIT & "' AND DVCD='000001' AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' GROUP BY COMP,UNIT " & _
      "UNION "
'RGP : RGP
QRY = QRY & "SELECT COMP,UNIT,'RGP' AS VTYP," & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DVCD='000001' AND VTYP IN ('RGP','ANX') AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' AND OPER='-' AND " & _
      " DBCD='000001' AND DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "') - " & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DVCD='000001' AND DBCD IN ('000003','000004') AND VTYP='IVR' AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' AND OPER='+' AND " & _
      "DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "') AS QNTY FROM STORETRAN WHERE COMP='" & compPth & _
      "' AND UNIT='" & LUNIT & "' AND DVCD='000001' AND RECSTAT<>'D' AND ICOD='XXXXXXXXXX' GROUP BY COMP,UNIT " & _
      "UNION "
'BOX : BOX
QRY = QRY & "SELECT COMP,UNIT,'BOX' AS VTYP,ISNULL(SUM(COPS),0) FROM BOXREGISTER " & _
      "WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & "' AND VTYP IN ('PPF','OPN') AND RECSTAT<>'D' AND " & _
      "VBDT<='" & Format(Opdt.Value, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT " & _
      "UNION "
'RETURNABLE PARTY : PTY
QRY = QRY & "SELECT COMP,UNIT,'PTY' AS VTYP," & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "' AND QNTY>0 AND RECSTAT<>'D' AND OPER='-') - " & _
      "(SELECT ISNULL(SUM(QNTY),0) FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & _
      "' AND DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "' AND QNTY>0 AND RECSTAT<>'D' AND OPER='+') AS QNTY " & _
      "FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & LUNIT & "' AND RECSTAT<>'D' " & _
      "AND DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "' GROUP BY COMP,UNIT " & _
      "UNION "
'WIP(WORK IN PROGRESS) : WIP
QRY = QRY & "SELECT STORETRAN.COMP,STORETRAN.UNIT,'WIP' AS VTYP," & _
      "(SELECT ISNULL(SUM(STORETRAN.QNTY),0) FROM STORETRAN WHERE STORETRAN.COMP='" & compPth & _
      "' AND STORETRAN.UNIT='" & LUNIT & "' AND STORETRAN.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & _
      "' AND STORETRAN.VTYP='ISS' AND STORETRAN.DVCD<>'000001' AND STORETRAN.RECSTAT<>'D' " & _
      "AND STORETRAN.ICOD='XXXXXXXXXX' AND STORETRAN.OPER='+' AND STORETRAN.DBCD='000001') - " & _
      "(SELECT ISNULL(SUM(BOXREGISTER.COPS),0) FROM BOXREGISTER WHERE BOXREGISTER.COMP='" & compPth & _
      "' AND BOXREGISTER.UNIT='" & LUNIT & "' AND BOXREGISTER.VBDT <= '" & Format(Opdt.Value, "MM/DD/YYYY") & _
      "' AND BOXREGISTER.RECSTAT<>'D') AS QNTY FROM STORETRAN WHERE STORETRAN.COMP='" & compPth & _
      "' AND STORETRAN.UNIT='" & LUNIT & "' AND STORETRAN.RECSTAT<>'D' " & _
      " AND STORETRAN.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & "' GROUP BY  STORETRAN.COMP,STORETRAN.UNIT"
       
CN.Execute "IF ( OBJECT_ID('VW_COPS_INVENTORY') IS NOT NULL ) DROP VIEW VW_COPS_INVENTORY "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox Err.Description
End Sub
