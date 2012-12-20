VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_ReturnableAddresswiseCopsSummary 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Returnable "
   ClientHeight    =   4425
   ClientLeft      =   2895
   ClientTop       =   2790
   ClientWidth     =   7050
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   6855
      Begin VB.OptionButton OPTAREA 
         Caption         =   "Area Wise"
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
         Left            =   5400
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Returnable Summary"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optDetail 
         Caption         =   "Returnable Detail Ledger"
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
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   6900
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   4800
         TabIndex        =   9
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
         Format          =   51642369
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1200
         TabIndex        =   7
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
         Format          =   51642369
         CurrentDate     =   38429
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
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
         TabIndex        =   6
         Top             =   308
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
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
         Left            =   3600
         TabIndex        =   8
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1410
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   6900
      Begin VB.TextBox TXTAREA 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   5085
      End
      Begin VB.TextBox txtDcod 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   5085
      End
      Begin VB.TextBox txtParty 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Area :"
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
         TabIndex        =   23
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee :"
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
         TabIndex        =   4
         Top             =   675
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name :"
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
         TabIndex        =   2
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   6825
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
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
         Left            =   1500
         TabIndex        =   13
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
         Left            =   3240
         TabIndex        =   14
         Top             =   200
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
         Image           =   "frmRPT_ReturnableAddresswiseCopsSummary.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4560
         TabIndex        =   15
         Top             =   200
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "frmRPT_ReturnableAddresswiseCopsSummary.frx":0452
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
         TabIndex        =   12
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6900
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5085
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
Attribute VB_Name = "frmRPT_ReturnableAddresswiseCopsSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As ADODB.Recordset
Dim rptsql As String
Public DBCR As String, PTYP As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = vbKeyReturn Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    dtFrom = GetMinDate
    dtTo = Date
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    lblArea.Enabled = False
    TXTAREA.Enabled = False
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()

If txtUNIT.Text = Empty Then
   MsgBox "Please Select Unit !!", vbInformation, "Key Field Unit Is Missing"
   txtUNIT.SetFocus
   Exit Sub
End If

ReportName = Empty

Call SetViewForCops

If optSummary.Value = True Then
   ReportName = App.PATH & "\Reports\AddresswiseOSReturnableCops.rpt"
Else
   ReportName = App.PATH & "\Reports\PartyAddresswiseOSReturnableCopsLedger.rpt"
End If


If OPTAREA.Value = True Then
   Call SetViewForArea
   ReportName = App.PATH & "\Reports\AreawiseOSReturnableCops.rpt"
End If

                                
CRPT.Reset
crptConnect CRPT

rptsql = Empty
rptsql = "{ACCMST.DRCR}='D'"

If TXTAREA <> Empty Then
rptsql = rptsql & " AND {VW_RETURNABLE_AREA_COPSSUMMARY.ARCD}  = '" & TXTAREA.Tag & "'"
End If

PERIOD = dtFrom & " To " & dtTo

If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If
    
If OPTAREA.Value = True Then
RPTN = "AREA WISE SUMMARY REPORT FOR RETURNABLE COPS "
Else
RPTN = Me.Caption
End If

    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT.Text & "'"
        .Formulas(5) = "OPDT=#" & Format(dtFrom.Value, "MM/DD/YYYY") & "#"
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
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub OPTAREA_Click()
lblArea.Enabled = True
TXTAREA.Enabled = True
End Sub

Private Sub optDetail_Click()
lblArea.Enabled = False
TXTAREA.Enabled = False
TXTAREA = Empty
End Sub

Private Sub optSummary_Click()
lblArea.Enabled = False
TXTAREA.Enabled = False
TXTAREA = Empty
End Sub

Private Sub TXTAREA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True:  M_DESC = Empty
        TXTAREA = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA = 'A' ", 0, Empty, "Select Area From List")
        TXTAREA.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTAREA = Empty
        TXTAREA.Tag = Empty
    End If
End Sub

Private Sub TXTDCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True:  M_DESC = Empty
        txtDcod = SearchList1("Select DISTINCT TOP 20 Code,Name From PADDMST", 0, Empty, "Select Consignee From List")
        txtDcod.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDcod = Empty
        txtDcod.Tag = Empty
    End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True:  M_DESC = Empty
        txtParty = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party From List")
        txtParty.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtParty = Empty
        txtParty.Tag = Empty
    End If
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
End Sub

Private Sub TXTZOOM_GotFocus():: txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTZOOM_LostFocus(): txtZoom.BackColor = vbWhite: End Sub

Private Sub SetViewForPallet()
Dim QRY As String
On Error GoTo VWERR
'CONVERT(NCHAR(30),'BOTTOM')
QRY = "CREATE VIEW VW_PARTY_AGENT_RETURNABLE AS " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER, " & _
   "ISNULL(SUM(BOTTOMPLY),0) AS PCS,CONVERT(CHAR(20),'BOTTOM') AS PLYNAM FROM PKGSTK " & _
   "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
   "' AND PKGSTK.RECSTAT<>'D' AND PKGSTK.BRCD IS NOT NULL " & _
   "  AND PKGSTK.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.BOTTOMPLY  > 0 " & _
   " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER " & _
   "UNION " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER, " & _
   "ISNULL(SUM(TOPPLY),0) AS PCS,CONVERT(CHAR(20),'TOP') AS PLYNAM FROM PKGSTK " & _
   "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
   "' AND PKGSTK.RECSTAT<>'D' AND PKGSTK.BRCD IS NOT NULL " & _
   "  AND PKGSTK.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.TOPPLY > 0 " & _
   " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER "
 
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not RS.EOF
    QRY = QRY & AddSQLForPallet(Trim(RS!NAME & ""))
RS.MoveNext
Loop
RS.Close
       
CN.Execute "IF ( OBJECT_ID('VW_PARTY_AGENT_RETURNABLE') IS NOT NULL ) DROP VIEW VW_PARTY_AGENT_RETURNABLE "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Function AddSQLForPallet(PLYNAM As String) As String

 AddSQLForPallet = "UNION " & _
          "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER," & _
          "ISNULL(SUM(" & PLYNAM & "),0) AS PCS,CONVERT(CHAR(20),'" & PLYNAM & "') AS PLYNAM FROM PKGSTK " & _
          "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
          "' AND PKGSTK.RECSTAT<>'D' AND PKGSTK.BRCD IS NOT NULL " & _
          "  AND PKGSTK.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
          "' AND (PKGSTK.TOPPLY > 0 OR PKGSTK.BOTTOMPLY  > 0 ) " & _
          " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.BRCD,PKGSTK.PCOD,PKGSTK.OPER "
          
End Function

Private Sub SetViewForCops()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_RETURNABLE_ADDRESSWISE_COPSSUMMARY AS " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.DCOD,PKGSTK.ADDRESS,PKGSTK.CHLN,PKGSTK.OPER," & _
   "ISNULL(SUM(PKGSTK.QNTY),0) AS COPS FROM PKGSTK Where PKGSTK.COMP='" & compPth & _
   "' AND PKGSTK.UNIT='" & txtUNIT.Tag & "' " & _
   "  AND PKGSTK.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.QNTY > 0 AND PKGSTK.RECSTAT<>'D' "
   
If txtParty <> Empty Then
  QRY = QRY & " AND PKGSTK.PCOD='" & txtParty.Tag & "' "
End If

If txtDcod <> Empty Then
  QRY = QRY & " AND PKGSTK.DCOD='" & txtDcod.Tag & "' "
End If

'If TXTAREA <> Empty Then
' QRY = QRY & " AND PADDMST.ARCD = '" & TXTAREA.Tag & "'"
'End If
   
   
QRY = QRY & "GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.DCOD,PKGSTK.ADDRESS,PKGSTK.CHLN,PKGSTK.OPER "

CN.Execute "IF ( OBJECT_ID('VW_RETURNABLE_ADDRESSWISE_COPSSUMMARY') IS NOT NULL ) DROP VIEW VW_RETURNABLE_ADDRESSWISE_COPSSUMMARY "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub

Private Sub SetFormula(COUNT As Long)

Dim ctr As Long: ctr = 0

With CRPT
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not RS.EOF
   COUNT = COUNT + 1
   ctr = ctr + 1
   If ctr = 7 Then Exit Sub 'AT PRESENT CONSIDER MAX FIVE PLY
   .Formulas(COUNT) = "PLY" & CStr(ctr) & "='" & Trim(RS!NAME & "") & "'"
RS.MoveNext
Loop
RS.Close
End With

End Sub
 
Private Sub txtUNIT_GotFocus(): txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtUNIT_LostFocus():  txtUNIT.BackColor = vbWhite: End Sub

Private Sub txtParty_GotFocus(): txtParty.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtParty_LostFocus():  txtParty.BackColor = vbWhite: End Sub

Private Sub TXTDCOD_GotFocus(): txtDcod.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTDCOD_LostFocus():  txtDcod.BackColor = vbWhite: End Sub


Private Sub SetViewForArea()
Dim QRY As String
On Error GoTo VWERR


QRY = "CREATE VIEW VW_RETURNABLE_AREA_COPSSUMMARY AS " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.DCOD,PKGSTK.ADDRESS,PADDMST.ARCD,PKGSTK.CHLN,PKGSTK.OPER," & _
   "ISNULL(SUM(PKGSTK.QNTY),0) AS COPS FROM PKGSTK INNER JOIN PADDMST ON PADDMST.CODE = PKGSTK.DCOD AND PADDMST.SRNO = PKGSTK.ADDRESS  Where PKGSTK.COMP='" & compPth & _
   "' AND PKGSTK.UNIT='" & txtUNIT.Tag & "' " & _
   "  AND PKGSTK.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.QNTY > 0 AND PKGSTK.RECSTAT<>'D' "
   
If txtParty <> Empty Then
  QRY = QRY & " AND PKGSTK.PCOD='" & txtParty.Tag & "' "
End If

If txtDcod <> Empty Then
  QRY = QRY & " AND PKGSTK.DCOD='" & txtDcod.Tag & "' "
End If

QRY = QRY & "GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.DCOD,PKGSTK.ADDRESS,PKGSTK.CHLN,PKGSTK.OPER,PADDMST.ARCD "

CN.Execute "IF ( OBJECT_ID('VW_RETURNABLE_AREA_COPSSUMMARY') IS NOT NULL ) DROP VIEW VW_RETURNABLE_AREA_COPSSUMMARY "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
End Sub


