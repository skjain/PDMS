VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRPT_ReturnableGroupwiseSummary 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Returnable "
   ClientHeight    =   3150
   ClientLeft      =   2895
   ClientTop       =   2790
   ClientWidth     =   5895
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
      Begin VB.OptionButton optpay 
         Caption         =   "Payable"
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
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton optrec 
         Caption         =   "Receiable"
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
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   5700
      Begin VB.TextBox txtParty 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4365
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
         TabIndex        =   13
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   5700
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Text            =   "100"
         Top             =   255
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4440
         TabIndex        =   5
         Top             =   255
         Width           =   1140
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2385
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         TabIndex        =   8
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   3000
      TabIndex        =   11
      Top             =   1560
      Width           =   2820
      Begin MSComCtl2.DTPicker Opdt 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
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
         Format          =   54198273
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
         TabIndex        =   6
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5700
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
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
         TabIndex        =   7
         Top             =   308
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmRPT_ReturnableGroupwiseSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As ADODB.Recordset
Dim rptsql As String
Public DBCR As String, PTYP As String, RPT_TYP As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = vbKeyReturn Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    Opdt = Date
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
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

If PTYP = "PALLET" Then
   Call SetViewForPallet
   If RPT_TYP = "GROUP" Then ReportName = App.PATH & "\Reports\PartyGroup+PartywiseOutstandingReturnablePallet.rpt"
   If RPT_TYP = "PARTY" Then ReportName = App.PATH & "\Reports\PartywiseOutstandingReturnablePallet.rpt"
ElseIf PTYP = "COPS" Then
   Call SetViewForCops
   ReportName = App.PATH & "\Reports\Agent+PartywiseOutstandingReturnableCops.rpt"
End If
                                
CRPT.Reset
crptConnect CRPT

rptsql = Empty
If optrec.Value = True Then
  If RPT_TYP = "PARTY" Then rptsql = "{ACCMST.DRCR}='D'"
 Else
  If RPT_TYP = "PARTY" Then rptsql = "{ACCMST.DRCR}='C'"
End If

PERIOD = "As on Date : " & Opdt

If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If
    
RPTN = Me.Caption
        
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT.Text & "'"
        .Formulas(5) = "OPDT=#" & Format(Opdt.Value, "MM/DD/YYYY") & "#"
         
         If PTYP = "PALLET" Then
            Call SetFormula(5)
         End If
         
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

Private Sub txtParty_GotFocus(): txtParty.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtParty_LostFocus():  txtParty.BackColor = vbWhite: End Sub

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

Private Sub txtUNIT_GotFocus(): txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtUNIT_LostFocus():  txtUNIT.BackColor = vbWhite: End Sub

Private Sub SetViewForPallet()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_GROUP_PARTY_RETURNABLE AS " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER, " & _
   "ISNULL(SUM(BOTTOMPLY),0) AS PCS,CONVERT(CHAR(20),'BOTTOM') AS PLYNAM FROM PKGSTK " & _
   "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
   "' AND PKGSTK.RECSTAT<>'D' " & _
   "  AND PKGSTK.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.BOTTOMPLY  > 0 "

    If txtParty <> Empty Then
      QRY = QRY & " AND PKGSTK.PCOD='" & txtParty.Tag & "' "
    End If

   
   QRY = QRY & " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER " & _
               "UNION " & _
               "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER, " & _
               "ISNULL(SUM(TOPPLY),0) AS PCS,CONVERT(CHAR(20),'TOP') AS PLYNAM FROM PKGSTK " & _
               "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
               "' AND PKGSTK.RECSTAT<>'D' " & _
               "  AND PKGSTK.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & _
               "' AND PKGSTK.TOPPLY > 0 "
   If txtParty <> Empty Then
      QRY = QRY & " AND PKGSTK.PCOD='" & txtParty.Tag & "' "
   End If
               
   QRY = QRY & " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER "
 
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PLYMST WHERE RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
Do While Not RS.EOF
    QRY = QRY & AddSQLForPallet(Trim(RS!NAME & ""))
RS.MoveNext
Loop
RS.Close
       
CN.Execute "IF ( OBJECT_ID('VW_GROUP_PARTY_RETURNABLE') IS NOT NULL ) DROP VIEW VW_GROUP_PARTY_RETURNABLE "
CN.Execute QRY

Exit Sub
VWERR:
MsgBox ERR.Description
Resume
End Sub

Private Function AddSQLForPallet(PLYNAM As String) As String
 Dim SUBQRY As String
 
 SUBQRY = "UNION " & _
          "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER," & _
          "ISNULL(SUM(" & PLYNAM & "),0) AS PCS,CONVERT(CHAR(20),'" & PLYNAM & "') AS PLYNAM FROM PKGSTK " & _
          "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & _
          "' AND PKGSTK.RECSTAT<>'D' " & _
          "  AND PKGSTK.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & _
          "' AND (PKGSTK.TOPPLY > 0 OR PKGSTK.BOTTOMPLY  > 0 ) "
          
          If txtParty <> Empty Then
             SUBQRY = SUBQRY & " AND PKGSTK.PCOD='" & txtParty.Tag & "' "
          End If
          
 SUBQRY = SUBQRY & " GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.PCOD,PKGSTK.OPER "
 
 AddSQLForPallet = SUBQRY
          
End Function

Private Sub SetViewForCops()
Dim QRY As String
On Error GoTo VWERR

QRY = "CREATE VIEW VW_RETURNABLE_DATEWISE_COPSSUMMARY AS " & _
   "SELECT PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.OPER,ISNULL(SUM(PKGSTK.QNTY),0) AS COPS FROM PKGSTK " & _
   "Where PKGSTK.COMP='" & compPth & "' AND PKGSTK.UNIT='" & txtUNIT.Tag & "' " & _
   "  AND PKGSTK.DATE<='" & Format(Opdt.Value, "MM/DD/YYYY") & _
   "' AND PKGSTK.QNTY > 0 AND PKGSTK.RECSTAT<>'D' " & _
   "GROUP BY PKGSTK.COMP,PKGSTK.UNIT,PKGSTK.DATE,PKGSTK.OPER "

CN.Execute "IF ( OBJECT_ID('VW_RETURNABLE_DATEWISE_COPSSUMMARY') IS NOT NULL ) DROP VIEW VW_RETURNABLE_DATEWISE_COPSSUMMARY "
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
 
 

