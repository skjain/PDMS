VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBoxSaleList 
   Caption         =   "Sale Help List"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   105
      TabIndex        =   5
      Top             =   345
      Width           =   8355
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4515
         TabIndex        =   7
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56360961
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56360961
         CurrentDate     =   38429
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date: "
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
         Left            =   255
         TabIndex        =   10
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
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
         Left            =   3315
         TabIndex        =   9
         Top             =   285
         Width           =   930
      End
   End
   Begin VB.Frame FramCont 
      Height          =   4995
      Left            =   120
      TabIndex        =   3
      Top             =   1050
      Width           =   8340
      Begin MSComctlLib.ListView lstBill 
         Height          =   4665
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   8229
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Greek"
            Size            =   11.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoice No."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total Quantity"
            Object.Width           =   2223
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Net Amount"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unique"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "VTYP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SRNO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   855
      Left            =   105
      TabIndex        =   0
      Top             =   6120
      Width           =   8415
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Label LBLDAYBOK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "DAYBOOK: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3945
      TabIndex        =   12
      Top             =   120
      Width           =   4320
   End
   Begin VB.Label LBLDIVNAM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "DIVISION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmBoxSaleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Public Sub cmdOk_Click()
  Dim SEL_SRNO As String
  SEL_SRNO = lstBill.SelectedItem.SubItems(1)
  If Trim(SEL_SRNO) = Empty Then
     lstBill.SetFocus
     Exit Sub
  End If
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim SQL As String
  SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & frmBoxSale.DIVCODE & "' AND DBCD='" & frmBoxSale.M_DBCD & "' AND VTYP='SAL' AND VBNO='" & SEL_SRNO & "' AND RECSTAT<>'D'"
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  End If
  With frmBoxSale
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM ACCMST WHERE CODE='" & EDTDAT!DRAC & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTDBAC = MSTDAT!NAME & ""
     Else
      .TXTDBAC.Text = Empty
    End If
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM PADDMST WHERE CODE='" & EDTDAT!DCOD & "' AND SRNO='" & EDTDAT!ADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTDLPTY = MSTDAT!NAME & ""
      .TXTADDRESS = MSTDAT!ADDR & ""
     Else
      .TXTDLPTY = Empty
    End If
    
    'TAX CATEGORY
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT NAME FROM TAXMST WHERE CODE ='" & EDTDAT!TXCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
        .TXTTAXNAM = Trim(MSTDAT!NAME & "")
    Else
        .TXTTAXNAM = Empty
    End If
    
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM REFMST WHERE CODE='" & EDTDAT!BRCD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
      .TXTBRNM = MSTDAT!NAME & ""
     Else
      .TXTBRNM = Empty
    End If
       
    .TXTRTORTAX = EDTDAT!TTYP & ""
    .TXTVBNO = EDTDAT!VBNO & ""
    .TXTVBDT = EDTDAT!Date
    .TXTCOMINV = EDTDAT!CVBN & ""
    .TXTCRDS = EDTDAT!CDAY
    .TXTRMRK = EDTDAT!BRMK & ""
    .TXTTPCS = EDTDAT!TPCS
    .TXTTQTY = Format(EDTDAT!TQTY, "######.000")
    .TXTITOT = Format(EDTDAT!ITOT, "##########.00")
    .TXTBNET = Format(EDTDAT!BNET, "##########.00")
    .TXTGDN = GetCode("LOCMST", EDTDAT!EXTRA4 & "", "CODE", "NAME")
    Dim I As Double
    Dim J As Double
    I = 0
    For I = 0 To .flexBTRM.Rows - 1
      J = 0
      For J = 0 To EDTDAT.Fields.COUNT - 1
        If Trim(EDTDAT.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(I, 0)) Then
            .flexBTRM.TextMatrix(I, 2) = Format(EDTDAT.Fields(J).Value, "#########.00")
        End If
        If Trim(EDTDAT.Fields(J).NAME) = "PER" & Trim(.flexBTRM.TextMatrix(I, 0)) Then
           .flexBTRM.TextMatrix(I, 1) = Format(EDTDAT.Fields(J).Value, "######.00")
        End If
      Next
    Next
    .FLEX.Rows = 2
     SQL = " SELECT SPTRAN.SRCH,SPTRAN.ICOD,SPTRAN.PCES,SPTRAN.QNTY,SPTRAN.RATE,SPTRAN.QORP,SPTRAN.AMNT,SPTRAN.GP_REMARKS,SPTRAN.VBNO,ITMMST.NAME  FROM SPTRAN INNER JOIN ITMMST ON " & _
          " ITMMST.CODE = SPTRAN.ICOD WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & "' AND SPTRAN.DBCD='" & frmBoxSale.M_DBCD & "' AND SPTRAN.DVCD = '" & frmBoxSale.DIVCODE & _
          "'AND SPTRAN.VTYP='SAL' AND SPTRAN.VBNO='" & SEL_SRNO & _
          "'AND SPTRAN.RECSTAT<>'D' ORDER BY SRCH"
          
     If EDTDAT.State = 1 Then EDTDAT.Close
     EDTDAT.Open SQL, CN, adOpenForwardOnly, adLockOptimistic
     I = 1
     Do While Not EDTDAT.EOF
    
    .FLEX.TextMatrix(I, 0) = EDTDAT!SRCH
     If Trim(EDTDAT!ICOD & "") <> "" Then
        .FLEX.TextMatrix(I, 1) = EDTDAT!NAME
     End If
     
     .FLEX.TextMatrix(I, 2) = EDTDAT!PCES
     .FLEX.TextMatrix(I, 3) = Format(EDTDAT!QNTY, "########.000")
     .FLEX.TextMatrix(I, 4) = Format(EDTDAT!RATE, "######.0000")
     .FLEX.TextMatrix(I, 5) = EDTDAT!QORP & ""
     .FLEX.TextMatrix(I, 6) = Format(EDTDAT!AMNT, "#########.00")
     .FLEX.TextMatrix(I, 7) = EDTDAT!GP_REMARKS & ""
     .FLEX.TextMatrix(I, 8) = EDTDAT!ICOD
     .FLEX.TextMatrix(I, 9) = FindBeamNo(Trim(EDTDAT!VBNO & ""))
     Call frmBoxSale.FillList(frmBoxSale.FLEX.ROW)
     Call Checkedroll(EDTDAT!VBNO & "")
     EDTDAT.MoveNext
     I = I + 1
     If Not EDTDAT.EOF Then
       If I > 1 Then
           .FLEX.Rows = .FLEX.Rows + 1
       End If
     End If
    Loop
    
    .TXTDBAC.Enabled = True
    .TXTBRNM.Enabled = True
    .TXTDLPTY.Enabled = True
    
  End With
  Unload Me
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me): Me.BackColor = RGB(RED, GREEN, BLUE)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
  LBLDIVNAM.Caption = DIVNAM
  LBLDAYBOK.Caption = frmBoxSale.Caption
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  Me.KeyPreview = True
  cmdOk.Enabled = False
  cmdCancel.Enabled = True
End Sub

Private Sub cmdGo_Click()
  lstBill.ListItems.Clear
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  SQL = Empty
  SQL = "SELECT DISTINCT BILLMAIN.*,ACCMST.NAME FROM BILLMAIN INNER JOIN ACCMST ON BILLMAIN.PCOD=ACCMST.CODE INNER JOIN SPTRAN ON BILLMAIN.COMP=SPTRAN.COMP AND BILLMAIN.UNIT=SPTRAN.UNIT AND BILLMAIN.DVCD=SPTRAN.DVCD AND BILLMAIN.DBCD=SPTRAN.DBCD AND BILLMAIN.VTYP=SPTRAN.VTYP AND BILLMAIN.VBNO=SPTRAN.VBNO WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.VTYP='SAL' AND BILLMAIN.DVCD='" & frmBoxSale.DIVCODE & "' AND BILLMAIN.DBCD='" & frmBoxSale.M_DBCD & "' AND BILLMAIN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND BILLMAIN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND (SPTRAN.VTYP='SAL') AND BILLMAIN.RECSTAT<>'D' AND BILLMAIN.UNIT='" & UNCD & "' AND BILLMAIN.BSTS='P' AND BILLMAIN.EXTRA5 ='BOX'  ORDER BY BILLMAIN.DATE,BILLMAIN.VBNO"
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
    MsgBox "No Record found for given criteria ", vbInformation
    txtToDate.SetFocus
    Exit Sub
  End If
  Do While Not EDTDAT.EOF
   Set lstItem = lstBill.ListItems.ADD
   lstItem.Text = Format(EDTDAT![Date], "dd/MM/yyyy")
   lstItem.SubItems(1) = EDTDAT![VBNO]
   lstItem.SubItems(2) = EDTDAT![NAME]
   lstItem.SubItems(3) = Format(EDTDAT!TQTY, "########.000")
   lstItem.SubItems(4) = Format(EDTDAT!BNET, "#############.00")
   lstItem.SubItems(6) = EDTDAT!VTYP
   lstItem.SubItems(7) = EDTDAT!SRNO
   EDTDAT.MoveNext
  Loop
  cmdOk.Enabled = True
  cmdOk.Default = True
 End Sub

Private Sub lstBill_GotFocus()
lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub lstBill_LostFocus()
 lstBill.BackColor = vbWhite
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub
Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Function FindBeamNo(GRNNO As String) As String
FindBeamNo = Empty
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

Dim M_BEAMNO As String

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT VBNO,GRSWGT,TRWGT,NTWGT AS QTY,RATE FROM TRDBOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND DVCD='" & frmBoxSale.DIVCODE & "' AND VTYP IN ('SAL') AND RECSTAT='A'  AND grnno ='" & lstBill.SelectedItem.SubItems(1) & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not GETRS.EOF
   If M_BEAMNO <> Empty Then M_BEAMNO = M_BEAMNO & ","
   M_BEAMNO = M_BEAMNO & "'" & Trim(GETRS!VBNO & "") & "'"
GETRS.MoveNext
Loop
GETRS.Close

FindBeamNo = M_BEAMNO

End Function

Private Function Checkedroll(GRNNO As String) As String
GetBoxDetails = Empty
Dim GETRS As ADODB.Recordset
Set GETRS = New ADODB.Recordset

Dim M_BEAMNO As String

If GETRS.State = 1 Then GETRS.Close
GETRS.Open "SELECT VBNO,GRSWGT,TRWGT,NTWGT AS QTY,RATE FROM TRDBOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND DVCD='" & DIVCODE & "' AND VTYP IN ('SAL') AND RECSTAT='A'  AND GRNNO ='" & lstBill.SelectedItem.SubItems(1) & "'", CN, adOpenDynamic, adLockOptimistic
Do While Not GETRS.EOF
   For I = 1 To frmBoxSale.lstRolls.ListItems.COUNT
       If Trim(GETRS!VBNO) = frmBoxSale.lstRolls.ListItems(I) Then
       frmBoxSale.lstRolls.ListItems(I).Checked = True
       End If
  Next
GETRS.MoveNext
Loop
GETRS.Close
End Function



