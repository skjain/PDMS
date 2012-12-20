VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmJobChallanList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Challan List"
   ClientHeight    =   6345
   ClientLeft      =   1080
   ClientTop       =   2385
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10695
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   10590
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   54591489
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1335
         TabIndex        =   3
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   54591489
         CurrentDate     =   38429
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
         Left            =   3540
         TabIndex        =   4
         Top             =   285
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
         Left            =   255
         TabIndex        =   2
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.Frame FramCont 
      Height          =   4635
      Left            =   75
      TabIndex        =   7
      Top             =   1005
      Width           =   10590
      Begin MSComctlLib.ListView lstBill 
         Height          =   4380
         Left            =   75
         TabIndex        =   8
         Top             =   165
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   7726
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Challan No."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Consinee Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Agent Name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "LotNo"
            Object.Width           =   2213
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Desc."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Grade"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "SubGrade"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Chln Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   75
      TabIndex        =   9
      Top             =   5625
      Width           =   10590
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
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
         Height          =   360
         Left            =   9240
         TabIndex        =   11
         Top             =   195
         Width           =   1035
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         TabIndex        =   10
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Label LBLDIVNAM 
      BackColor       =   &H00C0E0FF&
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
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmJobChallanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVCODE As String
Public M_DBCD As String

Private Sub cmdCancel_Click()
    frmJobChallan.CHALLAN = Empty
    Unload Me
End Sub

Public Sub CMDOK_Click()
    Dim CHLNNO As String
    CHLNNO = lstBill.SelectedItem.SubItems(1)
          
    If Trim(CHLNNO) = Empty Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    Dim EDTDAT As New ADODB.Recordset
    Dim MSTDAT As New ADODB.Recordset
    Set EDTDAT = New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset
    Dim SQL As String
    
    SQL = Empty
 
SQL = "SELECT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=SPTRAN.GRAD LEFT JOIN SUBGRDMST ON SPTRAN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND SPTRAN.UNIT = SUBGRDMST.UNIT AND SPTRAN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "SPTRAN.GRAD = SUBGRDMST.GRAD AND SPTRAN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = SPTRAN.DCOD AND PADDMST.SRNO = SPTRAN.ADDRESS "
SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.DVCD='" & DIVCODE & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT='A' AND SPTRAN.DBCD='" & M_DBCD & _
"' AND SPTRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND SPTRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
"' AND SPTRAN.VBNO='" & CHLNNO & "'"

    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmJobChallan
                       
    .CHALLAN = CHLNNO
    .lblBill.Caption = CHLNNO
    .txtCONSINEE.Text = EDTDAT!ACNM & ""
    .TXTVBDT = Format(EDTDAT!Date & "", "DD/MM/YYYY")
    .txtDCOD.Text = EDTDAT!CONSINEE & ""
    .TXTADDRESS.Text = EDTDAT!ADDRESS & ""
    .txtLTNO.Text = EDTDAT!LTNO & ""
    .txtLTNO.Tag = EDTDAT!LTNO & ""
    .TXTITM.Text = EDTDAT!ITNM & ""
    .TXTITM.Tag = EDTDAT!ITNM & ""
    .TXTGRAD.Text = EDTDAT!GRADE & ""
    .TXTGRAD.Tag = EDTDAT!GRADE & ""
    .TXTSUBGRD.Text = EDTDAT!SUBGRADE & ""
    .TXTSUBGRD.Tag = EDTDAT!SUBGRADE & ""
    .TXTPCS.Text = EDTDAT!PCES & ""
    .txtQTY.Text = EDTDAT!QNTY & ""
    .txtQTY.Tag = EDTDAT!QNTY & ""
    .TXTRATE.Text = EDTDAT!RATE & ""
    .TXTAMNT.Text = EDTDAT!AMNT & ""
    .BRMK = EDTDAT!extra1 & ""
    
    End With
    Unload Me
End Sub

Private Sub Form_Activate()
  LBLDIVNAM.Caption = frmJobChallan.TXTDVNM
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    txtFrDate = GetMinDate
    txtToDate = GetMaxDate
    Me.KeyPreview = True
    cmdOk.Enabled = False
    cmdCancel.Enabled = True
End Sub

Private Sub CMDGO_Click()
 lstBill.ListItems.Clear
 Dim EDTDAT As New ADODB.Recordset
 Set EDTDAT = New ADODB.Recordset
 Dim SQL As String
 SQL = Empty
 
SQL = "SELECT DISTINCT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=SPTRAN.GRAD INNER JOIN SUBGRDMST ON SPTRAN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND SPTRAN.UNIT = SUBGRDMST.UNIT AND SPTRAN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "SPTRAN.GRAD = SUBGRDMST.GRAD AND SPTRAN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = SPTRAN.DCOD AND PADDMST.SRNO = SPTRAN.ADDRESS "
SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.DVCD='" & DIVCODE & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT='A' AND SPTRAN.DBCD='" & M_DBCD & _
"' AND SPTRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND SPTRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'"

SQL = SQL & " ORDER BY SPTRAN.VBNO,SPTRAN.DATE"
 
If EDTDAT.State = 1 Then EDTDAT.Close
EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
If EDTDAT.EOF Then
   MsgBox "No Record found for given criteria ", vbInformation
   txtToDate.SetFocus
   Exit Sub
End If
  
 Do While Not EDTDAT.EOF
    Set lstitem = lstBill.ListItems.Add
    lstitem.Text = Format(EDTDAT![Date], "dd/MM/yyyy")
    lstitem.SubItems(1) = EDTDAT![VBNO]
    lstitem.SubItems(2) = EDTDAT![ACNM]
    lstitem.SubItems(3) = EDTDAT![CONSINEE]
    lstitem.SubItems(5) = EDTDAT![LTNO]
    lstitem.SubItems(6) = EDTDAT![ITNM]
    lstitem.SubItems(7) = EDTDAT![GRADE]    'GetCode("GRDMST", EDTDAT![grad], "CODE", "GRAD")
    lstitem.SubItems(8) = EDTDAT![SUBGRADE]
    lstitem.SubItems(9) = EDTDAT![QNTY]
    lstitem.SubItems(10) = EDTDAT![RATE]
    lstitem.SubItems(11) = EDTDAT![AMNT]
    EDTDAT.MoveNext
 Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmJobChallanList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub LSTBILL_GotFocus()
lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub LSTBILL_LostFocus()
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

