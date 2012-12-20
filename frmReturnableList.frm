VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReturnableList 
   Caption         =   "Returnable Issue / Receive Help"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   11
      Top             =   5685
      Width           =   10590
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
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
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
         TabIndex        =   5
         Top             =   195
         Width           =   1035
      End
   End
   Begin VB.Frame FramCont 
      Height          =   4635
      Left            =   120
      TabIndex        =   10
      Top             =   1065
      Width           =   10590
      Begin MSComctlLib.ListView lstBill 
         Height          =   4380
         Left            =   75
         TabIndex        =   3
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1835
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Slip No."
            Object.Width           =   1906
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   4763
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agent Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type"
            Object.Width           =   1589
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cops"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Wodden"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Pvc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Fibre"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Top"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Bottom"
            Object.Width           =   1235
         EndProperty
      End
   End
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   7
      Top             =   360
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
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   1
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1335
         TabIndex        =   0
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   38429
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
         TabIndex        =   9
         Top             =   285
         Width           =   1065
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
         TabIndex        =   8
         Top             =   285
         Width           =   885
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmReturnableList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVCODE As String
Public M_DBCD As String

Private Sub cmdCancel_Click()
    frmReturnable.CHALLAN = Empty
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
    
SQL = "SELECT PKGSTK.*,ACCMST.NAME AS ACNM,REFMST.NAME AS AGENT,PADDMST.NAME AS CONSIGNEE,PADDMST.ADDR AS ADDRESS " & _
      "FROM PKGSTK INNER JOIN ACCMST ON ACCMST.CODE=PKGSTK.PCOD " & _
      "INNER JOIN REFMST ON REFMST.CODE=PKGSTK.BRCD " & _
      "LEFT JOIN PADDMST ON PADDMST.CODE=PKGSTK.DCOD AND PADDMST.SRNO=PKGSTK.ADDRESS " & _
      "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND VTYP='RET' AND " & _
      "PKGSTK.RECSTAT='A' AND DBCD='" & M_DBCD & "' AND DATE >='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
      "' AND DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND CHLN = '" & CHLNNO & "'"

    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmReturnable
    .CHALLAN = CHLNNO
    .lblSlip.Caption = CHLNNO
    .txtName.Text = EDTDAT!ACNM & ""
    .TXTVBDT = Format(EDTDAT!Date & "", "DD/MM/YYYY")
    .TXTBRNM.Text = EDTDAT!AGENT & ""
    .TXTNOB.Text = EDTDAT!BOTTOMPLY & ""
    .TXTPallets = Val(EDTDAT!PALLETS & "")
    .TXTNOC.Text = EDTDAT!QNTY & ""
    .TXTNOT.Text = EDTDAT!TOPPLY & ""
    .TXTRMRK.Text = EDTDAT!BRMK & ""
    .TXTDCOD.Text = EDTDAT!CONSIGNEE & ""
    .TXTADDRESS.Text = EDTDAT!ADDRESS & ""
    
    Dim I As Long, J As Long
    I = 0
    For I = 1 To .FLEXPLY.Cols - 1
    J = 0
       For J = 0 To EDTDAT.Fields.COUNT - 1
          If Trim(EDTDAT.Fields(J).NAME) = Trim(.FLEXPLY.TextMatrix(0, I)) Then
             .FLEXPLY.TextMatrix(1, I) = Val(EDTDAT.Fields(J).Value)
          End If
       Next
    Next
    
    If Trim(EDTDAT![OPER] & "") = "+" Then
       .optRecieved.Value = True
    Else
       .optIssue.Value = True
    End If
    
    End With
    Unload Me
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

Private Sub cmdGo_Click()
 lstBill.ListItems.Clear
 Dim EDTDAT As New ADODB.Recordset
 Set EDTDAT = New ADODB.Recordset
 Dim SQL As String
 SQL = Empty
 
SQL = "SELECT DISTINCT PKGSTK.*,ACCMST.NAME AS ACNM,REFMST.NAME AS AGENT FROM PKGSTK INNER JOIN ACCMST ON ACCMST.CODE=PKGSTK.PCOD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE=PKGSTK.BRCD WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & DIVCODE & "' AND VTYP='RET' AND PKGSTK.RECSTAT='A' AND DBCD='" & M_DBCD & _
"' AND DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'"

SQL = SQL & " ORDER BY DATE DESC"
 
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
    lstItem.SubItems(1) = Trim(EDTDAT![chln] & "")
    lstItem.SubItems(2) = Trim(EDTDAT![ACNM] & "")
    lstItem.SubItems(3) = Trim(EDTDAT![AGENT] & "")
    If Trim(EDTDAT![OPER] & "") = "+" Then
       lstItem.SubItems(4) = "RECEIVE"
    Else
       lstItem.SubItems(4) = "ISSUE"
    End If
    lstItem.SubItems(5) = Trim(EDTDAT![QNTY] & "")
    lstItem.SubItems(9) = Trim(EDTDAT![TOPPLY] & "")
    lstItem.SubItems(10) = Trim(EDTDAT![BOTTOMPLY] & "")
    EDTDAT.MoveNext
 Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmReturnableList.Visible = True Then lstBill.SetFocus
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


