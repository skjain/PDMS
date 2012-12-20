VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReturnToStoreList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue List Help ItemWise"
   ClientHeight    =   6990
   ClientLeft      =   2130
   ClientTop       =   2400
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9992.263
   ScaleMode       =   0  'User
   ScaleWidth      =   9795
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   9555
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
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4185
         TabIndex        =   9
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   54263809
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
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
         Format          =   54263809
         CurrentDate     =   38429
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date: "
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
         Width           =   1005
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
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
         Left            =   3315
         TabIndex        =   3
         Top             =   285
         Width           =   825
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5235
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9540
      Begin MSComctlLib.ListView lstBill 
         Height          =   5025
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   8864
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            Text            =   "Slip No."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Machine No."
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4516
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qnty"
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
      Height          =   630
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   9495
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
         Left            =   5070
         TabIndex        =   7
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
         Left            =   6405
         TabIndex        =   8
         Top             =   195
         Width           =   1035
      End
   End
   Begin VB.Label LBLDIVNAM 
      Alignment       =   2  'Center
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
      Width           =   9720
   End
End
Attribute VB_Name = "frmReturnToStoreList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVISION As String

Private Sub CMDGO_Click()
  lstBill.ListItems.Clear
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  SQL = Empty
  SQL = "SELECT STORETRAN.*,MACMST.NAME FROM STORETRAN INNER JOIN MACMST ON (STORETRAN.PCOD=MACMST.CODE AND STORETRAN.UNIT=MACMST.UNIT AND STORETRAN.DVCD=MACMST.DVCD AND STORETRAN.COMP=MACMST.COMP) WHERE STORETRAN.COMP='" & compPth & "' AND STORETRAN.VTYP='RTI' AND STORETRAN.DVCD='" & DIVISION & "' AND STORETRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND STORETRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'  AND STORETRAN.RECSTAT<>'D'  AND STORETRAN.UNIT='" & UNCD & "' ORDER BY STORETRAN.DATE,STORETRAN.VBNO"
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
   lstitem.SubItems(2) = EDTDAT![NAME]
   lstitem.SubItems(3) = GetCode("ITMMST", EDTDAT!ICOD, "CODE", "NAME")
   lstitem.SubItems(4) = Format(EDTDAT!QNTY, "########.000")
   lstitem.SubItems(6) = EDTDAT!VTYP
   lstitem.SubItems(7) = EDTDAT!SRNO
   EDTDAT.MoveNext
  Loop
  
  cmdOk.Enabled = True
  cmdOk.Default = True
  If frmReturnToStoreList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub cmdCancel_Click()
  frmStoreIssue.M_SRNO = Empty
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim SEL_SRNO As String
  SEL_SRNO = lstBill.SelectedItem.SubItems(7)
  frmReturnToStore.M_SRNO = SEL_SRNO
  If Trim(SEL_SRNO) = Empty Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  Dim SQL As String
  SQL = "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='RTI' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'  AND STORETRAN.UNIT='" & UNCD & "' AND STORETRAN.DVCD='" & DIVISION & "'"
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  With frmReturnToStore
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM MACMST WHERE CODE='" & EDTDAT!PCOD & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVISION & "' AND COMP='" & compPth & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
        .txtMACHINE = MSTDAT!NAME & ""
    Else
        .txtMACHINE = Empty
    End If
    
    .TXTVBNO = EDTDAT!VBNO
    .TXTREQSLIP = EDTDAT!chln
    .TXTVBDT = EDTDAT!Date
    .ITMFLEX.Rows = 2
        
    I = 1
    Do While Not EDTDAT.EOF
     .ITMFLEX.TextMatrix(I, 0) = Trim(EDTDAT!ICOD)
     .ITMFLEX.TextMatrix(I, 1) = GetCode("ITMMST", Trim(EDTDAT!ICOD), "CODE", "NAME")
     '.ITMFLEX.TextMatrix(i, 2) = nstr(.GetItemStock() + Val(EDTDAT!QNTY), 12, 3)
     .ITMFLEX.TextMatrix(I, 3) = Trim(nstr(EDTDAT!QNTY, 12, 3))
     .ITMFLEX.TextMatrix(I, 4) = Trim(nstr(EDTDAT!RATE, 10, 3))
     .ITMFLEX.TextMatrix(I, 5) = Trim(nstr(EDTDAT!AMNT, 10, 2))
    EDTDAT.MoveNext
    If .ITMFLEX.Rows > 6 Then .ITMFLEX.TopRow = .ITMFLEX.TopRow + 2
    Loop
    
    .btn_sts (True)
    .cmdCancel.Cancel = True
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  LBLDIVNAM.Caption = frmReturnToStore.TXTFROMDIV
  DIVISION = GetDivCode(LBLDIVNAM.Caption)
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  cmdOk.Enabled = False
  cmdCancel.Enabled = True
  Me.KeyPreview = True
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
