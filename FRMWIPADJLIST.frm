VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRMWIPADJLIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WIP ADJUSTMENT HELP LIST"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   105
      TabIndex        =   9
      Top             =   6240
      Width           =   8280
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
         TabIndex        =   10
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
         Left            =   5040
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5235
      Left            =   105
      TabIndex        =   8
      Top             =   960
      Width           =   8220
      Begin MSComctlLib.ListView lstBill 
         Height          =   5025
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   8010
         _ExtentX        =   14129
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Adj. No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Division"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Remark"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DVCD"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Width           =   8235
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
         TabIndex        =   3
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4185
         TabIndex        =   2
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
         Format          =   86900737
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
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
         Format          =   86900737
         CurrentDate     =   38429
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
         TabIndex        =   7
         Top             =   285
         Width           =   825
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
         TabIndex        =   6
         Top             =   285
         Width           =   1005
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8280
   End
End
Attribute VB_Name = "FRMWIPADJLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVISION As String

Private Sub cmdGo_Click()
  lstBill.ListItems.Clear
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  SQL = Empty
  SQL = "SELECT DISTINCT STORETRAN.DATE,STORETRAN.DVCD,STORETRAN.VBNO,STORETRAN.SRNO,SUM(STORETRAN.QNTY),MACMST.NAME,DIVMST.NAME FROM STORETRAN INNER JOIN MACMST ON (STORETRAN.PCOD=MACMST.CODE AND " & _
  " STORETRAN.UNIT=MACMST.UNIT AND STORETRAN.DVCD=MACMST.DVCD AND STORETRAN.COMP=MACMST.COMP) INNER JOIN DIVMST ON " & _
  " STORETRAN.COMP = DIVMST.COMP AND STORETRAN.UNIT = DIVMST.UNIT AND STORETRAN.DVCD = DIVMST.CODE WHERE " & _
  " STORETRAN.COMP='" & compPth & "' AND STORETRAN.VTYP='WIP' AND STORETRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
  "' AND STORETRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'  AND STORETRAN.RECSTAT<>'D' " & _
  " AND STORETRAN.UNIT='" & UNCD & "' AND STORETRAN.SRCH = '1' GROUP BY STORETRAN.DATE,STORETRAN.DVCD,STORETRAN.SRNO,MACMST.NAME,STORETRAN.VBNO,DIVMST.NAME"
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
   lstItem.SubItems(2) = EDTDAT!NAME & ""
   lstItem.SubItems(3) = EDTDAT!SRNO & ""
   lstItem.SubItems(4) = EDTDAT!DVCD
   
  EDTDAT.MoveNext
  Loop
  
  cmdOk.Enabled = True
  cmdOk.Default = True
 ' If frmStoreIssList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub CMDCANCEL_Click()
  frmStoreIssue.M_SRNO = Empty
  Unload Me
End Sub

Private Sub CMDOK_Click()
  Dim SEL_SRNO As String
  SEL_SRNO = lstBill.SelectedItem.SubItems(3)
  FRMWIPADJ.M_SRNO = SEL_SRNO
  If Trim(SEL_SRNO) = Empty Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  Dim SQL As String
  SQL = "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND VTYP='WIP' AND SRNO='" & SEL_SRNO & "' AND RECSTAT<>'D'  AND STORETRAN.UNIT='" & UNCD & "' AND STORETRAN.DVCD='" & lstBill.SelectedItem.SubItems(4) & "'"
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  With FRMWIPADJ
  .TXTDVCD = GETDIVNAME(lstBill.SelectedItem.SubItems(4))
  .TXTRMRK = Trim(EDTDAT!ITEMRMRK & "")
    
    .TXTVBNO = EDTDAT!VBNO
    .TXTVBDT = EDTDAT!Date
    .ITMFLEX.Rows = 2
        
    I = 1
    Do While Not EDTDAT.EOF
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM MACMST WHERE CODE='" & EDTDAT!PCOD & "' AND UNIT='" & UNCD & "' AND DVCD='" & lstBill.SelectedItem.SubItems(4) & "' AND COMP='" & compPth & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
        .TXTMACHINE = Trim(MSTDAT!NAME & "")
    End If
    
     .ITMFLEX.TextMatrix(I, 0) = Trim(MSTDAT!NAME & "")
     .ITMFLEX.TextMatrix(I, 1) = GetCode("ITMMST", Trim(EDTDAT!ICOD), "CODE", "NAME")
     If (EDTDAT!OPER) = "-" Then
     .ITMFLEX.TextMatrix(I, 2) = nstr(Val(EDTDAT!QNTY & ""), 12, 3)
     Else
     .ITMFLEX.TextMatrix(I, 2) = nstr(Val(-(EDTDAT!QNTY & "")), 12, 3)
     End If
     .ITMFLEX.TextMatrix(I, 3) = (EDTDAT!PCOD)
     .ITMFLEX.TextMatrix(I, 4) = (EDTDAT!ICOD)
      I = I + 1
     If Not EDTDAT.EOF Then
       If I > 1 Then
           .ITMFLEX.Rows = .ITMFLEX.Rows + 1
       End If
     End If
     EDTDAT.MoveNext
    If .ITMFLEX.Rows > 5 Then .ITMFLEX.TopRow = .ITMFLEX.TopRow + 2
    Loop
    .ITMFLEX.Rows = .ITMFLEX.Rows - 1
    .btn_sts (True)
    .cmdCancel.CANCEL = True
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
 ' LBLDIVNAM.Caption = frmStoreIssue.TXTTODIV
 ' DIVISION = GetDivCode(LBLDIVNAM.Caption)
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

