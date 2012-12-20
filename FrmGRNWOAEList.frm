VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmGRNWOAEList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue List Help ItemWise"
   ClientHeight    =   7275
   ClientLeft      =   2130
   ClientTop       =   2400
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10399.67
   ScaleMode       =   0  'User
   ScaleWidth      =   8040
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   7875
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
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4320
         TabIndex        =   3
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
         Format          =   50200577
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1440
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
         Format          =   50200577
         CurrentDate     =   38429
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date: "
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
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
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5835
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   7860
      Begin MSComctlLib.ListView lstBill 
         Height          =   5625
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   9922
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GRN No."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Total Pcs"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Total Qty."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Total Amt."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   7815
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   195
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FrmGRNWOAEList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LVTYP As String
Dim I As Long

Private Sub cmdGo_Click()
  Dim lstITEM
  lstBill.ListItems.Clear
  frmGRNWOAcEffect.M_DBCD = "000005"
  
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  
  Dim SQL As String
  
  SQL = Empty
  SQL = "SELECT * FROM GRN " & _
        "WHERE GRN.COMP='" & compPth & "' AND GRN.UNIT='" & UNCD & _
        "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
        "' AND BSTS='P' " & _
        " AND VTYP='IVR' AND GRN.DBCD='" & frmGRNWOAcEffect.M_DBCD & "' AND GRN.RECSTAT<>'D' " & _
        "ORDER BY GRN.DATE,GRN.VBNO"
   
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
    MsgBox "No Record found for given criteria ", vbInformation
    txtToDate.SetFocus
    Exit Sub
  End If
  
  Do While Not EDTDAT.EOF
   Set lstITEM = lstBill.ListItems.ADD
   lstITEM.Text = Format(EDTDAT![Date], "dd/MM/yyyy")
   lstITEM.SubItems(1) = EDTDAT![VBNO]
   lstITEM.SubItems(2) = Trim(EDTDAT![TPCS] & "")
   lstITEM.SubItems(3) = Format(EDTDAT!TQTY, "########.000")
   lstITEM.SubItems(4) = Format(EDTDAT!ITOT, "########.00")
   EDTDAT.MoveNext
  Loop
  
  cmdOk.Enabled = True
  cmdOk.Default = True
  If frmGRNWOAcEffect.Visible = True Then lstBill.SetFocus
End Sub

Private Sub cmdCancel_Click()
  Unload Me
  frmGRNWOAcEffect.M_DBCD = Empty
End Sub

Private Sub CMDOK_Click()
  If lstBill.SelectedItem.SubItems(1) <> Empty And lstBill.ListItems.COUNT < 1 Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  With frmGRNWOAcEffect
  Dim VBNO As String
  
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  Dim SQL As String
  SQL = Empty
  
  SQL = "SELECT * FROM GRN " & _
        "WHERE GRN.COMP='" & compPth & "' AND GRN.UNIT='" & UNCD & _
        "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.VTYP='IVR' AND GRN.DBCD='" & frmGRNWOAcEffect.M_DBCD & "' AND GRN.RECSTAT<>'D' " & _
        "  AND GRN.VBNO='" & lstBill.SelectedItem.SubItems(1) & "' AND GRN.ACEFFECT = 'N'"
        
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  Else
     
     .TXTVBNO = EDTDAT!VBNO & ""                         'RECEIVE NO.
     .TXTVBDT = EDTDAT!Date & ""
     .TXTRMRK = Trim(EDTDAT!BRMK & "")
     .TXTITOT = Trim(nstr(EDTDAT!ITOT, 12, 2))
       
  End If
          
    'FIFO-----------------------------
         If RS.State = 1 Then RS.Close
         RS.Open "SELECT * FROM GRNTRAN WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
         "' AND DBCD='" & frmGRNWOAcEffect.M_DBCD & "' AND VBNO = '" & lstBill.SelectedItem.SubItems(1) & _
         "' AND GRN_QNTY <> BAL_QNTY ", CN, adOpenDynamic, adLockReadOnly
         If RS.EOF = False Then .M_ISSUE = "Y"
    '---------------------------------
    
    Dim DETRS As ADODB.Recordset
    Set DETRS = New ADODB.Recordset
    
    If DETRS.State = 1 Then DETRS.Close
    DETRS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND " & _
               "DBCD='" & frmGRNWOAcEffect.M_DBCD & "' AND VTYP='IVR' AND VBNO='" & lstBill.SelectedItem.SubItems(1) & _
               "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
               
    .Flex.Rows = 2
    
    I = 1
    Do While Not DETRS.EOF
     .Flex.TextMatrix(I, 0) = Trim(DETRS!SRCH)
     .Flex.TextMatrix(I, 1) = GetCodeERP("ITMMST", Trim(DETRS!ICOD & ""), "CODE", "NAME")
     .Flex.TextMatrix(I, 2) = Val(DETRS!PCES)
     .Flex.TextMatrix(I, 3) = Trim(nstr(DETRS!QNTY, 10, 3))
     
     .TXTTPCS = Val(.TXTTPCS) + Val(DETRS!PCES)
     .TXTTQTY = Val(.TXTTQTY) + Val(DETRS!QNTY)
     
     .Flex.TextMatrix(I, 4) = Trim(nstr(DETRS!RATE, 10, 2))
     .Flex.TextMatrix(I, 5) = Trim(nstr(DETRS!AMNT, 10, 2))
     .Flex.TextMatrix(I, 6) = Trim(DETRS!ICOD)
     
     .Flex.Rows = .Flex.Rows + 1
      I = I + 1
      DETRS.MoveNext
      If .Flex.Rows > 4 Then .Flex.TopRow = .Flex.TopRow + 2
    Loop
    
    .Flex.Rows = .Flex.Rows - 1
    .btn_sts (True)
    .cmdCancel.CANCEL = True
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
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

