VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmProcessedJobList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue List Help ItemWise"
   ClientHeight    =   7275
   ClientLeft      =   2130
   ClientTop       =   2400
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10399.67
   ScaleMode       =   0  'User
   ScaleWidth      =   10080
   Begin VB.Frame frmDTRNGE 
      Height          =   1320
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   9915
      Begin VB.OptionButton optJob 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&JOB WORK RECEIVED"
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
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton optRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&RETURNABLE RECEIVED"
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
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
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
         TabIndex        =   7
         Top             =   840
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4320
         TabIndex        =   6
         Top             =   840
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
         Format          =   50790401
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   840
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
         Format          =   50790401
         CurrentDate     =   38429
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of &Transaction    (A)                                                      (B)    "
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
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   6255
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
         TabIndex        =   3
         Top             =   840
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
         TabIndex        =   5
         Top             =   840
         Width           =   930
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5235
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   9900
      Begin MSComctlLib.ListView lstBill 
         Height          =   5025
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   9690
         _ExtentX        =   17092
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
         NumItems        =   7
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
            Text            =   "Party Name"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Service Desc."
            Object.Width           =   4516
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Service Amt"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Material Desc"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Material Cost"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   13
      Top             =   6600
      Width           =   9855
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   195
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FrmProcessedJobList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LVTYP As String

Private Sub CMDGO_Click()
  lstBill.ListItems.Clear
  
  Call SetVTYP
  
  Dim edtdat As New ADODB.Recordset
  Set edtdat = New ADODB.Recordset
  Dim sql As String
  
  sql = Empty
  
  sql = "SELECT GRN.*,ACCMST.NAME AS PARTY FROM GRN " & _
        "INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE " & _
        "WHERE GRN.COMP='" & compPth & "' AND GRN.UNIT='" & UNCD & _
        "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
        "' AND BSTS='P' " & _
        "  AND VTYP='IVR' AND GRN.DBCD='" & FrmProcessedJob.m_dbcd & "' AND GRN.RECSTAT<>'D' " & _
        "ORDER BY GRN.DATE,GRN.VBNO"
   
  If edtdat.State = 1 Then edtdat.Close
  edtdat.Open sql, CN, adOpenDynamic, adLockOptimistic
  If edtdat.EOF Then
    MsgBox "No Record found for given criteria ", vbInformation
    txtToDate.SetFocus
    Exit Sub
  End If
  
  Do While Not edtdat.EOF
   Set lstitem = lstBill.ListItems.ADD
   lstitem.Text = Format(edtdat![Date], "dd/MM/yyyy")
   lstitem.SubItems(1) = edtdat![VBNO]
   lstitem.SubItems(2) = edtdat![PARTY]
   lstitem.SubItems(3) = Trim(edtdat![SDESC] & "")
   lstitem.SubItems(4) = Format(edtdat!SAMT, "########.000")
   lstitem.SubItems(5) = Trim(edtdat![MDESC] & "")
   lstitem.SubItems(6) = Format(edtdat!ITOT, "########.000")
   edtdat.MoveNext
  Loop
  
  CMDOK.Enabled = True
  CMDOK.Default = True
  If FrmProcessedJobList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub cmdCancel_Click()
  Unload Me
  FrmProcessedJob.m_dbcd = Empty
End Sub

Private Sub cmdOk_Click()
  If lstBill.SelectedItem.SubItems(1) <> Empty And lstBill.ListItems.COUNT < 1 Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  With FrmProcessedJob
  Dim VBNO As String
  
  Dim edtdat As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set edtdat = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  Call SetVTYP
  
  Dim sql As String
  sql = Empty
  
   sql = "SELECT GRN.*,ACCMST.NAME AS PARTY,TRANSPORTMST.NAME AS TRANSPORT FROM GRN " & _
        "INNER JOIN ACCMST ON GRN.PCOD=ACCMST.CODE " & _
        "LEFT JOIN TRANSPORTMST ON GRN.TRCD=TRANSPORTMST.CODE " & _
        "WHERE GRN.COMP='" & compPth & "' AND GRN.UNIT='" & UNCD & _
        "' AND GRN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
        "' AND GRN.VTYP='IVR' AND GRN.DBCD='" & FrmProcessedJob.m_dbcd & "' AND GRN.RECSTAT<>'D' " & _
        "  AND GRN.VBNO='" & lstBill.SelectedItem.SubItems(1) & "' AND GRN.ACEFFECT <> 'N'"
    
        
  If edtdat.State = 1 Then edtdat.Close
  edtdat.Open sql, CN, adOpenDynamic, adLockOptimistic
  If edtdat.EOF Then
     lstBill.SetFocus
     Exit Sub
  Else
     .txtdbac = edtdat!PARTY & ""
     
     .TXTSCHLN = Trim(edtdat!chln & "")    'CHALLAN
     .TXTSCHLNDT = edtdat!CHDT & ""
     .TXTSBILLNO = Trim(edtdat!CVBN & "") 'BILL
     .TXTSBILLDATE = edtdat!GATD & ""
     .TXTVBNO = edtdat!VBNO & ""                         'RECEIVE NO.
     .TXTVBDT = edtdat!Date & ""
     
     .TXTLRNO = Trim(edtdat!LRNO & "")
     .TXTLRDT = IIf(edtdat!LRDT <> Null, edtdat!LRDT, Date)
     .TXTTRNM = edtdat!transport & ""
     .txtVHCL = Trim(edtdat!VHCL & "")
     .TXTRMRK = Trim(edtdat!BRMK & "")
     
     .optJob.Value = optJob.Value
     .optRGP.Value = optRGP.Value
     
     .TXTSAMT = Val(edtdat!SAMT)
     .TXTITOT = Val(edtdat!ITOT)
     .TXTSDESC = Trim(edtdat!SDESC & "")
     .TXTMDESC = Trim(edtdat!MDESC & "")
     
     
        Dim I As Double
        Dim J As Double
        I = 0
        For I = 0 To .flexBTRM.Rows - 1
          J = 0
          For J = 0 To edtdat.Fields.COUNT - 1
            If Trim(edtdat.Fields(J).NAME) = Trim(.flexBTRM.TextMatrix(I, 0)) Then
                .flexBTRM.TextMatrix(I, 2) = Format(edtdat.Fields(J).Value, "#########.00")
            End If
            If Trim(edtdat.Fields(J).NAME) = "PER" & Trim(.flexBTRM.TextMatrix(I, 0)) Then
               .flexBTRM.TextMatrix(I, 1) = Format(edtdat.Fields(J).Value, "######.00")
            End If
          Next
        Next
     
  End If
          
    'FIFO-----------------------------
      If FIFOREQ = "Y" Then
         If RS.State = 1 Then RS.Close
         RS.Open "SELECT * FROM GRNTRAN WHERE COMP ='" & compPth & "' AND UNIT ='" & UNCD & _
         "' AND DBCD='" & FrmProcessedJob.m_dbcd & "' AND VBNO = '" & edtdat!VBNO & "" & "' AND GRN_QNTY <> BAL_QNTY ", CN, adOpenDynamic, adLockReadOnly
         If RS.EOF = False Then .M_ISSUE = "Y"
      End If
    '---------------------------------
    
    Dim DETRS As ADODB.Recordset
    Set DETRS = New ADODB.Recordset
    
    'SQL = "SELECT JOBOUT.*,ITMMST.NAME AS ITEM,ACCMST.NAME AS PARTY,TRANSPORTMST.NAME AS TRANSPORT FROM JOBOUT " & _
  "INNER JOIN TRANSPORTMST ON JOBOUT.TRCD=TRANSPORTMST.CODE INNER JOIN ITMMST ON (JOBOUT.ICOD=ITMMST.CODE) " & _
  "INNER JOIN ACCMST ON (JOBOUT.PCOD=ACCMST.CODE) WHERE JOBOUT.COMP='" & compPth & "' AND JOBOUT.UNIT='" & UNCD & _
  "' AND VTYP='IVR' AND DBCD='" & FrmProcessedJob.M_DBCD & "' AND  JOBOUT.VBNO='" & lstBill.SelectedItem.SubItems(1) & _
  "' AND JOBOUT.RECSTAT<>'D' AND JOBOUT.CLRSTATUS='N' ORDER BY VBNO,SRCH"
    
    If DETRS.State = 1 Then DETRS.Close
    DETRS.Open "SELECT JOBOUT.*,ITMMST.NAME AS ITEM FROM JOBOUT " & _
               "INNER JOIN ITMMST ON JOBOUT.ICOD=ITMMST.CODE WHERE JOBOUT.COMP='" & compPth & _
               "' AND JOBOUT.UNIT='" & UNCD & "' AND " & _
               "DBCD='" & FrmProcessedJob.m_dbcd & "' AND VTYP='IVR' AND VBNO='" & lstBill.SelectedItem.SubItems(1) & _
               "' AND RECSTAT<>'D' ORDER BY VBNO,SRCH", CN, adOpenDynamic, adLockOptimistic
    
    .Flex.Rows = 2
    
    I = 1
    Do While Not DETRS.EOF
     .chkReturnable.Value = IIf(Trim(DETRS!Mode) = "Y", 1, 0)
     .FRMLRDTL.Tab = 0
     .Flex.TextMatrix(I, 0) = Trim(DETRS!SRCH)
     .Flex.TextMatrix(I, 1) = Trim(DETRS!RECNO)
     .Flex.TextMatrix(I, 2) = Trim(DETRS!Item)
     .Flex.TextMatrix(I, 3) = Trim(DETRS!ltno)      'Trim(nstr(DETRS!QNTY, 12, 3))
     .Flex.TextMatrix(I, 4) = Trim(DETRS!COPS)
     .Flex.TextMatrix(I, 5) = Val(DETRS!PCES)
     
     .TXTTPCS = Val(.TXTTPCS) + Val(DETRS!PCES)
     .TXTTQTY = Val(.TXTTQTY) + Val(DETRS!QNTY)
     .TXTTAMT = nstr(Val(.TXTTAMT) + Val(DETRS!AMNT), 10, 2)
     
     .Flex.TextMatrix(I, 6) = Trim(nstr(DETRS!QNTY, 10, 3))
     .Flex.TextMatrix(I, 7) = Trim(nstr(DETRS!RATE, 10, 2))
     .Flex.TextMatrix(I, 8) = Trim(nstr(DETRS!AMNT, 10, 2))
     .Flex.TextMatrix(I, 9) = Trim(DETRS!ICOD)
     .Flex.Rows = .Flex.Rows + 1
      I = I + 1
      DETRS.MoveNext
    If .Flex.Rows > 4 Then .Flex.TopRow = .Flex.TopRow + 2
    Loop
    
    .Flex.Rows = .Flex.Rows - 1
    .btn_sts (True)
    .cmdCancel.Cancel = True
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  CMDOK.Enabled = False
  cmdCancel.Enabled = True
  Me.KeyPreview = True
End Sub

Private Sub optJob_Click()
Call CMDGO_Click
End Sub

Private Sub optJob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub optNRGP_Click()
  Call CMDGO_Click
End Sub

Private Sub optRGP_Click()
  Call CMDGO_Click
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

Public Sub SetVTYP()
With FrmProcessedJob
    If optJob.Value = True Then
       .m_dbcd = "000003"
       .optJob.Value = True
    ElseIf optRGP.Value = True Then
       .m_dbcd = "000004"
       .optRGP.Value = True
    End If
End With
End Sub
