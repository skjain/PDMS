VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPurServicesList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Service List"
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
         Format          =   50331649
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
         Format          =   50331649
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GRN No."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Service Provider"
            Object.Width           =   5116
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bill No"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Bill Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   2919
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Unique"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "VTYP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "SRNO"
            Object.Width           =   0
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
   Begin VB.Label LBLUNTNAM 
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
Attribute VB_Name = "frmPurServicesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmPurchaseServices.M_VBNO = Empty
    Unload Me
End Sub

Public Sub CMDOK_Click()
    Dim SEL_VBNO As String
    
    SEL_VBNO = lstBill.SelectedItem.SubItems(1)
    
    If Trim(SEL_VBNO) = Empty Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    Dim EDTDAT As New ADODB.Recordset
    Dim MSTDAT As New ADODB.Recordset
    Set EDTDAT = New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset
    Dim SQL As String
    
    SQL = Empty
    SQL = "SELECT * FROM GRN " & _
          "WHERE GRN.COMP='" & compPth & "' AND GRN.UNIT='" & UNCD & "' AND GRN.DVCD='" & DIVCOD & _
          "' AND GRN.VTYP='PSR' AND GRN.DBCD='" & frmPurchaseServices.M_DBCD_DIRIVR & _
          "' AND GRN.VBNO='" & SEL_VBNO & "' AND GRN.RECSTAT<>'D'"
        
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmPurchaseServices
         
        If MSTDAT.State = 1 Then MSTDAT.Close
        MSTDAT.Open "SELECT * FROM ACCMST WHERE CODE='" & EDTDAT!DRAC & "'", CN, adOpenDynamic, adLockOptimistic
        If Not MSTDAT.EOF Then
            .txtdbac = MSTDAT!NAME & ""
            .txtdbac.Tag = EDTDAT!DRAC & ""
        Else
            .txtdbac.Text = Empty
        End If
        
        .cmbSelection = Trim(EDTDAT!TTYP & "")
        .M_VBNO = EDTDAT!VBNO
        .txtvbno = EDTDAT!VBNO & ""
        .TXTVBDT = EDTDAT!Date
        
        .TXTNARR = EDTDAT!SDESC & ""
        .TXTMDESC = Trim(EDTDAT!MDESC & "")
        .TXTRMRK = Trim(EDTDAT!BRMK & "")
        
        .TXTSAMT = Val(EDTDAT!SAMT & "")
        .TXTITOT = Val(EDTDAT!ITOT & "")
        .TXTBNET = Val(EDTDAT!BNET & "")
        .VBNO = Trim(EDTDAT!CVBN & "")
        .VBDT = IIf(EDTDAT!LRDT = Null, Date, Date)
  
    End With
    Unload Me
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    LBLUNTNAM.Caption = " Division : " & DIVNAM
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
    SQL = "SELECT DISTINCT GRN.*,ACCMST.NAME FROM GRN INNER JOIN ACCMST ON GRN.DRAC=ACCMST.CODE WHERE COMP='" & compPth & _
    "' AND UNIT='" & UNCD & "' AND BSTS='P' AND DVCD='" & DIVCOD & _
    "' AND VTYP='PSR' AND DBCD='" & frmPurchaseServices.M_DBCD_DIRIVR & _
    "' AND DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
    "' AND RECSTAT<>'D' ORDER BY DATE,VBNO"
    
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
        lstItem.SubItems(1) = EDTDAT![VBNO] & ""
        lstItem.SubItems(2) = EDTDAT![NAME] & ""
        lstItem.SubItems(3) = Trim(EDTDAT!CVBN & "")
        lstItem.SubItems(4) = Format(EDTDAT!GATD, "dd/MM/yyyy")
        lstItem.SubItems(5) = Trim(EDTDAT!SDESC & "")
        lstItem.SubItems(6) = Format(EDTDAT!ITOT, "#############.00")
        EDTDAT.MoveNext
    Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmPurServicesList.Visible = True Then lstBill.SetFocus
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

