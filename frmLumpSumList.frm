VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLumpSumList 
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
         Format          =   17760257
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
         Format          =   17760257
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Slip No."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LotNo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Finish Item"
            Object.Width           =   2717
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Grade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SubGrade"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Gross Weight"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Tare Weight"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Net Weight"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "VTYP"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
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
Attribute VB_Name = "frmLumpSumList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmLumpSumPacking.M_SRNO = Empty
    Unload Me
End Sub

Public Sub cmdOk_Click()
    Dim SEL_SRNO As String
    
    SEL_SRNO = lstBill.SelectedItem.SubItems(10)
    
    If Trim(SEL_SRNO) = Empty Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    Dim EDTDAT As New ADODB.Recordset
    Dim MSTDAT As New ADODB.Recordset
    Set EDTDAT = New ADODB.Recordset
    Set MSTDAT = New ADODB.Recordset
    Dim SQL As String
    
    SQL = Empty
      
    SQL = "SELECT PKGMAN.*,MACMST.NAME AS MACHINE,LOCMST.NAME AS LOCATION,SUBGRDMST.NAME AS SUBGRADE FROM PKGMAN "
    SQL = SQL & "INNER JOIN MACMST ON PKGMAN.COMP = MACMST.COMP AND PKGMAN.UNIT = MACMST.UNIT AND PKGMAN.DVCD = MACMST.DVCD "
    SQL = SQL & "INNER JOIN LOCMST ON PKGMAN.LOCCOD = LOCMST.CODE "
    SQL = SQL & "INNER JOIN SUBGRDMST ON PKGMAN.COMP = SUBGRDMST.COMP "
    SQL = SQL & "AND PKGMAN.UNIT = SUBGRDMST.UNIT AND PKGMAN.DVCD = SUBGRDMST.DVCD AND PKGMAN.GRAD = SUBGRDMST.GRAD AND "
    SQL = SQL & "PKGMAN.SUBGRAD = SUBGRDMST.SUBGRD WHERE PKGMAN.COMP='" & compPth & "' AND PKGMAN.UNIT='" & UNCD & _
    "' AND PKGMAN.DVCD='" & frmLumpSumPacking.LSDVCD & "' AND PKGMAN.VTYP='PPF' AND PKGMAN.DBCD='" & frmLumpSumPacking.M_DBCD & _
    "' AND PKGMAN.SRNO='" & SEL_SRNO & "' AND PKGMAN.RECSTAT<>'D'"
        
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmLumpSumPacking
         
    If MSTDAT.State = 1 Then MSTDAT.Close
    MSTDAT.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    "' AND DVCD='" & frmLumpSumPacking.LSDVCD & "' AND CODE='" & EDTDAT!FINITMCOD & "'", CN, adOpenDynamic, adLockOptimistic
    If Not MSTDAT.EOF Then
       .txtITEM = MSTDAT!NAME & ""
       .txtITEM.Tag = EDTDAT!FINITMCOD & ""
    Else
       .txtITEM.Text = Empty
    End If
        
        .M_SRNO = EDTDAT!SRNO
        .TXTMCCD = EDTDAT!MACHINE & ""
        .txtPCOD = GetCode("ACCMST", Trim(EDTDAT!PCOD & ""), "CODE", "NAME")
        .txtLoc = EDTDAT!LOCATION & ""
        .TXTSLIP = EDTDAT!SLIPNO & ""
        .TXTVBDT = EDTDAT!Date
        .TXTLOTNO = EDTDAT!LOTNO & ""
        .TXTGRAD = GetCode("GRDMST", EDTDAT!grad, "CODE", "GRAD")
        .TXTSUBGRD = EDTDAT!SUBGRADE & ""
        .TXTNOB = EDTDAT!NOB
        .txtCOPs = EDTDAT!CPB
        .TXTGRSWT = nstr(EDTDAT!GWPB, 12, 3)
        .TXTTAREWT = nstr(EDTDAT!TWPB, 12, 3)
        .TXTNETWT = nstr(EDTDAT!NWPB, 12, 3)
        .TXTCARTONNAME = GetCode("ITMMST", EDTDAT!BOX_COD, "CODE", "NAME")
        .TXTCOPSNAME = GetCode("ITMMST", EDTDAT!COPS_COD, "CODE", "NAME")
        .TXTIGRP = GetCode("ITMMST", EDTDAT!BOX_COD, "CODE", "IGCD")
        .TXTIGRP = GetCode("IGMMST", .TXTIGRP, "CODE", "NAME")
        .cmbPackaging.Text = GetCode("PKGNGMST", EDTDAT!PKGNG_COD & "", "CODE", "NAME")
        
        
        Dim i As Double, J As Double
        i = 0
        For i = 1 To .FLEXPLY.Cols - 1
           J = 0
           For J = 0 To EDTDAT.Fields.COUNT - 1
             If Trim(EDTDAT.Fields(J).NAME) = Trim(.FLEXPLY.TextMatrix(0, i)) Then
                 .FLEXPLY.TextMatrix(1, i) = Val(EDTDAT.Fields(J).Value)
             End If
           Next
        Next
     
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A' AND CODE='" & Trim(EDTDAT!PKGNG_COD & "") & "' AND PALLET='Y'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
 .FLEXPLY.Enabled = True
 .FLEXPLY.Tag = Val(Trim(RS!NOPLY & ""))
Else
 .FLEXPLY.Enabled = False
End If
RS.Close
    
End With
    
Unload Me
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    LBLUNTNAM.Caption = frmLumpSumPacking.TXTDVNM
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
    SQL = "SELECT DISTINCT PKGMAN.*,GRDMST.GRAD AS  GRADE,SUBGRDMST.NAME AS SUBGRADE,"
    SQL = SQL & "FINITMMST.NAME FROM PKGMAN INNER JOIN FINITMMST ON PKGMAN.COMP = FINITMMST.COMP "
    SQL = SQL & "AND PKGMAN.UNIT = FINITMMST.UNIT AND PKGMAN.DVCD = FINITMMST.DVCD AND "
    SQL = SQL & "PKGMAN.FINITMCOD = FINITMMST.CODE INNER JOIN GRDMST ON PKGMAN.GRAD = GRDMST.CODE "
    SQL = SQL & "INNER JOIN SUBGRDMST ON PKGMAN.COMP = SUBGRDMST.COMP "
    SQL = SQL & "AND PKGMAN.UNIT = SUBGRDMST.UNIT AND PKGMAN.DVCD = SUBGRDMST.DVCD AND PKGMAN.GRAD = SUBGRDMST.GRAD AND "
    SQL = SQL & "PKGMAN.SUBGRAD = SUBGRDMST.SUBGRD WHERE PKGMAN.COMP='" & compPth & _
    "' AND PKGMAN.UNIT='" & UNCD & "' AND PKGMAN.DVCD='" & frmLumpSumPacking.LSDVCD & "' AND VTYP='PPF' AND DBCD='" & frmLumpSumPacking.M_DBCD & _
    "' AND DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
    "' AND PKGMAN.RECSTAT<>'D' ORDER BY DATE,SLIPNO"
    
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
        lstItem.SubItems(1) = EDTDAT![SLIPNO]
        lstItem.SubItems(2) = EDTDAT![LOTNO]
        lstItem.SubItems(3) = Trim(EDTDAT![NAME])
        lstItem.SubItems(4) = EDTDAT![GRADE]
        lstItem.SubItems(5) = Trim(EDTDAT![SUBGRADE])
        lstItem.SubItems(6) = nstr(EDTDAT!GWPB, 12, 3)
        lstItem.SubItems(7) = nstr(EDTDAT!TWPB, 12, 3)
        lstItem.SubItems(8) = nstr(EDTDAT!NWPB, 12, 3)
        lstItem.SubItems(9) = EDTDAT!VTYP
        lstItem.SubItems(10) = EDTDAT!SRNO
        EDTDAT.MoveNext
    Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmLumpSumList.Visible = True Then lstBill.SetFocus
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

