VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCaptiveBoxChlnList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captive Challan List"
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
         Format          =   50462721
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
         Format          =   50462721
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
         Left            =   120
         TabIndex        =   11
         Top             =   120
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
            Text            =   "Challan No."
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "LotNo"
            Object.Width           =   2213
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Finish Item"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Raw Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Grade"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "SubGrade"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Chln Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   75
      TabIndex        =   8
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
         Left            =   8040
         TabIndex        =   9
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Label LBLTDIVNAM 
      Alignment       =   1  'Right Justify
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
      Left            =   5880
      TabIndex        =   12
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label LBLFDIVNAM 
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
      Width           =   4800
   End
End
Attribute VB_Name = "frmCaptiveBoxChlnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FDVCD As String
Public TDVCD As String
Public M_DBCD As String

Private Sub cmdCancel_Click()
    frmCaptiveBoxChallan.CHALLAN = Empty
    Unload Me
End Sub

Public Sub cmdOk_Click()
    Dim CHLNNO As String
    Dim RSDATA As ADODB.Recordset
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
 SQL = "SELECT SPTRAN.*,FINITMMST.NAME AS ITEM,ITMMST.NAME AS RAWITEM,SUBGRDMST.NAME AS SUBGRADE FROM SPTRAN "
 SQL = SQL & "INNER JOIN FINITMMST ON SPTRAN.COMP = FINITMMST.COMP AND SPTRAN.UNIT = FINITMMST.UNIT AND SPTRAN.DVCD = FINITMMST.DVCD AND SPTRAN.ICOD = FINITMMST.CODE "
 SQL = SQL & "LEFT JOIN SUBGRDMST ON SPTRAN.COMP = SUBGRDMST.COMP AND SPTRAN.UNIT = SUBGRDMST.UNIT AND SPTRAN.DVCD = SUBGRDMST.DVCD AND SPTRAN.GRAD = SUBGRDMST.GRAD AND SPTRAN.SUBGRD = SUBGRDMST.SUBGRD "
 SQL = SQL & "LEFT JOIN ITMMST ON SPTRAN.EXTRA2 = ITMMST.CODE "
 SQL = SQL & " WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
 "' AND SPTRAN.DVCD='" & FDVCD & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.DBCD='" & M_DBCD & _
 "' AND SPTRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND SPTRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
 "' AND SPTRAN.RECSTAT<>'D' AND SPTRAN.EXTRA3='" & TDVCD & "' AND SPTRAN.VBNO='" & CHLNNO & "'"
           
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmCaptiveBoxChallan
    
    'FIND MACHINE NAME IF EXIST
        If (EDTDAT!PCOD & "" <> "") Then .TXTMACHINE = FindMachineName(EDTDAT!PCOD & "")
        '---------------------------
    .CHALLAN = CHLNNO
    .lblBill.Caption = CHLNNO
    .TXTVBDT = Format(EDTDAT!Date & "", "DD/MM/YYYY")
    .txtLTNo.Text = EDTDAT!ltno & ""
    .TXTITM.Text = EDTDAT!Item & ""
    .TXTINAM.Text = EDTDAT!RAWITEM & ""
    .txtIGRP.Text = GetCode("IGMMST", GetCode("ITMMST", Trim(EDTDAT!EXTRA2 & ""), "CODE", "IGCD"), "CODE", "NAME")
    .RAWITMGRP = Trim(EDTDAT!EXTRA2 & "")
    .TXTGRAD.Text = GetCode("GRDMST", EDTDAT!grad & "", "CODE", "GRAD")
    .TXTSUBGRD.Text = EDTDAT!SUBGRADE & ""
    .TXTRATE.Text = EDTDAT!RATE & ""
    .BRMK = EDTDAT!EXTRA1 & ""
    
    
    '-----------------------------------------------------------------------------------------
    Set RSDATA = New ADODB.Recordset

SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & FDVCD & _
"'AND BOXREGISTER.LOTNO ='" & Trim(EDTDAT!ltno & "") & "' AND BOXREGISTER.ICOD = '" & Trim(EDTDAT!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(EDTDAT!grad & "") & _
"' AND RVTYP='DPF' AND BOXREGISTER.RECSTAT<>'D' AND RVBNO ='" & CHLNNO & "' AND RDBC = '" & EDTDAT!dbcd & "' ORDER BY VBDT"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   Exit Sub
End If

Dim COPS As Long: COPS = 0
Dim BOXES As Long: BOXES = 0
Dim NETWT As Double: NETWT = 0

  .lstBox.ListItems.Clear
  Do While Not RSDATA.EOF
   Set Item = .lstBox.ListItems.ADD
   Item.Text = RSDATA!VBNO
   Item.Checked = True
      BOXES = BOXES + 1
   Item.SubItems(1) = Val(Trim(RSDATA!COPS & ""))
      COPS = COPS + Val(Trim(RSDATA!COPS & ""))
   Item.SubItems(2) = nstr(RSDATA!NTWGT, 9, 3)
      NETWT = NETWT + Val(Trim(RSDATA!NTWGT & ""))
      
   Item.SubItems(2) = Trim(Item.SubItems(2))
   
   If Trim(RSDATA!SUBGRD & "") = "S" Or Trim(RSDATA!SUBGRD & "") = "Z" Or Trim(RSDATA!SUBGRD & "") = "0" Then
     Item.SubItems(3) = Trim(RSDATA!SUBGRD & "")
     .lstBox.ColumnHeaders(4).Text = "Twist"
   Else
     Item.SubItems(3) = Trim(RSDATA!SUBGRADE & "")
     .lstBox.ColumnHeaders(4).Text = "SubGrade"
   End If
   
   Item.SubItems(4) = nstr(RSDATA!GRSWGT, 9, 3)
   Item.SubItems(4) = Trim(Item.SubItems(4))
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5))
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = RSDATA!RMRK
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
  RSDATA.MoveNext
  Loop
  RSDATA.Close
  
  .txtCops = COPS
  .txtCTRN = BOXES
  .txtNTWT = NETWT
  .txtNTWT.Tag = NETWT
  
    '---------------------------------------
SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & FDVCD & _
"'AND BOXREGISTER.LOTNO ='" & Trim(EDTDAT!ltno & "") & "' AND BOXREGISTER.ICOD = '" & Trim(EDTDAT!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(EDTDAT!grad & "") & _
"' AND VTYP IN ('OPN','PPF')  AND BOXREGISTER.RECSTAT<>'D' AND RVBNO IS NULL "

SQL = SQL & " ORDER BY VBDT"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   Me.Hide
   Unload Me
   Exit Sub
End If
   Do While Not RSDATA.EOF
   Set Item = .lstBox.ListItems.ADD
   Item.Text = RSDATA!VBNO
   Item.SubItems(1) = RSDATA!COPS
   Item.SubItems(2) = nstr(RSDATA!NTWGT, 9, 3)
   Item.SubItems(2) = Trim(Item.SubItems(2))
   If Trim(RSDATA!SUBGRD & "") = "S" Or Trim(RSDATA!SUBGRD & "") = "Z" Or Trim(RSDATA!SUBGRD & "") = "0" Then
     Item.SubItems(3) = Trim(RSDATA!SUBGRD & "")
     If .lstBox.SelectedItem.ListSubItems.COUNT = 2 Then .lstBox.ColumnHeaders(4).Text = "Twist"
   Else
     Item.SubItems(3) = Trim(RSDATA!SUBGRADE & "")
     If .lstBox.ListItems.COUNT = 1 Then .lstBox.ColumnHeaders(4).Text = "SubGrade"
   End If
   
   Item.SubItems(4) = nstr(RSDATA!GRSWGT, 9, 3)
   Item.SubItems(4) = Trim(Item.SubItems(4))
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5))
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = RSDATA!RMRK
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
  RSDATA.MoveNext
  Loop
  RSDATA.Close
      
    End With
    Unload Me
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    LBLFDIVNAM.Caption = frmCaptiveBoxChallan.TXTFROMDIV
    LBLTDIVNAM.Caption = frmCaptiveBoxChallan.TXTTODIV
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
 
 SQL = "SELECT DISTINCT SPTRAN.*,FINITMMST.NAME AS ITEM,ITMMST.NAME AS RAWITEM,SUBGRDMST.NAME AS SUBGRADE FROM SPTRAN "
 SQL = SQL & "INNER JOIN FINITMMST ON SPTRAN.COMP = FINITMMST.COMP AND SPTRAN.UNIT = FINITMMST.UNIT AND SPTRAN.DVCD = FINITMMST.DVCD AND SPTRAN.ICOD = FINITMMST.CODE "
 SQL = SQL & "LEFT JOIN SUBGRDMST ON SPTRAN.COMP = SUBGRDMST.COMP AND SPTRAN.UNIT = SUBGRDMST.UNIT AND SPTRAN.DVCD = SUBGRDMST.DVCD AND SPTRAN.GRAD = SUBGRDMST.GRAD AND SPTRAN.SUBGRD = SUBGRDMST.SUBGRD "
 SQL = SQL & "LEFT JOIN ITMMST ON SPTRAN.EXTRA2 = ITMMST.CODE "
 SQL = SQL & " WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
 "' AND SPTRAN.DVCD='" & FDVCD & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.DBCD='" & M_DBCD & _
 "' AND SPTRAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND SPTRAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
 "' AND SPTRAN.RECSTAT<>'D' AND SPTRAN.EXTRA3='" & TDVCD & "' ORDER BY DATE,VBNO"
    
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
    lstItem.SubItems(2) = EDTDAT![ltno]
    lstItem.SubItems(3) = EDTDAT![Item]
    lstItem.SubItems(4) = Trim(EDTDAT![RAWITEM] & "")
    lstItem.SubItems(5) = GetCode("GRDMST", EDTDAT![grad], "CODE", "GRAD")
    lstItem.SubItems(6) = EDTDAT![SUBGRADE] & ""
    lstItem.SubItems(7) = EDTDAT![QNTY]
    lstItem.SubItems(8) = EDTDAT![RATE]
    lstItem.SubItems(9) = EDTDAT![AMNT]
    EDTDAT.MoveNext
 Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmCaptiveBoxChlnList.Visible = True Then lstBill.SetFocus
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

Private Function FindMachineName(CODE As String) As String
On Error GoTo LAST:
Dim TMPRS As ADODB.Recordset
Set TMPRS = New ADODB.Recordset
If TMPRS.State = 1 Then TMPRS.Close
TMPRS.Open "SELECT NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND DVCD='" & TDVCD & "' AND CODE='" & CODE & "'", CN, adOpenDynamic, adLockOptimistic
If Not TMPRS.EOF Then
   FindMachineName = Trim(TMPRS!NAME & "")
Else
   FindMachineName = ""
End If
TMPRS.Close
Exit Function
LAST:
MsgBox ERR.Description
End Function
