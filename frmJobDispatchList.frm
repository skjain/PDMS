VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmJobDispatchList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Delivery Challan"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDTRNGE 
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11295
      Begin VB.TextBox txtpcod 
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   5055
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
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24182785
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24182785
         CurrentDate     =   38429
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name : "
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
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1170
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
         TabIndex        =   2
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
         Left            =   240
         TabIndex        =   0
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.Frame frmIVR 
      Height          =   5130
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   11325
      Begin MSComctlLib.ListView lst 
         Height          =   4785
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8440
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr."
            Object.Width           =   443
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Chln Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Challan"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "A/c Party"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Consignee"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "LotNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Item Desc"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Grade"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "SubGrade"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Boxes"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Net Qnty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "dbcd"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&O.K"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmJobDispatchList.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Exit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmJobDispatchList.frx":0D8A
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Delivery Challan For Division : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmJobDispatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVCODE As String
Public DIVNAME As String
Public M_DBCD As String
Public VTCD As String
Public chln As String

Public Sub FillList()

lst.ListItems.Clear

Dim SQL As String
Dim M_ROW As Integer

Screen.MousePointer = vbHourglass
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=SPTRAN.GRAD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = SPTRAN.DCOD AND PADDMST.SRNO = SPTRAN.ADDRESS "
SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.DVCD='" & DIVCODE & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT='A' AND SPTRAN.DBCD='" & VTCD & _
"' AND SVBN IS NULL AND SPTRAN.EXTRA1 IS NULL "

SQL = SQL & " AND SPTRAN.DATE >= '" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
            "' AND SPTRAN.DATE <= '" & Format(txtToDate.Value, "MM/DD/YYYY") & "' "
If txtpcod <> Empty Then
   SQL = SQL & " AND SPTRAN.PCOD = '" & txtpcod.Tag & "'"
End If

SQL = SQL & " ORDER BY SPTRAN.VBNO,SPTRAN.DATE"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

Do While RECSET.EOF = False
    Set lstItm = lst.ListItems.ADD
    lstItm.Text = lst.ListItems.COUNT
    lstItm.SubItems(1) = CStr(Format(RECSET![Date], "dd/mm/yyyy"))
    lstItm.SubItems(2) = Trim(RECSET![VBNO] & "")
    lstItm.SubItems(3) = RECSET![ACNM]
    lstItm.SubItems(4) = Trim(RECSET!CONSINEE & "")
    lstItm.SubItems(5) = Trim(RECSET!ltno & "")
    lstItm.SubItems(6) = Trim(RECSET!ITNM & "")
    lstItm.SubItems(7) = Trim(RECSET!GRADE & "")
    lstItm.SubItems(8) = ""
    lstItm.SubItems(9) = Trim(RECSET!PCES & "")
    lstItm.SubItems(10) = Trim(RECSET!QNTY & "")
    lstItm.SubItems(11) = Trim(RECSET!dbcd & "")
    
    RECSET.MoveNext
Loop
     
     If lst.ListItems.COUNT > 0 Then
        lst.ListItems(1).Selected = True
        cmdOk.Default = True
     Else
        cmdOk.Default = False
    End If
Screen.MousePointer = vbNormal
End Sub

Private Sub CMDEXIT_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()
  Call FillList
End Sub

Private Sub CMDOK_Click()
If lst.ListItems.COUNT < 1 Then
 frmJobDispatch.chln = Empty
 Exit Sub
End If
Dim SQL As String

With frmJobDispatch
 .txtpcod = Empty: .txtCONSINEE = Empty: .TXTADDRESS = Empty: .TXTITEM = Empty: .txtLTNO = Empty: .TXTGRAD = Empty: .TXTSUBGRD = Empty
 .TXTRATE = Empty: .TXTRMRK = Empty
 .LBLS = 0: .LBLZ = 0: .LBL0 = 0
  
 chln = lst.SelectedItem.SubItems(2)
 .chln = lst.SelectedItem.SubItems(2)
 VTCD = lst.SelectedItem.SubItems(11)
 .VTCD = lst.SelectedItem.SubItems(11)
 
If Trim(chln) = Empty Then
     lstBill.SetFocus
     Exit Sub
End If
  
Set RECSET = New ADODB.Recordset

SQL = "SELECT SPTRAN.*,SUBGRDMST.NAME AS SUBGRADE,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=SPTRAN.GRAD "
SQL = SQL & "LEFT JOIN SUBGRDMST ON SPTRAN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND SPTRAN.UNIT = SUBGRDMST.UNIT AND SPTRAN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "SPTRAN.GRAD = SUBGRDMST.GRAD AND SPTRAN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = SPTRAN.DCOD AND PADDMST.SRNO = SPTRAN.ADDRESS "
SQL = SQL & "WHERE SPTRAN.COMP='" & compPth & "' AND SPTRAN.UNIT='" & UNCD & _
"' AND SPTRAN.DVCD='" & DIVCODE & "' AND SPTRAN.VTYP='DPF' AND SPTRAN.RECSTAT='A' AND SPTRAN.VBNO='" & chln & _
"' AND SPTRAN.DBCD = '" & VTCD & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found "
   Screen.MousePointer = vbNormal
   lstBill.SetFocus
   Exit Sub
End If
   
 If Trim(RECSET!ISRETURNABLE & "") = "Y" Then
    .chkReturnable.Value = 1
 Else
    .chkReturnable.Value = 0
 End If
   
 .LBLCHLN = Trim(RECSET!VBNO & "")
 .TXTVBDT = Format(RECSET!Date, "DD/MM/YYYY")
 .txtpcod = Trim(RECSET!ACNM & "")
 .txtCONSINEE = Trim(RECSET!CONSINEE & "")
 .TXTADDRESS = Trim(RECSET!ADDRESS & "")
 .TXTITEM = Trim(RECSET!ITNM & "")
 .txtLTNO = Trim(RECSET!ltno & "")
 .TXTGRAD = Trim(RECSET!GRADE & "")
 .txtLRNO = Trim(RECSET!LRNO & "")
 If Not IsNull(RECSET!LRDT) Then
    .LRDT = Format(RECSET!LRDT, "DD/MM/YYYY")
 End If
 .txtVHCL = GetCode("VHCLMST", Trim(RECSET!VEHICALNO & ""), "CODE", "NAME")
 .txtTransport = GetCode("TRANSPORTMST", Trim(RECSET!TRCD & ""), "CODE", "NAME")
 
 
 If Trim(RECSET!SUBGRD & "") = "S" Or Trim(RECSET!SUBGRD & "") = "Z" Or Trim(RECSET!SUBGRD & "") = "O" Then
   .TXTSUBGRD = Trim(RECSET!SUBGRD & "")
 Else
   .TXTSUBGRD = Trim(RECSET!SUBGRADE & "")
 End If
 
 
 .TXTRATE = Trim(RECSET!RATE & "")
 .TXTRMRK = Trim(RECSET!EXTRA4 & "")
 
 'INITIAL SET TOTAL BOX COPES
 .txtTTLCOPs = 0
 .txtTTLCTRN = 0
 .txtTTLNTWT = 0
 .txtRMNCOPs = 0
 .txtRMNCTRN = 0
 .txtRMNNTWT = 0
'===============================================
  
Set RSDATA = New ADODB.Recordset

SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & DIVCODE & _
"' AND RVTYP='DPF' AND BOXREGISTER.RECSTAT<>'D' AND RVBNO='" & Trim(RECSET!VBNO & "") & "' AND RDBC = '" & RECSET!dbcd & "' ORDER BY VBDT"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   'MsgBox "Boxes are not available for this criteria."
   'Exit Sub
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
   If Trim(RSDATA!SUBGRD & "") = "S" Or Trim(RSDATA!SUBGRD & "") = "Z" Or Trim(RSDATA!SUBGRD & "") = "0" Or Trim(RSDATA!SUBGRD & "") = "O" Then
     Item.SubItems(3) = Trim(RSDATA!SUBGRD & "")
     If .lstBox.SelectedItem.ListSubItems.COUNT = 2 Then .lstBox.ColumnHeaders(4).Text = "Twist"
     
     Select Case Trim(RSDATA!SUBGRD & "")
     Case "S"
       .LBLS = Val(.LBLS) + 1
       .LBLSWGT = Val(.LBLSWGT) + Val(Trim(RSDATA!NTWGT & ""))
     Case "Z"
       .LBLZ = Val(.LBLZ) + 1
       .LBLZWGT = Val(.LBLZWGT) + Val(Trim(RSDATA!NTWGT & ""))
     Case "0", "O"
       .LBL0 = Val(.LBL0) + 1
       .LBLOWGT = Val(.LBLOWGT) + Val(Trim(RSDATA!NTWGT & ""))
     End Select
        
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
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
   Item.SubItems(10) = Trim(RSDATA!Top & "")
   
   'PLY SETTING
   i = 0
   For i = 12 To .lstBox.ColumnHeaders.COUNT
      J = 0
      For J = 0 To RSDATA.Fields.COUNT - 1
        If Trim(RSDATA.Fields(J).NAME) = Trim(.lstBox.ColumnHeaders(i).Text) Then
            Item.SubItems(i - 1) = Val(RSDATA.Fields(J).Value)
        End If
      Next
   Next
   '--------------
      
    .txtTTLCOPs = Val(.txtTTLCOPs) + Val(RSDATA!COPS)
    .txtTTLCTRN = Val(.txtTTLCTRN) + 1
    .txtTTLNTWT = Val(.txtTTLNTWT) + Val(RSDATA!NTWGT)
      
  RSDATA.MoveNext
  Loop
  RSDATA.Close
  
  .txtCOPs = COPS
  .txtCTRN = BOXES
  .txtNTWT = NETWT
  .txtNTWT.Tag = NETWT
  
'---------------------------------------
SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & DIVCODE & _
"'AND BOXREGISTER.LOTNO ='" & Trim(RECSET!ltno & "") & "' AND BOXREGISTER.ICOD ='" & Trim(RECSET!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(RECSET!grad & "") & "' AND (VTYP='PPF' OR VTYP='OPN') AND " & _
"BOXREGISTER.RECSTAT<>'D' AND RVBNO IS NULL "

SQL = SQL & " AND BOXREGISTER.VBDT <= '" & Format(RECSET!Date, "MM/DD/YYYY") & "' "

      Dim XRS As ADODB.Recordset
      Set XRS = New ADODB.Recordset
        
      If XRS.State = 1 Then XRS.Close
      XRS.Open "SELECT TWSTREQ,CFGTYP FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND CODE='" & DIVCODE & "'", CN, adOpenDynamic, adLockOptimistic
      If Not XRS.EOF Then
         If Trim(XRS!CFGTYP) = "SG" And Trim(XRS!TWSTREQ) = "N" Then
            SQL = SQL & " AND BOXREGISTER.SUBGRD ='" & Trim(RECSET!SUBGRD & "") & "'"
         End If
      End If
      XRS.Close

If InStr(1, UCase(.cmbPackingType.Text), "JOB CHALLAN") <> 0 Then
   SQL = SQL & "AND DBCD='000005' AND PCOD='" & GetCode("ACCMST", .txtpcod, "NAME", "CODE") & "' ORDER BY VBDT"
ElseIf InStr(1, UCase(.cmbPackingType.Text), "CAPTIVE") <> 0 Then
   SQL = SQL & "AND DBCD='000001' ORDER BY VBDT"
ElseIf InStr(1, UCase(.cmbPackingType.Text), "EXPORT") <> 0 Then
   SQL = SQL & "AND DBCD='000002' ORDER BY VBDT"
Else
   SQL = SQL & "AND DBCD NOT IN ('000001','000002','000005') ORDER BY VBDT"
End If

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
   If Trim(RSDATA!SUBGRD & "") = "S" Or Trim(RSDATA!SUBGRD & "") = "Z" Or Trim(RSDATA!SUBGRD & "") = "0" Or Trim(RSDATA!SUBGRD & "") = "O" Then
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
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
   Item.SubItems(10) = Trim(RSDATA!Top & "")
   
   i = 0
   For i = 12 To .lstBox.ColumnHeaders.COUNT
      J = 0
      For J = 0 To RSDATA.Fields.COUNT - 1
        If Trim(RSDATA.Fields(J).NAME) = Trim(.lstBox.ColumnHeaders(i).Text) Then
            Item.SubItems(i - 1) = Val(RSDATA.Fields(J).Value)
        End If
      Next
   Next
   
   .txtTTLCOPs = Val(.txtTTLCOPs) + Val(RSDATA!COPS)
   .txtTTLCTRN = Val(.txtTTLCTRN) + 1
   .txtTTLNTWT = Val(.txtTTLNTWT) + Val(RSDATA!NTWGT)
   
  RSDATA.MoveNext
  Loop
  RSDATA.Close
  
  .txtRMNCOPs = Val(.txtTTLCOPs) - Val(.txtCOPs)
  .txtRMNCTRN = Val(.txtTTLCTRN) - Val(.txtCTRN)
  .txtRMNNTWT = Val(.txtTTLNTWT) - Val(.txtNTWT)
  
End With
Me.Hide
Unload Me
End Sub

Private Sub Form_Activate()
 txtFrDate.SetFocus
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
txtFrDate.Value = GetMinDate
 txtToDate.Value = GetMaxDate
 LBLHEAD.Caption = LBLHEAD.Caption + DIVNAME
'Call FillList
End Sub


Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtpcod = Empty
      Key = Empty
  ElseIf KeyCode = vbKeyF2 Then
     txtpcod = Empty:   Key = Empty:  NEW_VISIBLE = False
     txtpcod = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
     txtpcod.Tag = Key
  End If
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub
