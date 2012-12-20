VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmDOList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Order List For Box Wise Dispatch Against Order"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6282.436
   ScaleMode       =   0  'User
   ScaleWidth      =   14996.16
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPCOD 
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Width           =   5295
   End
   Begin VB.Frame frmIVR 
      Height          =   5010
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   13365
      Begin MSComctlLib.ListView lst 
         Height          =   4785
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   13095
         _ExtentX        =   23098
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr."
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "D.O No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "A/c Party"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivery Party"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Agent"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item Desc"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "LotNo"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Grade"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "SubGrade"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Quantity"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Tax/Retail "
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Remark"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "DBCD"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.TextBox TXTDONO 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      Image           =   "frmDOList.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   12120
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "E&xit"
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
      Image           =   "frmDOList.frx":0D8A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSEARCH 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Search"
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
      Image           =   "frmDOList.frx":11DC
      cBack           =   -2147483633
   End
   Begin VB.Label Label5 
      Caption         =   "Party &Name       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Delivery Order"
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
      Width           =   13335
   End
   Begin VB.Label LBLORD 
      Caption         =   "D.O.No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Press Enter To See the List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   6000
      Width           =   3015
   End
End
Attribute VB_Name = "frmDOList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim ORDN As String
Dim DONO As String
Public PKG_DBCD As String
Dim M_DBCD As String
Public DIVCODE As String
Dim RECSET As ADODB.Recordset
Dim RSDATA As ADODB.Recordset

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CMDOK_Click()
If lst.ListItems.COUNT < 1 Then
 Exit Sub
End If

Dim SQL As String
With frmBoxDispatch
 .txtpcod = Empty: .txtCONSINEE = Empty: .TXTADDRESS = Empty: .txtITEM = Empty: .txtLTNo = Empty: .TXTGRAD = Empty: .TXTSUBGRD = Empty
 .M_RTTX = Empty: .txtDONO = Empty: .txtQTY = Empty: .M_DORAT = Empty: .M_ARAT = Empty: .TXTRMRK = Empty
  
 ORDN = lst.SelectedItem.SubItems(1)
 .ORDN = lst.SelectedItem.SubItems(1)
 DONO = lst.SelectedItem.SubItems(3)
 .DONO = lst.SelectedItem.SubItems(3)
 M_DBCD = lst.SelectedItem.SubItems(14)
 .M_DBCD = lst.SelectedItem.SubItems(14)

If Trim(DONO) = Empty Then
     lstBill.SetFocus
     Exit Sub
End If
  
Set RECSET = New ADODB.Recordset

SQL = "SELECT ORDTRN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE,"
SQL = SQL & "REFMST.NAME AS AGENT,PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM ORDTRN INNER JOIN ACCMST "
SQL = SQL & "ON ACCMST.CODE=ORDTRN.PCOD INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND "
SQL = SQL & "FINITMMST.UNIT=ORDTRN.UNIT AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD LEFT JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND ORDTRN.UNIT = SUBGRDMST.UNIT AND ORDTRN.DVCD = SUBGRDMST.DVCD AND ORDTRN.GRAD = SUBGRDMST.GRAD "
SQL = SQL & "AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD INNER JOIN REFMST ON REFMST.CODE = ORDTRN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = ORDTRN.DCOD AND PADDMST.SRNO = ORDTRN.SRCH WHERE ORDTRN.COMP='" & compPth & _
"' AND ORDTRN.UNIT='" & UNCD & "' AND ORDTRN.VTYP='DOS' AND ORDTRN.DFLG<>'Y' AND ORDTRN.RECSTAT='A' AND ORDTRN.DOSTAT='Y' "
SQL = SQL & "AND ORDTRN.DONO='" & DONO & "' AND  ORDTRN.DBCD='" & M_DBCD & "' AND ORDTRN.ORDN='" & ORDN & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found "
   Screen.MousePointer = vbNormal
   lstBill.SetFocus
   Exit Sub
End If
   
 .txtpcod = Trim(RECSET!ACNM & "")
 .txtCONSINEE = Trim(RECSET!CONSINEE & "")
 .TXTADDRESS = Trim(RECSET!ADDRESS & "")
 .txtITEM = Trim(RECSET!ITNM & "")
 .txtLTNo = Trim(RECSET!ltno & "")
 .TXTGRAD = Trim(RECSET!GRADE & "")
 If Trim(RECSET!SUBGRD & "") = "S" Or Trim(RECSET!SUBGRD & "") = "Z" Or Trim(RECSET!SUBGRD & "") = "0" Then
   .TXTSUBGRD = Trim(RECSET!SUBGRD & "")
 Else
   .TXTSUBGRD = Trim(RECSET!SUBGRADE & "")
 End If
 .txtAgent = Trim(RECSET!AGENT & "")
 .M_RTTX = Trim(RECSET!TXRT & "")
 .TXTORDN = Trim(RECSET!ORDN & "")
 .txtDONO = Trim(RECSET!DONO & "")
 .dtDate = Trim(RECSET!DODT & "")
 .txtQTY = Trim(RECSET!QNTY & "")
 .M_DORAT = Trim(RECSET!RATE & "")
 .M_ARAT = Trim(RECSET!ARAT & "")
 .TXTRMRK = Trim(RECSET!BRMK & "")
 
 
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
"'AND BOXREGISTER.LOTNO ='" & Trim(RECSET!ltno & "") & "' AND BOXREGISTER.ICOD = '" & Trim(RECSET!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(RECSET!grad & "") & "' AND (VTYP='PPF' OR VTYP='OPN') AND BOXREGISTER.RECSTAT<>'D' AND RVBNO IS NULL AND DBCD NOT IN('000001','000005') "

If IsSubGradeReq(DIVCODE) Or IsTwistReq(DIVCODE) = "Y" Then
   If IsTwistReq(DIVCODE) = "Y" And Trim(RECSET!SUBGRD & "") = "0" Then
      
   Else
      SQL = SQL & " AND BOXREGISTER.SUBGRD ='" & Trim(RECSET!SUBGRD & "") & "' "
   End If
End If

'BOX DATE ARE LESS THEN OR EQUAL TO CHALLAN DATE
SQL = SQL & " AND VBDT <= '" & Format(frmBoxDispatch.TXTVBDT.Value, "MM/DD/YYYY") & "' "
'---------------------------------------------------

If frmBoxDispatch.CSVTable <> Empty Then
   SQL = SQL & "AND VBNO IN (SELECT * FROM " & frmBoxDispatch.CSVTable & ") "
End If

If PKG_DBCD = Empty Then
   SQL = SQL & " AND DBCD NOT IN('000002') "
Else
   SQL = SQL & " AND DBCD IN('000002') "
End If

SQL = SQL & " ORDER BY VBNO"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   Exit Sub
End If
  .lstBox.ListItems.Clear
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
   Item.SubItems(4) = Trim(Item.SubItems(4) & "")
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5) & "")
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = Trim(RSDATA!RMRK & "")
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
   Item.SubItems(10) = Trim(RSDATA![Top] & "")
   
   'FOR PALLETNO
   Item.SubItems(11) = Trim(RSDATA![PLTNO] & "")
   '--------------------------------------------------
   
   Dim i As Double, J As Double
   i = 0
   For i = 13 To .lstBox.ColumnHeaders.COUNT
      J = 0
      For J = 0 To RSDATA.Fields.COUNT - 1
        If Trim(RSDATA.Fields(J).NAME) = Trim(.lstBox.ColumnHeaders(i).Text) Then
            Item.SubItems(i - 1) = Val(RSDATA.Fields(J).Value & "")
        End If
      Next
   Next
   
     .txtTTLCOPs = Val(.txtTTLCOPs) + Val(RSDATA!COPS)
     .txtTTLCTRN = Val(.txtTTLCTRN) + 1
     .txtTTLNTWT = Val(.txtTTLNTWT) + Val(RSDATA!NTWGT)
   
  RSDATA.MoveNext
  Loop
  RSDATA.Close
  
  If SetIsShadeReq(DIVCODE) = "Y" Then
     .lstBox.ColumnHeaders(4).Text = "Shade"
  End If
  
End With
Me.Hide
Unload Me
End Sub

Private Sub cmdSearch_Click()
If txtpcod <> Empty Then
   Call FillList(" AND ORDTRN.PCOD ='" & Trim(txtpcod.Tag) & "'")
ElseIf Len(txtDONO) = 10 Then
   Call FillList(" AND ORDTRN.DONO ='" & Trim(txtDONO) & "'")
Else
   Call FillList
End If

End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
End Sub

Public Sub FillList(Optional FILTER As String)
Dim SQL As String
Dim M_ROW As Integer

Screen.MousePointer = vbHourglass
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT ORDTRN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE,REFMST.NAME AS AGENT, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM ORDTRN INNER JOIN ACCMST "
SQL = SQL & "ON ACCMST.CODE=ORDTRN.PCOD INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND "
SQL = SQL & "FINITMMST.UNIT=ORDTRN.UNIT AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD LEFT JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND ORDTRN.UNIT = SUBGRDMST.UNIT AND ORDTRN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "ORDTRN.GRAD = SUBGRDMST.GRAD AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE = ORDTRN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = ORDTRN.DCOD AND PADDMST.SRNO = ORDTRN.SRCH WHERE " & _
"ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & UNCD & "' AND ORDTRN.DVCD='" & DIVCODE & _
"' AND ORDTRN.VTYP='DOS' AND ORDTRN.DFLG<>'Y' AND ORDTRN.RECSTAT='A' AND ORDTRN.DOSTAT='Y' "

'DO DATE ARE LESS THEN OR EQUAL TO CHALLAN DATE
SQL = SQL & " AND ORDTRN.DODT <= '" & Format(frmBoxDispatch.TXTVBDT.Value, "MM/DD/YYYY") & "' "
'---------------------------------------------------

If FILTER <> Empty Then SQL = SQL & FILTER

SQL = SQL & "ORDER BY ORDTRN.DONO,ORDTRN.DODT"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
Else
   lst.ListItems.Clear
End If

Do While RECSET.EOF = False
   
Set lstItm = lst.ListItems.ADD()
    lstItm.Text = lst.ListItems.COUNT
    
    lstItm.SubItems(1) = RECSET![ORDN]
    lstItm.SubItems(2) = RECSET![DODT] 'CStr(Format(RECSET![DODt], "dd/mm/yyyy"))
    lstItm.SubItems(3) = RECSET![DONO]
    lstItm.SubItems(4) = Trim(RECSET!ACNM & "")
    lstItm.SubItems(5) = Trim(RECSET!CONSINEE & "")
    lstItm.SubItems(6) = Trim(RECSET!AGENT & "")
    
    lstItm.SubItems(7) = Trim(RECSET!ITNM & "")
    lstItm.SubItems(8) = Trim(RECSET!ltno & "")
    lstItm.SubItems(9) = Trim(RECSET!GRADE & "")
    lstItm.SubItems(10) = Trim(RECSET!SUBGRADE & "")
    
    lstItm.SubItems(11) = Val(Trim(RECSET!QNTY & ""))
    lstItm.SubItems(12) = RECSET!TXRT & ""
    lstItm.SubItems(13) = RECSET!BRMK & ""
    lstItm.SubItems(14) = Trim(RECSET!dbcd & "")
    RECSET.MoveNext
Loop
     
     If lst.ListItems.COUNT > 0 Then
        lst.ListItems(1).Selected = True
        lst.SetFocus
        cmdOk.Default = True
     Else
        cmdOk.Default = False
    End If
Screen.MousePointer = vbNormal
End Sub

Private Sub TXTDONO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtDONO = Empty Then Call FillList: Exit Sub
If Len(txtDONO) = 10 Then
   Call FillList(" AND ORDTRN.DONO ='" & Trim(txtDONO) & "'")
End If
End If
End Sub

Public Function IsSubGradeReq(DIVISIONCODE As String, Optional UNITCODE As String) As Boolean
IsSubGradeReq = False

If UNITCODE = Empty Then UNITCODE = UNCD

Dim DISPRS As ADODB.Recordset
Set DISPRS = New ADODB.Recordset

If DISPRS.State = 1 Then DISPRS.Close
DISPRS.Open "SELECT CFGTYP FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNITCODE & _
            "' AND CODE='" & DIVISIONCODE & "'", CN, adOpenDynamic, adLockOptimistic
If Not DISPRS.EOF Then
   IsSubGradeReq = IIf(Trim(DISPRS!CFGTYP & "") = "SG", True, False)
Else
   IsSubGradeReq = False
End If

End Function


Private Sub txtPCOD_GotFocus()
 txtpcod.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtpcod = SearchList1("Select Code,Name From ACCMST", 0, Empty, "Select Party From List")
        txtpcod.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtpcod = Empty
        txtpcod.Tag = Empty
    End If
End Sub

Private Sub txtPCOD_LostFocus()
 txtpcod.BackColor = vbWhite
End Sub
