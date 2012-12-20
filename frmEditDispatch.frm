VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmEditDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Challan Help For Edit"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7171.376
   ScaleMode       =   0  'User
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDTRNGE 
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11295
      Begin VB.TextBox txtpcod 
         Height          =   315
         Left            =   1320
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
         Left            =   1335
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
   End
   Begin VB.Frame frmIVR 
      Height          =   5250
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   11325
      Begin MSComctlLib.ListView lst 
         Height          =   5025
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8864
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
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Chln Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Challan"
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DONO"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "A/c Party"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivery Party"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Agent"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item Desc"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Quantity"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "OrderNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "DBCD"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "RDBC"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   6720
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
      Image           =   "frmEditDispatch.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   9720
      TabIndex        =   8
      Top             =   6720
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
      Image           =   "frmEditDispatch.frx":0D8A
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
Attribute VB_Name = "frmEditDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DIVCODE As String
Public DIVNAME As String
Public M_DBCD As String
Public VTCD As String
Public chln As String
Public PKG_DBCD As String

Public Sub FillList()
lst.ListItems.Clear
Dim SQL As String
Dim M_ROW As Integer
Dim lstItm

Screen.MousePointer = vbHourglass
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT ORDTRN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE,REFMST.NAME AS AGENT, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM ORDTRN INNER JOIN ACCMST "
SQL = SQL & "ON ACCMST.CODE=ORDTRN.PCOD INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND FINITMMST.UNIT=ORDTRN.UNIT AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD LEFT JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND ORDTRN.UNIT = SUBGRDMST.UNIT AND ORDTRN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "ORDTRN.GRAD = SUBGRDMST.GRAD AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE = ORDTRN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = ORDTRN.DCOD AND PADDMST.SRNO = ORDTRN.SRCH "
SQL = SQL & "WHERE ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & UNCD & _
"' AND ORDTRN.DVCD='" & DIVCODE & "' AND ORDTRN.VTYP='DPF' AND ORDTRN.RECSTAT='A' "

SQL = SQL & "AND ORDTRN.RDBC = '" & VTCD & "' "

SQL = SQL & "AND ORDTRN.SLIPDATE >= '" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
            "' AND ORDTRN.SLIPDATE <= '" & Format(txtToDate.Value, "MM/DD/YYYY") & "' "

'After Sale DO Doesnt Come : Condition Pending
'SQL = SQL & " "
If txtpcod <> Empty Then
   SQL = SQL & " AND ORDTRN.PCOD = '" & txtpcod.Tag & "'"
End If

SQL = SQL & " ORDER BY ORDTRN.SLIP,ORDTRN.SLIPDATE"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

Do While RECSET.EOF = False
   
Set lstItm = lst.ListItems.ADD()
    lstItm.Text = lst.ListItems.COUNT
    lstItm.SubItems(1) = CStr(Format(RECSET![SLIPDATE], "dd/mm/yyyy"))
    lstItm.SubItems(2) = Trim(RECSET![SLIP] & "")
    lstItm.SubItems(3) = RECSET![DONO]
    lstItm.SubItems(4) = Trim(RECSET!ACNM & "")
    lstItm.SubItems(5) = Trim(RECSET!CONSINEE & "")
    lstItm.SubItems(6) = Trim(RECSET!AGENT & "")
    lstItm.SubItems(7) = Trim(RECSET!ITNM & "")
    lstItm.SubItems(8) = nstr(Val(Trim(RECSET!QNTY & "")), 10, 3)
    lstItm.SubItems(8) = Trim(lstItm.SubItems(8))
    lstItm.SubItems(9) = Trim(RECSET!ORDN & "")
    lstItm.SubItems(10) = Trim(RECSET!DBCD & "")
    lstItm.SubItems(11) = Trim(RECSET!RDBC & "")
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
Dim i As Double, J As Double
Dim RSDATA As ADODB.Recordset
If lst.ListItems.COUNT < 1 Then
 frmBoxDispatch.chln = Empty
 Exit Sub
End If
Dim SQL As String

With frmBoxDispatch
 .txtpcod = Empty: .txtCONSINEE = Empty: .TXTADDRESS = Empty: .txtITEM = Empty: .txtLTNo = Empty: .TXTGRAD = Empty: .TXTSUBGRD = Empty
 .M_RTTX = Empty: .txtDONO = Empty: .txtQty = Empty: .M_DORAT = Empty: .M_ARAT = Empty: .TXTRMRK = Empty
  
 chln = lst.SelectedItem.SubItems(2)
 .chln = lst.SelectedItem.SubItems(2)
 M_DBCD = lst.SelectedItem.SubItems(10)
 .M_DBCD = lst.SelectedItem.SubItems(10)
 VTCD = lst.SelectedItem.SubItems(11)
 .VTCD = lst.SelectedItem.SubItems(11)

Call SetLRInfo(VTCD, chln)

If Trim(chln) = Empty Then
     lst.SetFocus
     Exit Sub
End If
  
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT ORDTRN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE,REFMST.NAME AS AGENT, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM ORDTRN INNER JOIN ACCMST "
SQL = SQL & "ON ACCMST.CODE=ORDTRN.PCOD INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND FINITMMST.UNIT=ORDTRN.UNIT AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD "
SQL = SQL & "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD LEFT JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND ORDTRN.UNIT = SUBGRDMST.UNIT AND ORDTRN.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "ORDTRN.GRAD = SUBGRDMST.GRAD AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE = ORDTRN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = ORDTRN.DCOD AND PADDMST.SRNO = ORDTRN.SRCH "
SQL = SQL & "WHERE ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & UNCD & _
"' AND ORDTRN.DVCD='" & DIVCODE & "' AND ORDTRN.VTYP='DPF' AND ORDTRN.RECSTAT='A' AND ORDTRN.SLIP='" & chln & _
"' AND ORDTRN.RDBC = '" & VTCD & "' AND ORDTRN.DBCD = '" & M_DBCD & "'"
'DBCD EXCHANGE VTCD


If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found "
   Screen.MousePointer = vbNormal
   'lstBill.SetFocus
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
 
 .TXTAGENT = Trim(RECSET!AGENT & "")
 .M_RTTX = Trim(RECSET!TXRT & "")
 .txtDONO = Trim(RECSET!DONO & "")
 .TXTORDN = Trim(RECSET!ORDN & "")
 .TXTVBDT = Format(RECSET!SLIPDATE, "DD/MM/YYYY")
 .dtDate = Trim(RECSET!DODT & "")
 .txtQty = Trim(RECSET!DELQNTY & "")
 .M_DORAT = Trim(RECSET!RATE & "")
 .M_ARAT = Trim(RECSET!ARAT & "")
 .TXTRMRK = Trim(RECSET!BRMK & "")
  
Set RSDATA = New ADODB.Recordset

SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & DIVCODE & _
"'AND BOXREGISTER.LOTNO ='" & Trim(RECSET!ltno & "") & "' AND BOXREGISTER.ICOD = '" & Trim(RECSET!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(RECSET!grad & "") & _
"' AND VTYP='DPF' AND BOXREGISTER.RECSTAT<>'D' AND RVBNO ='" & chln & "' AND RDBC = '" & Trim(VTCD) & "' "

If IsSubGradeReq(DIVCODE) Or IsTwistReq(DIVCODE) = "Y" Then
   If IsTwistReq(DIVCODE) = "Y" And Trim(RECSET!SUBGRD & "") = "0" Then
      
   Else
      SQL = SQL & " AND BOXREGISTER.SUBGRD ='" & Trim(RECSET!SUBGRD & "") & "' "
   End If
End If

SQL = SQL & " ORDER BY VBNO"

            
If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   MsgBox "Boxes are not available for this criteria."
   Exit Sub
End If

'INITIAL SET TOTAL BOX COPES
 .txtTTLCOPs = 0
 .txtTTLCTRN = 0
 .txtTTLNTWT = 0
 .txtRMNCOPs = 0
 .txtRMNCTRN = 0
 .txtRMNNTWT = 0
'===============================================

Dim COPS As Long: COPS = 0
Dim BOXES As Long: BOXES = 0
Dim NETWT As Double: NETWT = 0
Dim Item

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
   Item.SubItems(10) = Trim(RSDATA![Top] & "")
   Item.SubItems(11) = Trim(RSDATA![PLTNO] & "")
      
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
  
  .txtCops = COPS
  .txtCTRN = BOXES
  .txtNTWT = NETWT
  .txtNTWT.Tag = NETWT
  
  '---------------------------------------
SQL = "SELECT BOXREGISTER.*,SUBGRDMST.NAME AS SUBGRADE FROM BOXREGISTER LEFT JOIN SUBGRDMST ON BOXREGISTER.COMP = SUBGRDMST.COMP "
SQL = SQL & "AND BOXREGISTER.UNIT = SUBGRDMST.UNIT AND BOXREGISTER.DVCD = SUBGRDMST.DVCD AND "
SQL = SQL & "BOXREGISTER.GRAD = SUBGRDMST.GRAD AND BOXREGISTER.SUBGRD = SUBGRDMST.SUBGRD "
SQL = SQL & "WHERE BOXREGISTER.COMP = '" & compPth & _
"' AND BOXREGISTER.UNIT = '" & UNCD & "' AND BOXREGISTER.DVCD = '" & DIVCODE & _
"'AND BOXREGISTER.LOTNO ='" & Trim(RECSET!ltno & "") & "' AND BOXREGISTER.ICOD = '" & Trim(RECSET!ICOD & "") & _
"' AND BOXREGISTER.GRAD ='" & Trim(RECSET!grad & "") & _
"' AND VTYP IN ('PPF','OPN') AND BOXREGISTER.RECSTAT<>'D' AND RVBNO IS NULL AND DBCD NOT IN('000001','000005') "

SQL = SQL & " AND BOXREGISTER.VBDT <= '" & Format(RECSET!DODT, "MM/DD/YYYY") & "' "

If IsSubGradeReq(DIVCODE) Or IsTwistReq(DIVCODE) = "Y" Then
   If IsTwistReq(DIVCODE) = "Y" And Trim(RECSET!SUBGRD & "") = "0" Then
      
   Else
      SQL = SQL & " AND BOXREGISTER.SUBGRD ='" & Trim(RECSET!SUBGRD & "") & "' "
   End If
End If


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

If PKG_DBCD = Empty Then
   SQL = SQL & " AND DBCD NOT IN('000002') "
Else
   SQL = SQL & " AND DBCD IN('000002') "
End If

SQL = SQL & "  ORDER BY VBNO"

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open SQL, CN, adOpenDynamic, adLockOptimistic

If RSDATA.EOF Then
   'MsgBox "Boxes are not available for this criteria."
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
      'If .lstBox.SelectedItem.ListSubItems.COUNT = 2 Then .lstBox.ColumnHeaders(4).Text = "Twist"
   Else
      Item.SubItems(3) = Trim(RSDATA!SUBGRADE & "")
      'If .lstBox.ListItems.COUNT = 1 Then .lstBox.ColumnHeaders(4).Text = "SubGrade"
   End If
     
   Item.SubItems(4) = nstr(RSDATA!GRSWGT, 9, 3)
   Item.SubItems(4) = Trim(Item.SubItems(4))
   Item.SubItems(5) = nstr(RSDATA!TRWGT, 9, 3)
   Item.SubItems(5) = Trim(Item.SubItems(5))
   Item.SubItems(6) = Format(RSDATA!VBDT, "DD/MM/YYYY")
   Item.SubItems(7) = RSDATA!RMRK
   Item.SubItems(8) = Trim(RSDATA!PKG_STCOD & "")
   Item.SubItems(9) = Trim(RSDATA!ISRETURNABLE & "")
   Item.SubItems(10) = Trim(RSDATA![Top] & "")
   Item.SubItems(11) = Trim(RSDATA![PLTNO] & "")
         
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
  
  .txtRMNCOPs = Val(.txtTTLCOPs) - Val(.txtCops)
  .txtRMNCTRN = Val(.txtTTLCTRN) - Val(.txtCTRN)
  .txtRMNNTWT = Val(.txtTTLNTWT) - Val(.txtNTWT)

  
End With


Me.Hide
Unload Me
End Sub

Private Sub Form_Activate()
'  LBLHEAD.Caption = LBLHEAD.Caption + DIVNAME
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
'Call FillList
 txtFrDate.Value = GetMinDate
 txtToDate.Value = GetMaxDate
 LBLHEAD.Caption = LBLHEAD.Caption + DIVNAME
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
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtpcod = Empty) Then
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

Private Sub SetLRInfo(DBCD As String, VBNO As String)
Dim LRRS As ADODB.Recordset
Set LRRS = New ADODB.Recordset

With frmBoxDispatch

If LRRS.State = 1 Then LRRS.Close
LRRS.Open "SELECT LRNO,LRDT,VEHICALNO,TRCD FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND DVCD='" & DIVCODE & "' AND DBCD='" & DBCD & "' AND VBNO='" & VBNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not LRRS.EOF Then
    .TXTLRNO = Trim(LRRS!LRNO & "")
   If Not IsNull(LRRS!LRDT) Then
    .LRDT = Format(LRRS!LRDT, "DD/MM/YYYY")
   End If
    .TXTVHCL = GetCode("VHCLMST", Trim(LRRS!VEHICALNO & ""), "CODE", "NAME")
    .txtTransport = GetCode("TRANSPORTMST", Trim(LRRS!TRCD & ""), "CODE", "NAME")
End If
LRRS.Close

End With

End Sub
