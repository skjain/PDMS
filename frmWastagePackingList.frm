VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWastagePackingList 
   Caption         =   "Wastage Lumpsum Packing List"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   9
      Top             =   240
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
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   1
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1335
         TabIndex        =   0
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   50921473
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.Frame FramCont 
      Height          =   4635
      Left            =   120
      TabIndex        =   8
      Top             =   945
      Width           =   10590
      Begin MSComctlLib.ListView lstBill 
         Height          =   4380
         Left            =   75
         TabIndex        =   3
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1835
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Slip No."
            Object.Width           =   1906
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   3353
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Machine Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Location"
            Object.Width           =   3176
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Raw Item"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "NetWeight"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   7
      Top             =   5565
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Label LBLDIVNAM 
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
      TabIndex        =   6
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmWastagePackingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVCODE As String
Public M_DBCD As String
Public BOX_PKG_REQ As String
Public PKGSTCOD As String

Private Sub cmdCancel_Click()
    frmWastagePacking.CHALLAN = Empty
    Unload Me
End Sub

Public Sub CMDOK_Click()
   If BOX_PKG_REQ = "Y" Then
     Call FillDetailByBoxRegister
   Else
     Call FillDetailByPkgMan
   End If
   Exit Sub
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    txtFrDate = GetMinDate
    txtToDate = GetMaxDate
    Me.KeyPreview = True
    cmdOk.Enabled = False
    cmdCancel.Enabled = True
End Sub

Private Sub cmdGo_Click()
 lstBill.ListItems.Clear
 
 Dim SQL As String
 SQL = Empty

If BOX_PKG_REQ = "Y" Then
   Call FillListByBoxRegister
Else
   Call FillListByPkgMan
End If
 
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

Private Sub FillListByBoxRegister()
Dim EDTDAT As New ADODB.Recordset
Set EDTDAT = New ADODB.Recordset
Dim SQL As String

SQL = "SELECT DISTINCT BOXREGISTER.*,FINITMMST.NAME AS ITEM,MACMST.NAME AS MACHINE,LOCMST.NAME AS LOCATION FROM "
SQL = SQL & "BOXREGISTER INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND FINITMMST.UNIT=BOXREGISTER.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=BOXREGISTER.DVCD AND FINITMMST.CODE=BOXREGISTER.ICOD LEFT JOIN LOCMST ON BOXREGISTER.LOCCOD=LOCMST.CODE "
SQL = SQL & "LEFT JOIN MACMST ON BOXREGISTER.COMP =MACMST.COMP AND BOXREGISTER.UNIT =MACMST.UNIT AND BOXREGISTER.DVCD =MACMST.DVCD AND " & _
"MACMST.CODE=BOXREGISTER.MCCD WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
"' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND BOXREGISTER.PKG_STCOD='" & PKGSTCOD & "' AND BOXREGISTER.VTYP='PPF' AND BOXREGISTER.RECSTAT='A' AND BOXREGISTER.DBCD='" & M_DBCD & _
"' AND VBDT>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND VBDT<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "' AND BOXREGISTER.DSPWGT=0 "
SQL = SQL & " ORDER BY VBDT "

If EDTDAT.State = 1 Then EDTDAT.Close
EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
If EDTDAT.EOF Then
   MsgBox "No Record found for given criteria ", vbInformation
   txtToDate.SetFocus
   Exit Sub
End If
  
Do While Not EDTDAT.EOF
    Set lstItem = lstBill.ListItems.ADD
    lstItem.Text = Format(EDTDAT![VBDT], "dd/MM/yyyy")
    lstItem.SubItems(1) = Trim(EDTDAT![VBNO] & "")
    lstItem.SubItems(2) = Trim(EDTDAT![Item] & "")
    lstItem.SubItems(3) = Trim(EDTDAT![MACHINE] & "")
    lstItem.SubItems(4) = Trim(EDTDAT![LOCATION] & "")
    'lstItem.SubItems(5) = Trim(EDTDAT![LOTNO] & "")
    lstItem.SubItems(6) = Trim(EDTDAT![NTWGT] & "")
    lstItem.SubItems(6) = nstr(lstItem.SubItems(6), 12, 3)
    lstItem.SubItems(6) = Trim(lstItem.SubItems(6))
    EDTDAT.MoveNext
 Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmWastagePackingList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub FillListByPkgMan()
Dim EDTDAT As New ADODB.Recordset
Set EDTDAT = New ADODB.Recordset
Dim SQL As String

SQL = "SELECT DISTINCT PKGMAN.*,FINITMMST.NAME AS ITEM,MACMST.NAME AS MACHINE,LOCMST.NAME AS LOCATION FROM "
SQL = SQL & "PKGMAN INNER JOIN FINITMMST ON FINITMMST.COMP=PKGMAN.COMP AND FINITMMST.UNIT=PKGMAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=PKGMAN.DVCD AND FINITMMST.CODE=PKGMAN.FINITMCOD LEFT JOIN LOCMST ON PKGMAN.LOCCOD=LOCMST.CODE "
SQL = SQL & "LEFT JOIN MACMST ON MACMST.CODE=PKGMAN.MCCD WHERE PKGMAN.COMP='" & compPth & "' AND PKGMAN.UNIT='" & UNCD & _
"' AND PKGMAN.DVCD='" & DIVCODE & "' AND PKGMAN.VTYP='PPF' AND PKGMAN.PKG_STCOD = '" & PKGSTCOD & "' AND PKGMAN.RECSTAT='A' AND PKGMAN.DBCD='" & M_DBCD & _
"' AND PKGMAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND PKGMAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & "'"
SQL = SQL & " ORDER BY PKGMAN.DATE DESC"

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
    lstItem.SubItems(1) = Trim(EDTDAT![SLIPNO] & "")
    lstItem.SubItems(2) = Trim(EDTDAT![Item] & "")
    lstItem.SubItems(3) = Trim(EDTDAT![MACHINE] & "")
    lstItem.SubItems(4) = Trim(EDTDAT![LOCATION] & "")
    'lstItem.SubItems(5) = Trim(EDTDAT![LOTNO] & "")
    lstItem.SubItems(6) = Trim(EDTDAT![QNTY] & "")
    EDTDAT.MoveNext
 Loop
    
    cmdOk.Enabled = True
    cmdOk.Default = True
    If frmWastagePackingList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub FillDetailByBoxRegister()
    Dim CHLNNO As String
    CHLNNO = lstBill.SelectedItem.SubItems(1)
          
    If Trim(CHLNNO) = Empty Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    Dim EDTDAT As New ADODB.Recordset
    Set EDTDAT = New ADODB.Recordset
    
    Dim SQL As String
    
    SQL = Empty
    
SQL = "SELECT BOXREGISTER.*,FINITMMST.NAME AS ITEM,MACMST.NAME AS MACHINE,LOCMST.NAME AS LOCATION,GRDMST.GRAD AS GRADE FROM "
SQL = SQL & "BOXREGISTER INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND FINITMMST.UNIT=BOXREGISTER.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=BOXREGISTER.DVCD AND FINITMMST.CODE=BOXREGISTER.ICOD LEFT JOIN LOCMST ON BOXREGISTER.LOCCOD=LOCMST.CODE "
SQL = SQL & "LEFT JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD "
SQL = SQL & "LEFT JOIN MACMST ON MACMST.CODE=BOXREGISTER.MCCD WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
"' AND BOXREGISTER.DVCD='" & DIVCODE & "' AND BOXREGISTER.VTYP='PPF' AND BOXREGISTER.RECSTAT='A' AND BOXREGISTER.DBCD='" & M_DBCD & _
"' AND VBDT>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND VBDT<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
"' AND VBNO = '" & CHLNNO & "' "
  
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmWastagePacking
        .CHALLAN = CHLNNO
        .LBLSLIP.Caption = CHLNNO
        .txtITEM.Text = EDTDAT!Item & ""
        .TXTVBDT = Format(EDTDAT!VBDT & "", "DD/MM/YYYY")
        .TXTMCCD.Text = EDTDAT!MACHINE & ""
        .txtLoc.Text = EDTDAT!LOCATION & ""
        .TXTQNTY.Text = EDTDAT!NTWGT & ""
        .TXTQNTY.Text = nstr(.TXTQNTY.Text, 12, 3)
        .TXTQNTY.Text = Trim(.TXTQNTY.Text)
        .TXTRMRK.Text = Trim(EDTDAT!RMRK & "")
        .TXTGRAD.Text = Trim(EDTDAT!GRADE & "")
    End With
    Unload Me
End Sub

Private Sub FillDetailByPkgMan()
    Dim CHLNNO As String
    CHLNNO = lstBill.SelectedItem.SubItems(1)
          
    If Trim(CHLNNO) = Empty Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    Dim EDTDAT As New ADODB.Recordset
    Set EDTDAT = New ADODB.Recordset
    
    Dim SQL As String
    
    SQL = Empty
    
SQL = "SELECT PKGMAN.*,FINITMMST.NAME AS ITEM,MACMST.NAME AS MACHINE,LOCMST.NAME AS LOCATION FROM "
SQL = SQL & "PKGMAN INNER JOIN FINITMMST ON FINITMMST.COMP=PKGMAN.COMP AND FINITMMST.UNIT=PKGMAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=PKGMAN.DVCD AND FINITMMST.CODE=PKGMAN.FINITMCOD LEFT JOIN LOCMST ON PKGMAN.LOCCOD=LOCMST.CODE "
SQL = SQL & "LEFT JOIN MACMST ON MACMST.CODE=PKGMAN.MCCD WHERE PKGMAN.COMP='" & compPth & "' AND PKGMAN.UNIT='" & UNCD & _
"' AND PKGMAN.DVCD='" & DIVCODE & "' AND PKGMAN.VTYP='PPF' AND PKGMAN.RECSTAT='A' AND PKGMAN.DBCD='" & M_DBCD & _
"' AND PKGMAN.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND PKGMAN.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
"' AND PKGMAN.SLIPNO = '" & CHLNNO & "'"
   
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If EDTDAT.EOF Then
        lstBill.SetFocus
        Exit Sub
    End If
    
    With frmWastagePacking
        .CHALLAN = CHLNNO
        .LBLSLIP.Caption = CHLNNO
        .txtITEM.Text = EDTDAT!Item & ""
        .TXTVBDT = Format(EDTDAT!Date & "", "DD/MM/YYYY")
        .TXTMCCD.Text = EDTDAT!MACHINE & ""
        .txtLoc.Text = EDTDAT!LOCATION & ""
        .TXTQNTY.Text = EDTDAT!QNTY & ""
        .TXTQNTY.Text = nstr(.TXTQNTY.Text, 12, 3)
        .TXTQNTY.Text = Trim(.TXTQNTY.Text)
        .TXTRMRK.Text = EDTDAT!REMARKS & ""
    End With
    Unload Me
End Sub

