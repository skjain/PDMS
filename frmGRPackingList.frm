VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmGRPackingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Goods Return List"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7935
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
         Left            =   6840
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
         Format          =   24313857
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
         Format          =   24313857
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
      Height          =   3090
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   7965
      Begin MSComctlLib.ListView lst 
         Height          =   2745
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4842
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "VBNO"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "LotNo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Desc"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Grade"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Boxes"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Net Qnty"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3960
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
      Image           =   "frmGRPackingList.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   3960
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
      Image           =   "frmGRPackingList.frx":0D8A
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmGRPackingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DIVCODE As String
Public DIVNAME As String
Public CHALLAN As String
Public SUBGRD As String

Public Sub FillList()
Me.Caption = "HELLO"
lst.ListItems.Clear

Me.Caption = "HELLO121332212223"

Dim SQL As String
Dim M_ROW As Integer

Screen.MousePointer = vbHourglass
Set RECSET = New ADODB.Recordset

SQL = "SELECT DISTINCT GRPACKING.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE "
SQL = SQL & "FROM GRPACKING LEFT JOIN ACCMST ON ACCMST.CODE=GRPACKING.PCOD "
SQL = SQL & "LEFT JOIN FINITMMST ON FINITMMST.COMP=GRPACKING.COMP AND FINITMMST.UNIT=GRPACKING.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=GRPACKING.DVCD AND FINITMMST.CODE=GRPACKING.ICOD "
SQL = SQL & "LEFT JOIN GRDMST ON GRDMST.CODE=GRPACKING.GRAD "
SQL = SQL & "WHERE GRPACKING.COMP='" & compPth & "' AND GRPACKING.UNIT='" & UNCD & _
"' AND GRPACKING.DVCD='" & frmGRPacking.DIVCODE & "' AND GRPACKING.RECSTAT='A' AND GRPACKING.CRNO IS NULL "

SQL = SQL & " AND GRPACKING.VBDT >= '" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
            "' AND GRPACKING.VBDT <= '" & Format(txtToDate.Value, "MM/DD/YYYY") & "' "
SQL = SQL & " ORDER BY GRPACKING.VBNO,GRPACKING.VBDT"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

Do While RECSET.EOF = False
    Set lstItm = lst.ListItems.ADD
    lstItm.Text = CStr(Format(RECSET![VBDT], "dd/mm/yyyy"))
    lstItm.SubItems(1) = Trim(RECSET![VBNO] & "")
    lstItm.SubItems(2) = RECSET![ACNM] & ""
    lstItm.SubItems(3) = Trim(RECSET!LOTNO & "")
    lstItm.SubItems(4) = Trim(RECSET!ITNM & "")
    lstItm.SubItems(5) = Trim(RECSET!GRADE & "")
    lstItm.SubItems(6) = Trim(RECSET!BOXES & "")
    lstItm.SubItems(7) = Trim(RECSET!NETWGT & "")
    
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

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()
  Call FillList
End Sub

Private Sub CMDOK_Click()
If lst.ListItems.COUNT < 1 Then
 frmGRPacking.CHALLAN = Empty
 Exit Sub
End If
Dim SQL As String

With frmGRPacking
 .txtpcod = Empty: .TXTDENI = Empty: .txtLTNo = Empty: .TXTGRAD = Empty
 .TXTBOXES = Empty: .TXTCOP = Empty: .TXTGRWT = Empty
 .TXTTRWT = Empty: .TXTNTWT = Empty: SUBGRD = Empty
 .TXTSUBGRD = Empty: .TXTTWIST = Empty: .TWSTREQ = Empty
 
 CHALLAN = lst.SelectedItem.SubItems(1)
 .CHALLAN = lst.SelectedItem.SubItems(1)
 
If Trim(CHALLAN) = Empty Then
     'lstBill.SetFocus
     Exit Sub
End If
  
Set RECSET = New ADODB.Recordset

SQL = "SELECT GRPACKING.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,"
SQL = SQL & "PKGNGMST.NAME AS PKG FROM GRPACKING LEFT JOIN ACCMST ON ACCMST.CODE=GRPACKING.PCOD "
SQL = SQL & "LEFT JOIN FINITMMST ON FINITMMST.COMP=GRPACKING.COMP AND FINITMMST.UNIT=GRPACKING.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=GRPACKING.DVCD AND FINITMMST.CODE=GRPACKING.ICOD "
SQL = SQL & "LEFT JOIN GRDMST ON GRDMST.CODE=GRPACKING.GRAD "
SQL = SQL & "LEFT JOIN PKGNGMST ON PKGNGMST.CODE=GRPACKING.PKG_STCOD "
SQL = SQL & "WHERE GRPACKING.COMP='" & compPth & "' AND GRPACKING.UNIT='" & UNCD & _
"' AND GRPACKING.DVCD='" & frmGRPacking.DIVCODE & "' AND GRPACKING.RECSTAT='A' AND GRPACKING.VBNO='" & CHALLAN & "' "

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found "
   Screen.MousePointer = vbNormal
'   lstBill.SetFocus
   Exit Sub
Else
    .TXTVBNO = Trim(RECSET!VBNO & "")
    .CHALLAN = Trim(RECSET!VBNO & "")
    .TXTVBDT = Format(RECSET!VBDT, "DD/MM/YYYY")
    .txtpcod = Trim(RECSET!ACNM & "")
    .TXTDENI = Trim(RECSET!ITNM & "")
    .txtLTNo = Trim(RECSET!LOTNO & "")
    .TXTGRAD = Trim(RECSET!GRADE & "")
    .GRADE = Trim(RECSET!grad & "")
    .TXTBOXES = Trim(RECSET!BOXES & "")
    .TXTCOP = Trim(RECSET!COPS & "")
    .TXTGRWT = Trim(RECSET!GRSWGT & "")
    .TXTTRWT = Trim(RECSET!TRWGT & "")
    .TXTNTWT = Trim(RECSET!NETWGT & "")
    .txtRMK = Trim(RECSET!RMK & "")
    
    If .TWSTREQ = "Y" Then
        .TXTTWIST = Trim(RECSET!SUBGRD & "")
    Else
        SUBGRD = Trim(RECSET!SUBGRD & "")
    End If
    
    Dim TEMPRS As New ADODB.Recordset
    If TEMPRS.State = 1 Then TEMPRS.Close
    TEMPRS.Open "SELECT *FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                "' AND DVCD='" & .DIVCODE & "' AND SUBGRD='" & SUBGRD & "' AND GRAD='" & RECSET!grad & "'", CN, adOpenDynamic, adLockOptimistic
    If Not TEMPRS.EOF Then
        .TXTSUBGRD = Trim(TEMPRS!NAME & "")
    End If
End If

End With
Me.Hide
Unload Me
    
    Dim EDTDAT As New ADODB.Recordset
    If EDTDAT.State = 1 Then EDTDAT.Close
    EDTDAT.Open "SELECT * FROM GRPACKING WHERE Comp='" & compPth & "'  And UNIT='" & UNCD & _
            "' AND VBNO='" & CHALLAN & "' AND RECSTAT <> 'D' AND (FRESH>0 OR WASTAGE>0)", CN, adOpenDynamic, adLockOptimistic
    If Not EDTDAT.EOF Then
        MsgBox "Further Entry exist can not Modify / Delete it"
        frmGRPacking.CHALLAN = Empty
        frmGRPacking.TXTVBNO = Empty
        Exit Sub
    End If

End Sub

Private Sub Form_Activate()
 txtFrDate.SetFocus
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
txtFrDate.Value = GetMinDate
 txtToDate.Value = GetMaxDate
 
'Call FillList
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{TAB}"
End If
End Sub
