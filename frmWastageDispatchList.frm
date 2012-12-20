VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmWastageDispatchList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Wastage Challan No. from List"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDTRNGE 
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   9375
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
         Format          =   55705601
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
         Format          =   55705601
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
      Width           =   9405
      Begin MSComctlLib.ListView lst 
         Height          =   4785
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
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
         NumItems        =   8
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
            Text            =   "Item Desc"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Net Qnty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "dbcd"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   6360
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
      Image           =   "frmWastageDispatchList.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   7560
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
      Image           =   "frmWastageDispatchList.frx":0D8A
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
      Width           =   9375
   End
End
Attribute VB_Name = "frmWastageDispatchList"
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

SQL = "SELECT DISTINCT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
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
    lstItm.SubItems(5) = Trim(RECSET!ITNM & "")
    lstItm.SubItems(6) = Trim(RECSET!QNTY & "")
    lstItm.SubItems(7) = Trim(RECSET!dbcd & "")
    
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

Private Sub cmdOk_Click()
If lst.ListItems.COUNT < 1 Then
 frmWastageDispatch.CHALLAN = Empty
 Exit Sub
End If
Dim SQL As String

With frmWastageDispatch
 .txtConsinee = Empty: .txtDCOD = Empty: .TXTADDRESS = Empty: .TXTITM = Empty
 .TXTRATE = Empty: .txtQTY = Empty: .TXTAMNT = Empty
 chln = lst.SelectedItem.SubItems(2)
 .CHALLAN = lst.SelectedItem.SubItems(2)
  
If Trim(chln) = Empty Then
   lst.SetFocus
   Exit Sub
End If
  
Set RECSET = New ADODB.Recordset

SQL = "SELECT SPTRAN.*,ACCMST.NAME AS ACNM,FINITMMST.NAME AS ITNM, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM SPTRAN " & _
            "INNER JOIN ACCMST ON ACCMST.CODE=SPTRAN.PCOD "
SQL = SQL & "INNER JOIN FINITMMST ON FINITMMST.COMP=SPTRAN.COMP AND FINITMMST.UNIT=SPTRAN.UNIT AND "
SQL = SQL & "FINITMMST.DVCD=SPTRAN.DVCD AND FINITMMST.CODE=SPTRAN.ICOD "
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
    
 .lblBill = Trim(RECSET!VBNO & "")
 .TXTVBDT = Format(RECSET!Date, "DD/MM/YYYY")
 .txtConsinee = Trim(RECSET!ACNM & "")
 .txtDCOD = Trim(RECSET!CONSINEE & "")
 .TXTADDRESS = Trim(RECSET!ADDRESS & "")
 .TXTITM = Trim(RECSET!ITNM & "")
 .txtQTY = Trim(RECSET!QNTY & "")
 .TXTRATE = Trim(RECSET!RATE & "")
 .TXTAMNT = Trim(RECSET!AMNT & "")
 .BRMK = Trim(RECSET!EXTRA4 & "")
  
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
