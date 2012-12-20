VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmEditOpnCform 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening C-Form Bill Help For Edit"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7171.376
   ScaleMode       =   0  'User
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDTRNGE 
      Height          =   1080
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   8775
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
         Format          =   24313857
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
         Format          =   24313857
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
      Width           =   8805
      Begin MSComctlLib.ListView lst 
         Height          =   5025
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   8535
         _ExtentX        =   15055
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr."
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bill Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Bill No."
            Object.Width           =   2029
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "A/c Party"
            Object.Width           =   4235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Agent"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bill Amount"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin WelchButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   5040
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
      Image           =   "frmEditOpnCform.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   6240
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
      Image           =   "frmEditOpnCform.frx":0D8A
      cBack           =   -2147483633
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select Bill For Division : "
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
      Width           =   8775
   End
End
Attribute VB_Name = "frmEditOpnCform"
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

Public Sub FillList()
lst.ListItems.Clear
Dim SQL As String
Dim M_ROW As Integer
Dim lstItm

Screen.MousePointer = vbHourglass
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT BILLMAIN.*,ACCMST.NAME AS ACNM,REFMST.NAME AS AGENT, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM BILLMAIN "
SQL = SQL & "INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE = BILLMAIN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = BILLMAIN.DCOD AND PADDMST.SRNO = BILLMAIN.SRCH "
SQL = SQL & "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & UNCD & _
"' AND BILLMAIN.DVCD='" & DIVCODE & "' AND BILLMAIN.VTYP='OPC' AND BILLMAIN.RECSTAT='A' " & _
" AND BILLMAIN.DBCD='" & VTCD & "' "

SQL = SQL & "AND BILLMAIN.DATE >= '" & Format(txtFrDate.Value, "MM/DD/YYYY") & _
            "' AND BILLMAIN.DATE <= '" & Format(txtToDate.Value, "MM/DD/YYYY") & "' "

If txtpcod <> Empty Then
   SQL = SQL & " AND BILLMAIN.PCOD = '" & txtpcod.Tag & "'"
End If

SQL = SQL & " ORDER BY BILLMAIN.VBNO,BILLMAIN.DATE"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found ": Screen.MousePointer = vbNormal: Exit Sub
End If

Do While RECSET.EOF = False
   
Set lstItm = lst.ListItems.ADD()
    lstItm.Text = lst.ListItems.COUNT
    lstItm.SubItems(1) = CStr(Format(RECSET![Date], "dd/mm/yyyy"))
    lstItm.SubItems(2) = Trim(RECSET![VBNO] & "")
    lstItm.SubItems(3) = Trim(RECSET!ACNM & "")
    lstItm.SubItems(4) = Trim(RECSET!AGENT & "")
    lstItm.SubItems(5) = nstr(Val(Trim(RECSET!BNET & "")), 10, 2)
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
Dim i As Double, J As Double
Dim RSDATA As ADODB.Recordset
If lst.ListItems.COUNT < 1 Then
 frmOpeningCFormRecievable.TXTBILLNO = Empty
 Exit Sub
End If
Dim SQL As String

With frmOpeningCFormRecievable

 .M_PNAM = Empty: .txtDCOD = Empty: .TXTADDRESS = Empty
 .M_BRNM = Empty: .M_TXNM = Empty: .TXTBNET = Empty: .BRMK = Empty
  
 chln = lst.SelectedItem.SubItems(2)
 .TXTBILLNO = lst.SelectedItem.SubItems(2)
 
 If Trim(chln) = Empty Then
     lst.SetFocus
     Exit Sub
 End If
  
Dim RECSET As ADODB.Recordset
Set RECSET = New ADODB.Recordset

SQL = "SELECT BILLMAIN.*,ACCMST.NAME AS ACNM,REFMST.NAME AS AGENT,TAXMST.NAME AS TAX, "
SQL = SQL & "PADDMST.NAME AS CONSINEE,PADDMST.ADDR AS ADDRESS FROM BILLMAIN "
SQL = SQL & "INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD "
SQL = SQL & "INNER JOIN REFMST ON REFMST.CODE = BILLMAIN.BRCD "
SQL = SQL & "INNER JOIN PADDMST ON PADDMST.CODE = BILLMAIN.DCOD AND PADDMST.SRNO = BILLMAIN.SRCH "
SQL = SQL & "LEFT JOIN TAXMST ON TAXMST.CODE=BILLMAIN.TXCD "
SQL = SQL & "WHERE BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.UNIT='" & UNCD & _
"' AND BILLMAIN.VTYP='OPC' AND BILLMAIN.RECSTAT='A' AND BILLMAIN.DBCD='" & VTCD & _
"' AND BILLMAIN.VBNO='" & chln & "'"

If RECSET.State = 1 Then RECSET.Close
RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
If RECSET.EOF Then
   MsgBox "No Records Found "
   Screen.MousePointer = vbNormal
   Exit Sub
End If
   
  If Not IsNull(RECSET!Form) And Trim((RECSET!Form & "")) <> Empty Then
     MsgBox "Further Entry Exist", vbCritical
     Screen.MousePointer = vbNormal
     Exit Sub
  End If
   
 .M_PNAM = Trim(RECSET!ACNM & "")
 .txtDCOD = Trim(RECSET!CONSINEE & "")
 .TXTADDRESS = Trim(RECSET!ADDRESS & "")
 .M_BRNM = Trim(RECSET!AGENT & "")
 .TXTBNET = Trim(RECSET!BNET & "")
 .BRMK = Trim(RECSET!BRMK & "")
 .M_TXNM = Trim(RECSET!TAX & "")
 .TXTQNTY = Val(RECSET!TQTY & "")
 .TXTCHLN = Trim(RECSET!chln & "")
 
End With


Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
 Call ColorComponent(Me)
 txtFrDate.Value = FSDT - 1
 txtToDate.Value = FSDT - 1
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


