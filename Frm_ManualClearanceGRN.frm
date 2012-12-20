VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form Frm_ManualClearanceGRN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GRN Clearance for Inhouse JobWork"
   ClientHeight    =   6240
   ClientLeft      =   225
   ClientTop       =   735
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11250
   Begin VB.Frame Frmfilt 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.OptionButton OptClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptPending 
         Caption         =   "Pe&nding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TXTPTY 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   7695
      End
      Begin WelchButton.lvButtons_H CMDSRCH 
         Height          =   375
         Left            =   9480
         TabIndex        =   3
         Top             =   480
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "Frm_ManualClearanceGRN.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Name of &Party"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame FrmData 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   11055
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9840
         TabIndex        =   6
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdCLEAR 
         Caption         =   "Clea&r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   5
         Top             =   4440
         Width           =   975
      End
      Begin MSComctlLib.ListView lstGRNClearance 
         Height          =   4095
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "GRN Date"
            Object.Width           =   2382
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GRN No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Challan No."
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Challan Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Party Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Quantity"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remarks"
            Object.Width           =   3951
         EndProperty
      End
   End
End
Attribute VB_Name = "Frm_ManualClearanceGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LED_BALN As Double
Dim IGMNAM As String
Dim IGMCOD As String
Dim ORDDBNM As String
Dim ORDDBCD As String
Dim SQL As String

Private Sub cmdclose_Click()
  Unload Me
End Sub

Private Sub CMDSRCH_Click()
  lstGRNClearance.ListItems.Clear
  Dim SQL As String
  
  SQL = Empty
  SQL = "SELECT DISTINCT JOBGRN.DATE,JOBGRN.VBNO,JOBGRN.CHLN,JOBGRN.CHDT,ACCMST.NAME AS PARTY, " & _
  "JOBGRN.TQTY,JOBGRN.BRMK,JOBGRN.CLRSTATUS FROM JOBGRN INNER JOIN ACCMST ON (JOBGRN.PCOD=ACCMST.CODE) " & _
  "WHERE JOBGRN.COMP='" & compPth & "' AND JOBGRN.UNIT='" & UNCD & _
  "' AND VTYP='IVR' AND JOBGRN.DBCD='000002' AND JOBGRN.RECSTAT<>'D' "
  
  If OptPending.Value = True Then
    SQL = SQL & " AND JOBGRN.CLRSTATUS <> 'Y' "
  ElseIf OptClear.Value = True Then
    SQL = SQL & " AND JOBGRN.CLRSTATUS = 'Y' "
  End If
      
  If TXTPTY <> Empty Then
    SQL = SQL & " AND JOBGRN.PCOD='" & TXTPTY.Tag & "' "
  End If
    
  SQL = SQL & " ORDER BY JOBGRN.DATE,JOBGRN.VBNO"
          
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "No record found for given criteria"
    TXTPTY.SetFocus
    Exit Sub
  End If
  
  Dim Item As ListItem
  lstGRNClearance.ListItems.Clear
  Do While Not RS.EOF
    Set Item = lstGRNClearance.ListItems.ADD
    Item.Text = Format(RS!Date, "DD/MM/YYYY")
    If RS!CLRSTATUS = "Y" Then
      Item.Checked = True
     Else
      Item.Checked = False
    End If
    Item.SubItems(1) = RS!VBNO
    Item.SubItems(2) = RS!chln
    Item.SubItems(3) = Format(RS!CHDT, "DD/MM/YYYY")
    Item.SubItems(4) = RS!PARTY
    Item.SubItems(5) = Format(RS!TQTY, "########.000")
    Item.SubItems(6) = Trim(RS!BRMK & "")
    RS.MoveNext
  Loop
  lstGRNClearance.SetFocus
  RS.Close
End Sub

Private Sub cmdCLEAR_Click()
Dim LVTYP As String
Dim ITMCOD As String

If lstGRNClearance.ListItems.COUNT <= 0 Then
   MsgBox "Data not Found"
   TXTPTY.Enabled = True
   TXTPTY.SetFocus
   Exit Sub
End If

Dim I As Long
  For I = 1 To lstGRNClearance.ListItems.COUNT
    If lstGRNClearance.ListItems(I).Checked = True Then
      CN.Execute "UPDATE JOBGRN SET CLRSTATUS='Y' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='IVR' AND DBCD='000002' AND VBNO='" & lstGRNClearance.SelectedItem.SubItems(1) & _
      "' AND RECSTAT<>'D' AND CLRSTATUS='N'"
    End If
  Next
  
  Call CMDSRCH_Click
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me): Me.BackColor = RGB(RED, GREEN, BLUE)
 Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me): Me.BackColor = RGB(RED, GREEN, BLUE)
  Me.Left = Me.Left - 100
  FrmData.Enabled = True
  Frmfilt.Enabled = True
End Sub

Private Sub lstGRNClearance_LostFocus()
 lstGRNClearance.BackColor = vbWhite
End Sub

Private Sub OptAll_Click()
Call CMDSRCH_Click
End Sub

Private Sub OptPending_Click()
Call CMDSRCH_Click
End Sub

Private Sub OptClear_Click()
Call CMDSRCH_Click
End Sub

Private Sub TXTPTY_GotFocus()
  TXTPTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPTY_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    TXTPTY.Text = SearchList1("SELECT TOP 20 Code,NAME FROM ACCMST", 0, TXTPTY.Text, "SELECT A/C PARTY")
    TXTPTY.Tag = Key
  ElseIf KeyCode = vbKeyDelete Then
    TXTPTY.Text = Empty
    TXTPTY.Tag = Empty
  End If
End Sub

Private Sub TXTPTY_LostFocus()
  TXTPTY.BackColor = vbWhite
End Sub
