VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form Frm_ManualClearance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Clearance"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.OptionButton optRGP 
         BackColor       =   &H0080C0FF&
         Caption         =   "&RETURNABLE GATE PASS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   14
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton optJob 
         BackColor       =   &H0080C0FF&
         Caption         =   "&JOB WORK "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox TXTGRAD 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   8655
      End
      Begin VB.TextBox TXTITM 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   7815
      End
      Begin VB.TextBox TXTPTY 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   9255
      End
      Begin WelchButton.lvButtons_H CMDSRCH 
         Height          =   375
         Left            =   9600
         TabIndex        =   5
         Top             =   1200
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
         Image           =   "Frm_ManualClearanceJob.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of &Clearance    (A)                                             (B)     "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label3 
         Caption         =   "Grade"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Name of &Item"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
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
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame FrmData 
      Height          =   4215
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   11055
      Begin VB.CommandButton CMDREFRSH 
         Caption         =   "&Refresh"
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
         Left            =   9960
         TabIndex        =   9
         Top             =   3720
         Width           =   975
      End
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
         Left            =   8880
         TabIndex        =   8
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton cmdCLEAR 
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
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   3720
         Width           =   975
      End
      Begin MSComctlLib.ListView lstJobClearance 
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   5953
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Anx./Chln Date"
            Object.Width           =   2735
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Anx./Chln No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Quantity"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   3951
         EndProperty
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
End
Attribute VB_Name = "Frm_ManualClearance"
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

Private Sub CMDREFRSH_Click()
  Call CMDSRCH_Click
End Sub

Private Sub CMDSRCH_Click()
  Dim LVTYP As String
  
  lstJobClearance.ListItems.Clear
  
  If optJob.Value = True Then
    LVTYP = "ANX"
  Else
    LVTYP = "RGP"
  End If
  
  Dim SQL As String
  
  SQL = Empty
  SQL = "SELECT DISTINCT JOBOUT.DATE,JOBOUT.CLRSTATUS,JOBOUT.VBNO,JOBOUT.QNTY,JOBOUT.RMRK,ITMMST.NAME AS ITEM, " & _
  "ACCMST.NAME AS PARTY FROM JOBOUT INNER JOIN ITMMST " & _
  "ON (JOBOUT.ICOD=ITMMST.CODE) INNER JOIN ACCMST ON (JOBOUT.PCOD=ACCMST.CODE) WHERE JOBOUT.COMP='" & compPth & _
  "' AND JOBOUT.UNIT='" & UNCD & "' AND VTYP='" & LVTYP & "' AND JOBOUT.RECSTAT<>'D' AND CLRSTATUS='N' "
      
  If TXTPTY <> Empty Then
    SQL = SQL & " AND JOBOUT.PCOD='" & TXTPTY.Tag & "' "
  End If
  
  If Not TXTITM = Empty Then
    SQL = SQL & " AND JOBOUT.ICOD='" & TXTITM.Tag & "' "
  End If
  
    SQL = SQL & " ORDER BY JOBOUT.DATE,JOBOUT.VBNO"
          
  If RS.State = 1 Then RS.Close
  RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  'Fill List Details
  If RS.EOF Then
    MsgBox "No record found for given criteria"
    TXTPTY.SetFocus
    Exit Sub
  End If
  
  Dim Item As ListItem
  lstJobClearance.ListItems.Clear
  Do While Not RS.EOF
    Set Item = lstJobClearance.ListItems.ADD
    Item.Text = Format(RS!Date, "DD/MM/YYYY")
    
    If RS!CLRSTATUS = "Y" Then
      Item.Checked = True
     Else
      Item.Checked = False
    End If
    
    Item.SubItems(1) = RS!VBNO
    Item.SubItems(2) = RS!PARTY
    Item.SubItems(3) = RS!Item
    Item.SubItems(4) = Format(RS!QNTY, "########.000")
    Item.SubItems(5) = Trim(RS!RMRK & "")
    RS.MoveNext
  Loop
  lstJobClearance.SetFocus
  RS.Close
End Sub

Private Sub cmdCLEAR_Click()
Dim LVTYP As String
Dim ITMCOD As String

If lstJobClearance.ListItems.COUNT <= 0 Then
   MsgBox "Data not Found"
   TXTPTY.Enabled = True
   TXTPTY.SetFocus
   Exit Sub
End If

If RS.State = 1 Then RS.Close
RS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & lstJobClearance.SelectedItem.SubItems(3) & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
  ITMCOD = RS!CODE & ""
End If
RS.Close

If optJob.Value = True Then
    LVTYP = "ANX"
Else
    LVTYP = "RGP"
End If
Dim I As Long
  For I = 1 To lstJobClearance.ListItems.COUNT
    If lstJobClearance.ListItems(I).Checked = True Then
      CN.Execute "UPDATE JOBOUT SET CLRSTATUS='Y' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='" & LVTYP & "' AND VBNO='" & lstJobClearance.SelectedItem.SubItems(1) & _
      "' AND ICOD='" & ITMCOD & "' AND RECSTAT<>'D' AND CLRSTATUS='N'"
      
      CN.Execute "UPDATE JOBOUT SET MODE='Y' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='" & LVTYP & "' AND VBNO='" & lstJobClearance.SelectedItem.SubItems(1) & _
      "' AND RECSTAT<>'D'"
      
    End If
  Next
  
  Call CMDREFRSH_Click
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

Private Sub lstJobClearance_LostFocus()
 lstJobClearance.BackColor = vbWhite
End Sub

Private Sub optJob_Click()
Call CMDSRCH_Click
End Sub

Private Sub optRGP_Click()
Call CMDSRCH_Click
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    TXTGRAD.Text = SearchList1("select  TOP 20 grad,grad from grdmst", 0, TXTGRAD.Text, "SELECT GRADE FROM LIST")
  End If
  If KeyCode = vbKeyDelete Then
    TXTGRAD.Text = Empty
  End If
End Sub

Private Sub TXTITM_GotFocus()
  TXTITM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False
    TXTITM.Text = SearchList1("SELECT TOP 20 Code,NAME FROM ITMMST", 0, TXTITM.Text, "SELECT ITEM FROM LIST")
    TXTITM.Tag = Key
  ElseIf KeyCode = vbKeyDelete Then
    TXTITM.Text = Empty
  End If
End Sub

Private Sub TXTITM_LostFocus()
  TXTITM.BackColor = vbWhite
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
