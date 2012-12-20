VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_EXCISEList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Excise Entry No."
   ClientHeight    =   6555
   ClientLeft      =   450
   ClientTop       =   1185
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmDTRNGE 
      Height          =   720
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9690
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
         Height          =   375
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4320
         TabIndex        =   3
         Top             =   285
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53280769
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   255
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   53280769
         CurrentDate     =   38429
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
         Top             =   240
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
         Left            =   3360
         TabIndex        =   2
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5070
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   9720
      Begin MSComctlLib.ListView lstBill 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   8493
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Voucher No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Doc. No."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Exc. Register"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Supplier Type"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   9720
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Enabled         =   0   'False
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
         Left            =   5040
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
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
         Left            =   6360
         TabIndex        =   7
         Top             =   195
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_EXCISEList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VTYP As String
Dim STRSQL As String
Dim M_DBCD As String
Dim TEMPRS As ADODB.Recordset

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub CMDGO_Click()
Dim lstitem As ListItem
Dim PTYCOD As String
Dim SHDRQ As String
Dim O As Integer
        
    Screen.MousePointer = vbHourglass
    cmdGo.Default = False
    lstBill.ListItems.Clear
    
        If RS.State = 1 Then RS.Close
        STRSQL = "Select * from EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                 "' AND VTYP='EXD' AND DBCD='EXCREG' AND RECSTAT<>'D' AND DATE >= '" & Format(txtFrDate.Value, "MM/dd/YYYY") & _
                 "' and DATE <= '" & Format(txtToDate.Value, "MM/dd/YYYY") & "' ORDER BY DATE,SRNO" 'FOR BALANCE
        
        If RS.State = 1 Then RS.Close
        RS.Open STRSQL, CN, adOpenDynamic, adLockOptimistic
        STRSQL = Empty
        If RS.EOF = True Then
            MsgBox "There are no Record found.", vbInformation, App.Title
            cmdOk.Enabled = False
        Else
            Do While Not RS.EOF
                Set lstitem = lstBill.ListItems.ADD
                lstitem.Text = Format(RS!Date, "dd/MM/yyyy")
                lstitem.SubItems(1) = Trim(RS!SRNO & "")
                lstitem.SubItems(2) = Trim(RS!VBNO & "")
                lstitem.SubItems(3) = Trim(RS![TTYP] & "")
                lstitem.SubItems(4) = Trim(RS![EXTRA1] & "")
                lstitem.SubItems(5) = Trim(RS![EXTRA5] & "")
                RS.MoveNext
            Loop
            cmdOk.Enabled = True
            cmdOk.Default = True
            lstBill.SetFocus
        End If
        Screen.MousePointer = vbNormal
End Sub

Private Sub cmdOk_Click()
    With frm_servicetaxdb
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM EGPMAN WHERE COMP='" & compPth & "' AND VTYP='EXD' AND DBCD='EXCREG' AND UNIT='" & UNCD & _
              "' AND SRNO='" & Trim(lstBill.SelectedItem.SubItems(1)) & _
              "'  AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
         'Store All The Data In Respective Box
         .M_VBNO = Trim(lstBill.SelectedItem.SubItems(1))
         
         Select Case RS!TTYP
         Case "RG23-A"
           .EXCREG.ListIndex = 0
         Case "RG23-C"
           .EXCREG.ListIndex = 1
         Case Else
           .EXCREG.ListIndex = 2
         End Select
         
         Select Case RS!EXTRA1 & ""
          Case "Manufacturer"
             .SUPTYPE.ListIndex = 0
          Case "1st Stage Dealer"
             .SUPTYPE.ListIndex = 1
          Case Else
             .SUPTYPE.ListIndex = 2
         End Select
         
         .ENTRYNO = Trim(RS!SRNO)
         .ENTDAT.Value = RS!Date
         .VBNO.Text = Trim(RS!VBNO & "")
         .CENVAT.Text = RS!CENVAT
         .EDUCESS.Text = RS!EDUCESS
         .HEDUCESS.Text = RS!H_ED_CESS
         .ADUTY.Text = RS!A_DUTY
         .TXTEXTRA5.Text = Trim(RS!EXTRA5 & "")
         If IsNull(RS!CHDT) Then
            .txtodat.Value = RS!Date
         Else
            .txtodat.Value = RS!CHDT
         End If
      Else
         .M_VBNO = Empty
      End If
    End With
    
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)

    txtFrDate.Value = GetMinDate
    txtToDate.Value = GetMaxDate
    Set TEMPRS = New Recordset
    Me.KeyPreview = True
End Sub

Private Sub LSTBILL_GotFocus()
  lstBill.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub LSTBILL_LostFocus()
 lstBill.BackColor = vbWhite
End Sub


