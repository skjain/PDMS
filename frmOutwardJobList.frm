VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOutwardJobList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue List Help ItemWise"
   ClientHeight    =   7275
   ClientLeft      =   2130
   ClientTop       =   435
   ClientWidth     =   9705
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9705
   Begin VB.Frame framCmd 
      Height          =   630
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   9495
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
         Left            =   6405
         TabIndex        =   14
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
         Left            =   5040
         TabIndex        =   13
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame FramCont 
      Height          =   5235
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   9540
      Begin MSComctlLib.ListView lstBill 
         Height          =   5025
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   8864
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
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Slip No."
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   5397
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Item Description"
            Object.Width           =   4516
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qnty"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unique"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "VTYP"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame frmDTRNGE 
      Height          =   1320
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9555
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
         Top             =   840
         Width           =   915
      End
      Begin VB.OptionButton optJob 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&JOB WORK"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&RGP"
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
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optNRGP 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&NRGP"
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
         Left            =   7200
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   4440
         TabIndex        =   5
         Top             =   840
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   17694721
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   17694721
         CurrentDate     =   38429
      End
      Begin VB.Label lblToDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date: "
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
         Left            =   3480
         TabIndex        =   9
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date: "
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
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of &Transaction    (A)                                   (B)                         (C)"
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmOutwardJobList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LVTYP As String

Private Sub CMDGO_Click()
  lstBill.ListItems.Clear
  
  Call SetVTYP
  
  Dim EDTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Dim SQL As String
  
  SQL = Empty
  SQL = "SELECT JOBOUT.*,ITMMST.NAME AS ITEM,ACCMST.NAME AS PARTY FROM JOBOUT INNER JOIN ITMMST ON (JOBOUT.ICOD=ITMMST.CODE) " & _
  "INNER JOIN ACCMST ON (JOBOUT.PCOD=ACCMST.CODE) WHERE JOBOUT.COMP='" & compPth & "' AND JOBOUT.UNIT='" & UNCD & "' AND VTYP='" & LVTYP & _
  "' AND JOBOUT.DATE>='" & Format(txtFrDate.Value, "MM/DD/YYYY") & "' AND JOBOUT.DATE<='" & Format(txtToDate.Value, "MM/DD/YYYY") & _
  "'  AND JOBOUT.RECSTAT<>'D' AND JOBOUT.CLRSTATUS='N' AND MODE='N' ORDER BY JOBOUT.DATE,JOBOUT.VBNO"
   
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
    MsgBox "No Record found for given criteria ", vbInformation
    txtToDate.SetFocus
    Exit Sub
  End If
  
  Do While Not EDTDAT.EOF
   Set lstitem = lstBill.ListItems.Add
   lstitem.Text = Format(EDTDAT![Date], "dd/MM/yyyy")
   lstitem.SubItems(1) = EDTDAT![VBNO]
   lstitem.SubItems(2) = EDTDAT![PARTY]
   lstitem.SubItems(3) = EDTDAT![Item]
   lstitem.SubItems(4) = Format(EDTDAT!QNTY, "########.000")
   lstitem.SubItems(6) = EDTDAT!VTYP
   EDTDAT.MoveNext
  Loop
  
  cmdOk.Enabled = True
  cmdOk.Default = True
  If frmOutwardJobList.Visible = True Then lstBill.SetFocus
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If lstBill.SelectedItem.SubItems(1) <> Empty And lstBill.ListItems.COUNT < 1 Then
     lstBill.SetFocus
     Exit Sub
  End If
  
  With frmOutwardForJob
  Dim VBNO As String
  
  Dim EDTDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  Set EDTDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  
  Call SetVTYP
  .M_VTYP = LVTYP
  
  Dim SQL As String
  SQL = Empty
  SQL = "SELECT JOBOUT.*,ITMMST.NAME AS ITEM,ACCMST.NAME AS PARTY FROM JOBOUT INNER JOIN ITMMST ON (JOBOUT.ICOD=ITMMST.CODE) " & _
  "INNER JOIN ACCMST ON (JOBOUT.PCOD=ACCMST.CODE) WHERE JOBOUT.COMP='" & compPth & "' AND JOBOUT.UNIT='" & UNCD & _
  "' AND VTYP='" & LVTYP & "' AND JOBOUT.RECSTAT<>'D' AND VBNO='" & lstBill.SelectedItem.SubItems(1) & "'"
     
  If EDTDAT.State = 1 Then EDTDAT.Close
  EDTDAT.Open SQL, CN, adOpenDynamic, adLockOptimistic
  If EDTDAT.EOF Then
     lstBill.SetFocus
     Exit Sub
  Else
     .txtParty = EDTDAT!PARTY & ""
     .txtExpDays = EDTDAT!XDAYS & ""
     .TXTRMRK = EDTDAT!RMRK & ""
     .optJob.Value = optJob.Value
     .optRGP.Value = optRGP.Value
     .optNRGP.Value = optNRGP.Value
  End If
          
    .TXTVBNO = EDTDAT!VBNO
    .TXTVBDT = EDTDAT!Date
    .ITMFLEX.Rows = 2
    
    I = 1
    Do While Not EDTDAT.EOF
     .ITMFLEX.TextMatrix(I, 0) = Trim(EDTDAT!Item)
     .ITMFLEX.TextMatrix(I, 1) = Trim(EDTDAT!IDNO)
     .ITMFLEX.TextMatrix(I, 2) = nstr(Val(GetCode("ITMMST", Trim(EDTDAT!ICOD), "CODE", "BALQ")) + Val(EDTDAT!QNTY), 12, 3)
     .ITMFLEX.TextMatrix(I, 3) = Trim(nstr(EDTDAT!QNTY, 12, 3))
     .ITMFLEX.TextMatrix(I, 4) = Trim(nstr(EDTDAT!AMNT, 10, 2))
     .ITMFLEX.Rows = .ITMFLEX.Rows + 1
      I = I + 1
      EDTDAT.MoveNext
      
    If .ITMFLEX.Rows > 6 Then .ITMFLEX.TopRow = .ITMFLEX.TopRow + 2
    Loop
    
    .btn_sts (True)
    .cmdCancel.Cancel = True
  End With
  Unload Me
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  txtFrDate = GetMinDate
  txtToDate = GetMaxDate
  cmdOk.Enabled = False
  cmdCancel.Enabled = True
  Me.KeyPreview = True
End Sub

Private Sub optJob_Click()
Call CMDGO_Click
End Sub

Private Sub optJob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  SendKeys "{tab}"
End If
End Sub

Private Sub optNRGP_Click()
  Call CMDGO_Click
End Sub

Private Sub optRGP_Click()
  Call CMDGO_Click
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

Public Sub SetVTYP()
        If optJob.Value = True Then
           LVTYP = "ANX"
        ElseIf optRGP.Value = True Then
           LVTYP = "RGP"
        ElseIf optNRGP.Value = True Then
           LVTYP = "NGP"
        End If
End Sub


