VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmperiodicstockreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finish Stock Report Periodic"
   ClientHeight    =   6285
   ClientLeft      =   7005
   ClientTop       =   2400
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6930
   Begin VB.Frame Frame6 
      Height          =   675
      Left            =   120
      TabIndex        =   15
      Top             =   3915
      Width           =   6720
      Begin VB.TextBox txtGrade 
         Height          =   315
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtLTNo 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "&Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "&Lot No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame framFormat 
      Caption         =   "Report Format :"
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
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2355
      Width           =   6720
      Begin VB.ComboBox cboFormats 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   300
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "&Format :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   5475
      Width           =   6720
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   27
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   3000
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   4080
         TabIndex        =   28
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Pre&view"
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
         Image           =   "frmperiodicstockreport.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5400
         TabIndex        =   29
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
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
         Image           =   "frmperiodicstockreport.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         Caption         =   "Report &Zoom %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   1485
      End
   End
   Begin VB.Frame framDenier 
      Height          =   810
      Left            =   120
      TabIndex        =   12
      Top             =   3075
      Width           =   6720
      Begin VB.TextBox txtItem 
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   345
         Width           =   5490
      End
      Begin VB.Label Label2 
         Caption         =   "Den&ier :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   6720
      Begin MSComCtl2.DTPicker dtOPDT 
         Height          =   330
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   21889025
         CurrentDate     =   39343
      End
      Begin MSComCtl2.DTPicker dtenDT 
         Height          =   330
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   21889025
         CurrentDate     =   39343
      End
      Begin VB.Label Label5 
         Caption         =   "To Date"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
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
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6720
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   5565
      End
      Begin VB.Label Label8 
         Caption         =   "&Unit "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1515
      Width           =   6690
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmperiodicstockreport.frx":09EC
         Left            =   1080
         List            =   "frmperiodicstockreport.frx":09EE
         TabIndex        =   8
         Tag             =   "0"
         Text            =   "Select Type of Packing"
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "&Pkg Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.Frame Frame8 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   795
      Width           =   6690
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label14 
         Caption         =   "&Division :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmperiodicstockreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboFormats_GotFocus()
 cboFormats.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboFormats_LostFocus()
 cboFormats.BackColor = vbWhite
End Sub

Private Sub cboFormats_Validate(Cancel As Boolean)

 dtOPDT.Value = Now
 dtOPDT.Enabled = True

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()

Dim FLAG As Boolean
If RS.State = 1 Then RS.Close
RS.Open "SELECT PKGTYP FROM DIVMST  Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "' AND  NAME = '" & txtDVCD.Text & "'", CN, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then If Trim(RS!PKGTYP & "") = "L" Then FLAG = True Else FLAG = False
    

On Error GoTo errViewRepoRt
    If txtUNIT = Empty Or txtUNIT.Tag = Empty Then
        MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
        txtUNIT.SetFocus
    End If
  
    CRPT.Reset
    crptConnect CRPT
    
    Select Case cboFormats.ListIndex
        Case 0
          If FLAG Then MsgBox "Under Construction": Exit Sub
          RPTN = "Denier + Lot + Grade + Subgrade Wise Finish Stock Report "
          ReportName = App.PATH & "\Reports\Finish Stock report Periodic.rpt"
    End Select
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
                
    Call SetRPTSQL(FLAG)
    
    
    
        
    CRPT.ReplaceSelectionFormula rptsql
        
    Dim PERIOD As String
    PERIOD = "From Date " + Format(dtOPDT, "dd/mm/yyyy") + " To Date " + Format(dtenDT, "dd/mm/yyyy")
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT & "'"
         .Formulas(5) = "OPDT=#" & Format(dtOPDT, "MM/dd/yyyy") & "#"
        .Formulas(6) = "ENDT=#" & Format(dtenDT, "MM/dd/yyyy") & "#"
        .Formulas(7) = "PERIOD='" & PERIOD & "'"
        
        
                
        .DiscardSavedData = True
         RPTN = RPTN + Space(5) + ReportName
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
         txtUNIT.SetFocus
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
               
    Exit Sub
    
errViewRepoRt:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub Form_Activate()
   Call ColorComponent(Me)

    Me.KeyPreview = True
    
    If txtUNIT = Empty Then
        Call txtUnit_KeyDown(vbKeyF2, 0)
    End If
      
    If txtUNIT = Empty Then Unload Me: Exit Sub
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtDVCD" And txtDVCD = Empty Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
    Me.Left = Me.Left - 900
    dtOPDT = Now
    dtenDT = Now
    With cboFormats
            .AddItem "Denier+Lot+Grade+SubGrade Wise Finish Stock Report(Detail)"
            .ListIndex = 0
    End With
    dtOPDT.Enabled = True
    dtenDT.Enabled = True
    dtenDT = Now
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtUNIT) = Empty Or Trim(txtUNIT.Tag) = Empty Then txtUNIT.SetFocus: Exit Sub
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' and UNIT='" & txtUNIT.Tag & "' AND RECSTAT='A' AND CODE<>'000001'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If

End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtgrade_GotFocus()
 txtGrade.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtgrade_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtGrade = Empty
 ElseIf KeyCode = vbKeyF2 Then
    txtGrade.Text = SearchList1("select TOP 20 code,grad from grdmst", 0, txtGrade, "SELECT GRADE FROM MASTER")
    txtGrade.Tag = Key
 End If
End Sub

Private Sub txtgrade_LostFocus()
 txtGrade.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 txtItem.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtItem_LostFocus()
 txtItem.BackColor = vbWhite
End Sub

Private Sub txtltno_GotFocus()
 txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

Dim SQL As String
If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtLTNo = Empty
ElseIf KeyCode = vbKeyF2 Then
   M_DESC = Empty:   NEW_VISIBLE = False
   SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
   "' AND DVCD='" & txtDVCD.Tag & "' "
   
   If txtItem <> Empty Then
     SQL = SQL & "AND FICD='" & txtItem.Tag & "'"
   End If
   
   txtLTNo = SearchList(SQL)
End If
Me.KeyPreview = True

End Sub

Private Sub txtltno_LostFocus()
 txtLTNo.BackColor = vbWhite
End Sub

Private Sub txtUNIT_Change()
  txtDVCD = Empty
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
        Call SetPackingType
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtUNIT) = Empty Or Trim(txtUNIT.Tag) = Empty Then txtUNIT.SetFocus: Exit Sub
If Trim(txtDVCD) = Empty Or Trim(txtDVCD.Tag) = Empty Then txtDVCD.SetFocus: Exit Sub

    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtItem = SearchList1("Select TOP 20 Code,Name From FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
        "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Finish Item Form List")
        txtItem.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtItem = Empty
        txtItem.Tag = Empty
    End If
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
 txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
 txtZoom.BackColor = vbWhite
End Sub

Private Sub cmbPackingType_KeyPress(KeyAscii As Integer): KeyAscii = 0: End Sub

Private Sub SetPackingType()
cmbPackingType.Clear
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
'CAN BE DISTINGUISH AT PREVIEW BUTTON : SO USE UNCD
PKTYPRS.Open "SELECT SDESC FROM VTYPDESC WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
"' AND MVTYP='PPF' AND SDESC NOT LIKE '%WASTAGE%'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!SDESC)
 PKTYPRS.MoveNext
Loop
 
 cmbPackingType.AddItem "------ALL-PACKING------"
 
If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 2

End Sub

Private Function GetPackingCode() As String
GetPackingCode = ""
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close

PKTYPRS.Open "SELECT SCOD FROM VTYPDESC WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
"' AND MVTYP='PPF' AND SDESC='" & cmbPackingType.Text & "'", CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then
 GetPackingCode = Trim(PKTYPRS!SCOD & "")
End If

End Function

Private Sub SetRPTSQL(FLAG As Boolean)

        
    rptsql = Empty
    rptsql = "{BOXREGISTER.COMP}='" & compPth & "' AND {BOXREGISTER.UNIT}='" & txtUNIT.Tag & _
    "' AND {BOXREGISTER.RECSTAT}<>'D' AND {BOXREGISTER.VBDT}<=DATE(" & Year(dtenDT) & "," & Month(dtenDT) & "," & Day(dtenDT) & ") "
    
    If (cmbPackingType.ListIndex <> cmbPackingType.ListCount - 1) Then
       rptsql = rptsql & " AND {BOXREGISTER.DBCD}='" & GetPackingCode & "'"
       RPTN = "FINISH STOCK REPORT : (" & cmbPackingType.Text & " )"
    Else
       RPTN = "FINISH STOCK REPORT : (ALL PACKING TYPE)"
    End If
    
    If txtDVCD <> Empty Then rptsql = rptsql & " AND {BOXREGISTER.DVCD}='" & txtDVCD.Tag & "'"
    If txtLTNo <> Empty Then rptsql = rptsql & " AND {BOXREGISTER.LOTNO}='" & txtLTNo & "'"
    If txtGrade <> Empty Then rptsql = rptsql & " AND {BOXREGISTER.GRAD}=" & txtGrade.Tag & ""
    If txtItem <> Empty Then rptsql = rptsql & " AND {BOXREGISTER.ICOD}='" & txtItem.Tag & "'"
    
    
'    rptsql = rptsql & " AND {BOXREGISTER.VTYP} IN ['PPF','OPN'] "
    
    


End Sub


