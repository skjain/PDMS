VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRpt_Progss 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progressive Production Report"
   ClientHeight    =   3450
   ClientLeft      =   2925
   ClientTop       =   2580
   ClientWidth     =   6435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6435
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   75
      TabIndex        =   14
      Top             =   2040
      Width           =   6345
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   315
         Left            =   1005
         TabIndex        =   4
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18284545
         CurrentDate     =   37896
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   315
         Left            =   3765
         TabIndex        =   5
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18284545
         CurrentDate     =   37896
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   210
         Width           =   240
      End
      Begin VB.Label lblFrDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   75
      TabIndex        =   12
      Top             =   600
      Width           =   6345
      Begin VB.TextBox TXTPACKINGTYPE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   4890
      End
      Begin VB.TextBox txtMCNO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   590
         Width           =   4890
      End
      Begin VB.TextBox txtDVCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   210
         Width           =   4890
      End
      Begin VB.Label Label2 
         Caption         =   "Packing Type"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "M/C No "
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
         TabIndex        =   15
         Top             =   590
         Width           =   990
      End
      Begin VB.Label Label14 
         Caption         =   "Division "
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
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   75
      TabIndex        =   11
      Top             =   2640
      Width           =   6345
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport Crpt 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Preview"
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
         Image           =   "frmRpt_Progrss.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         Image           =   "frmRpt_Progrss.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         Caption         =   "&Zoom %"
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
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.Frame framCont 
      Height          =   615
      Left            =   75
      TabIndex        =   10
      Top             =   -60
      Width           =   6345
      Begin VB.TextBox txtUNIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Leave Blank To View All Unit Detail"
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label Label7 
         Caption         =   "Unit  "
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
         Height          =   210
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmRpt_Progss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    If txtUNIT = Empty Then
        MsgBox "Unit Is Missing !!", vbInformation, "Unit Alert"
        txtUNIT.SetFocus
        Exit Sub
    End If
    
    If txtDVCD = Empty Then
        MsgBox "Division Is Missing !!", vbInformation, "Division Alert"
        txtDVCD.SetFocus
        Exit Sub
    End If
    
    Call gatherData
    Crpt.WindowShowPrintSetupBtn = True
    Crpt.WindowShowProgressCtls = True
    Crpt.WindowShowSearchBtn = True
    Crpt.Destination = crptToWindow
    Crpt.WindowState = crptMaximized
    Crpt.Destination = crptToWindow
    Crpt.WindowTitle = "PROGRESSIVE PRODUCTION REPORT"
    Crpt.WindowState = crptMaximized
    
    Crpt.ACTION = 1
    Crpt.PageZoom Val(txtZoom)
End Sub

Private Sub Form_Activate()
 Call ColorComponent(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    txtFrDate.Value = FSDT
    txtToDate.Value = Now
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub gatherData()
Dim cPER As String
    Crpt.Reset
    crptConnect Crpt
    ReportName = App.PATH & "\Reports\RPT_Progressive.rpt"
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    Crpt.ReportFileName = ReportName
    
    rptsql = "({BOXREG.VBDT} >= Date(" & txtFrDate.Year & "," & txtFrDate.Month & "," & txtFrDate.Day & ") and {BOXREG.VBDT} <= Date(" & txtToDate.Year & "," & txtToDate.Month & "," & txtToDate.Day & ")) AND {BOXREG.UNIT}='" & txtUNIT.Tag & "' AND {BOXREG.DVCD}='" & txtDVCD.Tag & "' AND {BOXREG.RECSTAT}<>'D'"
    
    If txtMCNO <> Empty Then rptsql = rptsql & " AND {BOXREG.MCCD}='" & txtMCNO.Tag & "'"
    
    If Not TXTPACKINGTYPE = Empty Then rptsql = rptsql & " AND {BOXREG.DBCD}='" & TXTPACKINGTYPE.Tag & "' "
    
    Crpt.ReplaceSelectionFormula rptsql
    
    Crpt.DiscardSavedData = True
    
    cPER = Format(txtFrDate.Value, "dd/mm/yyyy") & " To " & Format(txtToDate.Value, "dd/mm/yyyy")
    
    Crpt.Formulas(1) = "PERIOD='" & CStr(cPER) & "'"
    Crpt.Formulas(2) = "STDT=#" & Format(txtFrDate.Value, "MM/dd/yyyy") & "#"
    Crpt.Formulas(3) = "ENDT=#" & Format(txtToDate.Value, "MM/dd/yyyy") & "#"
    Crpt.Formulas(4) = "DIVISION='" & txtDVCD & "'"
    Crpt.Formulas(5) = "UNIT='" & txtUNIT & "'"
    Crpt.Formulas(6) = "REPORTHEAD='Progressive Production Report'"
    Crpt.Formulas(7) = "COMPANY='" & compNm & "'"
End Sub

Private Sub txtDVCD_GotFocus()
txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        
        If txtUNIT = Empty Then
            MsgBox "Please Select Unit First !!", vbInformation, "Unit Missing !!"
            txtUnit_KeyDown vbKeyReturn, 0
            Exit Sub
        End If
        
        txtDVCD = SearchList1("SELECT TOP 20 Code,NAME From DIVMST Where COMP='" & compPth & "' and Unit='" & txtUNIT.Tag & "' AND CODE<>'000001' AND RECSTAT<>'D' ", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If

End Sub

Private Sub txtDVCD_LostFocus()
txtDVCD.BackColor = vbWhite
End Sub

Private Sub TXTMCNO_GotFocus()
txtMCNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMCNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtMCNO = Empty
    ElseIf KeyCode = vbKeyF2 Then
        M_DESC = Empty
        NEW_VISIBLE = False
        txtMCNO = SearchList1("SELECT TOP 20 Code,Name From MACMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select M/C FROM MASTER")
        txtMCNO.Tag = Key
    End If

End Sub

Private Sub TXTMCNO_LostFocus()
 txtMCNO.BackColor = vbWhite
End Sub

Private Sub TXTPACKINGTYPE_GotFocus()
  TXTPACKINGTYPE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPACKINGTYPE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTPACKINGTYPE = Empty) Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   TXTPACKINGTYPE = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & _
   "' AND UNIT='" & txtUNIT.Tag & "' AND VTYP='PPF' AND FYCD='" & FYCD & "'", 0, TXTPACKINGTYPE, "SELECT PACKING TYPE FROM LIST")
   TXTPACKINGTYPE.Tag = Key
ElseIf KeyCode = vbKeyDelete Then
   TXTPACKINGTYPE = Empty
   TXTPACKINGTYPE.Tag = Empty
End If
End Sub

Private Sub TXTPACKINGTYPE_LostFocus()
TXTPACKINGTYPE.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("SELECT TOP 20 Code,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If

End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub
