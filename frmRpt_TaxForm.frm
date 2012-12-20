VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRpt_TaxForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Sale Tax Form Register "
   ClientHeight    =   6015
   ClientLeft      =   4050
   ClientTop       =   2790
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6720
   Begin VB.Frame Frame5 
      Caption         =   "For Period :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   3
      Top             =   795
      Width           =   6495
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1365
         TabIndex        =   5
         Top             =   315
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
         Format          =   53215233
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   4920
         TabIndex        =   7
         Top             =   315
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
         Format          =   53215233
         CurrentDate     =   39538
      End
      Begin VB.Label Label9 
         Caption         =   "&To Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3420
         TabIndex        =   6
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "Fr&om Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   120
      TabIndex        =   26
      Top             =   4800
      Width           =   6495
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1800
         TabIndex        =   23
         Text            =   "100"
         Top             =   345
         Width           =   615
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
         Left            =   3840
         TabIndex        =   24
         Top             =   240
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
         Image           =   "frmRpt_TaxForm.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5160
         TabIndex        =   25
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmRpt_TaxForm.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         Caption         =   "Report &Zoom %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   27
         Top             =   345
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   6495
      Begin VB.ComboBox cboFormat 
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   4650
      End
      Begin VB.Label Label3 
         Caption         =   "&Format :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   1500
      Width           =   6495
      Begin VB.ComboBox CMBTYP 
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   4650
      End
      Begin VB.TextBox TXTBRCD 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2760
         Width           =   4650
      End
      Begin VB.TextBox TXTARCD 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2280
         Width           =   4650
      End
      Begin VB.TextBox txtREFCD 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1860
         Width           =   4650
      End
      Begin VB.ComboBox cboDivision 
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   645
         Width           =   4650
      End
      Begin VB.ComboBox cboUnit 
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   4650
      End
      Begin VB.TextBox txtPCOD 
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
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1455
         Width           =   4650
      End
      Begin VB.Label Label7 
         Caption         =   "Sale Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label6 
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1290
      End
      Begin VB.Label lblREF 
         Caption         =   "Ta&x Category :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   1290
      End
      Begin VB.Label Label2 
         Caption         =   "&Party Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   105
         TabIndex        =   15
         Top             =   1455
         Width           =   1170
      End
      Begin VB.Label Label14 
         Caption         =   "&Division :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Width           =   1170
      End
      Begin VB.Label Label8 
         Caption         =   "&Unit Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   105
         TabIndex        =   9
         Top             =   240
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmRpt_TaxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RPTN As String
Dim PERIOD As String
Dim m_unit As String
Dim M_DVCD As String
Dim M_DBCD As String
Dim M_CRAC As String

Private Sub cboDivision_Click()
    M_DVCD = Empty
    If cboDivision.ListIndex = -1 Then Exit Sub
End Sub

Private Sub cboDivision_GotFocus()
 cboDivision.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboDivision_LostFocus()
 cboDivision.BackColor = vbWhite
End Sub

Private Sub cboFormat_GotFocus()
cboFormat.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboFormat_LostFocus()
 cboFormat.BackColor = vbWhite
End Sub

Private Sub cboUnit_Click()
    m_unit = Empty
    cboDivision.Clear
    If cboUnit.ListIndex = -1 Then Exit Sub
    If cboUnit <> "<ALL>" Then
        m_unit = GetCode("UNTMST", cboUnit, "NAME", "CODE")
    End If
    Call FillCmb("SELECT NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & m_unit & "' AND RECSTAT<>'D'", cboDivision)
    cboDivision.AddItem "<ALL>"
End Sub

Private Sub cboUnit_GotFocus()
 cboUnit.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboUnit_LostFocus()
 cboUnit.BackColor = vbWhite
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & m_unit & "' AND NAME='" & cboDivision.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      M_DVCD = RS!CODE
    End If
    
    
    If cboFormat.ListIndex = -1 Then
        MsgBox "Please Select Report Format !!", vbInformation, "Report Format Missing"
        cboFormat.SetFocus
        Exit Sub
    End If
    
    If cboUnit.ListIndex = -1 Then
        MsgBox "Please Select Unit From List !!", vbInformation, "Unit is Empty"
        cboUnit.SetFocus
        Exit Sub
    End If
    
    Dim DBCD_CODE As String
    
    If CMBTYP = "<ALL>" Then
      DBCD_CODE = Empty
     Else
      If RS.State = 1 Then RS.Close
      RS.Open "select code from serialmaster where comp='" & compPth & "' and unit='" & m_unit & "' and name='" & Trim(CMBTYP) & "' and fycd='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
      If Not RS.EOF Then
        DBCD_CODE = Trim(RS!CODE)
       Else
        MsgBox "Invalid Sale Type"
        CMBTYP.SetFocus
       Exit Sub
     End If
   End If
    
    
    CRPT.Reset
    crptConnect CRPT
    
    ReportName = Empty
    rptsql = Empty
    Select Case cboFormat.ListIndex
        Case 0
            ReportName = App.PATH & "\Reports\SaleTaxFormRegister.rpt"
        Case Else
            MsgBox "Please Select Valid Report From List !!", vbInformation, "Invalid Report format"
    End Select
    
    RPTN = cboFormat.Text
    
    If Not DBCD_CODE = Empty Then
      RPTN = RPTN + " Book Name : " + CMBTYP.Text
    End If
            
    PERIOD = dtFrom & " To " & dtTo
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    
    rptsql = "{BILLMAIN.COMP}='" & compPth & "' AND {BILLMAIN.VTYP}='SAL' AND {BILLMAIN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {BILLMAIN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") And {BILLMAIN.RECSTAT}<>'D'"
        
    If cboUnit.ListIndex <> -1 And cboUnit.Text <> "<ALL>" Then rptsql = rptsql & "  AND {BILLMAIN.UNIT}='" & m_unit & "'"
    
    If cboDivision.ListIndex <> -1 And cboDivision.Text <> "<ALL>" Then rptsql = rptsql & " AND {BILLMAIN.DVCD}='" & M_DVCD & "'"
 
    If Not txtPCOD = Empty Then rptsql = rptsql & " AND {BILLMAIN.PCOD}='" & txtPCOD.Tag & "'"
    
    If Not TXTARCD = Empty Then rptsql = rptsql & " AND {ACCMST.ARCD}='" & TXTARCD.Tag & "'"

    If Not TXTBRCD = Empty Then rptsql = rptsql & " AND {BILLMAIN.BRCD}='" & TXTBRCD.Tag & "'"
    
    If Not DBCD_CODE = Empty Then rptsql = rptsql & " AND {BILLMAIN.DBCD}='" & DBCD_CODE & "'"
        
    
    If Not txtREFCD = Empty Then
        rptsql = rptsql & "  AND {BILLMAIN.TXCD}='" & txtREFCD.Tag & "'"
    End If
     cboFormat.SetFocus
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & cboUnit & "'"
        .Formulas(5) = "DIVISION='" & cboDivision & "'"
        .Formulas(6) = "SALETYPE='" & CMBTYP & "'"
        
        RPTN = RPTN + Space(5) + ReportName
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        If ReadConfigMaster("000051", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
    Exit Sub
    
errPreview:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub Form_Activate()
    Call ColorComponent(Me)
    Me.Caption = "SALES REGISTER"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
    'cboUnit.AddItem "<ALL>"
     Call FillCmb("SELECT NAME FROM UNTMST WHERE COMP='" & compPth & "'", cboUnit)
     If M_CUNT = "N" Then cboUnit.AddItem "<ALL>"
  '   Call FillCmb("SELECT NAME FROM ACCMST WHERE CODE IN (SELECT DISTINCT CRAC FROM BILLMAIN WHERE COMP='" & compPth & "' and VTYP='SAL')", cboBookName)
  '   cboBookName.AddItem "<ALL>"
     dtFrom = FSDT
     dtTo = GetMaxDate
     Me.Tag = RPTPARA
     With cboFormat
        .Clear
        .AddItem "SALE REGISTER DATE WISE"
        .AddItem "SALE REGISTER PARTY WISE"
        .AddItem "SALE REGISTER TAX CATEGORY WISE"
        .AddItem "SALE REGISTER AGENT WISE"
        .AddItem "SALE REGISTER CST SALE "
        .AddItem "SALE REGISTER WITH EXCISE DETAIL"
        .AddItem "SALE REGISTER DATE WISE2"
        .AddItem "SALE REGISTER DATE WISE(Excise Series)"
    End With
    
    Call FillCmb("SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND FYCD='" & FYCD & "' AND VTYP='SAL'", CMBTYP)
    CMBTYP.AddItem "<ALL>"
    CMBTYP.ListIndex = 0
    cboFormat.ListIndex = 0
End Sub

Private Sub txtARCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        TXTARCD = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA='A'", 0, Empty, "Select Area From List")
        TXTARCD.Tag = Key
        If TXTARCD <> Empty Then SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTARCD = Empty
    End If
End Sub

Private Sub TXTBRCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        TXTBRCD = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA='B'", 0, Empty, "Select Agent From List")
        TXTBRCD.Tag = Key
        If TXTBRCD <> Empty Then SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTBRCD = Empty
    End If
End Sub

Private Sub txtPCOD_GotFocus()
 txtPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        txtPCOD = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select Party From List")
        txtPCOD.Tag = Key
        If txtPCOD <> Empty Then SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtPCOD = Empty
    End If

End Sub

Private Sub txtPCOD_LostFocus()
 txtPCOD.BackColor = vbWhite
End Sub

Private Sub txtREFCD_GotFocus()
txtREFCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtREFCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        If lblREF.Caption = "&Area" Then
            txtREFCD.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM REFMST WHERE CATA='A'", 0, "", "List Of Areas")
        Else
            txtREFCD.Text = SearchList1("SELECT  TOP 20 CODE, NAME FROM TAXMST WHERE RECSTAT='A'", 0, "", "List Of Tax Forms")
        End If
            txtREFCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtREFCD = Empty
        txtREFCD.Tag = Empty
    End If

End Sub


Private Sub txtREFCD_LostFocus()
 txtREFCD.BackColor = vbWhite
End Sub




