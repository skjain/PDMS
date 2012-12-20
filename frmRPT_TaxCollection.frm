VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_TaxCollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Form Collection Report"
   ClientHeight    =   5805
   ClientLeft      =   4905
   ClientTop       =   2835
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   5760
   Begin Crystal.CrystalReport CRPT 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox CHKLET 
      Caption         =   "Letter Required"
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
      TabIndex        =   15
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   5505
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label Label8 
         Caption         =   "Unit Name :"
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
         TabIndex        =   0
         Top             =   285
         Width           =   990
      End
   End
   Begin VB.CheckBox chkBreak 
      Caption         =   "Page Break Required"
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
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   5205
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tax Form Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   5505
      Begin VB.OptionButton opTaxPart 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular "
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
         TabIndex        =   13
         Top             =   555
         Width           =   1245
      End
      Begin VB.OptionButton opTaxAll 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
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
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox txtTaxForm 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   3735
      End
   End
   Begin VB.TextBox txtZoom 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1080
      TabIndex        =   17
      Text            =   "100"
      Top             =   5280
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "Party Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   22
      Top             =   2175
      Width           =   5505
      Begin VB.OptionButton OPTPTYALL 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
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
         Left            =   150
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton OPTPTYPART 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular"
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
         Left            =   150
         TabIndex        =   10
         Top             =   555
         Width           =   1245
      End
      Begin VB.TextBox TXTPTYNAME 
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
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tax Form Date Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2700
      TabIndex        =   21
      Top             =   840
      Width           =   2895
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   1290
         TabIndex        =   8
         Top             =   765
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   57016321
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   1275
         TabIndex        =   6
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   57016321
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date :"
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
         Left            =   180
         TabIndex        =   5
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "End Date :"
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
         Left            =   195
         TabIndex        =   7
         Top             =   780
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tax Form Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   2505
      Begin VB.OptionButton OPTTAXALL 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
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
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.OptionButton OPTTAXPEND 
         Alignment       =   1  'Right Justify
         Caption         =   "Pending"
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
         TabIndex        =   3
         Top             =   555
         Width           =   2085
      End
      Begin VB.OptionButton OPTTAXCLR 
         Alignment       =   1  'Right Justify
         Caption         =   "Clear"
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
         TabIndex        =   4
         Top             =   840
         Width           =   2085
      End
   End
   Begin WelchButton.lvButtons_H CMDOK 
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   5160
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
      Image           =   "frmRPT_TaxCollection.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   5160
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
      Image           =   "frmRPT_TaxCollection.frx":0452
      cBack           =   -2147483633
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4200
      Width           =   5535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Zoom  :"
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
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Width           =   660
   End
End
Attribute VB_Name = "frmRPT_TaxCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_DBCD As String
Dim M_TXCD As String
Dim M_PCOD As String
Dim M_COMP As String
Dim M_BRCD As String

Private Sub CHKLET_Click()
  If CHKLET.Value = 1 Then chkBreak.Value = 1
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
On Error GoTo errPreview
Dim PERIOD As String
    
    PERIOD = CStr(txtFrDate) & " To " & CStr(txtToDate)
    

    
    If txtUNIT = Empty Then
        txtUnit_KeyDown vbKeyReturn, 0
    End If
    
    If txtUNIT = Empty Then Exit Sub
    
    CRPT.Reset
    
    crptConnect CRPT
    
    M_COMP = compPth
    txtUNIT.SetFocus
    rptsql = Empty
    If Me.Tag = "SAL" Then
        rptsql = "{BILLMAIN.COMP} IN ['" & M_COMP & "'] AND {BILLMAIN.BSTS}='A'  and {BILLMAIN.DATE}>=DATE(" & txtFrDate.Year & "," & txtFrDate.Month & "," & txtFrDate.Day & ") AND {BILLMAIN.DATE}<=DATE(" & txtToDate.Year & "," & txtToDate.Month & "," & txtToDate.Day & ") AND ({BILLMAIN.VTYP}='" & Me.Tag & "' OR {BILLMAIN.VTYP}='DBN' OR {BILLMAIN.VTYP}='OPC') AND {BILLMAIN.RECSTAT}<>'D'"
        
        If OPTTAXPEND.Value = True Then
            rptsql = rptsql & " AND (ISNULL({BILLMAIN.FORM}) OR TRIM({BILLMAIN.FORM})='')"
        ElseIf OPTTAXCLR.Value = True Then
            rptsql = rptsql & " AND TRIM({BILLMAIN.FORM})<>''"
        End If
        

        If opTaxPart.Value = True Then rptsql = rptsql & " AND {BILLMAIN.TXCD}='" & M_TXCD & "'"
        If OPTPTYPART.Value = True Then rptsql = rptsql & " AND {BILLMAIN.PCOD}='" & M_PCOD & "'"

        rptsql = rptsql & " AND {BILLMAIN.UNIT} IN [" & txtUNIT.Tag & "]"
    Else
        rptsql = "{PURMAN.COMP} IN ['" & M_COMP & "'] and {PURMAN.DATE}>=DATE(" & txtFrDate.Year & "," & txtFrDate.Month & "," & txtFrDate.Day & ") AND {PURMAN.DATE}<=DATE(" & txtToDate.Year & "," & txtToDate.Month & "," & txtToDate.Day & ") AND {PURMAN.VTYP}='" & Me.Tag & "' AND {PURMAN.RECSTAT}<>'D'"
        
        If OPTTAXPEND.Value = True Then
            rptsql = rptsql & " AND (ISNULL({PURMAN.FORM}) OR TRIM({PURMAN.FORM})='')"
        ElseIf OPTTAXCLR.Value = True Then
            rptsql = rptsql & " AND TRIM({PURMAN.FORM})<>''"
        End If
        

        If opTaxPart.Value = True Then rptsql = rptsql & " AND {PURMAN.TXCD}='" & M_TXCD & "'"
        If OPTPTYPART.Value = True Then rptsql = rptsql & " AND {PURMAN.PCOD}='" & M_PCOD & "'"

        
        rptsql = rptsql & " AND {PURMAN.UNIT} IN [" & txtUNIT.Tag & "]"
    
    End If
    If txtTaxForm <> Empty Then RPTN = RPTN & txtTaxForm
    If Me.Tag = "SAL" Then

            If CHKLET.Value = 1 Then
              ReportName = App.PATH & "\Reports\C-form Pending reminder.rpt"
            Else
              ReportName = App.PATH & "\Reports\c-form pending register.rpt"
            End If


    Else
       
       
            If CHKLET.Value = 1 Then
              ReportName = App.PATH & "\Reports\C-FORM PAYABLE LETTER (Purchase).rpt"
            Else
              ReportName = App.PATH & "\Reports\c-form pending register (Purchase).rpt"
            End If
       
    End If
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    
    CRPT.Formulas(1) = "PERIOD='" & PERIOD & "'"
    
    If OPTTAXALL.Value = True Then
       If Me.Tag = "SAL" Then
         RPTN = "TAX ALL FORM COLLECTION REGISTER"
        Else
         RPTN = "TAX ALL FORM ISSUE REGISTER"
       End If
    ElseIf OPTTAXCLR.Value = True Then
        If Me.Tag = "SAL" Then
          RPTN = "TAX CLEAR FORM COLLECTION REGISTER"
         Else
          If CHKLET.Value = 1 Then
            RPTN = "TAX CLEAR FORM ISSUE LETTER"
           Else
            RPTN = "TAX CLEAR FORM ISSUE REGISTER"
          End If
        End If
    Else
        If CHKLET.Value = 1 Then
          If Me.Tag = "SAL" Then
            RPTN = "SALE TAX PENDING FORM"
           Else
            RPTN = "SALE TAX PENDING FORM"
          End If
         Else
          If Me.Tag = "SAL" Then
            RPTN = "TAX PENDING FORM COLLECTION REGISTER"
           Else
            RPTN = "TAX PENDING FORM ISSUE REGISTER"
          End If
        End If
    End If
    
    If Me.Tag = "SAL" Then
        RPTN = "SALES " & RPTN
    Else
        RPTN = "PURCHASE " & RPTN
    End If
    

    
    CRPT.Formulas(3) = "PGBRK=" & chkBreak.Value
    CRPT.Formulas(4) = "UNIT='" & txtUNIT & "'"
    CRPT.WindowShowZoomCtl = True
    If ReadConfigMaster("000053", 8, "R") Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        Else
             CRPT.WindowShowPrintBtn = False
             CRPT.WindowShowPrintSetupBtn = False
        End If
    CRPT.WindowShowRefreshBtn = True
    CRPT.WindowShowSearchBtn = True
    CRPT.DiscardSavedData = True
    CRPT.ReplaceSelectionFormula rptsql
    CRPT.WindowState = crptMaximized
    CRPT.PageZoom Val(txtZoom)
    
    If OPTTAXALL.Value = True Then
        RPTN = "ALL TAX FORM COLLECTION REGISTER" & Space(10) & "Report : " & ReportName
    ElseIf OPTTAXCLR.Value = True Then
        RPTN = "CLEAR TAX FORM COLLECTION REGISTER" & Space(10) & "Report : " & ReportName
    Else
        RPTN = "PENDING TAX FORM COLLECTION REGISTER" & Space(10) & "Report : " & ReportName
    End If
    
    CRPT.WindowTitle = RPTN
    
    CRPT.ACTION = 1
    
    Exit Sub
    
errPreview:
    'Resume
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
    
End Sub

Private Sub Form_Activate()
  If RPTPARA = "SAL" Then
    Me.Caption = "SALES TAX FORM COLLECTION"
    RPTPARA = "SAL"
   Else
    Me.Caption = "SALES TAX FORM PAYABLE"
    RPTPARA = "PUR"
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)

    Me.Tag = RPTPARA
    
    TXTPTYNAME.Enabled = False
    txtTaxForm.Enabled = False
    txtFrDate.Value = FSDT
    txtToDate.Value = FEDT
    
    CHKLET.Visible = True
End Sub
Private Sub opTaxAll_Click()
    txtTaxForm.Enabled = False
    txtTaxForm = Empty
End Sub

Private Sub opTaxPart_Click()
    txtTaxForm.Enabled = True
    txtTaxForm.SetFocus
End Sub
Private Sub OPTPTYALL_Click()
    TXTPTYNAME.Enabled = False
    TXTPTYNAME = Empty
End Sub

Private Sub OPTPTYPART_Click()
    TXTPTYNAME.Enabled = True
    TXTPTYNAME.SetFocus
End Sub
Private Sub txtPTYName_GotFocus()
TXTPTYNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPTYName_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        M_DESC = Empty
        NEW_VISIBLE = False
        TXTPTYNAME = SearchList1("Select TOP 20 CODE,NAME From ACCMST", 0, Empty, "Select Party Account")
        M_PCOD = Key
    End If

End Sub

Private Sub txtPTYName_LostFocus()
 TXTPTYNAME.BackColor = vbWhite
End Sub

Private Sub txtTaxForm_GotFocus()
txtTaxForm.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtTaxForm_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        M_DESC = Empty
        NEW_VISIBLE = False
        txtTaxForm = SearchList1("Select TOP 20 CODE,NAME From TAXMST Where RECSTAT='A'", 0, Empty, "Select Tax Catagoery")
        M_TXCD = Key
    End If

End Sub

Private Sub txtTaxForm_LostFocus()
 txtTaxForm.BackColor = vbWhite
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
        LOAD frm_askunit
        If frm_askunit.LSTUNIT.ListCount > 0 Then
            frm_askunit.Show 1
        End If
        txtUNIT = sel_untnam
        txtUNIT.Tag = sel_untcod
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If


End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub
