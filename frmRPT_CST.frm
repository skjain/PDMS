VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_CST 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DVAT 31"
   ClientHeight    =   4095
   ClientLeft      =   7080
   ClientTop       =   3090
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   75
      TabIndex        =   16
      Top             =   60
      Width           =   5235
      Begin VB.CheckBox chkNewPage 
         Caption         =   "Print Summary On New Page"
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   7800
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Frame Frame4 
         Height          =   885
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   4980
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
            Left            =   960
            TabIndex        =   5
            Text            =   "100"
            Top             =   345
            Width           =   735
         End
         Begin WelchButton.lvButtons_H cmdPreview 
            Height          =   495
            Left            =   2280
            TabIndex        =   7
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
            Image           =   "frmRPT_CST.frx":0000
            cBack           =   -2147483633
         End
         Begin WelchButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   3600
            TabIndex        =   8
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
            Image           =   "frmRPT_CST.frx":0452
            cBack           =   -2147483633
         End
         Begin Crystal.CrystalReport CRPT 
            Left            =   2640
            Top             =   195
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label12 
            Caption         =   "Zoom %"
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
            Left            =   135
            TabIndex        =   14
            Top             =   345
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1125
         Left            =   135
         TabIndex        =   4
         Top             =   1680
         Width           =   4980
         Begin VB.ComboBox CMBTAX 
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
            TabIndex        =   6
            Top             =   720
            Width           =   3735
         End
         Begin VB.ComboBox cboBillType 
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
            TabIndex        =   3
            Top             =   300
            Width           =   3735
         End
         Begin VB.Label Label6 
            Caption         =   "Ag. Form"
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
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label5 
            Caption         =   "Bill Type :"
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
            Left            =   75
            TabIndex        =   12
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   135
         TabIndex        =   15
         Top             =   900
         Width           =   4980
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   330
            Left            =   3435
            TabIndex        =   2
            Top             =   240
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
            Format          =   53608449
            CurrentDate     =   38429
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   330
            Left            =   1050
            TabIndex        =   1
            Top             =   240
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
            Format          =   53608449
            CurrentDate     =   38429
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   11
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label1 
            Caption         =   "From :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   135
         TabIndex        =   17
         Top             =   150
         Width           =   4980
         Begin VB.TextBox txtUNIT 
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
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   255
            Width           =   3765
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
            TabIndex        =   9
            Top             =   285
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmRPT_CST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_DBCD As String

Private Sub cboBillType_GotFocus()
cboBillType.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboBillType_LostFocus()
cboBillType.BackColor = vbWhite
End Sub

Private Sub CMBTAX_GotFocus()
  CMBTAX.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub CMBTAX_LostFocus()
 CMBTAX.BackColor = vbWhite
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
Dim VATN As String
Dim TAXCD As String

    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM taxmst WHERE recstat='A' AND NAME='" & CMBTAX.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      TAXCD = RS!CODE
     Else
      TAXCD = Empty
    End If

    CRPT.Reset
    crptConnect CRPT
    
    If txtUNIT = Empty Then
        MsgBox "Please Select Unit !!", vbInformation
        txtUNIT.SetFocus
        Exit Sub
    End If

    txtUNIT.SetFocus
    rptsql = Empty
    ReportName = App.PATH & "\Reports\RPT_DVAT.rpt"
    rptsql = "{EGPMAN.COMP}='" & compPth & "' and ({EGPMAN.VTYP}='SAL' or {EGPMAN.VTYP}='DBN')  and {EGPMAN.RECSTAT}<>'D' AND {EGPMAN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {EGPMAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "
    
    If Not TAXCD = Empty Then rptsql = rptsql & " AND {EGPMAN.TXCD}='" & TAXCD & "'"
    If cboBillType.ListIndex >= 0 Then rptsql = rptsql & " AND {EGPMAN.RORT}='" & cboBillType & "'"
    If Not TAXCD = Empty Then rptsql = rptsql & " AND {EGPMAN.TXCD}='" & TAXCD & "'"
        
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    rptsql = rptsql & " AND {EGPMAN.UNIT} IN [" & txtUNIT.Tag & "]"
    
    CRPT.ReportFileName = ReportName

    If Not CMBTAX.Text = Empty Then
      RPTN = RPTN + " " + CMBTAX.Text
    End If
    
    RPTN = RPTN + " " + cboBillType.Text
    
    CRPT.ReplaceSelectionFormula rptsql
    PERIOD = dtFrom & " To " & dtTo
        
    With CRPT
        .Formulas(1) = "STDT=#" & Format(dtFrom, "MM/dd/yyyy") & "#"
        .Formulas(2) = "ENDT=#" & Format(dtTo, "MM/dd/yyyy") & "#"
                               
        RPTN = RPTN + Space(5) + ReportName
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If ReadConfigMaster("000053", 8, "R") Then
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
    
End Sub



Private Sub Form_Activate()
    If M_USRSECLEVL = "1" Then
      'If ReadConfigReport("0034", 7, "R") = False Then ModuleDeniedMessage_Report: Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
    
End Sub

Private Sub Form_Load()
    
Call ColorComponent(Me)
Call CenterChild(frm_Main, Me)
    
    dtFrom = GetMinDate
    dtTo = GetMaxDate
        
    Me.Tag = RPTPARA
    
    If Me.Tag = "PUR" Then
        Me.Caption = "VAT Form 201B"
    Else
        Me.Caption = "VAT Form 201A"
    End If
        
    If RS.State = 1 Then RS.Close
    RS.Open "select  distinct dgrq from igmmst where dgrq='Y'", CN, adOpenDynamic, adLockOptimistic

    CMBTAX.Clear
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM TAXMST WHERE RECSTAT='A' ORDER BY NAME", CN, adOpenDynamic, adLockOptimistic
    Do While Not RS.EOF
      CMBTAX.AddItem RS!NAME & ""
      RS.MoveNext
    Loop
    cboBillType.Clear
    cboBillType.AddItem "TAX INVOICE"
    cboBillType.AddItem "RETAIL INVOICE"
End Sub

Private Sub txtUNIT_GotFocus()
txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (txtUNIT = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
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


