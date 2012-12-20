VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_UserReport 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Detail Report"
   ClientHeight    =   3000
   ClientLeft      =   1995
   ClientTop       =   1935
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1965
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   6375
      Begin VB.TextBox TDSCT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   6
         ToolTipText     =   "Enter the Description of Item."
         Top             =   1200
         Width           =   3795
      End
      Begin VB.TextBox TXTUNIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   2
         ToolTipText     =   "Enter the Description of Item."
         Top             =   240
         Width           =   3795
      End
      Begin VB.TextBox TXTPCOD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   4
         ToolTipText     =   "Enter the Description of Item."
         Top             =   720
         Width           =   3795
      End
      Begin VB.OptionButton optMonthwise 
         BackColor       =   &H8000000B&
         Caption         =   "&User Rights Details"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   9
         Top             =   1680
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optSummary 
         BackColor       =   &H8000000B&
         Caption         =   "&Admin List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2040
         TabIndex        =   8
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Department 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Department :"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Company :"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Report Type :"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "User Name :"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1245
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      Height          =   885
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   6375
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Text            =   "100"
         Top             =   345
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2715
         Top             =   315
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3600
         TabIndex        =   13
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
         Image           =   "frmRPT_UserReport.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdcancel 
         Height          =   495
         Left            =   4920
         TabIndex        =   14
         Top             =   240
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
         Image           =   "frmRPT_UserReport.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Report &Zoom % :"
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
         Left            =   135
         TabIndex        =   11
         Top             =   405
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmRPT_UserReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    crpt.Reset
    crptConnect crpt
    If optSummary.Value = True Then
        ReportName = App.PATH & "\Reports\RPT_USERREPORT1.rpt"
    Else
       ReportName = App.PATH & "\Reports\RPT_USERREPORT.rpt"
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    
    txtUNIT.SetFocus
    rptsql = "{UserMast.COMP}='" & Trim(txtUNIT.Tag) & "'"
    
    If TDSCT.Text <> Empty Then rptsql = rptsql & " AND {UserMast.User_dept}='" & TDSCT.Tag & "'"
    If txtPCOD <> Empty Then rptsql = rptsql & " AND {UserMast.Uid}='" & txtPCOD.Tag & "'"
    
    crpt.ReportFileName = ReportName
    crpt.ReplaceSelectionFormula rptsql
    
    With crpt
    
      
        .DiscardSavedData = True
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
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "TDSCT" And TDSCT = Empty And KeyCode = vbKeyReturn Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    dtFrom = FSDT
    dtTo = Date
End Sub



Private Sub TDSCT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TDSCT.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM dept_MST WHERE COMP='" & compPth & "'", 0, Empty, "Select Department From List")
        TDSCT.Tag = Key
    End If
End Sub

Private Sub txtPCOD_GotFocus()
txtPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
     '   txtPCOD = SearchList1("Select TOP 20 uid, user_flnm From usermast", 0, Empty, "Select User From List")
     
         txtPCOD = SearchList1("Select  TOP 20 UID,USER_FLNM  From USERMAST", 0, Empty, "Select User From List")
        txtPCOD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtPCOD = Empty
    End If
End Sub

Private Sub txtPCOD_LostFocus()
txtPCOD.BackColor = vbWhite
End Sub

Private Sub TXTUNIT_GotFocus()
   txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (txtUNIT = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select TOP 20 COMP_PATH,COMP_NAME From COMPMAST ", 0, Empty, "Select Companey To View Report For ")
        txtUNIT.Tag = Key
'        Load frm_askunit
'        If frm_askunit.LSTUNIT.ListCount > 0 Then
'            frm_askunit.Show 1
'        End If
'        txtUNIT = sel_untnam
'        txtUNIT.Tag = sel_untcod
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub TXTUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub
