VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_VehicleEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Entry Register"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtUNIT 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "&Unit               "
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   5895
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/MM/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dtTo 
         Height          =   330
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/MM/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "&To Date       "
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
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "&From Date                "
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   5895
      Begin VB.TextBox txtTransport 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtVehicle 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Transport"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Vehicle No."
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
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   5895
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3240
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
         Image           =   "frmRPT_VehicleEntry.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4560
         TabIndex        =   8
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
         Image           =   "frmRPT_VehicleEntry.frx":0452
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
         Caption         =   "R&eport Zoom %"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRPT_VehicleEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RPTN As String
Dim m_unit As String
Dim ORDBOK As String
Dim ORDDBC As String
Dim sel_untcod As String
Dim SEL_DVCDNAM As String
Dim SEL_DVCDCOD As String
Dim M_DVCD As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    On Error GoTo errPreview
    
    If txtUNIT = Empty Then
       MsgBox "Pleas Select Unit", vbInformation
       txtUNIT.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtFrom) Then
       MsgBox "Pleas Select Correct Starting Date", vbInformation
       dtFrom.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(dtTo) Then
       MsgBox "Pleas Select Correct Ending Date", vbInformation
       dtTo.SetFocus
       Exit Sub
    End If
    
    rptsql = Empty
    RPTN = Empty
    CRPT.Reset
    crptConnect CRPT
    
    rptsql = "{VHCLENTRY.COMP}='" & compPth & "' AND {VHCLENTRY.UNIT} = '" & txtUNIT.Tag & "' AND {VHCLENTRY.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {VHCLENTRY.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")"
    
    If txtVehicle <> Empty Then rptsql = rptsql & " AND {VHCLENTRY.VHCD}='" & txtVehicle.Tag & "'"
    If txtTransport <> Empty Then rptsql = rptsql & " AND {VHCLENTRY.TRCD}='" & txtTransport.Tag & "'"
    
    ReportName = Empty
    
    ReportName = App.PATH & "\Reports\RPT_VEHICLENTRY.rpt"
    RPTN = "VEHICLE ENTRY REGISTER"
    
    CRPT.ReportFileName = ReportName
    
    PERIOD = dtFrom & " To " & dtTo
    
    CRPT.ReplaceSelectionFormula rptsql
    
     With CRPT
       .Formulas(1) = "COMPANY='" & compNm & "'"
       .Formulas(2) = "UNIT='" & txtUNIT & "'"
       .Formulas(3) = "REPORTHEAD='" & RPTN & "'"
       .Formulas(4) = "PERIOD='" & PERIOD & "'"
       
        RPTN = RPTN + Space(5) + ReportName
       .DiscardSavedData = True
       .WindowTitle = RPTN
       .Destination = crptToWindow
       .WindowState = crptMaximized
       .WindowShowPrintBtn = True
       .WindowShowPrintSetupBtn = True
       .WindowShowProgressCtls = True
       .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
         txtUNIT.SetFocus
        .ACTION = 1
        .PageZoom Val(txtZoom)
    End With
    
    Exit Sub

errPreview:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub dtFrom_GotFocus()
    dtFrom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub dtFrom_LostFocus()
    dtFrom.BackColor = vbWhite
End Sub

Private Sub dtTo_GotFocus()
    dtTo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub dtTo_LostFocus()
    dtTo.BackColor = vbWhite
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Call ColorComponent(Me)
 
    Call CenterChild(frm_Main, Me)
    
    dtFrom.Text = Format(Now, "dd/MM/yyyy")
    dtTo.Text = Format(Now, "dd/MM/yyyy")
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub txtTransport_GotFocus()
    txtTransport.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtTransport_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtTransport = SearchList1("Select  TOP 20 CODE,NAME From TRANSPORTMST ", 0, Empty, "Select Unit To View Report For ")
        txtTransport.Tag = Key
    
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtTransport = Empty
        txtTransport.Tag = Empty
    End If
End Sub

Private Sub txtTransport_LostFocus()
    txtTransport.BackColor = vbWhite
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
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtUNIT = Empty
        txtUNIT.Tag = Empty
    End If
End Sub

Private Sub txtUNIT_LostFocus()
    txtUNIT.BackColor = vbWhite
End Sub

Private Sub txtVehicle_GotFocus()
    txtVehicle.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtVehicle = SearchList1("Select  TOP 20 CODE,NAME From VHCLMST ", 0, Empty, "Select Unit To View Report For ")
        txtVehicle.Tag = Key
    
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtVehicle = Empty
        txtVehicle.Tag = Empty
    End If
End Sub

Private Sub txtVehicle_LostFocus()
    txtVehicle.BackColor = vbWhite
End Sub

Private Sub TXTZOOM_GotFocus()
    txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
    txtZoom.BackColor = vbWhite
End Sub
