VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRptPowerCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Consumption"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5850
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   5700
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4440
         TabIndex        =   9
         Top             =   255
         Width           =   1140
      End
      Begin VB.TextBox txtZoom 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         TabIndex        =   7
         Text            =   "100"
         Top             =   255
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2385
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Report Zoom %"
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
         TabIndex        =   6
         Top             =   315
         Width           =   1305
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5700
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   308
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   5700
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3720
         TabIndex        =   5
         Top             =   240
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
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   240
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
         Format          =   18284545
         CurrentDate     =   38429
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   308
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
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
         Height          =   195
         Left            =   2880
         TabIndex        =   4
         Top             =   308
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmRptPowerCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As ADODB.Recordset
Dim rptsql As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
    Call CenterChild(frm_Main, Me)
    dtFrom = GetMinDate
    dtTo = Date
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()

If txtUNIT.Text = Empty Then
   MsgBox "Please Select Unit !!", vbInformation, "Key Field Unit Is Missing"
   txtUNIT.SetFocus
   Exit Sub
End If

CRPT.Reset
crptConnect CRPT

ReportName = Empty
rptsql = Empty

PERIOD = dtFrom & " To " & dtTo

ReportName = App.PATH & "\Reports\PowerConsumption.rpt"

If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If

rptsql = "{POWERTRN.COMP}='" & compPth & "' AND {POWERTRN.UNIT}='" & txtUNIT.Tag & _
         "' And {POWERTRN.DATE}>=DATE(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") " & _
         "AND {POWERTRN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") "

RPTN = Me.Caption
        
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula rptsql
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT.Text & "'"
                           
         RPTN = RPTN + Space(5) + ReportName
                 
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
    
    Exit Sub
    
errPreview:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show 1
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
End Sub

Private Sub TXTZOOM_GotFocus():: txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub TXTZOOM_LostFocus(): txtZoom.BackColor = vbWhite: End Sub

Private Sub txtUNIT_GotFocus(): txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub
Private Sub txtUNIT_LostFocus():  txtUNIT.BackColor = vbWhite: End Sub


