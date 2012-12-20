VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_DispatchRegWithoutDO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dispatch Register"
   ClientHeight    =   7155
   ClientLeft      =   2940
   ClientTop       =   870
   ClientWidth     =   7020
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framDBCD 
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   6810
      Begin VB.ComboBox cmbDispatchType 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   5145
      End
      Begin VB.Label Label4 
         Caption         =   "&Dispatch Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   6855
      Begin VB.OptionButton optClear 
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
         Height          =   255
         Left            =   4800
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optPending 
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
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Challan Status"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   120
      TabIndex        =   36
      Top             =   2520
      Width           =   6855
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   5145
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Rpt &Format"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   6855
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   29
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
         Image           =   "frmRPT_DispatchRegWithoutDO.frx":0000
         cBack           =   -2147483633
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   2880
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5040
         TabIndex        =   30
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
         Image           =   "frmRPT_DispatchRegWithoutDO.frx":0452
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
         TabIndex        =   27
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Width           =   6855
      Begin VB.TextBox txtConsignee 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   5085
      End
      Begin VB.TextBox txtsubgrad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   2085
      End
      Begin VB.TextBox txtltno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   5085
      End
      Begin VB.TextBox txtgrade 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   1845
      End
      Begin VB.TextBox txtParty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   5085
      End
      Begin VB.TextBox txtAgent 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5085
      End
      Begin VB.TextBox TXTITEM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   5085
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LBLitem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      Height          =   585
      Left            =   120
      TabIndex        =   33
      Top             =   1920
      Width           =   6855
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   180
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
         Left            =   4080
         TabIndex        =   9
         Top             =   180
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
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         Left            =   3240
         TabIndex        =   8
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&From Date"
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
         TabIndex        =   6
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   6855
      Begin VB.TextBox txtDVCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   5205
      End
      Begin VB.Label LBLDIV 
         BackStyle       =   0  'Transparent
         Caption         =   "&Division Name            "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Frame Frame7 
      Height          =   585
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   6855
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
         Height          =   300
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   5205
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Unit Name              "
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
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRPT_DispatchRegWithoutDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RPTN As String
Dim PERIOD As String

Private Sub dtFrom_Validate(Cancel As Boolean)
   If Not IsDate(dtFrom) And dtFrom <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtFrom.SetFocus
    End If
End Sub

Private Sub dtTo_Validate(Cancel As Boolean)
If Not IsDate(dtTo) And dtTo <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtTo.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Not txtUNIT = Empty Then Exit Sub
    Call txtUnit_KeyDown(vbKeyF2, 0)
    If txtUNIT = Empty Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtDVCD" And txtDVCD = Empty Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom = FSDT
    dtTo = FEDT
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    
    Call SetReportFormat
    Call SetDispatchType
End Sub

Private Sub cmdpreview_Click()
On Error GoTo errPreview
     
    CRPT.Reset
    crptConnect CRPT
    
    rptsql = Empty
    ReportName = Empty
    
    If cboFormats.ListIndex = -1 Then
        MsgBox "Invalid Report format !! Choose From List !!", vbInformation
        cboFormats.SetFocus
        SendKeys "%{DOWN}"
        Exit Sub
    End If
    
    If txtUNIT = Empty Then
        MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field : Missing"
        txtUNIT.SetFocus
        Exit Sub
    End If
    
    'If txtDVCD = Empty Then
        'MsgBox "Please Select Division !!", vbInformation, "Division Is Key Field : Missing"
        'txtDVCD.SetFocus
        'Exit Sub
    'End If
    
    If cmbDispatchType = Empty Then
        MsgBox "Please Select Dispatch Type !!", vbInformation, "Dispatch Is Key Field : Missing"
        cmbDispatchType.SetFocus
        Exit Sub
    End If
    
    'GetDispatchCode
    
    If cboFormats.ListIndex = -1 Then cboFormats.ListIndex = 0
    
    rptsql = "{SPTRAN.COMP}='" & compPth & "' AND {SPTRAN.VTYP}='DPF' And {SPTRAN.DATE}>=DATE(" & Year(dtFrom) & _
    "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {SPTRAN.DATE}<=DATE(" & Year(dtTo) & "," & Month(dtTo) & _
    "," & Day(dtTo) & ") And {SPTRAN.RECSTAT}<>'D' AND {SPTRAN.UNIT}='" & txtUNIT.Tag & _
    "' AND {SPTRAN.DBCD}='" & GetDispatchCode & "'"
    
    If Not txtDVCD = Empty Then rptsql = rptsql & " AND {SPTRAN.DVCD}='" & txtDVCD.Tag & "'   "
    If Not TXTITEM = Empty Then rptsql = rptsql & " AND {SPTRAN.ICOD}='" & TXTITEM.Tag & "' "
    If Not txtAgent = Empty Then rptsql = rptsql & " AND {SPTRAN.BRCD}='" & txtAgent.Tag & "' "
    If Not txtParty = Empty Then rptsql = rptsql & " AND {SPTRAN.PCOD}='" & txtParty.Tag & "' "
    If Not txtConsignee = Empty Then rptsql = rptsql & " AND {PADDMST.NAME}='" & txtConsignee & "' "
    
    If Not txtltno = Empty Then rptsql = rptsql & " AND {SPTRAN.LTNO}='" & txtltno & "' "
    If Not txtgrade = Empty Then rptsql = rptsql & " AND {SPTRAN.GRAD}='" & txtgrade.Tag & "' "
    If Not txtsubgrad = Empty Then rptsql = rptsql & " AND {SPTRAN.subgrd}='" & txtsubgrad.Tag & "' "
        
    If optPending.Value = True Then rptsql = rptsql & " AND ISNULL({SPTRAN.SVBN}) "
    If optClear.Value = True Then rptsql = rptsql & " AND NOT ISNULL({SPTRAN.SVBN}) "
                       
    Select Case cboFormats.ListIndex
       Case 0
           ReportName = App.PATH & "\Reports\Dispatch Without DO.rpt"
           RPTN = "Dispatch Register "
       Case 1
           ReportName = App.PATH & "\Reports\BoxWiseDispatchRegister.rpt"
           RPTN = "Dispatch Register with Box Details"
       Case 2
           ReportName = App.PATH & "\Reports\DatewiseDispatchRegisterWithoutDO.rpt"
           RPTN = "Date wise Dispatch Register "
    End Select
        
            
    If ReportName = Empty Then
        ReportErrorMessage 0
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
       
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
        
    CRPT.ReplaceSelectionFormula rptsql
    
    PERIOD = dtFrom & " To " & dtTo
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        If txtDVCD <> Empty Then
           .Formulas(3) = "DIVISION='" & txtDVCD & "'"
        Else
           .Formulas(3) = "DIVISION='ALL DIVISION'"
        End If
        .Formulas(4) = "PERIOD='" & PERIOD & "'"
        .Formulas(5) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(6) = "LBLVIEW='" & LabelDisplay(txtDVCD.Tag, txtUNIT.Tag) & "'"
        .Formulas(7) = "DISPATCH='" & cmbDispatchType & "'"
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000067", 8, "R") Then
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

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub SetReportFormat()
    With cboFormats
         .Clear
         .AddItem "Dispatch Register"
         .AddItem "Dispatch Register with Box Details"
         .AddItem "Datewise Dispatch Register"
    End With
    If cboFormats.ListCount > 0 Then cboFormats.ListIndex = 0
End Sub

Private Sub txtConsignee_GotFocus()
 txtConsignee.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtConsignee_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtConsignee = Empty
  ElseIf KeyCode = vbKeyF2 Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtConsignee = SearchList1("Select DISTINCT CODE,NAME From PADDMST WHERE RECSTAT='A'", 0, Empty, "Select Consignee Name ")
  End If
 Me.KeyPreview = True
End Sub

Private Sub txtConsignee_LostFocus()
  txtConsignee.BackColor = vbWhite
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       If txtUNIT.Tag = Empty Or txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub
       If txtDVCD.Tag = Empty Or txtDVCD = Empty Then txtDVCD.Enabled = True: txtDVCD.SetFocus: Exit Sub
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       TXTITEM = SearchList1("Select  TOP 20 Code,Name From FINITMMST where COMP='" & compPth & _
       "' and UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Item From List")
       TXTITEM.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       TXTITEM = Empty
       TXTITEM.Tag = Empty
    End If
End Sub


Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' and UNIT='" & txtUNIT.Tag & "' AND CODE <>'000001' AND RECSTAT<>'D'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If

End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_LostFocus()
 txtZoom.BackColor = vbWhite
End Sub

Private Sub txtItem_LostFocus()
 TXTITEM.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 TXTITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtAgent_LostFocus()
 txtAgent.BackColor = vbWhite
End Sub

Private Sub txtAgent_GotFocus()
 txtAgent.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtParty_LostFocus()
 txtParty.BackColor = vbWhite
End Sub

Private Sub txtParty_GotFocus()
 txtParty.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTZOOM_GotFocus()
 txtZoom.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboFormats_GotFocus()
    cboFormats.BackColor = RGB(BRED, BGREEN, BBLUE)
    SendKeys "%{DOWN}"
End Sub

Private Sub cboFormats_LostFocus()
 cboFormats.BackColor = vbWhite
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       txtParty = SearchList1("Select  TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party From List")
       txtParty.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       txtParty = Empty
       txtParty.Tag = Empty
    End If
End Sub

Private Sub txtAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       txtAgent = SearchList1("Select  TOP 20 Code,Name From REFMST where CATA='B'", 0, Empty, "Select Agent From List")
       txtAgent.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       txtAgent = Empty
       txtAgent.Tag = Empty
    End If
End Sub

Private Sub SetDispatchType()
    Dim PKTYPRS As ADODB.Recordset
    Set PKTYPRS = New ADODB.Recordset
    If PKTYPRS.State = 1 Then PKTYPRS.Close
    PKTYPRS.Open "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    "' AND VTYP='DPF' ", CN, adOpenDynamic, adLockOptimistic
    'AND NAME NOT LIKE '%WASTAGE%' AND NAME NOT LIKE '%CAPTIVE%'
    
    If Not PKTYPRS.EOF Then VTCD = Trim(PKTYPRS!CODE)
    Do While Not PKTYPRS.EOF
     cmbDispatchType.AddItem Trim(PKTYPRS!NAME)
    PKTYPRS.MoveNext
    Loop
    If cmbDispatchType.ListCount > 1 Then cmbDispatchType.ListIndex = 0
End Sub

Private Function GetDispatchCode() As String
GetDispatchCode = ""
    Dim PKTYPRS As ADODB.Recordset
    Set PKTYPRS = New ADODB.Recordset
    If PKTYPRS.State = 1 Then PKTYPRS.Close
    PKTYPRS.Open "SELECT DISTINCT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
    "' AND VTYP='DPF' AND NAME ='" & cmbDispatchType.Text & "'", CN, adOpenDynamic, adLockOptimistic
    'AND NAME NOT LIKE '%WASTAGE%' AND NAME NOT LIKE '%CAPTIVE%'
    If Not PKTYPRS.EOF Then GetDispatchCode = Trim(PKTYPRS!CODE)
End Function
Private Sub txtltno_GotFocus()
 txtltno.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       If txtUNIT.Tag = Empty Or txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub
       If txtDVCD.Tag = Empty Or txtDVCD = Empty Then txtDVCD.Enabled = True: txtDVCD.SetFocus: Exit Sub
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       txtltno = SearchList1("Select  TOP 20 LTNO,LTNO From TXULOT where COMP='" & compPth & _
       "' and UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "'", 0, Empty, "Select Lot No. From List")
       txtltno.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       txtltno = Empty
       txtltno.Tag = Empty
    End If
End Sub
Private Sub txtltno_LostFocus()
 txtltno.BackColor = vbWhite
End Sub
Private Sub txtgrade_GotFocus()
 txtgrade.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub txtgrade_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       txtgrade = SearchList1("Select  TOP 20 CODE,GRAD From GRDMST", 0, Empty, "Select Grade From List")
       txtgrade.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       txtgrade = Empty
       txtgrade.Tag = Empty
    End If
End Sub
Private Sub txtgrade_LostFocus()
 txtgrade.BackColor = vbWhite
End Sub




Private Sub txtsubgrad_GotFocus()
 txtsubgrad.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub txtsubgrad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
       If txtgrade.Text = Empty Then Exit Sub
       NEW_VISIBLE = False
       CANCEL_VISIBLE = True
       M_DESC = Empty
       txtsubgrad.Text = SearchList1("Select  TOP 20 SUBGRD,NAME From SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "' AND DVCD='" & txtDVCD.Tag & "' AND GRAD='" & txtgrade.Tag & "'", 0, Empty, "Select Sub-Grade From List")
       txtsubgrad.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
       txtsubgrad = Empty
       txtsubgrad.Tag = Empty
    End If
End Sub
Private Sub txtsubgrad_LostFocus()
 txtsubgrad.BackColor = vbWhite
End Sub





