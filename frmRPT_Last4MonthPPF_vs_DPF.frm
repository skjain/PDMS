VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_Last4MonthPPF_vs_DPF 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   Last 4 Month Production V/s Dispatch Summary"
   ClientHeight    =   4155
   ClientLeft      =   2940
   ClientTop       =   870
   ClientWidth     =   6900
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   6735
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   6735
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
         TabIndex        =   11
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3840
         TabIndex        =   12
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
         Image           =   "frmRPT_Last4MonthPPF_vs_DPF.frx":0000
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
         Left            =   5160
         TabIndex        =   13
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
         Image           =   "frmRPT_Last4MonthPPF_vs_DPF.frx":0452
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
         TabIndex        =   10
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   6735
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
         TabIndex        =   7
         Top             =   240
         Width           =   5085
      End
      Begin VB.Label LBLAGENT 
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
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      Height          =   585
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   2895
      Begin MSMask.MaskEdBox dtAson 
         Height          =   330
         Left            =   1560
         TabIndex        =   9
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
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&As on Date"
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
         Left            =   360
         TabIndex        =   8
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      Height          =   585
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   6735
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
         Width           =   5085
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
         ForeColor       =   &H80000006&
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
      TabIndex        =   14
      Top             =   120
      Width           =   6735
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
         Width           =   5085
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
Attribute VB_Name = "frmRPT_Last4MonthPPF_vs_DPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RPTN As String
Dim PERIOD As String
Public M1SD As String
Public M2SD As String
Public M3SD As String
Public M4SD As String

Private Sub dtAson_Validate(Cancel As Boolean)
    If Not IsDate(dtAson) And dtAson <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtAson.SetFocus
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
    dtAson = Format(Now(), "dd/mm/yyyy")
    Call SetReportFormat
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub cmdpreview_Click()

On Error GoTo errPreview
   Call Find4Month_Sdate
   
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
        MsgBox "Please Select Unit !!", vbInformation, "Unit Is Key Field Missing"
        txtUNIT.SetFocus
    End If
        
    If cboFormats.ListIndex = -1 Then cboFormats.ListIndex = 0
    
    Call SetReport
            
    If ReportName = Empty Then
        ReportErrorMessage 0
        Exit Sub
    End If
  
    
    CRPT.ReportFileName = ReportName
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    ReportName = RPTN + Space(5) + ReportName
        
    CRPT.ReplaceSelectionFormula rptsql
    
    PERIOD = dtFrom & " To " & dtTo
    
    With CRPT
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "UNIT='" & txtUNIT & "'"
        .Formulas(3) = "ASONDAT=#" & Format(dtAson, "MM/dd/yyyy") & "#"
 
        
        .Formulas(4) = "M1SD=#" & Format(M1SD, "MM/dd/yyyy") & "#"
        .Formulas(5) = "M2SD=#" & Format(M2SD, "MM/dd/yyyy") & "#"
        .Formulas(6) = "M3SD=#" & Format(M3SD, "MM/dd/yyyy") & "#"
        .Formulas(7) = "M4SD=#" & Format(M4SD, "MM/dd/yyyy") & "#"
                
        .Formulas(8) = "REPORTHEAD='" & RPTN & "'"
        
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        
        If cUName = "ADMIN" Then
             CRPT.WindowShowPrintBtn = True
             CRPT.WindowShowPrintSetupBtn = True
        ElseIf ReadConfigMaster("000069", 8, "R") Then
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
         .AddItem "Division+Item wise Production V/s Dispatch Summary"
    End With
    If cboFormats.ListCount > 0 Then cboFormats.ListIndex = 0
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
If txtUNIT.Tag = Empty Or txtUNIT = Empty Then txtUNIT.Enabled = True: txtUNIT.SetFocus: Exit Sub
If txtDVCD.Tag = Empty Or txtDVCD = Empty Then txtDVCD.Enabled = True: txtDVCD.SetFocus: Exit Sub

    If KeyCode = vbKeyF2 Then
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
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' and UNIT='" & txtUNIT.Tag & "' AND RECSTAT<>'D' AND CODE<>'000001'", 0, Empty, "Select Division")
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

Private Sub Find4Month_Sdate()

Dim CURMONTH As Long
CURMONTH = Month(dtAson)

Dim TMPSTR  As String

If CURMONTH > 3 Then

        'FIRST MONTH
        TMPSTR = CStr(CURMONTH - 3)
        If Len(TMPSTR) = 1 Then TMPSTR = "0" & TMPSTR
        
        M1SD = "01/" & TMPSTR & "/" & CStr(Year(dtAson))
        
        'SECOND MONTH
        TMPSTR = CStr(CURMONTH - 2)
        If Len(TMPSTR) = 1 Then TMPSTR = "0" & TMPSTR
        
        M2SD = "01/" & TMPSTR & "/" & CStr(Year(dtAson))
        
        'THIRD MONTH
        TMPSTR = CStr(CURMONTH - 1)
        If Len(TMPSTR) = 1 Then TMPSTR = "0" & TMPSTR
        
        M3SD = "01/" & TMPSTR & "/" & CStr(Year(dtAson))

ElseIf CURMONTH = 3 Then
       M3SD = "01/02/" & CStr(Year(dtAson))
       M2SD = "01/01/" & CStr(Year(dtAson))
       
       TMPSTR = CStr(Year(dtAson) - 1)
       M1SD = "01/12/" & TMPSTR
       
ElseIf CURMONTH = 2 Then
       M3SD = "01/01/" & CStr(Year(dtAson))
             
       TMPSTR = CStr(Year(dtAson) - 1)
       M2SD = "01/12/" & TMPSTR
       M1SD = "01/11/" & TMPSTR
ElseIf CURMONTH = 1 Then
                 
       TMPSTR = CStr(Year(dtAson) - 1)
       M3SD = "01/12/" & TMPSTR
       M2SD = "01/11/" & TMPSTR
       M1SD = "01/10/" & TMPSTR
End If
 
TMPSTR = CStr(Month(dtAson))
If Len(TMPSTR) = 1 Then TMPSTR = "0" & TMPSTR
 
M4SD = "01/" & TMPSTR & "/" & Year(dtAson)
 
End Sub

Private Sub SetReport()
Dim PKGTYP As String
rptsql = Empty
    
'---Find Lumpsum or carton
 If txtDVCD = Empty Then
    Dim CHKRS As ADODB.Recordset
    Set CHKRS = New ADODB.Recordset
    
    If CHKRS.State = 1 Then CHKRS.Close
    CHKRS.Open "SELECT DISTINCT PKGTYP FROM DIVMST  Where COMP='" & compPth & _
    "' And UNIT='" & txtUNIT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not CHKRS.EOF Then
       If CHKRS.RecordCount = 1 Then
          PKGTYP = Trim(CHKRS!PKGTYP & "")
       Else
          frmMsgPackType.Show 1
          PKGTYP = frmMsgPackType.Tag
          Unload frmMsgPackType
       End If
    End If
 Else
    PKGTYP = GetDivisionType(txtUNIT.Tag, txtDVCD.Tag)
 End If
'---------------------------
    
 'SET REPORT SQL
 If PKGTYP = "C" Then
    rptsql = "{BOXREGISTER.COMP}='" & compPth & "' And {BOXREGISTER.UNIT}='" & txtUNIT.Tag & "' And {BOXREGISTER.RECSTAT}='A' "
    
    If Not txtDVCD.Text = Empty Then rptsql = rptsql & "  AND {BOXREGISTER.DVCD}='" & txtDVCD.Tag & "'"
    If Not TXTITEM = Empty Then rptsql = rptsql & " AND {BOXREGISTER.ICOD}='" & TXTITEM.Tag & "'"
 Else
    rptsql = "{PKGMAN.COMP}='" & compPth & "' And {PKGMAN.UNIT}='" & txtUNIT.Tag & "' And {PKGMAN.RECSTAT}='A' "
    
    If Not txtDVCD.Text = Empty Then rptsql = rptsql & "  AND {PKGMAN.DVCD}='" & txtDVCD.Tag & "' "
    If Not TXTITEM = Empty Then rptsql = rptsql & " AND {PKGMAN.ICOD}='" & TXTITEM.Tag & "' "
 End If
 '--------------------------------
 
 'SET REPORT NAME
 Select Case cboFormats.ListIndex
 Case 0
       RPTN = "Division+Item Wise Production Vs Dispatch"
       If PKGTYP = "C" Then
          ReportName = App.PATH & "\Reports\Division+Item Wise Production Vs Dispatch.rpt"
       Else
          ReportName = App.PATH & "\Reports\Division+Item Wise Production Vs Dispatch(LUMPSUM).rpt"
       End If
 End Select
End Sub
