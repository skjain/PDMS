VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_LRPrinting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LR Printing"
   ClientHeight    =   6780
   ClientLeft      =   2940
   ClientTop       =   480
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      Height          =   630
      Left            =   90
      TabIndex        =   28
      Top             =   3315
      Width           =   7545
      Begin VB.FileListBox flInvoice 
         Height          =   285
         Left            =   6120
         Pattern         =   "LR*.RPT"
         TabIndex        =   30
         Top             =   210
         Visible         =   0   'False
         Width           =   840
      End
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
         Left            =   1800
         TabIndex        =   15
         Top             =   195
         Width           =   4125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Available Format :"
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
         TabIndex        =   29
         Top             =   225
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   630
      Left            =   90
      TabIndex        =   27
      Top             =   750
      Width           =   7530
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
         Height          =   330
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   5580
      End
      Begin VB.Label Label14 
         Caption         =   "Division :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   3840
      TabIndex        =   26
      Top             =   2685
      Width           =   3780
      Begin VB.OptionButton opPrePrinted 
         Caption         =   "Pre Printed"
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
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   1410
      End
      Begin VB.OptionButton opPlain 
         Caption         =   "&Plain"
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
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.Label Label7 
         Caption         =   "Vie&w :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   90
      TabIndex        =   25
      Top             =   2685
      Width           =   3645
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "frmRPT_LRPrinting.frx":0000
         Left            =   1800
         List            =   "frmRPT_LRPrinting.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   195
         Width           =   1740
      End
      Begin VB.Label Label6 
         Caption         =   "View Mode :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   90
      TabIndex        =   24
      Top             =   2040
      Width           =   7530
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   195
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
         Top             =   195
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
      Begin VB.Label Label5 
         Caption         =   "&To Date : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "&From Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.Frame framDBCD 
      Height          =   615
      Left            =   90
      TabIndex        =   23
      Top             =   1410
      Width           =   7530
      Begin VB.ComboBox cmbSaleType 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   210
         Width           =   5625
      End
      Begin VB.Label Label2 
         Caption         =   "Type of Sale :"
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
   Begin VB.Frame framDIVISION 
      Height          =   690
      Left            =   90
      TabIndex        =   22
      Top             =   60
      Width           =   7530
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   5565
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame Frame5 
      Height          =   795
      Left            =   120
      TabIndex        =   20
      Top             =   5880
      Width           =   7530
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
         Left            =   3720
         TabIndex        =   17
         Text            =   "100"
         Top             =   285
         Width           =   735
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   7080
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   4560
         TabIndex        =   18
         Top             =   195
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
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5880
         TabIndex        =   19
         Top             =   195
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
         cBack           =   -2147483633
      End
      Begin VB.Label Label13 
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
         Left            =   2880
         TabIndex        =   21
         Top             =   300
         Width           =   900
      End
   End
   Begin MSComctlLib.ListView lstInvoice 
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Inv. No"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Party Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bill Amount"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmRPT_LRPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim M_VBNO As String
Dim M_BILLDT As Date
Dim M_DBCD As String
Dim M_PARTY As String
Dim M_PAD1 As String
Dim M_PAD2 As String
Dim M_PAD3 As String
Dim M_DELPTY As String
Dim M_DAD1 As String
Dim M_DAD2 As String
Dim M_DAD3 As String

Dim M_LTNO As String
Dim M_EXCOM As String
Dim M_LRNO As String
Dim M_LRDT As String
Dim M_VHCL As String
Dim M_DESTINATION As String

Dim M_TQTY As Double
Dim rsSource As Recordset
Dim RSREF As Recordset

Dim NEW_UNIT As String
Dim NEW_DVCD As String
Dim NEW_DBCD As String

Private Sub cmbSaleType_GotFocus()
    cmbSaleType.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cmbSaleType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbSaleType_LostFocus()
cmbSaleType.BackColor = vbWhite
End Sub

Private Sub cboFormat_GotFocus()
cboFormat.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboFormat_LostFocus()
cboFormat.BackColor = vbWhite
End Sub

Public Sub cboStatus_Click()
    cboFormat.Clear
    If cboStatus.ListIndex = 0 Then
        opPlain.Enabled = True
        opPrePrinted.Enabled = True
        opPlain.Value = True
        cboFormat.Enabled = False
    Else
        opPlain.Enabled = True
        opPrePrinted.Enabled = False
        opPlain.Value = True
        cboFormat.Enabled = True
    End If
    
    Call CreateFormatList
    
End Sub

Private Sub cboStatus_GotFocus()
cboStatus.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboStatus_LostFocus()
    cboStatus.BackColor = vbWhite
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Public Sub cmdpreview_Click()
Dim M_CHLN As String
Dim ctr As Long
Dim i As Integer
    
    CRPT.Reset
    crptConnect CRPT
    
    M_CHLN = Empty
    
    M_DBCD = GetDBCDPDMS("CODE", "NAME", cmbSaleType, txtUNIT.Tag, "SAL")
    
    If cboFormat.ListIndex = -1 And cboFormat.Enabled Then
        MsgBox "Please Select Format From List !!", vbInformation, "Format Missing!!"
        cboFormat.SetFocus
        Exit Sub
    End If
    
    If cboStatus.ListIndex = 1 Then
        
        For i = 1 To lstInvoice.ListItems.COUNT
            If lstInvoice.ListItems(i).Checked = True Then
                If M_CHLN <> Empty Then M_CHLN = M_CHLN & ","
                M_CHLN = M_CHLN & "'" & Trim(lstInvoice.ListItems(i)) & "'"
            End If
        Next
        
        If M_CHLN = Empty Then
            MsgBox "No Item Selected !!", vbInformation, "No Information Found !!"
            If lstInvoice.ListItems.COUNT < 1 Then Call lstInvoice_GotFocus
            Exit Sub
        End If
        
        rptsql = "{BILLMAIN.COMP}='" & compPth & "' AND {SPTRAN.VTYP}='SAL' AND {BILLMAIN.RECSTAT}<>'D' " & _
                 "AND {BILLMAIN.VBNO} IN [" & M_CHLN & "] AND {BILLMAIN.UNIT}='" & txtUNIT.Tag & _
                 "' AND {BILLMAIN.DVCD}='" & txtDVCD.Tag & "' AND {BILLMAIN.DBCD}='" & M_DBCD & "'"
        
        If txtDVCD.Tag <> "000001" Then
           If cboFormat <> Empty Then
            ReportName = App.PATH & "\Reports\" & cboFormat
           Else
            ReportName = App.PATH & "\Reports\LR_GEN.RPT"
           End If
        End If
                
        If Dir(ReportName, vbNormal) = Empty Then
            ReportErrorMessage 1001
            Exit Sub
        End If
        
        CRPT.ReportFileName = ReportName
        
        RPTN = "Invoice : "
        
        CRPT.ReplaceSelectionFormula rptsql
    
        With CRPT
            RPTN = RPTN + Space(5) + ReportName
        
            .DiscardSavedData = True
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowShowProgressCtls = True
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
             txtDVCD.SetFocus
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .PageLast
            .PageFirst
            .ACTION = 1
            .PageZoom Val(txtZoom)
            Exit Sub
        End With
    Else
        
        If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
        Close #1
        Open "C:\DOSPRINT\" & ComputerName & ".TXT" For Output As #1
        
        For i = 1 To lstInvoice.ListItems.COUNT
            If lstInvoice.ListItems(i).Checked Then
                Select Case M_COMPBILL
                Case "TEX"
                    If opPlain.Value = True Then
                        Call LR_TEX(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    Else
                        Call LR_TEX1(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    End If
                Case "MCS"
                    If opPlain.Value = True Then
                        Call LR_MCS(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    Else
                        Call LR_MCS1(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    End If
                Case "MCK"
                    If opPlain.Value = True Then
                        Call LR_MCK(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    Else
                        Call LR_MCK1(txtUNIT.Tag, txtDVCD.Tag, M_DBCD, Trim(lstInvoice.ListItems(i)))
                    End If
                End Select
            End If
        Next
        
        Close #1
    End If
        
    If Not BILLPRINTONLINE Then
        frmRPT_DosViewer.Show
        frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & ".TXT")
    Else
        LOAD frmRPT_DosViewer
        frmRPT_DosViewer.Hide
        frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & ".TXT")
        frmRPT_DosViewer.PrintSlip
    End If
    
End Sub

Private Sub dtFrom_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub dtFrom_Validate(Cancel As Boolean)
    If Not IsDate(dtFrom) And dtFrom <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtFrom.SetFocus
    End If
End Sub

Private Sub dtTo_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_Validate(Cancel As Boolean)
    If Not IsDate(dtTo) And dtTo <> "__/__/____" Then
        Cancel = True
        MsgBox "Please Enter Valid Date !!", vbInformation, "Date Format Checking !!"
        dtTo.SetFocus
    End If
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  'cboFormat.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

    End Sub

Private Sub Form_Load()
    Call CenterChild(frm_Main, Me)
    dtFrom = Format(Now, "DD/MM/YYYY")
    dtTo = Format(Now, "DD/MM/YYYY")
    cboStatus.ListIndex = 1
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    txtUNIT.Enabled = False
    Call SetSaleType
End Sub

Public Sub lstInvoice_GotFocus()
 lstInvoice.BackColor = RGB(BRED, BGREEN, BBLUE)
 Call GenInvList("")
End Sub

Private Sub lstInvoice_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  If Item.Checked = True Then
    SendKeys "{DOWN}"
  End If
End Sub

Private Sub lstInvoice_LostFocus()
  lstInvoice.BackColor = vbWhite
End Sub

Private Sub txtDVCD_GotFocus()
  txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Public Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("SELECT TOP 20 Code,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & txtUNIT.Tag & "' AND RECSTAT<>'D'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
    
    If cmbSaleType.ListCount > 0 Then cmbSaleType.ListIndex = 0

End Sub

Private Sub txtDVCD_LostFocus()
txtDVCD.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Public Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtDVCD = Empty
        txtUNIT = SearchList1("SELECT TOP 20 Code,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
End Sub

Public Sub InvoicePrint(dbcd As String, unit As String, DVCD As String)
 
    frmRPT_DosViewer.Show
    frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & "-CHALLAN.TXT")
    
    Exit Sub
 
errChallanPrint:
    Close #1
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show
End Sub

Private Sub CreateFormatList()
Dim FY_FSO As FileSystemObject
Dim fldr As Folder
Dim fil As File
Dim FLS As Files
Dim fName As String
Dim fPath As String
Dim i As Integer
    
    fPath = App.PATH
    If Right(fPath, 1) = "\" Then fPath = Left(fPath, Len(fPath) - 1)
    flInvoice.PATH = fPath & "\Reports"
    
    For i = 0 To flInvoice.ListCount - 1
        fName = UCase("LR_" & M_COMPBILL & ".rpt")
        
           If fName = UCase(flInvoice.List(i)) Or UCase(flInvoice.List(i)) = UCase("LR_GEN.RPT") Then
              cboFormat.AddItem flInvoice.List(i)
           End If
    Next
    
End Sub

Public Sub GenInvList(VBNO As String)
Dim SQL As String
Static LSQL As String
Dim Item As ListItem
    
    M_DBCD = GetDBCDPDMS("CODE", "NAME", cmbSaleType, txtUNIT.Tag, "SAL")
    
    SQL = "SELECT BILLMAIN.date,BILLMAIN.vbno,accmst.name,BILLMAIN.tpcs,BILLMAIN.BNET from BILLMAIN,accmst where BILLMAIN.pcod=accmst.code and BILLMAIN.COMP='" & compPth & "' AND BILLMAIN.vtyp='SAL' and BILLMAIN.dbcd='" & M_DBCD & "' AND RECSTAT<>'D' AND DVCD='" & txtDVCD.Tag & "' AND BILLMAIN.UNIT='" & UNCD & "' "
    
    If VBNO <> Empty Then
        SQL = SQL & " AND VBNO = '" & VBNO & "' "
    End If
    
    If OnlineBillNum <> Empty Then
        SQL = SQL & " AND VBNO = '" & OnlineBillNum & "' "
    Else
        SQL = SQL & " AND BILLMAIN.date>='" & Format(dtFrom, "MM/dd/yyyy") & "' and BILLMAIN.date<= '" & Format(dtTo, "MM/dd/yyyy") & "' "
    End If
    
    SQL = SQL & " ORDER BY BILLMAIN.VBNO,BILLMAIN.DATE "
        
    LSQL = SQL
    Set rsTemp = New Recordset
    rsTemp.Open LSQL, CN
    
    lstInvoice.ListItems.Clear
    
    Do While Not rsTemp.EOF
        Set Item = lstInvoice.ListItems.ADD
        Item.Text = rsTemp!VBNO
        If VBNO <> Empty Then Item.Checked = True: M_VBNO = Item.Text
        If OnlineBillNum <> Empty Then Item.Checked = True: OnlineBillNum = Item.Text
        Item.SubItems(1) = rsTemp!Date
        Item.SubItems(2) = rsTemp!NAME
        Item.SubItems(3) = Format(rsTemp!BNET, "#######.000")
        rsTemp.MoveNext
    Loop
    rsTemp.Close

End Sub

Private Sub txtUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub

Private Sub SetSaleType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open "SELECT NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
Do While Not PKTYPRS.EOF
 cmbSaleType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop

If cmbSaleType.ListCount >= 1 Then cmbSaleType.ListIndex = 0

End Sub

Public Function LR_TEX(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_TEX
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,SPTRAN.LTNO,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3 " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(40), 40)
        M_PAD1 = Left(rsSource!POAD1 & Space(40), 40)
        M_PAD2 = Left(rsSource!POAD2 & Space(40), 40)
        M_PAD3 = Left(rsSource!POAD3 & Space(40), 40)
        
        M_DELPTY = Left(rsSource!UNAME & Space(40), 40)
        M_DAD1 = Left(rsSource!DFAD1 & Space(40), 40)
        M_DAD2 = Left(rsSource!DFAD2 & Space(40), 40)
        M_DAD3 = Left(rsSource!DFAD3 & Space(40), 40)
            
        M_LTNO = Left(rsSource!ltno & Space(10), 10)
            
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(13), 13)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(26), 26)
        
        Print #1, Space(60) + DCA + M_LRNO + DCI
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(2) + DCA + CWA + M_DELPTY + Space(35) + M_PARTY + CWI + DCI
        Print #1, Space(2) + CWA + M_DAD1 + Space(35) + M_PAD1 + CWI
        Print #1, Space(2) + CWA + M_DAD2 + Space(35) + M_PAD2 + CWI
        Print #1, Space(2) + CWA + M_DAD3 + Space(35) + M_PAD3 + CWI
        Print #1,
        Print #1, Space(6) + "SILVASA " + Space(20) + M_DESTINATION + Space(12) + M_LRDT
        Print #1,
        Print #1,
        Print #1,
        Print #1, nstr(rsSource!TPCS, 5, 0) + Space(1) + CWA + M_EXCOM + CWI + Space(1) + nstr(rsSource!TQTY, 8, 3)
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(6) + "CHLN NO.: " + M_VBNO
        Print #1, Space(6) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(6) + "LOT NO. : " + M_LTNO
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        
 End If
End Function

Public Function LR_TEX1(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_TEX
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,SPTRAN.LTNO,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,VHCLMST.NAME AS VEHICLE,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3  " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        "LEFT JOIN VHCLMST ON SPTRAN.VEHICALNO = VHCLMST.CODE " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(32), 32)
        M_DELPTY = Left(rsSource!UNAME & Space(32), 32)
        
        M_LTNO = Left(rsSource!ltno & Space(10), 10)
        
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(10), 10)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        M_VHCL = Left(rsSource!VEHICLE & Space(10), 10)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(30), 30)
        
        
        Print #1,
        Print #1, Space(70) + DCA + M_LRNO + DCI
        Print #1,
        Print #1, Space(70) + DCA + M_VHCL + DCI
        Print #1,
        Print #1, Space(10) + DCA + M_DELPTY + Space(7) + M_PARTY + DCI
        Print #1,
        Print #1,
        Print #1, Space(63) + M_LRDT
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(2) + nstr(rsSource!TPCS, 5, 0) + Space(2) + CWA + M_EXCOM + CWA + Space(6) + nstr(rsSource!TQTY, 8, 3)
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(9) + "CHLN NO.: " + M_VBNO
        Print #1, Space(9) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(9) + "LOT NO. : " + M_LTNO
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
 End If
End Function

Public Function LR_MCS(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_MCS
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,VHCLMST.NAME AS VHCL,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3 " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "LEFT JOIN VHCLMST ON SPTRAN.VEHICALNO = VHCLMST.CODE " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(40), 40)
        M_PAD1 = Left(rsSource!POAD1 & Space(40), 40)
        M_PAD2 = Left(rsSource!POAD2 & Space(40), 40)
        M_PAD3 = Left(rsSource!POAD3 & Space(40), 40)
        
        M_DELPTY = Left(rsSource!UNAME & Space(40), 40)
        M_DAD1 = Left(rsSource!DFAD1 & Space(40), 40)
        M_DAD2 = Left(rsSource!DFAD2 & Space(40), 40)
        M_DAD3 = Left(rsSource!DFAD3 & Space(40), 40)
        
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(13), 13)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        M_VHCL = Left(rsSource!VHCL & Space(15), 15)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(30), 30)
        
        Print #1,
        Print #1,
        Print #1, Space(61) + DCA + M_LRNO + DCI
        Print #1, Space(61) + DCA + M_LRDT + DCI
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(2) + DCA + CWA + M_DELPTY + Space(12) + "    SURANGI        " + Space(12) + M_PARTY + CWI + DCI
        Print #1, Space(2) + CWA + M_DAD1 + Space(43) + M_PAD1 + CWI
        Print #1, Space(2) + CWA + M_DAD2 + Space(16) + M_DESTINATION + Space(12) + M_PAD2 + CWI
        Print #1, Space(2) + CWA + M_DAD3 + Space(43) + M_PAD3 + CWI
        
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(11) + nstr(rsSource!TPCS, 5, 0) + Space(4) + CWA + M_EXCOM + CWI + Space(7) + nstr(rsSource!TQTY, 10, 3)
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(20) + "CHLN NO.: " + M_VBNO
        Print #1, Space(20) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(20) + "VHCL NO.: " + M_VHCL
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
 End If
End Function

Public Function LR_MCS1(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_MCS
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,VHCLMST.NAME AS VHCL,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3 " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "LEFT JOIN VHCLMST ON SPTRAN.VEHICALNO = VHCLMST.CODE " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(40), 40)
        M_PAD1 = Left(rsSource!POAD1 & Space(40), 40)
        M_PAD2 = Left(rsSource!POAD2 & Space(40), 40)
        M_PAD3 = Left(rsSource!POAD3 & Space(40), 40)
        
        M_DELPTY = Left(rsSource!UNAME & Space(40), 40)
        M_DAD1 = Left(rsSource!DFAD1 & Space(40), 40)
        M_DAD2 = Left(rsSource!DFAD2 & Space(40), 40)
        M_DAD3 = Left(rsSource!DFAD3 & Space(40), 40)
        
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(13), 13)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        M_VHCL = Left(rsSource!VHCL & Space(15), 15)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(15), 15)
        
        Print #1, Space(60) + DCA + M_LRNO + DCI
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        
        Print #1, Space(2) + DCA + CWA + M_DELPTY + Space(30) + M_PARTY + CWI + DCI
        Print #1, Space(2) + CWA + M_DAD1 + Space(30) + M_PAD1 + CWI
        Print #1, Space(2) + CWA + M_DAD2 + Space(30) + M_PAD2 + CWI
        Print #1, Space(2) + CWA + M_DAD3 + Space(30) + M_PAD3 + CWI
        Print #1,
        Print #1, Space(6) + "SURANGI " + Space(20) + M_DESTINATION + Space(12) + M_LRDT
        Print #1,
        Print #1,
        Print #1, nstr(rsSource!TPCS, 5, 0) + Space(2) + M_EXCOM + Space(3) + nstr(rsSource!TQTY, 8, 3)
        Print #1,
        Print #1,
        Print #1, Space(6) + "CHLN NO.: " + M_VBNO
        Print #1, Space(6) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(6) + "VHCL NO.: " + M_VHCL
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
 End If
End Function

Public Function LR_MCK(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_MCS
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,VHCLMST.NAME AS VHCL,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3 " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "LEFT JOIN VHCLMST ON SPTRAN.VEHICALNO = VHCLMST.CODE " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(40), 40)
        M_PAD1 = Left(rsSource!POAD1 & Space(40), 40)
        M_PAD2 = Left(rsSource!POAD2 & Space(40), 40)
        M_PAD3 = Left(rsSource!POAD3 & Space(40), 40)
        
        M_DELPTY = Left(rsSource!UNAME & Space(40), 40)
        M_DAD1 = Left(rsSource!DFAD1 & Space(40), 40)
        M_DAD2 = Left(rsSource!DFAD2 & Space(40), 40)
        M_DAD3 = Left(rsSource!DFAD3 & Space(40), 40)
        
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(13), 13)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        M_VHCL = Left(rsSource!VHCL & Space(15), 15)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(15), 15)
        
        Print #1,
        Print #1,
        Print #1, Space(61) + DCA + M_LRNO + DCI
        Print #1, Space(61) + DCA + M_LRDT + DCI
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(2) + DCA + CWA + M_DELPTY + Space(12) + "SURANGI        " + Space(12) + M_PARTY + CWI + DCI
        Print #1, Space(2) + CWA + M_DAD1 + Space(39) + M_PAD1 + CWI
        Print #1, Space(2) + CWA + M_DAD2 + Space(12) + M_DESTINATION + Space(12) + M_PAD2 + CWI
        Print #1, Space(2) + CWA + M_DAD3 + Space(39) + M_PAD3 + CWI
        
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(11) + nstr(rsSource!TPCS, 5, 0) + Space(2) + M_EXCOM + Space(8) + nstr(rsSource!TQTY, 10, 3)
        Print #1,
        Print #1,
        Print #1,
        Print #1, Space(18) + "CHLN NO.: " + M_VBNO
        Print #1, Space(18) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(18) + "VHCL NO.: " + M_VHCL
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
 End If
End Function

Public Function LR_MCK1(unit As String, DVCD As String, dbcd As String, INVNO As String)
'*******************************************************************************************
'Module Name : LR_MCS
'
'Company Name:
'
'Dev. Date :
'
'Parameter : UNIT , DIVISION , DAYBOK CODE , INVOICE NO
'
'Developed By : DHARMENDRA
'
'Module Purpose : LR Generation
'********************************************************************************************

    Dim spc_invhead As String

    Call CollectData(unit)
    
    Set RS = New Recordset
    Set rsSource = New Recordset
    
    SQL = "Select BILLMAIN.*,SPTRAN.CHLN,SPTRAN.CHDT,ACCMST.NAME AS PARTY,ACCMST.POAD1,ACCMST.POAD2,ACCMST.POAD3," & _
        "CITYMASTER.NAME AS CITY,PADDMST.NAME AS CNAME,PADDMST.ADD1 AS CADD1,PADDMST.ADD2 AS CADD2," & _
        "PADDMST.ADD3 AS CADD3,DIVMST.EXCOMMODITY,VHCLMST.NAME AS VHCL,UNTMST.NAME AS UNAME,UNTMST.DFAD1,UNTMST.DFAD2,UNTMST.DFAD3 " & _
        " FROM BILLMAIN INNER JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP " & _
        " AND SPTRAN.VTYP=BILLMAIN.VTYP AND SPTRAN.DBCD=BILLMAIN.DBCD AND " & _
        "SPTRAN.UNIT=BILLMAIN.UNIT AND SPTRAN.DVCD=BILLMAIN.DVCD AND " & _
        "SPTRAN.VBNO=BILLMAIN.VBNO INNER JOIN ACCMST ON ACCMST.CODE=BILLMAIN.PCOD " & _
        "LEFT JOIN DIVMST ON DIVMST.CODE=BILLMAIN.DVCD AND " & _
        "DIVMST.COMP=BILLMAIN.COMP AND DIVMST.UNIT=BILLMAIN.UNIT " & _
        "LEFT JOIN VHCLMST ON SPTRAN.VEHICALNO = VHCLMST.CODE " & _
        "LEFT JOIN CITYMASTER ON BILLMAIN.DSTN = CITYMASTER.CODE " & _
        "INNER JOIN UNTMST ON UNTMST.CODE=BILLMAIN.UNIT AND UNTMST.COMP=BILLMAIN.COMP " & _
        "INNER JOIN PADDMST ON PADDMST.CODE=SPTRAN.DCOD AND PADDMST.SRNO=SPTRAN.ADDRESS " & _
        " WHERE BILLMAIN.UNIT='" & unit & "' AND BILLMAIN.DVCD='" & DVCD & "' AND BILLMAIN.DBCD='" & dbcd & _
        "' AND BILLMAIN.VBNO ='" & INVNO & "' AND BILLMAIN.COMP='" & compPth & _
        "' AND SPTRAN.RECSTAT<>'D' AND BILLMAIN.VTYP='SAL'"
    
    If rsSource.State = 1 Then rsSource.Close
    rsSource.Open SQL, CN
            
    If rsSource.EOF = False Then
        M_PARTY = Left(rsSource!PARTY & Space(40), 40)
        M_PAD1 = Left(rsSource!POAD1 & Space(40), 40)
        M_PAD2 = Left(rsSource!POAD2 & Space(40), 40)
        M_PAD3 = Left(rsSource!POAD3 & Space(40), 40)
        
        M_DELPTY = Left(rsSource!UNAME & Space(40), 40)
        M_DAD1 = Left(rsSource!DFAD1 & Space(40), 40)
        M_DAD2 = Left(rsSource!DFAD2 & Space(40), 40)
        M_DAD3 = Left(rsSource!DFAD3 & Space(40), 40)
        
        M_VBNO = Left(rsSource!chln & Space(6), 6)
        M_BILLDT = Left(CStr(rsSource!CHDT) & Space(10), 10)
        M_LRNO = Left(rsSource!LRNO & Space(13), 13)
        M_LRDT = Left(CStr(IIf(IsNull(rsSource!LRDT), "", rsSource!LRDT)) & Space(10), 10)
        M_DESTINATION = Left(rsSource!CITY & Space(15), 15)
        M_VHCL = Left(rsSource!VHCL & Space(15), 15)
        
        M_EXCOM = Left(rsSource!EXCOMMODITY & Space(15), 15)
        
        Print #1, Space(60) + DCA + M_LRNO + DCI
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        
        Print #1, Space(2) + DCA + CWA + M_DELPTY + Space(30) + M_PARTY + CWI + DCI
        Print #1, Space(2) + CWA + M_DAD1 + Space(30) + M_PAD1 + CWI
        Print #1, Space(2) + CWA + M_DAD2 + Space(30) + M_PAD2 + CWI
        Print #1, Space(2) + CWA + M_DAD3 + Space(30) + M_PAD3 + CWI
        Print #1,
        Print #1, Space(6) + "SURANGI " + Space(20) + M_DESTINATION + Space(12) + M_LRDT
        Print #1,
        Print #1,
        Print #1, nstr(rsSource!TPCS, 5, 0) + Space(2) + M_EXCOM + Space(3) + nstr(rsSource!TQTY, 8, 3)
        Print #1,
        Print #1,
        Print #1, Space(6) + "CHLN NO.: " + M_VBNO
        Print #1, Space(6) + "CHLN DT.: " + CStr(M_BILLDT)
        Print #1, Space(6) + "VHCL NO.: " + M_VHCL
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
 End If
End Function



