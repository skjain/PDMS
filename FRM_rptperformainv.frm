VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRM_rptperformainv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proforma Invoice"
   ClientHeight    =   6825
   ClientLeft      =   3720
   ClientTop       =   1770
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7575
   Begin VB.Frame Frame4 
      Height          =   2760
      Left            =   120
      TabIndex        =   22
      Top             =   3180
      Width           =   7425
      Begin MSComctlLib.ListView lstInvoice 
         Height          =   2415
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   4260
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
            Text            =   "Order No."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Height          =   795
      Left            =   120
      TabIndex        =   24
      Top             =   5940
      Width           =   7410
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
         Left            =   1080
         TabIndex        =   26
         Text            =   "100"
         Top             =   285
         Width           =   735
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
         TabIndex        =   28
         Top             =   240
         Width           =   1140
      End
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
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   1140
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   5640
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
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
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame framDIVISION 
      Height          =   690
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   5205
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
         TabIndex        =   1
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Frame framDBCD 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   630
      Width           =   7410
      Begin VB.ComboBox cboDaybok 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         Top             =   210
         Width           =   5145
      End
      Begin VB.Label Label2 
         Caption         =   "&SalesMan :"
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
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   7410
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1215
         TabIndex        =   8
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
         Left            =   3960
         TabIndex        =   10
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
         TabIndex        =   7
         Top             =   255
         Width           =   885
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
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1905
      Width           =   2940
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FRM_rptperformainv.frx":0000
         Left            =   1200
         List            =   "FRM_rptperformainv.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   195
         Width           =   1500
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
         TabIndex        =   12
         Top             =   225
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   3165
      TabIndex        =   14
      Top             =   1905
      Width           =   4380
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
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
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
         TabIndex        =   17
         Top             =   240
         Width           =   1410
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame Frame7 
      Height          =   630
      Left            =   120
      TabIndex        =   18
      Top             =   2535
      Width           =   7425
      Begin VB.ComboBox cboFormat 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   195
         Width           =   4365
      End
      Begin VB.FileListBox flInvoice 
         Height          =   285
         Left            =   5640
         Pattern         =   "INVOICE*.RPT"
         TabIndex        =   21
         Top             =   210
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Format :"
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
         TabIndex        =   19
         Top             =   225
         Width           =   705
      End
   End
End
Attribute VB_Name = "FRM_rptperformainv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_VBNO As String
Dim M_DBCD As String

Private Sub cboDaybok_GotFocus()
cboDaybok.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub cboDaybok_LostFocus()
cboDaybok.BackColor = vbWhite
End Sub

Private Sub cboFormat_GotFocus()
cboFormat.BackColor = RGB(BRED, BGREEN, BBLUE)
cboFormat.ListIndex = 0
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdpreview_Click()
Dim M_CHLN As String
Dim ctr As Long
    
    
    CRPT.Reset
    crptConnect CRPT
    
    M_CHLN = Empty
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM SALMANMST WHERE NAME='" & cboDaybok.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      M_DBCD = RS!CODE
     Else
      M_DBCD = Empty
    End If
    
    
    If cboFormat.ListIndex = -1 And cboFormat.Enabled Then
        MsgBox "Please Select Format From List !!", vbInformation, "Format Missing!!"
        cboFormat.SetFocus
        Exit Sub
    End If
    
    If cboStatus.ListIndex = 1 Then
        
        For I = 1 To lstInvoice.ListItems.COUNT
            If lstInvoice.ListItems(I).Checked = True Then
                If M_CHLN <> Empty Then M_CHLN = M_CHLN & ","
                M_CHLN = M_CHLN & "'" & Trim(lstInvoice.ListItems(I)) & "'"
            End If
        Next
        
        If M_CHLN = Empty Then
            MsgBox "No Item Selected !!", vbInformation, "No Information Found !!"
            If lstInvoice.ListItems.COUNT < 1 Then Call lstInvoice_GotFocus
            Exit Sub
        End If
        
        rptsql = "{ORDMAN.COMP}='" & compPth & "'  AND {ORDMAN.ORDN} IN [" & M_CHLN & "] AND {ORDMAN.DBCD}='" & M_DBCD & "'"
        
        If cboFormat <> Empty Then
            ReportName = App.PATH & "\Reports\" & cboFormat
        Else
            ReportName = App.PATH & "\Reports\PERFORMAINVOICEGEN.RPT"
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
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .PageLast
            .PageFirst
            .ACTION = 1
            .PageZoom Val(txtZoom)
            Exit Sub
        End With
    Else
        MsgBox "No Format  !!", vbInformation, ""
        If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
        Close #1
    End If
        
    If Not BILLPRINTONLINE Then
        frmRPT_DosViewer.Show
        frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & "INV.TXT")
    Else
        LOAD frmRPT_DosViewer
        frmRPT_DosViewer.Hide
        frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & "INV.TXT")
        frmRPT_DosViewer.PrintSlip
    End If
    
End Sub

Private Sub dtFrom_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub dtTo_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If txtDVCD = Empty And ActiveControl.NAME = "txtDVCD" And KeyCode = 13 Then Exit Sub
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = 13 Then Exit Sub
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"

    End Sub

Private Sub Form_Load()
    Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom = Format(Now, "DD/MM/YYYY")
    dtTo = Format(Now, "DD/MM/YYYY")
    cboStatus.ListIndex = 1
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
    txtUNIT.Enabled = False
    cboDaybok.Clear
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT NAME FROM SALMANMST", CN, adOpenDynamic, adLockOptimistic
    Do While Not RS.EOF
     cboDaybok.AddItem RS!NAME & ""
     RS.MoveNext
    Loop
    If cboDaybok.ListCount > 0 Then
      cboDaybok.ListIndex = 0
    End If
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
    
    fPath = App.PATH
    If Right(fPath, 1) = "\" Then fPath = Left(fPath, Len(fPath) - 1)
    flInvoice.PATH = fPath & "\Reports"
    
    For I = 0 To flInvoice.ListCount - 1
        fName = UCase("PERFORMAINVOICE" & M_COMPBILL & ".rpt")
        If UNTEMBDRQ = "N" Then
            If fName = UCase(flInvoice.List(I)) Or UCase(flInvoice.List(I)) = UCase("PERFORMAINVOICEGEN.RPT") Then
                cboFormat.AddItem flInvoice.List(I)
            End If
        Else
            If fName = UCase(flInvoice.List(I)) Or UCase(flInvoice.List(I)) = UCase("PERFORMAINVOICEGEN.RPT") Then
                cboFormat.AddItem flInvoice.List(I)
            End If
        End If
    Next
    
    cboFormat.AddItem "PERFORMAINVOICEGEN.RPT"
    
End Sub
Public Sub GenInvList(VBNO As String)
Dim SQL As String
Static LSQL As String
Dim Item As ListItem

    If Not IsDate(dtFrom) Then
     dtFrom.SetFocus
     Exit Sub
    ElseIf Not IsDate(dtTo) Then
     dtTo.SetFocus
     Exit Sub
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT CODE FROM SALMANMST WHERE NAME='" & cboDaybok.Text & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
      M_DBCD = RS!CODE
     Else
      M_DBCD = Empty
    End If
                  
    SQL = "SELECT ORDN AS VBNO,ORDT AS DATE,ACCMST.NAME AS NAME,ISNULL(SUM(ORDMAN.AMNT),0) AS BNET FROM ORDMAN " & _
          "INNER JOIN ACCMST ON ACCMST.CODE = ORDMAN.PCOD " & _
          "Where ORDMAN.COMP='" & compPth & "' AND ORDMAN.UNIT='" & txtUNIT.Tag & _
          "' AND ORDMAN.DBCD ='" & M_DBCD & "' AND ORDMAN.ORDT>='" & Format(dtFrom, "MM/dd/yyyy") & _
          "' AND ORDMAN.ORDT<= '" & Format(dtTo, "MM/dd/yyyy") & "' GROUP BY ORDMAN.ORDN,ORDMAN.ORDT,ACCMST.NAME"
                     
    LSQL = SQL
    Set rsTemp = New Recordset
    rsTemp.Open LSQL, CN
    lstInvoice.ListItems.Clear
    Do While Not rsTemp.EOF
        Set Item = lstInvoice.ListItems.ADD
        Item.Text = rsTemp!VBNO & ""
        If VBNO <> Empty Then Item.Checked = True: M_VBNO = Item.Text
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

