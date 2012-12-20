VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmRPT_GRPackPrinting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GR Packing Printing"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   135
      TabIndex        =   12
      Top             =   660
      Width           =   6375
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
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
         CurrentDate     =   39447
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
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
         CurrentDate     =   39447
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "&From Date "
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
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4170
      Left            =   120
      TabIndex        =   11
      Top             =   1380
      Width           =   8220
      Begin MSComctlLib.ListView lstGR 
         Height          =   3840
         Left            =   120
         TabIndex        =   4
         Top             =   195
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6773
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
            Text            =   "GR No"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Party Name"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   600
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame4 
      Height          =   900
      Left            =   120
      TabIndex        =   10
      Top             =   5580
      Width           =   8220
      Begin VB.OptionButton OPTPLN 
         Caption         =   "Plain GRN"
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
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton OPTPRNT 
         Caption         =   "Pre-Printed "
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
         Left            =   4680
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   360
         TabIndex        =   5
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
         Image           =   "frmRPT_GRPackPrinting.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   1680
         TabIndex        =   6
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
         Image           =   "frmRPT_GRPackPrinting.frx":0452
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame7 
      Height          =   660
      Left            =   135
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      Begin VB.TextBox txtUNIT 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   4965
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmRPT_GRPackPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SEL_VBNO As String
Dim SQL As String

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    If OPTPRNT.Value = True Then
        Call Pre_PrintGRN
        Exit Sub
    End If
    
    Dim M_VBNO As String

    CRPT.Reset
    crptConnect CRPT
    
    ReportName = App.PATH & "\Reports\GR_GEN.rpt"
        
    M_VBNO = Empty
    Dim i As Double
    For i = 1 To lstGR.ListItems.COUNT
        If lstGR.ListItems(i).Checked Then
            If M_VBNO <> "" Then M_VBNO = M_VBNO & ","
            M_VBNO = M_VBNO & "'" & lstGR.ListItems(i) & "'"
        End If
    Next
    
    If M_VBNO = Empty Then
        MsgBox "Please Select Valid GR Packing Entry !!", vbInformation, "Select GR Packing From List!!"
        lstGR.SetFocus
        Exit Sub
    End If
    
    If Dir(ReportName, vbNormal) = Empty Then
        MsgBox "Report File is Missing From Report Folder !!", vbInformation, "Check Report File & " & ReportName
        Exit Sub
    End If
    
    SQL = "{GRPACKING.COMP}='" & compPth & "' AND {GRPACKING.UNIT}='" & txtUNIT.Tag & "' AND {GRPACKING.VBNO} IN [" & M_VBNO & "] AND {GRPACKING.RECSTAT}<>'D'"
    CRPT.ReportFileName = ReportName
    CRPT.ReplaceSelectionFormula SQL
    
    With CRPT
        RPTN = RPTN + Space(5) + ReportName
        If M_CUNT = "Y" Then
            CRPT.Formulas(1) = "UNCMP='" & txtUNIT & "'"
        Else
            CRPT.Formulas(1) = "UNCMP='" & compNm & "'"
        End If
        .DiscardSavedData = True
        .WindowTitle = RPTN
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowProgressCtls = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
         txtUNIT.SetFocus
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .PageLast
        .PageFirst
        .ACTION = 1
        .PageZoom Val(100)
    End With
End Sub

Private Sub Form_Activate()
   Call ColorComponent(Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = vbKeyReturn Then Exit Sub
        
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
    Call CenterChild(frm_Main, Me)
    dtFrom = GetMinDate
    dtTo = GetMaxDate
    
    txtUNIT = UntNm
    txtUNIT.Tag = UNCD
End Sub

Private Sub lstGR_GotFocus()
 lstGR.BackColor = RGB(BRED, BGREEN, BBLUE)
On Error GoTo errGOTFocus
    Dim Item As ListItem

    If txtUNIT = Empty Then
        MsgBox "Please Select Unit From List !!", vbInformation, "Unit Missing!!"
        txtUNIT.SetFocus
        Exit Sub
    End If
                     
    SQL = "Select GRPACKING.VBNO,GRPACKING.VBDT,SUM(ISNULL(GRPACKING.NETWGT,0)) AS NTWGT,ACCMST.NAME AS PARTY From GRPACKING INNER JOIN ACCMST ON ACCMST.CODE=GRPACKING.PCOD Where GRPACKING.COMP='" & compPth & "' AND GRPACKING.UNIT='" & txtUNIT.Tag & "'"
    SQL = SQL & " AND GRPACKING.VBDT>='" & Format(dtFrom, "MM/dd/yyyy") & "' AND GRPACKING.VBDT<='" & Format(dtTo, "MM/dd/yyyy") & "'"
    SQL = SQL & " GROUP BY GRPACKING.VBNO,GRPACKING.VBDT,ACCMST.NAME "
    
    Set rsTemp = New Recordset
    rsTemp.Open SQL, CN
    
    lstGR.ListItems.Clear
    
    Do While rsTemp.EOF = False
        Set Item = lstGR.ListItems.ADD
        Item.Text = rsTemp!VBNO
        Item.SubItems(1) = rsTemp!VBDT
        Item.SubItems(2) = rsTemp!PARTY
        Item.SubItems(3) = rsTemp!NTWGT
        rsTemp.MoveNext
    Loop
    
    rsTemp.Close
    
    Exit Sub
    
errGOTFocus:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub lstGR_LostFocus()
lstGR.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or Trim(txtUNIT) = Empty Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("SELECT TOP 20 Code,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtUNIT = Empty
    End If
End Sub

Private Sub Pre_PrintGRN()

    Dim i As Double
    i = 1
  
    If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
    Close #1
  
    Open "C:\DOSPRINT\" & ComputerName & ".TXT" For Output As #1
  
    For i = 1 To lstGR.ListItems.COUNT
        If lstGR.ListItems(i).Checked Then
            SEL_VBNO = lstGR.ListItems(i)
            Select Case M_COMPBILL
                Case "CMC"
                    
                Case "GSL"
                    
                Case Else
                    Print #1, "Format Not Exist"
            End Select
        End If
    Next
    
    Close #1
  
    LOAD frmRPT_DosViewer
    frmRPT_DosViewer.Hide
    frmRPT_DosViewer.LoadDocument ("C:\DOSPRINT\" & ComputerName & ".TXT")
    frmRPT_DosViewer.Show
End Sub

Private Sub txtUNIT_LostFocus()
txtUNIT.BackColor = vbWhite
End Sub

