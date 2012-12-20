VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmMergeMst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge Master"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7140
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   6975
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frmMergeMst.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Save"
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
         Image           =   "frmMergeMst.frx":059A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdprint 
         Height          =   495
         Left            =   3720
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
         Image           =   "frmMergeMst.frx":0B34
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmMergeMst.frx":0F86
         cBack           =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.TextBox TXTBASE 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtitm 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtmrg 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin FramePlusCtl.FramePlus frameActive 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         HighlightColor  =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         Begin VB.OptionButton optDeactive 
            Caption         =   "Deactive"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1200
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin Crystal.CrystalReport CRPT 
         Left            =   6480
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label3 
         Caption         =   "Raw Material"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Base Party"
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Merge No."
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMergeMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim SQL As String

Private Sub cmdDelete_Click()
    If txtmrg = Empty Then Exit Sub
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND LTNO='" & txtmrg & "' AND RECSTAT<>'D'"
    If Not RS.EOF Then
       MsgBox "Further Entry Exist !! Can't Delete ", vbCritical, "Alert"
       txtmrg.Enabled = True
       txtmrg.SetFocus
       Exit Sub
    End If
    RS.Close
    
    If MsgBox("Are You Sure ? Want To Delete Merge ?", vbQuestion + vbYesNo, "Delete Merge") = vbYes Then
       CN.BeginTrans
       CN.Execute "DELETE FROM MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                  "' AND MRGN='" & txtmrg & "'"
                  
       CN.Execute "INSERT INTO DAILYSTAT (COMP,VTYP,SRNO,PCOD,DBCD,VBNO,QNTY,AMNT,CUSR,DTTM,ACTN) VALUES('" & compPth & "','MRG','XXXXXXXXXXXXX','" & txtmrg & "',NULL,'',0,0,'" & cUName & "','" & Format(Now, "MM/dd/yyyy HH:MM:SS AMPM") & "','D')"
       CN.CommitTrans
       MsgBox "SUCCESS FULLY DELETED ", vbInformation
    End If
    
    txtmrg = Empty
    TXTBASE = Empty
    txtitm = Empty
    txtmrg.Enabled = True
    txtmrg.SetFocus
    
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub cmdprint_Click()
On Error GoTo errPrintReport
Dim M_SelCode  As String
Dim i As Long
    
    CRPT.Reset
    crptConnect CRPT
    
    ReportName = App.PATH & "\Reports\rpt_MERGELST.rpt"
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    rptsql = Empty
    
    rptsql = "{MRGMST.RECSTAT} <> 'D'"
    CRPT.ReportFileName = ReportName
    CRPT.DiscardSavedData = True
    CRPT.ReplaceSelectionFormula rptsql
    CRPT.WindowState = crptMaximized
    
    CRPT.WindowTitle = "Merge Master Report" & Space(5) & "Report : " & ReportName
    
    CRPT.WindowShowPrintBtn = True
    CRPT.WindowShowPrintSetupBtn = True
    CRPT.WindowShowSearchBtn = True
    CRPT.WindowShowExportBtn = True
    CRPT.WindowShowRefreshBtn = True
    
    CRPT.ACTION = 1
    
    Exit Sub
    
errPrintReport:
    
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description & vbCrLf & " Error In Report " & ReportName
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(ActiveControl.NAME) = "TXTMRG" Then
      If txtmrg = Empty Then Exit Sub
   End If
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
End Sub

Private Sub TXTBASE_GotFocus()
  TXTBASE.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTBASE_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(TXTBASE.Text) = Empty Then
    Key = Empty
    TXTBASE.Text = SearchList1("SELECT TOP 20 Code,NAME FROM ACCMST", 0, TXTBASE.Text, "SELECT BASE PARTY FROM LIST")
    TXTBASE.Tag = Key
  End If
End Sub

Private Sub TXTBASE_LostFocus()
  TXTBASE.BackColor = vbWhite
End Sub

Private Sub TXTITM_GotFocus()
   txtitm.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTITM_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(txtitm.Text) = Empty Then
  Key = Empty
    txtitm.Text = SearchList1("SELECT TOP 20 CODE,NAME FROM VWITEM", 0, txtitm.Text, "SELECT RAW ITEM FROM LIST")
    txtitm.Tag = Key
  End If
End Sub

Private Sub TXTITM_LostFocus()
txtitm.BackColor = vbWhite
End Sub

Private Sub txtmrg_GotFocus()
   txtmrg.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtmrg_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Or (txtmrg = Empty And KeyCode = 13) Then
      SQL = "SELECT DISTINCT MRGN,MRGN FROM MRGMST WHERE COMP='" & compPth & _
          "' AND UNIT='" & UNCD & "' "
     txtmrg = SearchList1(SQL, 0, txtmrg, "SELECT MERGE NO. FROM LIST")
  End If
End Sub

Private Sub txtmrg_LostFocus()
  txtmrg.BackColor = vbWhite
  
  If Trim(txtmrg) = Empty Then Exit Sub
  
  Dim ITMCOD As String
  Dim PTYCOD As String
  
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "select * from mrgmst where comp='" & compPth & "' and unit='" & UNCD & _
          "' and mrgn='" & txtmrg & "' ", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Merge No. Does Not Exist"
   ' txtmrg.SetFocus
   ' Exit Sub
  Else
   ITMCOD = RS!ICOD
   PTYCOD = RS!PCOD & ""
   optActive.Value = IIf(Trim(RS!ACTIVE & "") = "Y", True, False)
   optDeactive.Value = IIf(Trim(RS!ACTIVE & "") = "N", True, False)
  End If
  
  'ITMCOD = RS!ICOD
  'PTYCOD = RS!PCOD & ""
  'optActive.Value = IIf(Trim(RS!ACTIVE & "") = "Y", True, False)
  'optDeactive.Value = IIf(Trim(RS!ACTIVE & "") = "N", True, False)
  
  If RS.State = 1 Then RS.Close
  RS.Open "select * from itmmst where code='" & ITMCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    txtitm.Text = Empty
  Else
    txtitm.Text = RS!NAME & ""
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE CODE='" & PTYCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    TXTBASE.Text = Empty
  Else
    TXTBASE.Text = RS!NAME & ""
  End If
  
End Sub

Private Sub cmdSave_Click()
  On Error GoTo LAST
  Dim ITMCOD As String
  
  Set RS = New ADODB.Recordset
  If RS.State = 1 Then RS.Close
  RS.Open "select * from mrgmst where mrgn='" & txtmrg & "' and comp='" & compPth & "' and unit='" & UNCD & "' ", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
   ' MsgBox "Merge No. Does Not Exist"
   ' txtmrg.SetFocus
   ' Exit Sub
  End If

  If RS.State = 1 Then RS.Close
  RS.Open "select * from itmmst where name='" & txtitm & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Missing in Master"
    txtmrg.SetFocus
    Exit Sub
  End If
  ITMCOD = RS!CODE
    
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM ACCMST WHERE NAME='" & TXTBASE.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Base Party is missing"
    TXTBASE.SetFocus
    Exit Sub
  End If
  TXTBASE.Tag = RS!CODE
  
  CN.Execute "delete from mrgmst where comp='" & compPth & "' and unit='" & UNCD & "' and mrgn='" & txtmrg & "'"
  If RS.State = 1 Then RS.Close
  RS.Open "select * from mrgmst where 1=2", CN, adOpenDynamic, adLockOptimistic
  RS.AddNew
    RS!COMP = compPth
    RS!unit = UNCD
    RS!MRGN = txtmrg
    RS!PCOD = TXTBASE.Tag
    RS!ICOD = ITMCOD
    RS!ACTIVE = IIf(optActive.Value = True, "Y", "N")
    RS!RECSTAT = "A"
  RS.Update
    
  Call ClsData(frmMergeMst)
  txtmrg.SetFocus
  Exit Sub
LAST:
  MsgBox ERR.Description
  If RS.State = 1 Then RS.CancelUpdate
  Exit Sub
End Sub



