VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_Anx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANNEXTURE 'IV'"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5880
   Begin VB.Frame Frame3 
      Height          =   795
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   5700
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
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3000
         TabIndex        =   8
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
         Image           =   "frmRPT_Anx.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4320
         TabIndex        =   9
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
         Image           =   "frmRPT_Anx.frx":0452
         cBack           =   -2147483633
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
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   11
      Top             =   1200
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
         Format          =   56688641
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
         Format          =   56688641
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
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   120
      TabIndex        =   10
      Top             =   480
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
   Begin VB.Label LBLHEAD 
      Caption         =   "Annexture 'IV' (To Be Maintained By Jobworker / Processor)"
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
      TabIndex        =   13
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmRPT_Anx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptsql As String

Private Sub Form_Activate()
If Me.Tag = "4" Then
       Me.Caption = "ANNEXTURE 'IV'"
       LBLHEAD = "Annexture 'IV' (To Be Maintained By Assessee)"
    ElseIf Me.Tag = "5" Then
       Me.Caption = "ANNEXTURE 'V'"
       LBLHEAD = "Annexture 'V' (To Be Maintained By Jobworker / Processor)"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl.NAME = "txtUNIT" And txtUNIT = Empty And KeyCode = vbKeyReturn Then Exit Sub
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
Call SetView

If Me.Tag = "4" Then
   ReportName = App.PATH & "\Reports\Annexture4.rpt"
   RPTN = "ANNEXTURE IV"
ElseIf Me.Tag = "5" Then
   ReportName = App.PATH & "\Reports\Annexture5.rpt"
   RPTN = "ANNEXTURE V"
End If
   
If Dir(ReportName, vbNormal) = Empty Then
   ReportErrorMessage 1001
   Exit Sub
End If
    
CRPT.ReportFileName = ReportName
CRPT.ReplaceSelectionFormula rptsql

'CRPT.SubreportToChange = ""
'CRPT.SubreportToChange = "Rec_Anx5.rpt"
'CRPT.Connect = "DSN=" & ServerName & ";UID=sa;PWD= " & DefaultPassword_live & ";DSQ=" & CN.DefaultDatabase
       
'CRPT.SubreportToChange = ""
'CRPT.SubreportToChange = "Iss_Anx5.rpt"
'CRPT.Connect = "DSN=" & ServerName & ";UID=sa;PWD= " & DefaultPassword_live & ";DSQ=" & CN.DefaultDatabase
    
    With CRPT
        .DiscardSavedData = True
        .Formulas(1) = "COMPANY='" & compNm & "'"
        .Formulas(2) = "REPORTHEAD='" & RPTN & "'"
        .Formulas(3) = "PERIOD='" & PERIOD & "'"
        .Formulas(4) = "UNIT='" & txtUNIT.Text & "'"
         
         RPTN = RPTN + Space(5) + ReportName
                 

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
        txtUNIT.SetFocus
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

Private Sub SetView()
On Error Resume Next
Dim SQL As String


If Me.Tag = "4" Then
CN.Execute "DROP VIEW ANNEXTURE4"

SQL = "CREATE VIEW ANNEXTURE4 AS " & _
      "SELECT JOBOUT.COMP,JOBOUT.UNIT,JOBOUT.VBNO AS GRN,JOBOUT.VTYP,JOBOUT.VBNO,JOBOUT.DATE,ACCMST.NAME AS PARTY," & _
      "ITMMST.NAME AS ITEM,0 AS GRNQTY,QNTY AS QNTY,0 AS WASTAGE,JOBOUT.RMRK AS BRMK FROM JOBOUT " & _
      "INNER JOIN ACCMST ON ACCMST.CODE=JOBOUT.PCOD INNER JOIN ITMMST ON ITMMST.CODE=JOBOUT.ICOD " & _
      "WHERE JOBOUT.COMP='" & compPth & "' AND JOBOUT.UNIT='" & UNCD & "' AND JOBOUT.VTYP='ANX' AND " & _
      "JOBOUT.RECSTAT<>'D' AND JOBOUT.DATE>='" & Format(dtFrom.Value, "MM/DD/YYYY") & _
      "' AND JOBOUT.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
      "' Union " & _
      "SELECT JOBOUT.COMP,JOBOUT.UNIT,JOBOUT.RECNO AS GRN,JOBOUT.VTYP,JOBOUT.VBNO,JOBOUT.DATE,ACCMST.NAME AS PARTY," & _
      "ITMMST.NAME AS ITEM,0 AS GRNQTY,QNTY,0 AS WASTAGE,'' AS BRMK FROM JOBOUT " & _
      "INNER JOIN ACCMST ON ACCMST.CODE=JOBOUT.PCOD " & _
      "INNER JOIN ITMMST ON ITMMST.CODE=JOBOUT.ICOD " & _
      "WHERE JOBOUT.COMP='" & compPth & "' AND JOBOUT.UNIT='" & UNCD & _
      "' AND JOBOUT.DBCD='000003' AND JOBOUT.VTYP='IVR' AND JOBOUT.RECSTAT<>'D' AND " & _
      "JOBOUT.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & "'"

CN.Execute SQL

ElseIf Me.Tag = "5" Then

CN.Execute "DROP VIEW ANX5"

SQL = "CREATE VIEW ANX5 AS " & _
      "SELECT JOBIN.COMP,JOBIN.UNIT,JOBIN.VBNO AS GRN,JOBIN.VTYP,JOBIN.CHLN,JOBIN.CHDT,ACCMST.NAME AS PARTY," & _
      "ITMMST.NAME AS ITEM,QNTY AS GRNQTY,QNTY AS QNTY,JOBGRN.BRMK FROM JOBIN " & _
      "INNER JOIN ACCMST ON ACCMST.CODE=JOBIN.PCOD " & _
      "INNER JOIN ITMMST ON ITMMST.CODE=JOBIN.ICOD " & _
      "INNER JOIN JOBGRN ON JOBGRN.COMP=JOBIN.COMP AND JOBGRN.UNIT=JOBIN.UNIT AND JOBGRN.VTYP=JOBIN.VTYP " & _
      "AND JOBGRN.DBCD=JOBIN.DBCD AND JOBGRN.VBNO=JOBIN.VBNO WHERE JOBIN.VTYP='IVR' AND JOBIN.RECSTAT<>'D' AND " & _
      "JOBIN.DATE>='" & Format(dtFrom.Value, "MM/DD/YYYY") & _
      "' AND JOBIN.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & _
      "' Union " & _
      "SELECT JOBIN.COMP,JOBIN.UNIT,JOBIN.GRNNO AS GRN,JOBIN.VTYP,JOBIN.VBNO,JOBIN.DATE,ACCMST.NAME AS PARTY," & _
      "FINITMMST.NAME AS ITEM,JOBGRN.TQTY AS GRNQTY,QNTY,'' AS BRMK FROM JOBIN " & _
      "INNER JOIN ACCMST ON ACCMST.CODE=JOBIN.PCOD " & _
      "LEFT JOIN FINITMMST ON FINITMMST.CODE=JOBIN.ICOD AND FINITMMST.COMP=JOBIN.COMP AND FINITMMST.UNIT=JOBIN.UNIT AND FINITMMST.DVCD=JOBIN.DVCD " & _
      "INNER JOIN JOBGRN ON JOBGRN.COMP=JOBIN.COMP AND JOBGRN.UNIT=JOBIN.UNIT AND JOBGRN.VTYP='IVR' " & _
      "AND JOBGRN.VBNO=JOBIN.GRNNO WHERE JOBIN.VTYP='DPF' AND JOBIN.RECSTAT<>'D' " & _
      "AND JOBIN.DATE<='" & Format(dtTo.Value, "MM/DD/YYYY") & "'"

CN.Execute SQL

End If
End Sub
