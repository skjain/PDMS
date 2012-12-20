VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRPT_DailyEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Entry Status Report"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6780
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   6615
      Begin VB.CheckBox chkAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkEdt 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chkDel 
         Caption         =   "Deleted"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin Crystal.CrystalReport crpt 
         Left            =   1800
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         DiscardSavedData=   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   6615
      Begin VB.OptionButton OPTSYS 
         Caption         =   "System Date"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton OPTPST 
         Caption         =   "Posting Date"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "User Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   6615
      Begin VB.OptionButton USROPTALL 
         Alignment       =   1  'Right Justify
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
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton USROPTPAR 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular"
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
         Left            =   90
         TabIndex        =   8
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox USRCMB 
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
         ItemData        =   "frmRpt_DailyEntry.frx":0000
         Left            =   1440
         List            =   "frmRpt_DailyEntry.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   510
         Width           =   4875
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Company Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox CMBCOMP 
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
         ItemData        =   "frmRpt_DailyEntry.frx":0004
         Left            =   1440
         List            =   "frmRpt_DailyEntry.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   4875
      End
      Begin VB.OptionButton OPTCOMPSELC 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton OPTCOMPALL 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   6615
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
         TabIndex        =   19
         Text            =   "100"
         Top             =   240
         Width           =   735
      End
      Begin WelchButton.lvButtons_H cmdpreview 
         Height          =   495
         Left            =   3720
         TabIndex        =   14
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
         Image           =   "frmRpt_DailyEntry.frx":0008
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frmRpt_DailyEntry.frx":045A
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
         TabIndex        =   20
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.Frame Frame10 
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   6615
      Begin MSComCtl2.DTPicker perd2 
         Height          =   330
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   52625409
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker perd1 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   52625409
         CurrentDate     =   38429
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
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
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "From Date"
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
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   285
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmRPT_DailyEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim M_STAT As String
Dim M_COMPPATH As String
Dim RPTN As String
Dim PERIOD As String
Dim TRSY As String
    CN.Execute "UPDATE DAILYSTAT SET TRANDTTM =DTTM WHERE TRANDTTM IS NULL"
    crpt.Reset
    crptConnect crpt
    perd1.SetFocus
    rptsql = Empty
    
    RPTN = "Daily Entry Status Report For Record : Added /Modifed / Deleted"
    
    PERIOD = CStr(perd1.Value) & " To " & CStr(perd2.Value)
    
    M_COMPPATH = getCompPath(CMBCOMP.Text)
    
    M_STAT = ""
    
    If OPTPST.Value = True Then
      TRSY = "1"
     Else
      TRSY = "0"
    End If
    
    If chkAdd.Value = 1 Then M_STAT = "'N','E'"
    
    If chkEdt.Value = 1 And M_STAT = "" Then
       M_STAT = "'M'"
      Else
        If chkEdt.Value = 1 Then
          If M_STAT <> Empty Then M_STAT = M_STAT & ","
          M_STAT = M_STAT & "'M'"
        End If
    End If
    
    If chkDel.Value = 1 And M_STAT = Empty Then
        M_STAT = "'D'"
       Else
        If chkDel.Value = 1 Then
          If M_STAT <> Empty Then M_STAT = M_STAT & ","
          M_STAT = M_STAT & "'D'"
        End If
    End If
    
    
    If M_STAT = "" Then
      MsgBox "Please Select Some Value", vbCritical
      Exit Sub
    End If
    
    If OPTPST.Value = True Then
      M_STAT = M_STAT & "," + "'T','N','F','0','G','1'"
    End If
    
    crpt.ReportFileName = App.PATH & "\Reports\Daily Entry Status.rpt"

   
      
    If USROPTALL.Value = True Then
        rptsql = "{DAILYSTAT.COMP}='" & compPth & "' AND {DAILYSTAT.ACTN} IN [" & M_STAT & "] AND {DAILYSTAT.TRANDTTM}>=DATE(" & perd1.Year & "," & perd1.Month & "," & perd1.Day & ") AND {DAILYSTAT.TRANDTTM}<=DATE(" & perd2.Year & "," & perd2.Month & "," & perd2.Day & ")"
       Else
        rptsql = "{DAILYSTAT.COMP}='" & compPth & "' AND {DAILYSTAT.ACTN} IN [" & M_STAT & "] AND {DAILYSTAT.TRANDTTM}>=DATE(" & perd1.Year & "," & perd1.Month & "," & perd1.Day & ") AND {DAILYSTAT.TRANDTTM}<=DATE(" & perd2.Year & "," & perd2.Month & "," & perd2.Day & ") AND {DAILYSTAT.CUSR}='" & USRCMB.Text & "'"
    End If
    
    
    
    
    crpt.ReplaceSelectionFormula rptsql
   
    crpt.DiscardSavedData = True
 
    crpt.Formulas(1) = "RPTN='" & RPTN & "'"
    crpt.Formulas(2) = "PERIOD='" & PERIOD & "'"
    crpt.Formulas(3) = "TRSY='" & TRSY & "'"
    crpt.WindowTitle = "Daily Entry Status Report"
    RPTN = RPTN + Space(5) + ReportName
    crpt.Destination = crptToWindow
    crpt.WindowState = crptMaximized
    crpt.WindowShowProgressCtls = True
    crpt.WindowShowPrintBtn = True
    crpt.WindowShowPrintSetupBtn = True
    crpt.WindowShowRefreshBtn = True
    crpt.WindowShowSearchBtn = True
    crpt.ACTION = 1
    
    
End Sub

Private Sub CMBCOMP_GotFocus()
CMBCOMP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub CMBCOMP_LostFocus()
CMBCOMP.BackColor = vbWhite
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call ColorComponent(Me)
Me.BackColor = RGB(RED, GREEN, BLUE)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call CenterChild(frm_Main, Me)
    Call FillCmb("Select COMP_NAME From COMPMAST", CMBCOMP)
    Call FillCmb("Select UID From USERMAST", USRCMB)
    perd1.Value = Now
    perd2.Value = Now
  Exit Sub

errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
    
End Sub

Private Sub OPTCOMPALL_Click()
    CMBCOMP.ListIndex = -1
End Sub

Private Sub OPTCOMPSELC_Click()
    If CMBCOMP.ListCount > 0 Then CMBCOMP.ListIndex = 0
End Sub

Private Sub OPTPST_Click()
   If OPTPST.Value = True Then
     USROPTALL.Value = True
     USROPTPAR.Value = False
     
   End If
End Sub

Private Sub USRCMB_GotFocus()
USRCMB.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub USRCMB_LostFocus()
USRCMB.BackColor = vbWhite
End Sub

Private Sub USROPTPAR_Click()
   If USRCMB.ListCount > 0 Then USRCMB.ListIndex = 0
End Sub

