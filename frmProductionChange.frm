VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmProductionChange 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Packing Station Production Date Change"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6510
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6330
      Begin MSMask.MaskEdBox dtFrom 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
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
         Left            =   4320
         TabIndex        =   4
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
         Caption         =   "F&rom Date "
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
         TabIndex        =   6
         Top             =   195
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "T&o Date "
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
         Left            =   3360
         TabIndex        =   5
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.TextBox TXTPKGSTATION 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4860
   End
   Begin WelchButton.lvButtons_H cmdUpdate 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Update"
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
      Image           =   "frmProductionChange.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "&Exit"
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
      Image           =   "frmProductionChange.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "Pkg &Station"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "frmProductionChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub cmdupdate_Click()
On Error GoTo LAST

If TXTPKGSTATION = Empty Then
   TXTPKGSTATION.SetFocus
   Exit Sub
Else
   If RS.State = 1 Then RS.Close
   RS.Open "SELECT CODE FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
   "' AND NAME = '" & TXTPKGSTATION & "'", CN, adOpenDynamic, adLockOptimistic
   If Not RS.EOF Then
      TXTPKGSTATION.Tag = Trim(RS!CODE & "")
   Else
      TXTPKGSTATION.SetFocus
      Exit Sub
   End If
End If

'CHECKING COORECT DATA

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND RECSTAT<>'D' AND VBDT = '" & Format(dtFrom, "MM/DD/YYYY") & _
        "' AND PKG_STCOD = '" & TXTPKGSTATION.Tag & "'"
If RS.EOF Then
   MsgBox "Record Not Found", vbCritical
   Exit Sub
End If


If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND RECSTAT<>'D' AND VBDT > '" & Format(dtFrom, "MM/DD/YYYY") & _
        "' AND PKG_STCOD = '" & TXTPKGSTATION.Tag & "'"
If Not RS.EOF Then
   MsgBox "Further Date Entry Exist", vbCritical
   Exit Sub
End If

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND RECSTAT<>'D' AND VBDT > '" & Format(dtTo, "MM/DD/YYYY") & _
        "' AND VBDT < '" & Format(dtFrom, "MM/DD/YYYY") & _
        "' AND PKG_STCOD = '" & TXTPKGSTATION.Tag & "'"
If Not RS.EOF Then
   MsgBox "In Between Date Production Exist", vbCritical
   Exit Sub
End If

If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND RECSTAT<>'D' AND PKG_STCOD = '" & TXTPKGSTATION.Tag & "' ORDER BY VBDT DESC"
If Not RS.EOF Then
   If Format(RS!VBDT, "DD/MM/YYYY") <> Format(dtFrom, "DD/MM/YYYY") Then
      MsgBox "NOT A LAST Date PRODUCTION", vbCritical
      Exit Sub
   End If
End If

Dim AYS
AYS = MsgBox("Are You Sure To Change Production Date ? ", vbYesNo + vbQuestion, "Change ?")
If AYS = VBNO Then
   Me.Caption = "Change??"
   Exit Sub
End If


Dim L As Long

CN.BeginTrans

CN.Execute "UPDATE STORETRAN SET DATE='" & Format(dtTo, "MM/DD/YYYY") & "' WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND RECSTAT<>'D' AND DATE = '" & Format(dtFrom, "MM/DD/YYYY") & _
           "' AND SRNO = '" & TXTPKGSTATION.Tag & "' AND VTYP='PPF' ", L

CN.Execute "UPDATE BOXREGISTER SET VBDT='" & Format(dtTo, "MM/DD/YYYY") & "' WHERE COMP='" & compPth & _
           "' AND UNIT='" & UNCD & "' AND RECSTAT<>'D' AND VBDT = '" & Format(dtFrom, "MM/DD/YYYY") & _
           "' AND PKG_STCOD = '" & TXTPKGSTATION.Tag & "'", L
           
CN.Execute "UPDATE PCKMST SET LSTPCKDT = '" & Format(dtTo, "MM/DD/YYYY") & _
           "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & TXTPKGSTATION.Tag & "'", L
                      
CN.CommitTrans

MsgBox "success"

Me.Caption = "SUCCESS"

Exit Sub
LAST:
MsgBox ERR.Description
Resume
CN.RollbackTrans
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  dtFrom = Format(Now, "DD/MM/YYYY")
  dtTo = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   frm_Main.mnuMiscRepoOp1(12).Visible = False
End Sub

Private Sub TXTPKGSTATION_GotFocus()
 TXTPKGSTATION.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPKGSTATION_KeyDown(KeyCode As Integer, Shift As Integer)
    If TXTPKGSTATION.Text = Empty Or KeyCode = vbKeyF2 Then
        NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
        TXTPKGSTATION = SearchList1("SELECT CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A'", 0, TXTPKGSTATION, "SELECT PACKING STATION FROM LIST")
        TXTPKGSTATION.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTPKGSTATION = Empty
        TXTPKGSTATION.Tag = Empty
    End If
End Sub

Private Sub TXTPKGSTATION_LostFocus()
 TXTPKGSTATION.BackColor = vbWhite
End Sub



