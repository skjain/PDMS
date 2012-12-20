VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmRetBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box Return"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10275
   Begin FramePlusCtl.FramePlus Frm1 
      Height          =   6975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12303
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   8438015
      BackColor       =   16777215
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      Begin VB.TextBox TXTRMRK 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   13
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox TXTLOTNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TXTREQSLIP 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TXTMACHINE 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox TXTFROMDIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox TXTTODIV 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox TXTVBNO 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8520
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TXTCOST 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   5640
         Width           =   3615
      End
      Begin VB.TextBox TXTINAM 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox TXTPCOD 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1920
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Top             =   1080
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56688641
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   6120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "&Add"
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
         Image           =   "frmRetBox.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1560
         TabIndex        =   14
         Top             =   6120
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
         Image           =   "frmRetBox.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   6120
         Width           =   1095
         _ExtentX        =   1931
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
         Image           =   "frmRetBox.frx":1124
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   3960
         TabIndex        =   17
         Top             =   6120
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
         Image           =   "frmRetBox.frx":1576
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H CMDHELP 
         Height          =   375
         Left            =   7440
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Search"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
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
      Begin MSComctlLib.ListView lstBox 
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   10485760
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman Greek"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Box No"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Gross Wt."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tare Wt."
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Net Wt."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lot No."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "RATE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "GRNNO"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label LBLFIFO 
         BackStyle       =   0  'Transparent
         Caption         =   "Edit && Delete are not allowed. (FIFO Applied)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5280
         TabIndex        =   30
         Top             =   6240
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks  :"
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
         Left            =   5280
         TabIndex        =   29
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Slip No. :"
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
         Left            =   6960
         TabIndex        =   28
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "From Division "
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
         Left            =   240
         TabIndex        =   27
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Division    "
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
         Left            =   240
         TabIndex        =   26
         Top             =   1125
         Width           =   1380
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. Slip No. :"
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
         Left            =   6960
         TabIndex        =   25
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ret. Da&te :"
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
         Left            =   6960
         TabIndex        =   24
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   6735
         Left            =   120
         Top             =   120
         Width           =   10095
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Return to Store Division from  Another Division"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No."
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
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
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
         Left            =   6240
         TabIndex        =   20
         Top             =   1920
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name :"
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
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10200
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   10200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   10200
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00BEE7FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Head :"
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
         Left            =   240
         TabIndex        =   18
         Top             =   5640
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmRetBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public M_DBCD As String
Public M_DVCD As String
Public M_DVNM As String
Public M_SRNO As String
Dim SAVEFLAG As Boolean
Dim GRWT As String
Dim TRWT As String
Dim NTWT As String
Dim ROWNO As Long
Dim GRNNO As String
Dim Item As String
Dim BOXNO As String
Dim SWITCH As Boolean
'-------------------------------------------------------------------------------------------
' FORM EVENTS
'-------------------------------------------------------------------------------------------

Private Sub cmdCancel_Click()
  TXTFROMDIV.Tag = TXTFROMDIV
  ClsData (Me)
  TXTFROMDIV = TXTFROMDIV.Tag
 
  btn_sts (True)
  
  cmdAdd.SetFocus
  M_SRNO = Empty

  SWITCH = False
  TXTVBDT.Enabled = True
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub CMDITMDEL_Click()
End Sub

Private Sub CMDHELP_Click()
Call SearchBoxHelp
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST

 If CHKSAVEDATA = True Then
    Exit Sub
 End If
  
'Genrate Sr. No.
 If M_SRNO = Empty Then
    M_SRNO = pubGenSrNoSTR(TXTVBDT, "RTI")
 End If
    
 If SAVEFLAG = True Then
     TXTVBNO = GenVNO("RTI", M_DBCD)
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT VBNO FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='RTI' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
     If Not RS.EOF Then
        MsgBox "Duplicate Slip No. !!!! ", vbCritical
        Exit Sub
     End If
 End If
    
 Call SAVERTI
 
 If SAVEFLAG = True Then
    MsgBox "Your Issue Slip No. is " + TXTVBNO.Text
 End If
    Call cmdCancel_Click
    lstBox.ListItems.Clear
 Exit Sub
    
LAST:
    MsgBox ERR.Description
    If RS.State = 1 Then
        RS.CancelUpdate
    End If
    CN.RollbackTrans
    Exit Sub

End Sub

Private Sub Form_Activate()
' Call ColorComponent(Me)
' Me.BackColor = RGB(RED, GREEN, BLUE)
 btn_sts (True)
 'FIFO-------------------------------------
  If FIFOREQ = "Y" Then
     LBLFIFO.Visible = True
  End If
  '------------------------------------------
End Sub

Private Sub Form_Load()
 FIFOREQ = "Y"
 Call CenterChild(frm_Main, Me)
 Me.KeyPreview = True
 Me.Tag = zoomflag
 M_DBCD = "000001"
 If Not zoomflag = True Then
    M_SRNO = Empty
 End If
 M_DVCD = "000001"
 TXTVBDT = Now
 TXTVBDT.MaxDate = FEDT
 TXTVBDT.MinDate = FSDT

 TXTTODIV = GETDIVNAME("000001")
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = "TXTRMRK" And KeyAscii = vbKeyReturn Then cmdSave.SetFocus: Exit Sub
 If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
 End If
End Sub
'-------------------------------------------------------------------------------------------
' BUTTON EVENTS
'-------------------------------------------------------------------------------------------
Private Sub cmdAdd_Click()
    zoomflag = False
    btn_sts (False)
    M_SRNO = Empty
    TXTVBNO = GenVNO("RTI", M_DBCD)
    SAVEFLAG = True
    TXTFROMDIV.Enabled = True
    TXTFROMDIV.SetFocus
End Sub

Private Sub CMDOK_Click()
 Dim INDEX As Long
 
End Sub
'-------------------------------------------------------------------------------------------
' LOCAL PROCEDURE
'-------------------------------------------------------------------------------------------
Public Sub btn_sts(Yes As Boolean)
    cmdSave.Enabled = Not Yes
    cmdCancel.Enabled = Not Yes
    cmdAdd.Enabled = Yes
    TXTMACHINE.Enabled = Not Yes
    TXTVBDT.Enabled = Not Yes
    TXTREQSLIP.Enabled = Not Yes
    TXTRMRK.Enabled = Not Yes

End Sub
'-------------------------------------------------------------------------------------------

Private Sub ITMFLEX_Click()
   
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TXTCOST_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTCOST = Empty
    ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And TXTCOST = Empty) Then
        M_DESC = Empty
        NEW_VISIBLE = True
        Key = Empty
        TXTCOST = SearchList1("Select  TOP 20 Code,Name From REFMST WHERE CATA='N' AND NAME NOT LIKE '%DISABLE%'", 0, Empty, "Select COSTING HEAD FROM MASTER")
        TXTCOST.Tag = Key
        If key_PressNew = True Then
            M_DESC = ""
            Ref_Cat = "N"
            LOAD Frm_Ref_FAS
            Frm_Ref_FAS.Tag = Ref_Cat
            Frm_Ref_FAS.Show
        End If
    End If
    Me.KeyPreview = True
End Sub

'-------------------------------------------------------------------------------------------
' CODE FOR CURSOR POSITION ON MODULE
'-------------------------------------------------------------------------------------------

Private Sub TXTFROMDIV_GotFocus()
 TXTFROMDIV.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}":
End Sub

Private Sub TXTFROMDIV_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
          
    If KeyCode = vbKeyF2 Or (Trim(TXTFROMDIV) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTFROMDIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, TXTFROMDIV.Text, "SELECT DIVISION FROM LIST")
        TXTFROMDIV.Tag = Key
        M_DVNM = TXTFROMDIV
        M_DVCD = Key
    End If
        
    Me.KeyPreview = True
End Sub

Private Sub TXTFROMDIV_LostFocus()
 TXTFROMDIV.BackColor = vbWhite
End Sub


Private Sub TXTINAM_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False

If KeyCode = vbKeyF2 Or (Trim(TXTINAM) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        TXTINAM.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM ITMMST ", 0, Empty, "SELECT ITEM FROM LIST")
        TXTINAM.Tag = Key
If key_PressNew = True Then
          M_DESC = ""
          TXTINAM = Empty
          frm_Item.Show
          End If
End If
    Me.KeyPreview = True

End Sub


Private Sub TXTPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtpcod = Empty
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtpcod = Empty) Then
     M_DESC = Empty:   NEW_VISIBLE = False
     Key = Empty
     txtpcod = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
     txtpcod.Tag = Key
  End If
  
  Me.KeyPreview = True

End Sub

Private Sub TXTTODIV_GotFocus()
 TXTTODIV.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTTODIV_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
          
    If KeyCode = vbKeyF2 Or (Trim(TXTTODIV) = Empty And KeyCode = vbKeyReturn) Then
        NEW_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        TXTTODIV.Text = SearchList1("SELECT TOP 20 CODE, NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001'  AND RECSTAT='A'", 0, TXTTODIV.Text, "SELECT DIVISION FROM LIST")
        TXTTODIV.Tag = Key
        M_DVNM = TXTTODIV
        M_DVCD = Key
    End If
        
    Me.KeyPreview = True
End Sub

Private Sub TXTTODIV_LostFocus()
 TXTTODIV.BackColor = vbWhite
End Sub

Private Sub txtMachine_GotFocus()
 TXTMACHINE.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMACHINE_LostFocus()
 TXTMACHINE.BackColor = vbWhite
End Sub

Private Sub TXTREQSLIP_GotFocus()
 TXTREQSLIP.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTREQSLIP_LostFocus()
 TXTREQSLIP.BackColor = vbWhite
End Sub



Private Sub TXTRMRK_GotFocus()
 TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
 SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTRMRK_LostFocus()
 TXTRMRK.BackColor = vbWhite
End Sub

Private Sub txtMachine_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.KeyPreview = False
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTMACHINE = Empty
    ElseIf KeyCode = vbKeyF2 Then
        M_DESC = Empty
        NEW_VISIBLE = False
        TXTMACHINE = SearchList1("Select  TOP 20 Code,Name From MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & M_DVCD & "'", 0, Empty, "Select M/C FROM MASTER")
    End If
    Me.KeyPreview = True
End Sub

Private Sub TXTMACHINE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And TXTMACHINE = Empty Then
        Call txtMachine_KeyDown(vbKeyF2, 0)
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub CLEARDATA()
        
        TXTINAM = Empty
        TXTFROMDIV = Empty
        End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(TXTINAM) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        If TXTINAM.Enabled Then TXTINAM.SetFocus
        CheckData = True
        Exit Function
    End If
    

    
End Function


Private Function CHKSAVEDATA() As Boolean
If TXTFROMDIV = Empty Then
  MsgBox "Enter Source Division then Save"
  CHKSAVEDATA = True
  If TXTFROMDIV.Enabled Then TXTFROMDIV.SetFocus
  Exit Function
End If

If TXTTODIV = Empty Then
  MsgBox "Enter Destination Division then Save"
  CHKSAVEDATA = True
  If TXTTODIV.Enabled Then TXTTODIV.SetFocus
  Exit Function
End If

If TXTMACHINE = Empty Then
  MsgBox "Enter Machine Number then Save"
  CHKSAVEDATA = True
  If TXTMACHINE.Enabled Then TXTMACHINE.SetFocus
  Exit Function
End If

If TXTREQSLIP = Empty Then
  MsgBox "Enter Requision Slip Number !!!", vbInformation
  CHKSAVEDATA = True
  If TXTREQSLIP.Enabled Then TXTREQSLIP.SetFocus
  Exit Function
End If

If TXTINAM = Empty Then
MsgBox "Please Select Lot No. !!! ", vbInformation
CHKSAVEDATA = True
  If TXTINAM.Enabled Then TXTINAM.SetFocus
  Exit Function
 End If
 
 If TXTCOST = Empty Then
 MsgBox "Please Select Cost Head !!! ", vbInformation
 CHKSAVEDATA = True
  If TXTCOST.Enabled Then TXTCOST.SetFocus
  Exit Function
 End If
 
 

End Function

Private Sub SAVERTI()
  
  On Error GoTo LAST
  Dim SQL As String
  
  Dim SAVDAT As New ADODB.Recordset
  Dim MSTDAT As New ADODB.Recordset
  
  Set SAVDAT = New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
      
  
  CN.BeginTrans
  Call DELETERTI
  SQL = Empty
  
  If SAVDAT.State = 1 Then SAVDAT.Close
  SAVDAT.Open "SELECT * FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND VTYP='RTI' AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "'", CN, adOpenDynamic, adLockOptimistic
  
  Dim AI As String
  Dim BQ As Double
  Dim i As Long
  Dim DVCOD As String
  DVCOD = GetDivCode(TXTFROMDIV)
 
 '--------------------------------------------------------------------------------
  Dim FIFORATE As Double
  i = 0
   For i = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(i).Checked = True Then
      BOXNO = lstBox.ListItems(i).Text
      NTWT = lstBox.ListItems(i).SubItems(3)
      
    ''------------------------
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "RTI"
    SAVDAT!SRNO = M_SRNO
    SAVDAT!SRCH = i
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTREQSLIP
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CHEAD = TXTCOST.Tag
    SAVDAT!ICOD = TXTINAM.Tag: AI = TXTINAM.Tag
    SAVDAT!PCES = 0
    SAVDAT!QNTY = Val(NTWT): BQ = Val(NTWT)
    SAVDAT!RATE = FindFIFORate(GetCode("ITMMST", Trim(TXTINAM), "NAME", "CODE"))
    SAVDAT!AMNT = Val(NTWT) * FindFIFORate(GetCode("ITMMST", Trim(TXTINAM), "NAME", "CODE"))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    SAVDAT!OPER = "-"
     
    SAVDAT!PCOD = GetMachineCode(DVCOD, TXTMACHINE)
    SAVDAT!DVCD = DVCOD
    
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
        
    SAVDAT.AddNew
    SAVDAT!COMP = compPth
    SAVDAT!VTYP = "RTI"
    SAVDAT!SRNO = M_SRNO
    SAVDAT!SRCH = i + (lstBox.ListItems.COUNT)
    SAVDAT!VBNO = TXTVBNO.Text
    SAVDAT!chln = TXTREQSLIP
    SAVDAT!Date = Format(TXTVBDT, "YYYY/MM/DD")
    SAVDAT!dbcd = M_DBCD
    SAVDAT!CHEAD = TXTCOST.Tag
    SAVDAT!ICOD = TXTINAM.Tag: AI = TXTINAM.Tag
    SAVDAT!PCES = 0
    SAVDAT!QNTY = Val(NTWT): BQ = Val(NTWT)
    SAVDAT!RATE = FindFIFORate(GetCode("ITMMST", Trim(TXTINAM), "NAME", "CODE"))
    SAVDAT!AMNT = Val(NTWT) * FindFIFORate(GetCode("ITMMST", Trim(TXTINAM), "NAME", "CODE"))
    SAVDAT!QORP = "Q"
    SAVDAT![User] = cUName
    
    If SAVEFLAG = True Then
      SAVDAT!SYSR = "N"
     Else
      SAVDAT!SYSR = "U"
    End If
    
    SAVDAT!OPER = "+"
    SAVDAT!PCOD = GetMachineCode(DVCOD, TXTMACHINE)
    SAVDAT!DVCD = GetDivCode(TXTTODIV)
    SAVDAT!unit = UNCD
    SAVDAT!RECSTAT = "A"
    SAVDAT.Update
   End If
  Next
  
'-------------------------------------------------------
' UPDATION AND INSERTION IN TRDBOXREGISTER
i = 0
   For i = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(i).Checked = True Then
      BOXNO = lstBox.ListItems(i).Text
      GRWT = lstBox.ListItems(i).SubItems(1)
      TRWT = lstBox.ListItems(i).SubItems(2)
      NTWT = lstBox.ListItems(i).SubItems(3)
      GRNNO = lstBox.ListItems(i).SubItems(6)
 CN.Execute "UPDATE TRDBOXREGISTER SET RVTYP = 'RTI',RVBNO = '" & BOXNO & "',RVBDT = '" & Format(TXTVBDT, "YYYY/MM/DD") & "', RDBC = '" & M_DBCD & _
             "' WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' AND DVCD = '" & GetDivCode(TXTFROMDIV) & "' AND VBNO = '" & BOXNO & _
             "' AND GRNNO = '" & GRNNO & "' AND VTYP = 'ISS' AND OPER = '-'"
 '-----------------------------------------------------------------------------
 
 If SAVDAT.State = 1 Then SAVDAT.Close
 SAVDAT.Open "SELECT * FROM TRDBOXREGISTER WHERE COMP='" & compPth & "' AND  UNIT = '" & UNCD & "' AND DVCD = '" & DVCOD & "' AND  VTYP='RTI' AND GRNNO ='" & Trim(TXTVBNO) & "'", CN, adOpenDynamic, adLockOptimistic
 SAVDAT.AddNew
 
     SAVDAT!COMP = compPth
     SAVDAT!VTYP = "RTI"
     SAVDAT!GRNNO = Trim(TXTVBNO)
     SAVDAT!VBDT = Format(TXTVBDT, "YYYY/MM/DD")
     SAVDAT!dbcd = M_DBCD
     SAVDAT!ICOD = TXTINAM.Tag
     SAVDAT!PCOD = GetMachineCode(DVCOD, TXTMACHINE)
     SAVDAT!VBNO = Trim(BOXNO)
     SAVDAT!GRSWGT = Val(GRWT)
     SAVDAT!TRWGT = Val(TRWT)
     SAVDAT!NTWGT = Val(NTWT)
     SAVDAT!DVCD = GetDivCode(TXTTODIV)
     SAVDAT!unit = UNCD
     SAVDAT!RECSTAT = "A"
     SAVDAT!OPER = "+"
     SAVDAT!LOTNO = lstBox.ListItems(i).SubItems(4)
     SAVDAT!RATE = lstBox.ListItems(i).SubItems(5)
     SAVDAT.Update
     
 End If
 Next
 
'UPDATE VOUCHER TYPE MASTER
  If SAVEFLAG = True Then
    Call SetSRNO(TXTVBNO, "RTI", M_DBCD)
  End If
  
  '------------------------------------
  'DAILYENTRY Status
  ' Call DAILYSTATUS("ISS", GetMachineCode(DVCOD, TXTMACHINE), M_DBCD, Val(ITMFLEX.TextMatrix(1, 3)), TXTVBNO, Val(ITMFLEX.TextMatrix(1, 5)), cUName, "N", Now, TXTVBDT)
  '-------------------------------------
  'FIFO
    If SAVEFLAG = True And FIFOREQ = "Y" Then
       Call SetFIFOUP
    End If
  '----------------------
  
  CN.CommitTrans
  Exit Sub
LAST:
 MsgBox ERR.Description
 If SAVDAT.State = 1 Then
   SAVDAT.CancelUpdate
   SAVDAT.Close
 End If
 CN.RollbackTrans
End Sub

Private Sub UPDATESTATUS()
  Dim DLYSTA As New ADODB.Recordset
  If DLYSTA.State = adStateOpen Then DLYSTA.Close
  DLYSTA.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenKeyset, adLockPessimistic
  DLYSTA.AddNew
  DLYSTA!COMP = compPth & ""
  DLYSTA!VTYP = "RTI"
  DLYSTA!dbcd = M_DBCD
  DLYSTA!QNTY = 0
  DLYSTA!VBNO = TXTVBNO & ""
  DLYSTA!AMNT = 0
  DLYSTA!CUSR = cUName
  If SAVEFLAG = True Then
    DLYSTA!ACTN = "E"
   Else
    DLYSTA!ACTN = "M"
  End If
  DLYSTA!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS AMPM")
  DLYSTA.Update
  DLYSTA.Close
End Sub

Private Sub DELETERTI()
  CN.Execute "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='RTI' " & _
             " AND DBCD='" & M_DBCD & "' AND VBNO='" & TXTVBNO & "' "
End Sub

Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If TXTREQSLIP.Enabled Then TXTREQSLIP.SetFocus
End If
End Sub

Private Sub SearchBoxHelp()

Dim IND As Integer
Dim lstItem As ListItem
Dim icode As String
Dim dcode As String
Dim SQL As String
    
Screen.MousePointer = vbHourglass
lstBox.ListItems.Clear
Dim RECSET As New ADODB.Recordset
If RS.State = 1 Then RS.Close
SQL = Empty
If RECSET.State = 1 Then RECSET.Close
            
            
SQL = "Select TRDBOXREGISTER.VBNO,TRDBOXREGISTER.GRSWGT,TRDBOXREGISTER.TRWGT,TRDBOXREGISTER.NTWGT,TRDBOXREGISTER.LOTNO,TRDBOXREGISTER.RATE,TRDBOXREGISTER.GRNNO from TRDBOXREGISTER INNER JOIN  ITMMST ON " & _
        " ITMMST.CODE = TRDBOXREGISTER.ICOD  where TRDBOXREGISTER.COMP='" & compPth & "' AND TRDBOXREGISTER.UNIT='" & UNCD & _
        "' AND TRDBOXREGISTER.DVCD ='" & TXTFROMDIV.Tag & "' AND  RECSTAT<>'D'" & _
        " AND OPER = '-' AND VTYP <> 'SAL' AND ICOD = '" & TXTINAM.Tag & "' AND RVTYP IS NULL AND RVBNO IS NULL AND RDBC IS NULL"
            
            
   If RECSET.State = 1 Then RECSET.Close
   RECSET.Open SQL, CN, adOpenDynamic, adLockOptimistic
   
   If txtpcod <> Empty Then
      SQL = SQL & " AND TRDBOXREGISTER.PCOD = '" & txtpcod.Tag & "'"
   End If
   
   If TXTLOTNO <> Empty Then
      SQL = SQL & " TRDBOXREGISTER.LOTNO = '" & TXTLOTNO & "'"
   End If
   
    With RECSET
    If RECSET.EOF = True Then
                MsgBox "There are no Record found.", vbInformation, App.TITLE
    Else
            Do While Not RECSET.EOF
            Set lstItem = lstBox.ListItems.ADD
            lstItem.Text = Trim(RECSET![VBNO])
            lstItem.SubItems(1) = Trim(!GRSWGT)
            lstItem.SubItems(2) = Trim(!TRWGT)
            lstItem.SubItems(3) = Trim(!NTWGT)
            lstItem.SubItems(4) = Trim(!LOTNO & "")
            lstItem.SubItems(5) = Trim(!RATE & "")
            lstItem.SubItems(6) = Trim(!GRNNO)
            RECSET.MoveNext
        Loop
            
    End If
    End With
RECSET.Close
Screen.MousePointer = vbNormal
lstBox.SetFocus


End Sub

'FIFO----------------------
Private Sub SetFIFOUP()
On Error GoTo FIFOERR

'VARIABLE DECLARATION
Dim ICOD As String, Item As String, INDEX As Long
Dim BALQNTY As Double, TMPQTY As Double
Dim FIFORS As ADODB.Recordset: Set FIFORS = New ADODB.Recordset
Dim i As Long
'-------------------------------------------------------------
 i = 0
   For i = 1 To lstBox.ListItems.COUNT
    If lstBox.ListItems(i).Checked = True Then
      BOXNO = lstBox.ListItems(i).Text
      GRWT = lstBox.ListItems(i).SubItems(1)
      TRWT = lstBox.ListItems(i).SubItems(2)
      NTWT = lstBox.ListItems(i).SubItems(3)
      GRNNO = lstBox.ListItems(i).SubItems(6)
'INITIALISE
Item = TXTINAM.Text
BALQNTY = Val(NTWT)
'-------------------

If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & Item & "'", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then ICOD = Trim(FIFORS!CODE & "")
FIFORS.Close

'EITHER CASE :IF PENDING GRN EXIST
If FIFORS.State = 1 Then FIFORS.Close
FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not FIFORS.EOF Then
       If BALQNTY > 0 Then
          FIFORS!RET_DPT_QNTY = Val(FIFORS!RET_DPT_QNTY) + BALQNTY
          FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
          BALQNTY = 0
       End If
          FIFORS.Update
End If
'----------------------------------------------
'OR CASE : IF NO PENDING GRN EXIST
If BALQNTY > 0 Then
    If FIFORS.State = 1 Then FIFORS.Close
    FIFORS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD='" & ICOD & "' ORDER BY DATE,VBNO DESC", CN, adOpenDynamic, adLockOptimistic
    If Not FIFORS.EOF Then
       If BALQNTY > 0 Then
          FIFORS!RET_DPT_QNTY = Val(FIFORS!RET_DPT_QNTY) + BALQNTY
          FIFORS!BAL_QNTY = Val(FIFORS!GRN_QNTY) - Val(FIFORS!ISS_QNTY) - Val(FIFORS!RET_PTY_QNTY) + Val(FIFORS!RET_DPT_QNTY)
       End If
          FIFORS.Update
    End If
 End If
'-----------------------------
End If
Next i
Exit Sub
FIFOERR:
MsgBox ERR.Description
End Sub

Private Function FindFIFORate(ICOD As String) As Double
'DEFAULT
FindFIFORate = 0
'---------------

Dim F1RS As ADODB.Recordset
Set F1RS = New ADODB.Recordset

If F1RS.State = 1 Then F1RS.Close
F1RS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND ICOD='" & ICOD & "' AND BAL_QNTY > 0 ORDER BY DATE,VBNO", CN, adOpenDynamic, adLockOptimistic
If Not F1RS.EOF Then
   FindFIFORate = Val(F1RS!RATE)
   F1RS.Close
   Exit Function
Else 'SPECIAL CASE : : IF NO PENDING GRN EXIST

    If F1RS.State = 1 Then F1RS.Close
    F1RS.Open "SELECT * FROM GRNTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND ICOD='" & ICOD & "' ORDER BY DATE,VBNO DESC", CN, adOpenDynamic, adLockOptimistic
    If Not F1RS.EOF Then
       FindFIFORate = Val(F1RS!RATE)
       F1RS.Close
    End If
    
End If

End Function


