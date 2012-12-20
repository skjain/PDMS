VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form Frm_positionmaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position Master"
   ClientHeight    =   3135
   ClientLeft      =   3570
   ClientTop       =   3225
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7440
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   3255
      Left            =   -120
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5741
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
      Begin VB.TextBox txtend 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1920
         Width           =   570
      End
      Begin VB.TextBox txtdvcd 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   5490
      End
      Begin VB.TextBox txtpostioncode 
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   1
         Top             =   480
         Width           =   450
      End
      Begin VB.TextBox txtmac 
         Appearance      =   0  'Flat
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   5490
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   4920
         TabIndex        =   8
         Top             =   2520
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
         Image           =   "Frm_positionmaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6240
         TabIndex        =   9
         Top             =   2520
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
         Image           =   "Frm_positionmaster.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Ends"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lbldvcd 
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Position No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label LBLHEADING1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Position Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   120
         Width           =   1935
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   5160
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frm_positionmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT CODE FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND NAME='" & txtdvcd & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Division "
    txtdvcd.SetFocus
    Exit Sub
  End If
  txtdvcd.Tag = RS!CODE
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "'  AND NAME='" & txtmac.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid M/c Reference "
    txtmac.SetFocus
    Exit Sub
  End If
  txtmac.Tag = RS!CODE
  
  If Val(txtend) = 0 Then
    MsgBox "Invalid Ends"
    txtend.SetFocus
    Exit Sub
  End If
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM POSITIONMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & txtpostioncode & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!unit = UNCD
  RS!DVCD = txtdvcd.Tag
  RS!CODE = Val(txtpostioncode)
  RS!MCCD = txtmac.Tag
  RS!EndS = Val(txtend)
  RS.Update
  MsgBox "Data Save Succefuly"
  Call ClsData(Frm_positionmaster)
  txtpostioncode.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Frm_positionmaster.KeyPreview = True
End Sub
Private Sub txtdvcd_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(txtdvcd) = Empty Then
    txtdvcd = SearchList1("SELECT CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, txtdvcd, "SEELCT DIVISION FROM LIST")
    txtdvcd.Tag = Key
  End If
End Sub

Private Sub txtend_Validate(Cancel As Boolean)
  If Val(txtend) = 0 Then
    MsgBox "Numeric Value is Allowed"
    Cancel = True
  End If
End Sub

Private Sub txtmac_KeyDown(KeyCode As Integer, Shift As Integer)
  If RS.State = 1 Then RS.Close
  RS.Open "select code from divmst where comp='" & compPth & "' and unit='" & UNCD & "' and name='" & txtdvcd.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtdvcd.Tag = RS!CODE
   Else
    txtdvcd.Tag = Empty
  End If
  RS.Close
  If KeyCode = vbKeyF2 Or Trim(txtmac) = Empty Then
    txtmac = SearchList1("SELECT CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "'", 0, txtmac, "SEELCT M/c FROM LIST")
    txtmac.Tag = Key
  End If
End Sub

Private Sub txtpostioncode_LostFocus()
  If txtdvcd.Enabled = True Then
    txtdvcd.SetFocus
  End If
End Sub

Private Sub txtpostioncode_Validate(Cancel As Boolean)
  If txtpostioncode = Empty Then Exit Sub

  If Val(txtpostioncode) = 0 Then
     MsgBox "Postion should be numeric"
     Cancel = True
     Exit Sub
  End If
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT POSITIONMASTER.DVCD,MCCD,DIVMST.NAME AS DVNM, MACMST.NAME AS MCNM,ENDS FROM POSITIONMASTER " & _
          "INNER JOIN DIVMST ON DIVMST.COMP=POSITIONMASTER.COMP AND DIVMST.UNIT=POSITIONMASTER.UNIT AND " & _
          "DIVMST.CODE=POSITIONMASTER.DVCD " & _
          "INNER JOIN MACMST ON MACMST.COMP=POSITIONMASTER.COMP AND POSITIONMASTER.UNIT=MACMST.UNIT AND " & _
          "MACMST.DVCD=POSITIONMASTER.DVCD AND MACMST.CODE=POSITIONMASTER.MCCD  " & _
          "WHERE POSITIONMASTER.COMP='" & compPth & "' AND POSITIONMASTER.UNIT='" & UNCD & "' AND POSITIONMASTER.CODE='" & txtpostioncode & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtdvcd = RS!DVNM
    txtmac = RS!MCNM
    txtend = RS!EndS
    txtdvcd.Enabled = False
    txtmac.Enabled = False
   Else
    txtdvcd.Enabled = True
    txtmac.Enabled = True
    txtdvcd.SetFocus
  End If
End Sub
