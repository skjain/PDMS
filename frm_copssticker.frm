VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frm_copssticker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cops Sticker Printing Module"
   ClientHeight    =   3855
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7110
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   3855
      Left            =   -120
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6800
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
      Begin Crystal.CrystalReport CRPT 
         Left            =   3120
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtdoff 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   10
         Top             =   5160
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox txtgrad 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   3090
      End
      Begin VB.TextBox TXTTIME 
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
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   3
         Top             =   840
         Width           =   1290
      End
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2280
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   1
         Top             =   840
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   5490
      End
      Begin VB.TextBox txtltno 
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
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2280
         Width           =   1410
      End
      Begin VB.TextBox txtdeni 
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   1410
      End
      Begin MSComCtl2.DTPicker txtdate 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   40435713
         CurrentDate     =   40740
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   4680
         TabIndex        =   11
         Top             =   2880
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
         Image           =   "frm_copssticker.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   6000
         TabIndex        =   12
         Top             =   2880
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
         Image           =   "frm_copssticker.frx":059A
         cBack           =   -2147483633
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Doff No."
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
         Index           =   7
         Left            =   4680
         TabIndex        =   23
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade "
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
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Time :"
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
         Index           =   5
         Left            =   4800
         TabIndex        =   21
         Top             =   840
         Width           =   975
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
         TabIndex        =   20
         Top             =   2280
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
         TabIndex        =   19
         Top             =   1320
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
         TabIndex        =   18
         Top             =   840
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
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label LBLHEADING1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Printing of Sticker"
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
         Left            =   3120
         TabIndex        =   0
         Top             =   120
         Width           =   3855
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lot No. :"
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
         Index           =   2
         Left            =   2280
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
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
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Denier :"
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
         Index           =   4
         Left            =   4680
         TabIndex        =   14
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_copssticker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  
  If Trim(TXTTIME.Text) = Empty Then
    MsgBox "Invalid Time "
    TXTTIME.SetFocus
    Exit Sub
  End If
  
  'If Trim(txtdoff.Text) = Empty Then
  '  MsgBox "Invalid Doff "
  '  txtdoff.SetFocus
  '  Exit Sub
  'End If
  
  If Trim(txtdoff.Text) = Empty Then
     txtdoff.Text = ""
  End If
  
  'If Trim(txtgrad.Text) = Empty Then
  '  txtgrad.SetFocus
  '  Exit Sub
  'End If
  
  If txtgrad.Text = Empty Then
    txtgrad.Text = "N/A"
  End If
  
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
  RS.Open "SELECT LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "' AND LTNO='" & txtltno & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Lot No. "
    txtltno.SetFocus
    Exit Sub
  End If
  Dim ITMCOD As String
  
  If RS.State = 1 Then RS.Close
  RS.Open "select code from finitmmst where comp='" & compPth & "' and unit='" & UNCD & "' and dvcd='" & txtdvcd.Tag & "' and name='" & txtdeni & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    MsgBox "Invalid Finish Item"
    txtdeni.SetFocus
    Exit Sub
  End If
  ITMCOD = RS!CODE
  
  'If RS.State = 1 Then RS.Close
  'RS.Open "select code from grdmst where grad='" & txtgrad & "'", CN, adOpenDynamic, adLockOptimistic
  'If RS.EOF Then
  '  MsgBox "Invalid Grade "
  '  txtgrad.SetFocus
  '  Exit Sub
  'End If
 '
  'txtgrad.Tag = RS!CODE
  
  txtgrad.Tag = 0
  
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DOFFTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND POCD='" & txtpostioncode & "' and date='" & Format(txtdate, "mm/dd/yyyy") & "' AND TIME ='" & TXTTIME & "'", CN, adOpenDynamic, adLockOptimistic
  If RS.EOF Then
    RS.AddNew
  End If
  RS!COMP = compPth
  RS!unit = UNCD
  RS!DVCD = txtdvcd.Tag
  RS!POCD = txtpostioncode
  RS!MCCD = txtmac.Tag
  RS!Date = Format(txtdate, "YYYY/MM/DD")
  RS!Time = TXTTIME.Text
  RS!DFNO = txtdoff
  RS!grad = txtgrad.Tag
  RS!ltno = txtltno
  RS!ICOD = ITMCOD
  RS!EndS = Val(txtend)
  RS.Update
  MsgBox "Data Save Succefuly"
  
  Call PRINTDATA
  
  Call ClsData(frm_copssticker)
  txtpostioncode.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  frm_copssticker.KeyPreview = True
  txtdate.MinDate = FSDT
  txtdate.MaxDate = FEDT
  txtdate = Now
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtdate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TXTTIME.SetFocus
  End If
End Sub

Private Sub txtdate_LostFocus()
  Call findata
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False: Key = Empty
    txtgrad.Text = SearchList1("SELECT TOP 20 CODE,GRAD FROM GRDMST", 0, txtgrad, "SELECT Grad From Master")
    txtgrad.Tag = Key
    
  End If
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
  If RS.State = 1 Then RS.Close
  RS.Open "select code from divmst where comp='" & compPth & "' and unit='" & UNCD & "' and name='" & txtdvcd.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtdvcd.Tag = RS!CODE
   Else
    txtdvcd.Tag = Empty
  End If
  RS.Close
  Dim SQL As String: Me.KeyPreview = False
  
  
  
  
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtltno = Empty
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtltno = Empty) Then
     M_DESC = Empty:   NEW_VISIBLE = False: Key = Empty
     SQL = "SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "'"
     txtltno = SearchList(SQL)
  End If
  If txtltno <> Empty Then FindFinishItem
  Me.KeyPreview = True
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
  If txtpostioncode = Empty Then Exit Sub
  Call findata
  If txtltno = Empty Then
    MsgBox "Lot No. Not Defined"
    txtpostioncode.SetFocus
    Exit Sub
  End If
  If txtdeni = Empty Then
    MsgBox "Invalid Item detail "
    txtdeni.SetFocus
    Exit Sub
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
    MsgBox "Invalid Position No."
    Cancel = True
  End If
End Sub


Private Sub FindFinishItem()
Dim RSITM As ADODB.Recordset: Set RSITM = New ADODB.Recordset
If RS.State = 1 Then RS.Close
  RS.Open "select code from divmst where comp='" & compPth & "' and unit='" & UNCD & "' and name='" & txtdvcd.Text & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    txtdvcd.Tag = RS!CODE
   Else
    txtdvcd.Tag = Empty
  End If
  RS.Close
Dim FICD As String

If RSITM.State = 1 Then RSITM.Close
RSITM.Open "SELECT FICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "' AND LTNO='" & txtltno & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSITM.EOF Then FICD = RSITM!FICD
RSITM.Close

If FICD <> Empty Then
  If RSITM.State = 1 Then RSITM.Close
  RSITM.Open "SELECT NAME,ISRETURNABLE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & txtdvcd.Tag & "' AND CODE='" & FICD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RSITM.EOF Then
     txtdeni = RSITM!NAME

  Else
     txtdeni = Empty

  End If
  RSITM.Close
End If
End Sub
Private Sub findata()
 If RS.State = 1 Then RS.Close
 RS.Open "SELECT LTNO,ICOD,FINITMMST.NAME AS NAME FROM POSITIONTRAN " & _
         "INNER JOIN FINITMMST ON FINITMMST.COMP=POSITIONTRAN.COMP AND FINITMMST.UNIT=POSITIONTRAN.UNIT AND " & _
         "FINITMMST.DVCD=POSITIONTRAN.DVCD AND FINITMMST.CODE=POSITIONTRAN.ICOD WHERE POSITIONTRAN.COMP='" & compPth & "' AND " & _
         "POSITIONTRAN.UNIT='" & UNCD & "' AND POSITIONTRAN.POCD='" & txtpostioncode & "' AND POSITIONTRAN.DATE<='" & Format(txtdate, "MM/DD/YYYY") & "' ORDER BY POSITIONTRAN.DATE DESC", CN, adOpenDynamic, adLockOptimistic
         
 If Not RS.EOF Then
   
   txtltno = RS!ltno
   txtdeni = RS!NAME
  Else
   txtltno = Empty
   txtdeni = Empty
 End If
End Sub


Private Sub finddata1()
 If RS.State = 1 Then RS.Close
 RS.Open "SELECT DFNO,GRDMST.GRAD AS GRAD,LTNO,ICOD,FINITMMST.NAME AS NAME FROM DOFFTRAN " & _
         "INNER JOIN FINITMMST ON FINITMMST.COMP=DOFFTRAN.COMP AND FINITMMST.UNIT=DOFFTRAN.UNIT AND " & _
         "FINITMMST.DVCD=DOFFTRAN.DVCD AND FINITMMST.CODE=DOFFTRAN.ICOD " & _
         "INNER JOIN GRDMST ON GRDMST.CODE=DOFFTRAN.GRAD " & _
         "WHERE DOFFTRAN.COMP='" & compPth & "' AND " & _
         "DOFFTRAN.UNIT='" & UNCD & "' AND DOFFTRAN.POCD='" & txtpostioncode & "' AND DOFFTRAN.DATE='" & Format(txtdate, "MM/DD/YYYY") & "' AND TIME='" & TXTTIME & "'", CN, adOpenDynamic, adLockOptimistic
 If Not RS.EOF Then
   txtdoff = RS!DFNO
   txtgrad = RS!grad
  Else
   txtdoff = Empty
   txtgrad = Empty
 End If
End Sub


Private Sub TXTTIME_GotFocus()
  TXTTIME = Format(Time(), "HH:MM")
End Sub

Private Sub TXTTIME_KeyDown(KeyCode As Integer, Shift As Integer)
  Call finddata1
End Sub

Private Sub PRINTDATA()
  Dim MSTDAT As New ADODB.Recordset
  Set MSTDAT = New ADODB.Recordset
  CN.Execute "DELETE FROM COPSSTICKER WHERE COMPUTERNAME='" & ComputerName & "'"
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DOFFTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND POCD='" & txtpostioncode & "' and date='" & Format(txtdate, "mm/dd/yyyy") & "' AND TIME ='" & TXTTIME & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
    Dim CNTR
    CNTR = 1
    For CNTR = 1 To Val(txtend)
     If MSTDAT.State = 1 Then MSTDAT.Close
     MSTDAT.Open "SELECT * FROM COPSSTICKER WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
     MSTDAT.AddNew
     MSTDAT!COMP = RS!COMP
     MSTDAT!unit = RS!unit
     MSTDAT!DVCD = RS!DVCD
     MSTDAT!MCCD = RS!MCCD
     MSTDAT!ICOD = RS!ICOD
     MSTDAT!Date = RS!Date
     MSTDAT!Time = RS!Time
     MSTDAT!ltno = RS!ltno
     MSTDAT!POCD = RS!POCD
     MSTDAT!ComputerName = ComputerName
     MSTDAT.Update
    Next
    
    
    
    CRPT.Reset
    crptConnect CRPT
    
    Dim rptfile
    rptfile = "copsstickerprinting" & "_" & "SLS" & ".RPT"
    
    'copsstickerprinting_SLS.rpt
    
    ReportName = App.PATH & "\Reports\" & rptfile
    
    If Dir(ReportName, vbNormal) = Empty Then
        ReportErrorMessage 1001
        Exit Sub
    End If
    
    CRPT.ReportFileName = ReportName
    CRPT.DiscardSavedData = True
    
    CRPT.ReplaceSelectionFormula "{COPSSTICKER.comp}='" & compPth & "' and {COPSSTICKER.computername}='" & ComputerName & "'"
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "Cops Sticker Printing Report"
    CRPT.WindowShowPrintBtn = True
    CRPT.WindowShowPrintSetupBtn = True
    CRPT.WindowShowSearchBtn = True
    CRPT.WindowShowExportBtn = True
    CRPT.WindowShowRefreshBtn = True
    CRPT.Destination = crptToPrinter
    CRPT.ACTION = 1
    
    txtpostioncode.SetFocus
    
    
  End If
  
  
  
  
  
End Sub
