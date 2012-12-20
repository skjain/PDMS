VERSION 5.00
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FRM_TRNEXPORD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Contract Details"
   ClientHeight    =   6780
   ClientLeft      =   2955
   ClientTop       =   2400
   ClientWidth     =   8580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8580
   Begin VB.ComboBox TXTPAYMENT 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   12
      Top             =   3840
      Width           =   6255
   End
   Begin VB.TextBox TXTEXRATE 
      Height          =   285
      Left            =   480
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox TXTCIFFOB 
      Height          =   315
      Left            =   6480
      Style           =   1  'Simple Combo
      TabIndex        =   19
      Top             =   5640
      Width           =   1695
   End
   Begin VB.ComboBox TXTVSLNO 
      Height          =   315
      Left            =   480
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox TXTMARKS 
      Height          =   315
      Left            =   1080
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Top             =   5640
      Width           =   1335
   End
   Begin VB.ComboBox TXTPAYMENTIRR 
      Height          =   315
      Left            =   3840
      Style           =   1  'Simple Combo
      TabIndex        =   16
      Top             =   5280
      Width           =   4335
   End
   Begin VB.ComboBox TXTREMARK3 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   15
      Top             =   4920
      Width           =   6255
   End
   Begin VB.ComboBox TXTREMARK2 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   14
      Top             =   4560
      Width           =   6255
   End
   Begin VB.ComboBox TXTREMARK1 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   13
      Top             =   4200
      Width           =   6255
   End
   Begin VB.ComboBox TXTFNLDES 
      Height          =   315
      Left            =   480
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   3120
      Width           =   2535
   End
   Begin VB.ComboBox TXTPORTOFDIS 
      Height          =   315
      Left            =   5640
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox TXTPORTOFLOD 
      Height          =   315
      Left            =   3120
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox TXTPLACEOFRCPT 
      Height          =   315
      Left            =   5640
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox TXTPRECARIAGE 
      Height          =   315
      Left            =   3120
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox TXTTERMS 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   11
      Top             =   3480
      Width           =   6255
   End
   Begin VB.ComboBox TXTCNTRYFNLDES 
      Height          =   315
      Left            =   5640
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox TXTCNTRYOFORIGIN 
      Height          =   315
      ItemData        =   "FRM_TRNEXPORD.frx":0000
      Left            =   3120
      List            =   "FRM_TRNEXPORD.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox TXTPKGTYP 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   18
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox TXTBANKDTL 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox TXTEXPORTREF 
      Height          =   285
      Left            =   480
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
   Begin WelchButton.lvButtons_H CMDOK 
      Height          =   495
      Left            =   7080
      TabIndex        =   20
      Top             =   6240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&O.k"
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
      Image           =   "FRM_TRNEXPORD.frx":0004
      cBack           =   -2147483633
   End
   Begin VB.Label Label18 
      Caption         =   "Terms of Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "Exchange Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "FoB/CIF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   37
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label LBLDIV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Export Contract Detail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   8655
   End
   Begin VB.Label Label16 
      Caption         =   "Package Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2520
      TabIndex        =   35
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "Marks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Payment By irrevocable LC at sight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Label Label13 
      Caption         =   "Remark (If Any)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Bank Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   31
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "Final Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Port of Discharge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   29
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Port of Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Vessel/Flight No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Place of Receipt By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   26
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Pre-Carriage By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Terms of Delivery"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Country of Final Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Country of Origin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Export Ref No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   5415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   8295
   End
End
Attribute VB_Name = "FRM_TRNEXPORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
  Me.Hide
End Sub

Private Sub Form_Activate()
  Me.Caption = "SALE ORDER BOOKING FOR " + ORDBOK
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  Call FillCombo("Select Distinct CNTRYOFORGIN from EXPORD where CNTRYOFORGIN is not null or CNTRYOFORGIN<>''", TXTCNTRYOFORIGIN)
  Call FillCombo("Select Distinct CNTRYOFFINALDES from EXPORD where CNTRYOFFINALDES is not null or CNTRYOFFINALDES<>''", TXTCNTRYFNLDES)
  Call FillCombo("Select Distinct TRMSOFDLRY from EXPORD where TRMSOFDLRY is not null or TRMSOFDLRY<>''", TXTTERMS)
  Call FillCombo("Select Distinct TRMSOFPYMT from EXPORD where TRMSOFPYMT is not null or TRMSOFPYMT<>''", TXTPAYMENT)
  Call FillCombo("Select Distinct PRECARIGBY from EXPORD where PRECARIGBY is not null or PRECARIGBY<>''", TXTPRECARIAGE)
  Call FillCombo("Select Distinct PLACEOFRCPT from EXPORD where PLACEOFRCPT is not null or PLACEOFRCPT<>''", TXTPLACEOFRCPT)
  Call FillCombo("Select Distinct VSLFLTNO from EXPORD where VSLFLTNO is not null or VSLFLTNO<>''", TXTVSLNO)
  Call FillCombo("Select Distinct PORTOFLOAD from EXPORD where PORTOFLOAD is not null or PORTOFLOAD<>''", TXTPORTOFLOD)
  Call FillCombo("Select Distinct PORTOFDISCHARG from EXPORD where PORTOFDISCHARG is not null or PORTOFDISCHARG<>''", TXTPORTOFDIS)
  Call FillCombo("Select Distinct FINALDEST from EXPORD where FINALDEST is not null or FINALDEST<>''", TXTFNLDES)
  Call FillCombo("Select Distinct REMARK1 from EXPORD where REMARK1 is not null or REMARK1<>''", TXTREMARK1)
  Call FillCombo("Select Distinct REMARK2 from EXPORD where REMARK2 is not null or REMARK2<>''", TXTREMARK2)
  Call FillCombo("Select Distinct REMARK3 from EXPORD where REMARK3 is not null or REMARK3<>''", TXTREMARK3)
  Call FillCombo("Select Distinct MARKNO from EXPORD where MARKNO is not null or MARKNO<>''", TXTMARKS)
  Call FillCombo("Select Distinct PAYMENTBYIRRLC from EXPORD where PAYMENTBYIRRLC is not null or PAYMENTBYIRRLC<>''", TXTPAYMENTIRR)
  If TXTEXPORTREF.Visible = True Then
    TXTEXPORTREF.SetFocus
  End If
End Sub



Private Sub TXTBANKDTL_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(TXTBANKDTL) = Empty Then
    Key = Empty: NEW_VISIBLE = False: M_DESC = Empty
    
    TXTBANKDTL = SearchList1("SELECT CODE,NAME FROM REFMST WHERE CATA='L'", 0, TXTBANKDTL, "SELECT LC BANK DETAIL")
    If key_PressNew = True Then
       M_DESC = "":  Key = "":  TXTBANKDTL.Text = ""
       FRM_LCMST.Show 1
    End If
  ElseIf KeyCode = vbKeyDelete Then
        TXTBANKDTL = Empty
  End If
End Sub

Private Sub TXTCNTRYFNLDES_GotFocus()
    TXTCNTRYFNLDES.Height = 1155
    TXTCNTRYFNLDES.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCNTRYFNLDES_LostFocus()
    TXTCNTRYFNLDES.BackColor = vbWhite
    TXTCNTRYFNLDES.Height = 325
End Sub
Private Sub TXTCNTRYOFORIGIN_GotFocus()
    TXTCNTRYOFORIGIN.Height = 1155
    TXTCNTRYOFORIGIN.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTCNTRYOFORIGIN_LostFocus()
    TXTCNTRYOFORIGIN.BackColor = vbWhite
    TXTCNTRYOFORIGIN.Height = 325
End Sub

Private Sub TXTEXRATE_GotFocus()
  TXTEXRATE.BackColor = RGB(BRED, BGREEN, BBLUE)
  TXTEXRATE.SelStart = 0
  TXTEXRATE.SelLength = Len(TXTEXRATE)
End Sub

Private Sub TXTEXRATE_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTEXRATE, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTEXRATE_LostFocus()
  TXTEXRATE.BackColor = vbWhite
End Sub

Private Sub TXTFNLDES_GotFocus()
    TXTFNLDES.Height = 1155
    TXTFNLDES.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTFNLDES_LostFocus()
    TXTFNLDES.BackColor = vbWhite
    TXTFNLDES.Height = 325
End Sub

Private Sub TXTMARKS_GotFocus()
    TXTMARKS.Height = 1155
    TXTMARKS.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMARKS_LostFocus()
    TXTMARKS.BackColor = vbWhite
    TXTMARKS.Height = 325
End Sub

Private Sub TXTPAYMENTIRR_GotFocus()
    TXTPAYMENTIRR.Height = 1155
    TXTPAYMENTIRR.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPAYMENTIRR_LostFocus()
    TXTPAYMENTIRR.BackColor = vbWhite
    TXTPAYMENTIRR.Height = 325
End Sub

Private Sub TXTPKGTYP_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Or Trim(TXTPKGTYP) = Empty Then
    TXTPKGTYP = SearchList1("SELECT CODE,NAME FROM PKGNGMST WHERE STATUS='A' AND RECSTAT='A'", 0, TXTPKGTYP, "SELECT PACKAGING TYPE FROM LIST")
    TXTPKGTYP.Tag = Key
  End If
End Sub

Private Sub TXTPLACEOFRCPT_GotFocus()
    TXTPLACEOFRCPT.Height = 1155
    TXTPLACEOFRCPT.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTPLACEOFRCPT_LostFocus()
    TXTPLACEOFRCPT.BackColor = vbWhite
    TXTPLACEOFRCPT.Height = 325
End Sub

Private Sub TXTPORTOFDIS_GotFocus()
    TXTPORTOFDIS.Height = 1155
    TXTPORTOFDIS.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPORTOFDIS_LostFocus()
    TXTPORTOFDIS.BackColor = vbWhite
    TXTPORTOFDIS.Height = 325
End Sub
Private Sub TXTPORTOFLOD_GotFocus()
    TXTPORTOFLOD.Height = 1155
    TXTPORTOFLOD.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTPORTOFLOD_LostFocus()
    TXTPORTOFLOD.BackColor = vbWhite
    TXTPORTOFLOD.Height = 325
End Sub
Private Sub TXTPRECARIAGE_GotFocus()
    TXTPRECARIAGE.Height = 1155
    TXTPRECARIAGE.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTPRECARIAGE_LostFocus()
    TXTPRECARIAGE.BackColor = vbWhite
    TXTPRECARIAGE.Height = 325
End Sub
Private Sub TXTREMARK1_GotFocus()
    TXTREMARK1.Height = 1155
    TXTREMARK1.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTREMARK1_LostFocus()
    TXTREMARK1.BackColor = vbWhite
    TXTREMARK1.Height = 325
End Sub
Private Sub TXTTERMS_GotFocus()
    TXTTERMS.Height = 1155
    TXTTERMS.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTTERMS_LostFocus()
    TXTTERMS.BackColor = vbWhite
    TXTTERMS.Height = 325
End Sub
Private Sub TXTPAYMENT_GotFocus()
    TXTPAYMENT.Height = 1155
    TXTPAYMENT.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTPAYMENT_LostFocus()
    TXTPAYMENT.BackColor = vbWhite
    TXTPAYMENT.Height = 325
End Sub
Private Sub TXTVSLNO_GotFocus()
    TXTVSLNO.Height = 1155
    TXTVSLNO.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTVSLNO_LostFocus()
    TXTVSLNO.BackColor = vbWhite
    TXTVSLNO.Height = 325
End Sub
Private Sub TXTREMARK2_GotFocus()
    TXTREMARK2.Height = 1155
    TXTREMARK2.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTREMARK2_LostFocus()
    TXTREMARK2.BackColor = vbWhite
    TXTREMARK2.Height = 325
End Sub
Private Sub TXTREMARK3_GotFocus()
    TXTREMARK3.Height = 1155
    TXTREMARK3.ZOrder
    ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub
Private Sub TXTREMARK3_LostFocus()
    TXTREMARK3.BackColor = vbWhite
    TXTREMARK3.Height = 325
End Sub
