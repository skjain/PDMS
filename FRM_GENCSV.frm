VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form FRM_GENCSV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate CSV File"
   ClientHeight    =   4920
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8055
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   7815
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   5925
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Export File Path "
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
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Tag             =   "S"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   7815
      Begin VB.ComboBox txtDONO 
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
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TXTSUBGRD 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtGrade 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtLTNo 
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin MSMask.MaskEdBox dtDate 
         Height          =   285
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Format          =   "dd-mm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "DO Qnty."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Tag             =   "S"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "&Sub Grade"
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
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "&Grade"
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
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "&Lot No"
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
         Top             =   600
         Width           =   960
      End
      Begin VB.Label LBLDO 
         BackStyle       =   0  'Transparent
         Caption         =   "D.O. No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "S"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LBLDODT 
         BackStyle       =   0  'Transparent
         Caption         =   "D.O. Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Tag             =   "S"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame framDIVISION 
      Height          =   630
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   7770
      Begin VB.TextBox txtUNIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   195
         Width           =   6420
      End
      Begin VB.Label lblUnit 
         Caption         =   "&Unit Name :"
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
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   1080
      End
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   7770
      Begin VB.TextBox txtDVCD 
         Height          =   330
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   165
         Width           =   6420
      End
      Begin VB.Label Label14 
         Caption         =   "Division :"
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
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   885
      End
   End
   Begin WelchButton.lvButtons_H CMDPENDBOX 
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "&Pending Box"
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
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H CMDPNDDO 
      Height          =   735
      Left            =   3360
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "Pending &DO"
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
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H CMDDSP 
      Height          =   735
      Left            =   4560
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      Caption         =   "Dispatch &Box"
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
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "FRM_GENCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QRY As String
Dim SQL As String

Private Sub CMDDSP_Click()
 QRY = Empty
 QRY = "SELECT BOXREGISTER.VBNO AS BOXN,ORDTRN.DONO AS DONO,BOXREGISTER.RVBDT,PADDMST.NAME AS DELPTY," & _
       "FINITMMST.NAME AS ITNM,BOXREGISTER.LOTNO,GRDMST.GRAD,SUBGRDMST.NAME AS SUBGRAD,BOXREGISTER.NTWGT," & _
       "ORDTRN.ARAT,BOXREGISTER.RVBNO " & _
       "From BOXREGISTER " & _
       "INNER JOIN ORDTRN ON ORDTRN.COMP=BOXREGISTER.COMP AND ORDTRN.UNIT=BOXREGISTER.UNIT " & _
       "AND ORDTRN.DVCD=BOXREGISTER.DVCD AND ORDTRN.RDBC=BOXREGISTER.RDBC " & _
       "AND ORDTRN.SLIP=BOXREGISTER.RVBNO " & _
       "INNER JOIN PADDMST ON PADDMST.CODE=ORDTRN.DCOD AND PADDMST.SRNO=ORDTRN.SRCH " & _
       "INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND FINITMMST.UNIT=BOXREGISTER.UNIT " & _
       "AND FINITMMST.DVCD=BOXREGISTER.DVCD AND FINITMMST.CODE=BOXREGISTER.ICOD " & _
       "INNER JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD " & _
       "LEFT JOIN SUBGRDMST ON SUBGRDMST.COMP=BOXREGISTER.COMP AND SUBGRDMST.UNIT=BOXREGISTER.UNIT " & _
       "AND SUBGRDMST.DVCD=BOXREGISTER.DVCD AND SUBGRDMST.GRAD=BOXREGISTER.GRAD " & _
       "AND SUBGRDMST.SUBGRD=BOXREGISTER.SUBGRD " & _
       "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
       "' AND BOXREGISTER.DVCD='" & txtDVCD.Tag & "' AND  BOXREGISTER.VTYP='DPF' AND " & _
       "BOXREGISTER.RVBDT='" & Format(Now, "MM/DD/YYYY") & "'"
       
 If RS.State = 1 Then RS.Close
 RS.Open QRY, CN, adOpenDynamic, adLockOptimistic
 Call SetFile(QRY, "C:\DOSPRINT\Dispatch.txt")
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub CMDPENDBOX_Click()

Dim FILENAME As String
QRY = Empty

If M_COMPBILL = "SLS" Then
 If txtLTNo = Empty Then txtLTNo.SetFocus: Exit Sub
 If txtGrade = Empty Then txtGrade.SetFocus: Exit Sub
 If TXTSUBGRD = Empty Then TXTSUBGRD.SetFocus: Exit Sub
 
 FILENAME = Trim(txtDONO)
 
 If FILENAME = Empty Then
    FILENAME = Trim(txtLTNo) + "_" + Trim(txtGrade) + "_" + Trim(TXTSUBGRD)
 End If
 
 If Val(txtQTY) > 0 Then
    FILENAME = FILENAME + "_" + Trim(nstr(txtQTY, 12, 3))
 End If
 
 QRY = "SELECT BOXREGISTER.VBNO AS BOXN FROM BOXREGISTER " & _
     "INNER JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD AND GRDMST.GRAD='" & txtGrade & "'" & _
     "INNER JOIN SUBGRDMST ON SUBGRDMST.COMP=BOXREGISTER.COMP AND SUBGRDMST.UNIT=BOXREGISTER.UNIT " & _
     "AND SUBGRDMST.DVCD=BOXREGISTER.DVCD AND SUBGRDMST.GRAD=BOXREGISTER.GRAD AND " & _
     "SUBGRDMST.SUBGRD=BOXREGISTER.SUBGRD AND SUBGRDMST.NAME='" & TXTSUBGRD & _
     "' WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
     "' AND BOXREGISTER.DVCD='" & txtDVCD.Tag & "' AND BOXREGISTER.VTYP='PPF' AND " & _
     "BOXREGISTER.RECSTAT='A' AND BOXREGISTER.LOTNO='" & txtLTNo & "' ORDER BY VBDT,VBNO"
Else
    If txtLTNo = Empty Then txtLTNo.SetFocus: Exit Sub
    
    If Trim(txtDONO) <> Empty Or Trim(txtDONO) <> "" Then
        FILENAME = Trim(txtDONO)
    ElseIf (Trim(txtDONO) = Empty Or Trim(txtDONO) = "") And (Trim(txtGrade) = Empty Or Trim(txtGrade) = "") And (Trim(txtLTNo) <> Empty Or Trim(txtLTNo) <> "") Then
        FILENAME = Trim(txtLTNo)
    ElseIf (Trim(txtDONO) = Empty Or Trim(txtDONO) = "") And (Trim(txtGrade) <> Empty Or Trim(txtGrade) <> "") And (Trim(txtLTNo) <> Empty Or Trim(txtLTNo) <> "") Then
        FILENAME = Trim(txtLTNo) + Trim(txtGrade)
    End If

 QRY = "SELECT VBNO AS BOXN FROM BOXREGISTER " & _
     "INNER JOIN FINITMMST ON FINITMMST.COMP=BOXREGISTER.COMP AND FINITMMST.UNIT=BOXREGISTER.UNIT " & _
     "AND FINITMMST.DVCD=BOXREGISTER.DVCD AND FINITMMST.CODE=BOXREGISTER.ICOD " & _
     "INNER JOIN GRDMST ON GRDMST.CODE=BOXREGISTER.GRAD " & _
     "LEFT JOIN SUBGRDMST ON SUBGRDMST.COMP=BOXREGISTER.COMP AND SUBGRDMST.UNIT=BOXREGISTER.UNIT " & _
     "AND SUBGRDMST.DVCD=BOXREGISTER.DVCD AND SUBGRDMST.GRAD=BOXREGISTER.GRAD " & _
     "AND SUBGRDMST.SUBGRD=BOXREGISTER.SUBGRD " & _
     "WHERE BOXREGISTER.COMP='" & compPth & "' AND BOXREGISTER.UNIT='" & UNCD & _
     "' AND BOXREGISTER.DVCD='" & txtDVCD.Tag & "' AND BOXREGISTER.VTYP='PPF' AND BOXREGISTER.RECSTAT='A' "
 
 If (Trim(txtGrade) = Empty Or Trim(txtGrade) = "") And (Trim(txtLTNo) <> Empty Or Trim(txtLTNo) <> "") Then
        QRY = QRY & " AND BOXREGISTER.LOTNO='" & Trim(txtLTNo) & "' "
 ElseIf (Trim(txtGrade) <> Empty Or Trim(txtGrade) <> "") And (Trim(txtLTNo) <> Empty Or Trim(txtLTNo) <> "") Then
        QRY = QRY & " AND BOXREGISTER.LOTNO='" & Trim(txtLTNo) & "' AND GRDMST.GRAD='" & Trim(txtGrade) & "' "
 End If
 
 QRY = QRY & " ORDER BY VBDT,VBNO"
End If

 'If RS.State = 1 Then RS.Close
 'RS.Open QRY, CN, adOpenDynamic, adLockOptimistic
 
 If Dir(txtFile, vbNormal) = Empty Then
    MsgBox "Invalid Export Drive Path. Check Unit Configuration", vbCritical, "Invalid Drive Path"
    Exit Sub
 End If
 Call SetFile(QRY, txtFile & FILENAME & ".txt")
 
End Sub

Private Sub CMDPNDDO_Click()
 QRY = Empty
 QRY = "SELECT ORDTRN.DONO,ORDTRN.DODT,PADDMST.NAME AS PTYNAM,REFMST.NAME AS AGTNAM,FINITMMST.NAME AS ITNM,GRDMST.GRAD AS GRADE,ORDTRN.LTNO,SUBGRDMST.NAME AS SUBGRADE,ORDTRN.RATE,ORDTRN.QNTY,0 AS DELQTY,ORDTRN.QNTY AS BALQTY FROM ORDTRN " & _
     "INNER JOIN FINITMMST ON FINITMMST.COMP=ORDTRN.COMP AND FINITMMST.UNIT=ORDTRN.UNIT " & _
     "AND FINITMMST.DVCD=ORDTRN.DVCD AND FINITMMST.CODE=ORDTRN.ICOD " & _
     "INNER JOIN PADDMST ON PADDMST.CODE=ORDTRN.DCOD AND PADDMST.SRNO=ORDTRN.SRCH " & _
     "INNER JOIN REFMST ON REFMST.CODE=ORDTRN.BRCD " & _
     "INNER JOIN ACCMST ON ACCMST.CODE=ORDTRN.PCOD " & _
     "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD " & _
     "LEFT JOIN SUBGRDMST ON SUBGRDMST.COMP=ORDTRN.COMP AND SUBGRDMST.UNIT=ORDTRN.UNIT " & _
     "AND SUBGRDMST.DVCD=ORDTRN.DVCD AND SUBGRDMST.GRAD=ORDTRN.GRAD " & _
     "AND SUBGRDMST.SUBGRD=ORDTRN.SUBGRD " & _
     "WHERE ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & UNCD & _
     "' AND ORDTRN.DVCD='" & txtDVCD.Tag & _
     "' AND ORDTRN.VTYP='DOS' AND ORDTRN.DOSTAT='Y' AND DFLG<>'Y' AND ORDTRN.RECSTAT='A' ORDER BY DODT,DONO"
 
 If RS.State = 1 Then RS.Close
 RS.Open QRY, CN, adOpenDynamic, adLockOptimistic
 Call SetFile(QRY, "C:\DOSPRINT\DO.csv")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  txtUNIT = UntNm
  txtUNIT.Tag = UNCD
  SendKeys "{TAB}"
  
   
End Sub

Private Sub txtDONO_GotFocus()
  Call FillDetail
  ActiveControl.BackColor = RGB(BRED, BGREEN, BBLUE)
  txtDONO.Height = 1155
  txtDONO.ZOrder
End Sub

Private Sub TXTDONO_LostFocus()
  txtDONO.BackColor = vbWhite
  txtDONO.Height = 325
  
    SQL = "SELECT LTNO,ORDTRN.GRAD,GRDMST.GRAD AS GRADE,SUBGRDMST.NAME AS SUBGRADE," & _
          "ORDTRN.SUBGRD,ORDTRN.DODT,ORDTRN.QNTY FROM ORDTRN " & _
          "INNER JOIN GRDMST ON GRDMST.CODE=ORDTRN.GRAD " & _
          "LEFT JOIN SUBGRDMST ON ORDTRN.COMP = SUBGRDMST.COMP AND ORDTRN.UNIT = SUBGRDMST.UNIT AND " & _
          "ORDTRN.DVCD = SUBGRDMST.DVCD AND ORDTRN.GRAD = SUBGRDMST.GRAD AND ORDTRN.SUBGRD = SUBGRDMST.SUBGRD " & _
          "WHERE ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & txtUNIT.Tag & "' AND ORDTRN.DVCD='" & txtDVCD.Tag & _
          "' AND DONO='" & txtDONO & "' AND ORDTRN.VTYP='DOS' AND ORDTRN.DFLG<>'Y' AND " & _
          "ORDTRN.RECSTAT='A' AND ORDTRN.DOSTAT='Y' "
               
    If RS.State = 1 Then RS.Close
    RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       txtLTNo = Trim(RS!ltno & "")
       txtGrade = Trim(RS!GRADE & "")
       txtGrade.Tag = Trim(RS!grad & "")
       TXTSUBGRD = Trim(RS!SUBGRADE & "")
       TXTSUBGRD.Tag = Trim(RS!SUBGRD & "")
       dtDate = Format(Trim(RS!DODT & ""), "DD/MM/YYYY")
       txtQTY = Trim(RS!QNTY & "")
    End If
  
End Sub

Private Sub txtDVCD_GotFocus()
 txtDVCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtDVCD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtDVCD = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = False
        M_DESC = Empty
        Key = Empty
        txtDVCD = SearchList1("Select  TOP 20 CODE,NAME From DIVMST Where COMP='" & compPth & "' And UNIT='" & UNCD & "' AND RECSTAT<>'D' AND CODE<>'000001'", 0, Empty, "Select Division")
        txtDVCD.Tag = Key
    ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        txtDVCD = Empty
        txtDVCD.Tag = Empty
    End If
End Sub

Private Sub txtDVCD_LostFocus()
 txtDVCD.BackColor = vbWhite
End Sub

Private Sub SetFile(QRY As String, TXTFLE)
 On Error GoTo LAST
  
  If Dir("C:\DOSPRINT", vbDirectory) = Empty Then MkDir ("C:\DOSPRINT")
   
  Close #1
  Open TXTFLE For Output As #1
  If M_COMPBILL <> "SLS" Then
     Select Case TXTFLE
      Case "C:\DOSPRINT\stockfile.csv"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
      Case "C:\DOSPRINT\do.csv"
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
    End Select
End If


  Dim PRTSTG As String
  Dim CNTR As Long
  
  If RS.State Then RS.Close
  RS.Open QRY, CN, adOpenDynamic, adLockOptimistic
  CNTR = 0
  Do While Not RS.EOF
   PRTSTG = Empty
   CNTR = CNTR + 1
   Dim J As Double
   J = 0
   Dim NUMVAL As Double

   For J = 0 To RS.Fields.COUNT - 1
    If RS.Fields(J).Type = 3 Then
      NUMVAL = RS.Fields(J).Value
      PRTSTG = PRTSTG + nstr(NUMVAL, 12, 3) '+ ","
     ElseIf RS.Fields(J).Type = 135 Then
      PRTSTG = PRTSTG + CStr(RS.Fields(J).Value) '+ ","
     ElseIf RS.Fields(J).Type = 131 Then
      NUMVAL = RS.Fields(J).Value
      PRTSTG = PRTSTG + nstr(NUMVAL, 12, 3) '+ ","
     Else
      PRTSTG = PRTSTG + RS.Fields(J).Value & "" '+ ","
     End If
    Next
    Print #1, PRTSTG
    RS.MoveNext
  Loop
  Close #1
  MsgBox "File Successful Create at " + TXTFLE
  Exit Sub
LAST:
 If ERR.Number = 70 Then
    MsgBox "File is Currently Opened", vbCritical, "Access Denied"
 Else
    MsgBox ERR.Description
 End If
 
End Sub


Private Sub txtgrade_GotFocus()
txtGrade.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtgrade_KeyDown(KeyCode As Integer, Shift As Integer)
If txtDONO <> Empty Then Exit Sub

If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtGrade = Empty) Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   txtGrade = SearchList1("SELECT DISTINCT CODE,GRAD FROM GRDMST", 0, txtGrade, "SELECT GRADE FROM LIST")
   txtGrade.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   txtGrade = Empty
   txtGrade.Tag = Empty
End If
End Sub

Private Sub txtgrade_LostFocus()
 txtGrade.BackColor = vbWhite
End Sub

Private Sub txtltno_GotFocus()
txtLTNo.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtltno_KeyDown(KeyCode As Integer, Shift As Integer)
If txtDONO <> Empty Then Exit Sub

If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtLTNo = Empty) Then
   NEW_VISIBLE = False: CANCEL_VISIBLE = True: M_DESC = Empty: Key = Empty
   txtLTNo = SearchList1("SELECT DISTINCT LTNO,LTNO FROM TXULOT WHERE COMP='" & compPth & _
   "' AND UNIT='" & txtUNIT.Tag & _
   "' AND DVCD = '" & txtDVCD.Tag & "'", 0, txtLTNo, "SELECT LOTNO FROM LIST")
   txtLTNo.Tag = Key
ElseIf KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
   txtLTNo = Empty
   txtLTNo.Tag = Empty
End If
End Sub


Private Sub txtltno_LostFocus()
 txtLTNo.BackColor = vbWhite
End Sub

Private Sub TXTSUBGRD_GotFocus()
  TXTSUBGRD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSUBGRD_KeyDown(KeyCode As Integer, Shift As Integer)
If txtDONO <> Empty Then Exit Sub
If txtGrade = Empty Then Exit Sub

Me.KeyPreview = False

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     TXTSUBGRD = Empty
     TXTSUBGRD.Tag = Empty
ElseIf KeyCode = vbKeyF2 Or (KeyCode = 13 And TXTSUBGRD = Empty) Then
   Key = Empty: M_DESC = Empty:  NEW_VISIBLE = False
   Dim SQL As String
   SQL = "SELECT DISTINCT SUBGRD,NAME FROM SUBGRDMST WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & _
         "' AND DVCD='" & txtDVCD.Tag & "' AND GRAD='" & GetCode("GRDMST", txtGrade, "GRAD", "CODE") & "'"
         
   TXTSUBGRD = SearchList1(SQL, 0, Empty)
   TXTSUBGRD.Tag = Key
End If

Me.KeyPreview = True
End Sub

Private Sub TXTSUBGRD_LostFocus()
TXTSUBGRD.BackColor = vbWhite
End Sub

Private Sub txtUNIT_GotFocus()
 txtUNIT.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtUNIT = Empty) Then
        NEW_VISIBLE = False
        CANCEL_VISIBLE = True
        M_DESC = Empty
        Key = Empty
        txtUNIT = SearchList1("Select  TOP 20 CODE,NAME From UNTMST Where COMP='" & compPth & "'", 0, Empty, "Select Unit To View Report For ")
        txtUNIT.Tag = Key
    End If
    
End Sub

Private Sub txtUNIT_LostFocus()
 txtUNIT.BackColor = vbWhite
 If txtUNIT <> Empty Then
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT SCANEXP FROM UNTCFG WHERE COMP='" & compPth & "' AND UNIT='" & txtUNIT.Tag & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       txtFile = Trim(RS!SCANEXP & "")
    End If
    RS.Close
 End If
End Sub

Private Sub FillDetail()

  
  SQL = "SELECT DISTINCT ORDTRN.DONO FROM ORDTRN WHERE " & _
        "ORDTRN.COMP='" & compPth & "' AND ORDTRN.UNIT='" & txtUNIT.Tag & "' AND ORDTRN.DVCD='" & txtDVCD.Tag & _
        "' AND ORDTRN.VTYP='DOS' AND ORDTRN.DFLG<>'Y' AND ORDTRN.RECSTAT='A' AND ORDTRN.DOSTAT='Y' " & _
        "ORDER BY ORDTRN.DONO"
  
  Call FillCombo(SQL, txtDONO)
  
End Sub
