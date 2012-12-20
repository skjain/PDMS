VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHB~1.OCX"
Begin VB.Form frmWastagePacking 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wastage Packing Module"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10335
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   5115
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9022
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
      Begin VB.TextBox TXTGRAD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TXTRMRK 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   3720
         Width           =   7215
      End
      Begin VB.TextBox TXTMCCD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox TXTLOC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox TXTPKGSTATION 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox TXTDVNM 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TXTITEM 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox TXTQNTY 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   315
         Left            =   8160
         TabIndex        =   6
         Top             =   1080
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
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
         Format          =   18350081
         CurrentDate     =   39383
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   1440
         TabIndex        =   0
         Top             =   4440
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
         Image           =   "frmWastagePacking.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   5400
         TabIndex        =   3
         Top             =   4440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "E&dit"
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
         Image           =   "frmWastagePacking.frx":039A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6720
         TabIndex        =   4
         Top             =   4440
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
         Image           =   "frmWastagePacking.frx":0734
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   4440
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
         Image           =   "frmWastagePacking.frx":0ACE
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   4440
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
         Image           =   "frmWastagePacking.frx":1858
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   4440
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
         Image           =   "frmWastagePacking.frx":1CAA
         cBack           =   -2147483633
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4440
         X2              =   4440
         Y1              =   1800
         Y2              =   2640
      End
      Begin VB.Label LBLCFG 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
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
         Left            =   5040
         TabIndex        =   28
         Top             =   1920
         Width           =   735
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   5160
         X2              =   5160
         Y1              =   2640
         Y2              =   3480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10200
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10200
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine No. / Name"
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
         Left            =   6600
         TabIndex        =   27
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Storage Location"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks(If Any) :"
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
         Left            =   360
         TabIndex        =   25
         Tag             =   "S"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wastage Slip No.   :"
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
         Left            =   6240
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6000
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label LBLSLIP 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXXX"
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
         Left            =   8160
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label LBLHEAD 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Wastage Packing"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   0
         Width           =   4455
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   6840
         X2              =   6840
         Y1              =   1800
         Y2              =   2640
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   2415
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   10095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Division Name"
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
         Left            =   360
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Packing Station"
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
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10200
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1455
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of  Packing"
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
         Left            =   6240
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wastage"
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
         Left            =   480
         TabIndex        =   18
         Tag             =   "S"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   10200
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   17
         Tag             =   "S"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   8160
         TabIndex        =   16
         Tag             =   "S"
         Top             =   1920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWastagePacking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim INDEX As Long
Dim GRADE As String
Dim MCCD As String
Dim DIVCODE As String
Dim DIVNAME As String
Dim M_DBCD As String
Dim PKG_SCOD As String
Dim BOX_PKG_REQ As String
Dim LOAD As String
Dim FICD As String, LOCCOD As String
Public CHALLAN As String

Private Sub Form_Activate()
If DIVCODE = Empty Or DIVNAME = Empty Or TXTPKGSTATION = Empty Or PKG_SCOD = Empty Then
   Unload Me
   Exit Sub
End If

  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  If key_PressNew Then cmdAdd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(ActiveControl.NAME) = "TXTQNTY" And Val(TXTQNTY) = 0 Then Exit Sub
If KeyAscii = 13 Then
  SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
On Error GoTo errLoad

LOAD = "Y"
M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
DIVCODE = Empty: DIVNAME = Empty
  
If DIVCODE = Empty Then
   DIVNAME = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
   DIVCODE = Key
End If

  TXTDVNM = UCase(DIVNAME)
  
'DEFAULT
BOX_PKG_REQ = "Y"
  
'-------PACKING STATION MASTER
M_DESC = Empty:  Key = Empty:  NEW_VISIBLE = False: PKG_SCOD = Empty
TXTPKGSTATION = SearchList1("SELECT TOP 20 CODE,NAME FROM PCKMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "'", 0, "", "SELECT PACKING STATION FROM MASTER LIST")
If Key = Empty Then Exit Sub
PKG_SCOD = Key
'---------------------------
  
  LBLSLIP.Caption = GenPackSlipNo(PKG_SCOD)
  
  M_DBCD = "000006"
  Call btn_sts(True)
  TXTVBDT = Now
  cmdExit.Cancel = True
  Me.Show
JUMP:
  Exit Sub
  
errLoad:
    ErrNumber = ERR.Number
    ErrMessage = ERR.Description
    frm_ErrorHandler.Show vbModal
End Sub

Private Sub cmdAdd_Click()
    Call btn_sts(False)
    TXTVBDT = Now
    If TXTITEM.Enabled Then TXTITEM.SetFocus
    LBLSLIP = GenPackSlipNo(PKG_SCOD)
    SAVEFLAG = True
    cmdCancel.Cancel = True
    cmdDelete.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    cmdExit.Cancel = True
    Call btn_sts(True)
    
    TXTDVNM.Tag = TXTDVNM
    TXTPKGSTATION.Tag = TXTPKGSTATION
    ClsData (Me)
    TXTDVNM = TXTDVNM.Tag
    TXTPKGSTATION = TXTPKGSTATION.Tag
    
    LBLSLIP = GenPackSlipNo(PKG_SCOD)
    TXTVBDT = Now
    cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim ANS As String, SQL As String, TEMPRS As New ADODB.Recordset
    
    If LBLSLIP.Caption = "" Then Exit Sub
    If SAVEFLAG = True Then Exit Sub
    
    ANS = MsgBox("Do you Want to Delete this record?", vbYesNo + vbQuestion, App.TITLE)
    
    If ANS = vbYes Then
    
    If BOX_PKG_REQ = "Y" Then
       SQL = "UPDATE BOXREGISTER SET RECSTAT = 'D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
        "' AND VBNO='" & LBLSLIP & "' AND PKG_STCOD='" & PKG_SCOD & "' AND RECSTAT<>'D' AND VTYP='PPF' and DBCD='" & M_DBCD & "'"
    Else
       SQL = "UPDATE PKGMAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
       "' AND DBCD='" & M_DBCD & "' AND SLIPNO='" & LBLSLIP & "' AND VTYP='PPF' AND RECSTAT='A'"
    End If
    
       CN.BeginTrans
       CN.Execute SQL
       
       
     CN.Execute "UPDATE STORETRAN SET RECSTAT='D' WHERE COMP='" & compPth & "' AND VTYP='PPF' AND UNIT='" & UNCD & _
     "' AND DBCD='" & M_DBCD & "' AND VBNO='" & LBLSLIP & "'"
              
       CN.CommitTrans
       MsgBox "Data are Successfully Deleted."
       
     End If
     
    Call cmdCancel_Click
End Sub

Private Sub cmdEdit_Click()
    SAVEFLAG = False
    frmWastagePackingList.DIVCODE = DIVCODE
    frmWastagePackingList.M_DBCD = M_DBCD
    frmWastagePackingList.BOX_PKG_REQ = BOX_PKG_REQ
    frmWastagePackingList.PKGSTCOD = PKG_SCOD
    CHALLAN = Empty
    frmWastagePackingList.Show 1
    
    If CHALLAN = Empty Or CHALLAN = "" Then
       btn_sts (True)
       cmdAdd.Enabled = True
       SAVEFLAG = True
       cmdAdd.SetFocus
    Else
       btn_sts (False)
    End If
    
End Sub

Private Sub cmdExit_Click()
    key_PressNew = False
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo errPRIMARYKEY

    Dim SQL As String, FICD As String, LOCCOD As String
    
    Dim TEMPRS As New ADODB.Recordset
    Set TEMPRS = New ADODB.Recordset
         
    If SAVEFLAG = True Then
        LBLSLIP = GenPackSlipNo(PKG_SCOD)
        If IsBoxExistInUnit(LBLSLIP) Then
           MsgBox "Slip No. " & LBLSLIP.Caption & " Already Exist.", vbCritical
           Exit Sub
        End If
    End If
    
    If Trim(TXTITEM.Text) = "" Then
       MsgBox "Please Enter Valid Item Name.", vbInformation, App.TITLE
       If Trim(TXTITEM.Text) = "" Then TXTITEM = Trim(TXTITEM): TXTITEM.SetFocus
       Exit Sub
    End If
    
    
    If TXTGRAD = Empty Then
       MsgBox "Please Select Grade", vbOKOnly
       Exit Sub
    End If
    
                       
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT CODE FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
   "' AND NAME ='" & TXTITEM & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      FICD = Trim(TEMPRS!CODE)
   End If
   TEMPRS.Close
   
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT CODE FROM LOCMST WHERE NAME ='" & TXTLOC & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      LOCCOD = Trim(TEMPRS!CODE)
   End If
   TEMPRS.Close
      
   If TEMPRS.State = 1 Then TEMPRS.Close
   TEMPRS.Open "SELECT CODE FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
        "' AND DVCD='" & DIVCODE & "' AND NAME ='" & TXTMCCD & "'", CN, adOpenDynamic, adLockOptimistic
   If Not TEMPRS.EOF Then
      MCCD = Trim(TEMPRS!CODE)
   End If
   TEMPRS.Close
      
   GRADE = GetCode("GRDMST", TXTGRAD, "GRAD", "CODE")
      
   CN.BeginTrans
      
   If SAVEFLAG = True Then
     On Error GoTo errPRIMARYKEY
               
     If BOX_PKG_REQ = "Y" Then
      SQL = "INSERT INTO BOXREGISTER(COMP,UNIT,DVCD,DBCD,VTYP,VBNO,VBDT,CHLN,PKG_STCOD,PKGNG_COD,"
      SQL = SQL & "LOCCOD,PCOD,ISRETURNABLE,LOTNO,ICOD,GRAD,SUBGRD,MCCD,COPS,BOXWGT,COPSWGT,GRSWGT,TRWGT,"
      SQL = SQL & "NTWGT,RMRK,RECSTAT)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
      "','" & M_DBCD & "','PPF','" & LBLSLIP & "','" & Format(TXTVBDT, "MM/DD/YYYY") & "','" & LBLSLIP & _
      "','" & PKG_SCOD & "','WASTE','" & LOCCOD & "','WASTE','N','WASTE','" & FICD & "','" & GRADE & "','0','" & MCCD & _
      "','0','0','0','" & Val(TXTQNTY) & "','0','" & Val(TXTQNTY) & "','" & Trim(TXTRMRK) & "','A')"
     Else
      SQL = "INSERT INTO PKGMAN(COMP,UNIT,DVCD,DBCD,VTYP,SLIPNO,DATE,PKG_STCOD,PCOD,BOX_COD,COPS_COD,LOTNO,"
      SQL = SQL & "FINITMCOD,GRAD,SUBGRAD,LOCCOD,MCCD,NOB,CPB,GWPB,TWPB,NWPB,QNTY,SYSR,OPER,[USER],REMARKS,RECSTAT) VALUES('" & compPth & _
      "','" & UNCD & "','" & DIVCODE & "','" & M_DBCD & "','PPF','" & LBLSLIP & "','" & Format(TXTVBDT, "MM/DD/YYYY") & _
      "','" & PKG_SCOD & "','WASTE','WASTE','WASTE','WASTE','" & FICD & _
      "','" & GRADE & "','0','" & LOCCOD & "','" & MCCD & "','0','0','0','0','0','" & Val(TXTQNTY) & "','N','+','" & cUName & "','" & TXTRMRK & "','A')"
     End If
      
      CN.Execute SQL
      CN.Execute "UPDATE PCKMST SET [LBNO]='" & LBLSLIP & "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE='" & PKG_SCOD & "'"
          
   Else
          
    If BOX_PKG_REQ = "Y" Then
    
        SQL = "UPDATE BOXREGISTER SET LOTNO='WASTE',MCCD='" & MCCD & "',LOCCOD='" & LOCCOD & "',ICOD='" & FICD & _
        "',GRSWGT='" & Val(TXTQNTY) & "',NTWGT='" & Val(TXTQNTY) & "',RMRK='" & Trim(TXTRMRK) & "',GRAD='" & GRADE & _
        "' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & _
        "' AND VBNO='" & LBLSLIP & "' AND PKG_STCOD='" & PKG_SCOD & "' AND RECSTAT<>'D' AND VTYP='PPF' and DBCD='" & M_DBCD & "'"

    Else
        
        SQL = "UPDATE PKGMAN SET LOTNO='WASTE',MCCD='" & MCCD & "',LOCCOD='" & LOCCOD & "',FINITMCOD='" & FICD & _
        "',GWPB='" & Val(TXTQNTY) & "',NWPB='" & Val(TXTQNTY) & "',QNTY='" & Val(TXTQNTY) & "',REMARKS='" & Trim(TXTRMRK) & "',GRAD='" & GRADE & _
        "' WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND DBCD='" & M_DBCD & "' AND SLIPNO='" & LBLSLIP & _
        "' AND VTYP='PPF' AND RECSTAT='A'"
    End If
    
    CN.Execute SQL
   End If
   
   'DAILYSTATUS ENTRY
   
   If SAVEFLAG = True Then
      Call DAILYSTATUS("PPF", FICD, M_DBCD, Val(TXTQNTY), LBLSLIP, 0, cUName, "N", Now, TXTVBDT)
    Else
      Call DAILYSTATUS("PPF", FICD, M_DBCD, Val(TXTQNTY), LBLSLIP, 0, cUName, "M", Now, TXTVBDT)
   End If
    
    CN.CommitTrans
    
    If SAVEFLAG Then
       MsgBox "Your Wastage Slip No. " & LBLSLIP & " are Successfully Saved."
    Else
       MsgBox "Your Wastage Slip No. " & LBLSLIP & " are Successfully Edited."
    End If
    
    Call btn_sts(True)
    Call cmdCancel_Click
    TXTVBDT = Now
    LBLSLIP.Caption = GenPackSlipNo(PKG_SCOD)
    cmdAdd.SetFocus
    cmdExit.Cancel = True
    Exit Sub

errPRIMARYKEY:
MsgBox ERR.Description
End Sub

Private Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = Not bool
    cmdCancel.Enabled = Not bool
    cmdAdd.Enabled = bool
    cmdEdit.Enabled = bool
    cmdDelete.Enabled = Not bool
    TXTITEM.Enabled = Not bool
    TXTQNTY.Enabled = Not bool
    TXTLOC.Enabled = Not bool
    TXTMCCD.Enabled = Not bool
    TXTRMRK.Enabled = Not bool
End Sub

Private Sub TXTGRAD_GotFocus()
    TXTGRAD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTGRAD_KeyDown(KeyCode As Integer, Shift As Integer)
'If COUNTER > 0 Then Exit Sub
  If Trim(TXTGRAD.Text) = Empty Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = False: Key = Empty
    TXTGRAD.Text = SearchList1("SELECT TOP 20 CODE,GRAD FROM GRDMST", 0, TXTGRAD, "SELECT GRADE")
    TXTGRAD.Tag = Key
  End If
End Sub

Private Sub TXTGRAD_LostFocus()
    TXTGRAD.BackColor = vbWhite
End Sub

Private Sub txtItem_GotFocus()
 TXTITEM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtitem_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or Trim(TXTITEM.Text) = Empty Then
        NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
        TXTITEM.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM FINITMMST  WHERE COMP='" & compPth & _
        "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, TXTITEM, "List of Finish Item Name")
   ElseIf KeyCode = vbKeyDelete Then
        TXTITEM = Empty
   End If
Me.KeyPreview = True

End Sub

Private Sub txtItem_LostFocus()
 TXTITEM.BackColor = vbWhite
End Sub

Private Sub TXTLOC_GotFocus()
  TXTLOC.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTLOC_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTLOC.Text) = Empty) Or KeyCode = vbKeyF2 Then
    TXTLOC.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM LOCMST", 0, TXTLOC, "SELECT LOCATION FROM MASTER")
  End If
End Sub

Private Sub TXTLOC_LostFocus()
  TXTLOC.BackColor = vbWhite
End Sub

Private Sub TXTMCCD_GotFocus()
  TXTMCCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMCCD_KeyDown(KeyCode As Integer, Shift As Integer)
Me.KeyPreview = False
   If KeyCode = vbKeyF2 Or (KeyCode = 13 And TXTMCCD = Empty) Then
        NEW_VISIBLE = False:  M_DESC = Empty:   Key = Empty
        TXTMCCD.Text = SearchList1("SELECT  TOP 20 CODE,NAME FROM MACMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'", 0, TXTMCCD, "List of Machine Name")
   ElseIf KeyCode = vbKeyDelete Then
        TXTMCCD = Empty
   End If
Me.KeyPreview = True
End Sub

Private Sub TXTMCCD_LostFocus()
TXTMCCD.BackColor = vbWhite
End Sub

Private Sub TXTQNTY_GotFocus()
  TXTQNTY.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTQNTY_KeyPress(KeyAscii As Integer)
  If CheckNumericKey(KeyAscii, TXTQNTY, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub TXTQNTY_LostFocus()
  TXTQNTY.BackColor = vbWhite
End Sub

Private Sub TXTRMRK_GotFocus()
  TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTRMRK_LostFocus()
  TXTRMRK.BackColor = vbWhite
End Sub


Private Function IsBoxExistInUnit(BOXNUM As String) As Boolean
IsBoxExistInUnit = False

Dim CHKRS As ADODB.Recordset
Set CHKRS = New ADODB.Recordset
If CHKRS.State = 1 Then CHKRS.Close
CHKRS.Open "SELECT * FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND VBNO='" & BOXNUM & "'", CN, adOpenDynamic, adLockOptimistic
If Not CHKRS.EOF Then
   IsBoxExistInUnit = True
End If
CHKRS.Close
End Function


