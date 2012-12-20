VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form frmLotMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lot / Batch  Master Creation / Editing"
   ClientHeight    =   6285
   ClientLeft      =   2925
   ClientTop       =   2010
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9000
   Begin VB.TextBox TXTMRGN 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox SUBPKGNG 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox TXTSHCD 
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox TXTPER 
      Height          =   285
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   17
      Top             =   2880
      Width           =   1035
   End
   Begin FramePlusCtl.FramePlus frameActive 
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1320
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active"
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
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame frameMain 
      Height          =   3135
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   8535
      Begin VB.TextBox M_SRCH 
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   14
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox M_RINM 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   2535
      End
      Begin WelchButton.lvButtons_H CMDOK 
         Height          =   375
         Left            =   7200
         TabIndex        =   18
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
         Image           =   "frmLotMaster.frx":0000
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdRemove 
         Height          =   375
         Left            =   7200
         TabIndex        =   21
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Remove"
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
         Image           =   "frmLotMaster.frx":039A
         Enabled         =   0   'False
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid FLEX 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBMRGN 
         Alignment       =   2  'Center
         Caption         =   "Merge No."
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
         Left            =   3840
         TabIndex        =   32
         Tag             =   "S"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPer 
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7680
         TabIndex        =   30
         Top             =   2760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   960
         X2              =   960
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Sr No."
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
         Left            =   240
         TabIndex        =   25
         Tag             =   "S"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Raw Material"
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
         Left            =   1440
         TabIndex        =   24
         Tag             =   "S"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Percentage (%)"
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
         Left            =   5640
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         Height          =   1095
         Left            =   120
         Top             =   240
         Width           =   8295
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000080&
         X1              =   7080
         X2              =   7080
         Y1              =   240
         Y2              =   1320
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7080
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   5520
         X2              =   5520
         Y1              =   240
         Y2              =   1320
      End
   End
   Begin VB.TextBox M_LTNO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   1800
      MaxLength       =   12
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox M_FINM 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   3855
   End
   Begin WelchButton.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":07EC
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":0B86
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":0F20
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":12BA
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdCancel 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":2044
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   5400
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
      Image           =   "frmLotMaster.frx":2496
      cBack           =   -2147483633
   End
   Begin VB.Label LBLMRGN 
      Caption         =   "Merge :"
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
      Left            =   4080
      TabIndex        =   31
      Tag             =   "S"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label LBLSUBPKG 
      Caption         =   "Sub Packaging :"
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
      TabIndex        =   10
      Tag             =   "S"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label LBLSHCD 
      Caption         =   "Shade  :"
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
      Left            =   5760
      TabIndex        =   12
      Tag             =   "S"
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LBLDIVISION 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Division : SIZING DIVISION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   240
      Width           =   4935
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   0
      X2              =   8520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   8640
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Lot / Batch Master"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Height          =   6135
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label2 
      Caption         =   "Lot/Batch No.  :"
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
      TabIndex        =   6
      Tag             =   "S"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Finish Item  :"
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
      TabIndex        =   8
      Tag             =   "S"
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmLotMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVEFLAG As Boolean
Dim M_FICD As String
Dim M_SHCD As String
Dim M_FIGC As String
Dim M_DVCD As String
Dim M_RICD As String
Dim M_BSCD As String
Dim DIVCODE As String
Dim DIVNAME As String
Dim SAVFLG As Boolean
Dim CFGTYP As String
Dim ROWNO As Long
Dim INFORS As New ADODB.Recordset
Dim SUBPKGCODE As String
Dim SWITCH As Boolean
Dim BOXREQ As String

Private Sub cmdAdd_Click()
Dim Ctrl As Control
  For Each Ctrl In Me
    If TypeOf Ctrl Is TextBox Then
       Ctrl = Replace(Ctrl, "'", "", 1)
    End If
Next
Call btn_sts(True)
M_LTNO.Enabled = True: M_LTNO.SetFocus
SAVEFLAG = True
cmdCancel.Cancel = True
End Sub

Private Sub cmdEdit_Click()
  If M_USRSECLEVL = "1" Then
    If ReadConfigMaster("000009", 5, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  SAVEFLAG = False
  NEW_VISIBLE = False
  M_LTNO = SearchList1("select DISTINCT LTNO AS LOT, LTNO from TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND RECSTAT='A'", 0, "", "List Of LOT/BATCH MASTER")
    
  frameActive.Visible = True
  If M_LTNO <> Empty Then Call FILLFLEX
  
  If isFurtherEntryExist("LOT", M_LTNO, DIVCODE) Then
     MsgBox "Further Entry Exist"
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT ISNULL(RICD,'') AS RICD FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
             "' AND DVCD='" & DIVCODE & "' AND LTNO='" & M_LTNO & "' AND RECSTAT <>'D' AND ACTIVE='Y'", CN, adOpenDynamic, adLockOptimistic
     If Not RS.EOF Then
        If Trim(RS!RICD & "") = Empty Then
           GoTo MOVE
        End If
     End If
     
     If optActive.Value Then
        If MsgBox("Want to Deactive Lot.", vbYesNo + vbDefaultButton2) = vbYes Then
           CN.Execute "UPDATE TXULOT SET ACTIVE='N' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                      "' AND DVCD='" & DIVCODE & "' AND LTNO='" & M_LTNO & "' AND RECSTAT <>'D'"
        End If
     Else
        If MsgBox("Want to Active Lot. ", vbYesNo + vbDefaultButton2) = vbYes Then
           CN.Execute "UPDATE TXULOT SET ACTIVE='Y' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                      "' AND DVCD='" & DIVCODE & "' AND LTNO='" & M_LTNO & "' AND RECSTAT <>'D'"
        End If
     End If
     
     Call RESETALL
     Call btn_sts(False)
     Call cmdCancel_Click
     M_SRCH.Enabled = True
     cmdAdd.SetFocus
     Exit Sub
 End If
  
  'CHECK FOR TRANSACTION EXIST
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT TOP 1 LOTNO FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND DVCD='" & DIVCODE & "' AND LOTNO='" & M_LTNO & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     MsgBox "Lot Can't be Edit : Entry Exist in Transaction"
     Call RESETALL
     Call btn_sts(False)
     Call cmdCancel_Click
     M_SRCH.Enabled = True
     cmdAdd.SetFocus
     Exit Sub
  End If
  '----------------------------
'FOR THOSE WHICH IS DIRECTLY TRANSFERED BUT RAW NOT PROPERLY DEFINED AT THE TIME OF TRANSFERING
MOVE:
  If M_FINM.Enabled = True Then M_FINM.SetFocus
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub CMDOK_Click()
 Dim INDEX As Long
 
 If Val(M_SRCH) <= 0 Then
   Call M_SRCH_GotFocus
 End If
 
 'COMMENT BECAUSE LESS/GAIN FOR WIP ADJUSTMENT
 'If ChkPercentage(SWITCH) Then Exit Sub
 
 If Not SWITCH Then
      ROWNO = FLEX.Rows - 1
 End If
 
 If CheckData(ROWNO) Then Exit Sub
 
    FLEX.TextMatrix(ROWNO, 0) = Trim(M_SRCH)
    FLEX.TextMatrix(ROWNO, 1) = Trim(M_RINM)
    
    FLEX.TextMatrix(ROWNO, 2) = TXTPER
    
    FLEX.TextMatrix(ROWNO, 3) = TXTMRGN
        
    Call SetPercentage
           
    If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If FLEX.TextMatrix(FLEX.Rows - 1, 1) <> "" Then
             FLEX.Rows = FLEX.Rows + 1
          End If
           If FLEX.Rows > 6 Then FLEX.TopRow = FLEX.TopRow + 2
           M_RINM.SetFocus
    Else
        cmdSave.Enabled = True: cmdSave.SetFocus
    End If
    
    If FLEX.Rows - 1 <= 9 Then
       M_SRCH = "0" & CStr(FLEX.Rows - 1)
    Else
       M_SRCH = CStr(FLEX.Rows - 1)
    End If
    
    'REMOVE BELOW COMMENT BLOCK WHEN ITEMS PROCESS ARE GOING TO MULTIPLE
    Call CLEARDATA
    cmdOk.Caption = "&Add"
    SWITCH = False
End Sub

Private Sub cmdRemove_Click()
Dim CURSOR As Long
Dim J As Long

For J = ROWNO To FLEX.Rows - 2
 FLEX.TextMatrix(J, 1) = FLEX.TextMatrix(J + 1, 1)
 FLEX.TextMatrix(J, 2) = FLEX.TextMatrix(J + 1, 2)
Next J

FLEX.Rows = FLEX.Rows - 1

Call CLEARDATA

If MsgBox("Want to Add More Item ", vbYesNo + vbDefaultButton2) = vbYes Then
          If FLEX.TextMatrix(FLEX.Rows - 1, 1) <> "" Then
             FLEX.Rows = FLEX.Rows + 1
          End If
           If FLEX.Rows > 6 Then FLEX.TopRow = FLEX.TopRow + 2
           M_RINM.SetFocus
Else
        cmdSave.Enabled = True: cmdSave.SetFocus
End If

If FLEX.Rows - 1 <= 9 Then
   M_SRCH = "0" & CStr(FLEX.Rows - 1)
Else
   M_SRCH = CStr(FLEX.Rows - 1)
End If

SWITCH = False
M_RINM.SetFocus
cmdOk.Caption = "&Add"
CMDREMOVE.Enabled = False
End Sub

Private Sub Flex_Click()

If FLEX.Rows > 1 And FLEX.TextMatrix(FLEX.ROW, 1) <> Empty Then
    cmdOk.Caption = "Upd&ate"
    CMDREMOVE.Enabled = True
    ROWNO = FLEX.ROW
    M_SRCH = FLEX.TextMatrix(ROWNO, 0)
    M_RINM = FLEX.TextMatrix(ROWNO, 1)
    TXTPER = FLEX.TextMatrix(ROWNO, 2)
    TXTMRGN = FLEX.TextMatrix(ROWNO, 3)
    SWITCH = True
  End If
    
   If Val(FLEX.ROW) > 0 Then
     M_SRCH = FLEX.TextMatrix(FLEX.ROW, 0)
     Call M_SRCH_LostFocus
     If M_RINM.Enabled = True Then M_RINM.SetFocus
   End If
End Sub


Private Sub Form_Load()
  Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  DIVCODE = Empty:  DIVNAME = Empty:  NEW_VISIBLE = False
  DIVNAME = SearchList1("SELECT TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RECSTAT='A' AND CODE<>'000001'", 0, "", "SELECT DIVISION MASTER")
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND NAME='" & DIVNAME & "' AND CODE<>'000001'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
     DIVCODE = RS!CODE
     DIVNAME = RS!NAME
     Me.Tag = DIVCODE
  End If
  
  '------------------------------------------------------------------------
 'SUB PACKAGING TYPE
 If INFORS.State = 1 Then INFORS.Close
    INFORS.Open "SELECT * FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE = '" & DIVCODE & "' AND RECSTAT='A' ", CN, adOpenDynamic, adLockOptimistic
 If Not INFORS.EOF Then
    CFGTYP = INFORS!CFGTYP & ""
 End If
 
 If CFGTYP = "SG" Then
    LBLSUBPKG.Enabled = True
    SUBPKGNG.Enabled = True
 Else
    LBLSUBPKG.Enabled = False
    SUBPKGNG.Enabled = False
 End If
 '------------------------------------------------------------------------
  
 ' Dim BRS As New ADODB.Recordset
 ' If BRS.State = 1 Then BRS.Close
 ' BRS.Open "SELECT * FROM UNTCFG WHERE COMP = '" & compPth & "' AND UNIT = '" & UNCD & "' ", CN, adOpenDynamic, adLockOptimistic
 ' If Not BRS.EOF Then
 ' BOXREQ = Trim(BRS!BOXREQ)
 ' End If
  
 ' If BOXREQ = "Y" Then
 '    LBLMRGN.Visible = True
 '    TXTMRGN.Visible = True
 ' Else
 '    LBLMRGN.Visible = False
 '    TXTMRGN.Visible = False
 ' End If
  
  
  LBLDIVISION.Caption = "Division : " & UCase(DIVNAME)
  M_DVCD = DIVCODE
  Call RESETALL
  
  If LabelDisplay(M_DVCD, UNCD) = "Shade" Then
     LBLSHCD.Visible = True
     TXTSHCD.Visible = True
  End If
  
  Call btn_sts(False)
  Me.KeyPreview = True
End Sub

Private Sub Form_Activate()
  If DIVNAME = Empty Or DIVCODE = Empty Then
   MsgBox "SELECT DIVISION"
   Unload Me
  End If

  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  DIVCODE = Me.Tag
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If ActiveControl.NAME = "M_FINM" And M_FINM = Empty Then Exit Sub
 If ActiveControl.NAME = "M_LTNO" And M_LTNO = Empty Then Exit Sub
 
 If ActiveControl.NAME = "M_RINM" And M_RINM = Empty Then Exit Sub
  If KeyAscii = vbKeyReturn Then
    SendKeys "{TAB}"
  End If
End Sub

Private Sub cmdCancel_Click()
 Call RESETALL
 Call btn_sts(False)
 M_SRCH.Enabled = True
 cmdOk.Caption = "&Add"
 cmdOk.Enabled = True
 optActive.Value = True
 frameActive.Visible = False
 CMDREMOVE.Enabled = False
 cmdAdd.SetFocus
 lblPer.Caption = "0 %"
End Sub

Private Sub cmdDelete_Click()
  If M_USRSECLEVL = 1 Then
     If ReadConfigMaster("000009", 6, "M") = False Then ModuleDeniedMessage: Exit Sub
  End If
  On Error GoTo LAST
  
  frameActive.Visible = True
  
  SAVEFLAG = False
  NEW_VISIBLE = False
  M_LTNO = SearchList1("select DISTINCT LTNO AS LOT, LTNO from TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND RECSTAT = 'A'", 0, "", "List Of LOT/BATCH MASTER")
  
  If M_LTNO <> Empty Then Call FILLFLEX
  
  If isFurtherEntryExist("LOT", M_LTNO, DIVCODE) Then
     MsgBox "Further Entry Exist"
     Call RESETALL
     Call btn_sts(False)
     Call cmdCancel_Click
     M_SRCH.Enabled = True
     cmdAdd.SetFocus
     Exit Sub
 End If
     
  'CHECK FOR TRANSACTION EXIST
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT TOP 1 LOTNO FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND DVCD='" & DIVCODE & "' AND LOTNO='" & M_LTNO & "' AND RECSTAT='A'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then
     MsgBox "Lot Can't be Edit : Entry Exist in Transaction"
     Call RESETALL
     Call btn_sts(False)
     Call cmdCancel_Click
     M_SRCH.Enabled = True
     cmdAdd.SetFocus
     Exit Sub
  End If
  '----------------------------
  
  Dim AYS
  AYS = MsgBox("Are You Sure to Delete It ? ", vbYesNo)
  If AYS = vbYes Then
    CN.BeginTrans
    CN.Execute "DELETE FROM TXULOT WHERE LTNO='" & M_LTNO & "' AND COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "'"
    '---------------------
    'DAILYSTATUS
     Call DAILYSTATUS("LOT", M_LTNO, "", 0, "", 0, cUName, "D", Now, Now)
    CN.CommitTrans
  End If
  Call RESETALL
  Call cmdCancel_Click
  Call btn_sts(False)
  If cmdAdd.Enabled Then cmdAdd.SetFocus
  Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub RESETALL()
  M_LTNO = Empty
  M_FINM = Empty
  M_RINM = Empty
  M_FICD = Empty
  M_RICD = Empty
  M_BSCD = Empty
  M_SRCH = Empty
  TXTPER = Empty
  TXTSHCD = Empty
  SUBPKGNG = Empty
  
  FLEX.Clear
  FLEX.Rows = 1
  FLEX.Rows = 2
  FLEX.ColWidth(0) = 800
  FLEX.ColWidth(1) = 4350
  FLEX.ColWidth(2) = 1800
  FLEX.ColWidth(3) = 0
  
  FLEX.Clear
  FLEX.TextMatrix(0, 0) = "SrNo."
  FLEX.TextMatrix(0, 1) = "Raw Material"
  FLEX.TextMatrix(0, 2) = "Percentage"
  FLEX.TextMatrix(0, 3) = "Merge"
    
  FLEX.ColAlignment(0) = vbLeftJustify
  FLEX.ColAlignment(1) = vbLeftJustify
  FLEX.ColAlignment(2) = vbRightJustify
End Sub

Private Sub M_FINM_GotFocus()
M_FINM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_FINM_KeyDown(KeyCode As Integer, Shift As Integer)
    
    M_DESC = Empty
    Key = Empty
    
    If KeyCode = vbKeyF2 Or Trim(M_FINM) = Empty Then
        NEW_VISIBLE = False
        M_FINM.Text = SearchList1("Select TOP 20 CODE,NAME from FINITMMST WHERE COMP='" & compPth & _
                      "' AND UNIT='" & UNCD & "' AND DVCD='" & M_DVCD & "'", 0, M_FINM, "List Of FINISH ITEM MASTER")
        M_FICD = Key
    End If
    
    'If key_PressNew = True Then
    '    M_DESC = ""
    '    frm_FinItmMst.ONLINEITEM = True
    '    mod_Var.DIVCOD = DIVCODE
    '    mod_Var.DIVNAM = DIVNAME
    '    frm_FinItmMst.Show
    'End If
    
End Sub

Private Sub M_FINM_LostFocus()
M_FINM.BackColor = vbWhite
End Sub

Private Sub M_LTNO_GotFocus()
M_LTNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_LTNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 32, 39
            KeyAscii = 0
        Case Else
            
    End Select
End Sub

Private Sub M_LTNO_LostFocus()
M_LTNO.BackColor = vbWhite
End Sub

Private Sub M_RINM_GotFocus()
M_RINM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub M_RINM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RS As New ADODB.Recordset
Dim SPECI As String
Dim MRGN As String
Dim igcd As String


    M_DESC = Empty
    Key = Empty
    If KeyCode = vbKeyF2 Or Trim(M_RINM) = Empty Then
        NEW_VISIBLE = True
        M_RINM.Text = SearchList1("select DISTINCT code, name from ITMMST", 0, M_RINM, "List Of RAW-ITEM MASTER")
        M_RICD = Key
    End If
    If key_PressNew = True Then
        M_DESC = ""
        frm_Item.Show
    End If
    

    If RS.State = 1 Then RS.Close
       RS.Open "SELECT *  FROM ITMMST WHERE CODE = '" & GetCode("ITMMST", M_RINM, "NAME", "CODE") & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       igcd = RS!igcd
    End If
    
    If RS.State = 1 Then RS.Close
       RS.Open "SELECT * FROM IGMMST WHERE CODE = '" & igcd & "'", CN, adOpenDynamic, adLockOptimistic
    If Not RS.EOF Then
       SPECI = RS!SPECIFICATION
       MRGN = RS!MERGE
    End If
    If MRGN = "Y" Then
       LBMRGN.Visible = True
       TXTMRGN.Visible = True
    Else
       LBMRGN.Visible = False
       TXTMRGN.Visible = False
    End If
    'Call M_SRCH_GotFocus
    'Call M_SRCH_LostFocus
End Sub

Private Sub cmdSave_Click()
  On Error GoTo LAST
  Dim INDEX As Long
    
  'If lblPer.Caption <> "100 %" Then MsgBox "Percentage Should be 100%": Exit Sub
  
  If M_LTNO = "WASTE" Then
     MsgBox "WASTE IS A RESERVE WORD IN ERP, PLEASE USE ANOTHER ONE", vbCritical
     If M_LTNO.Enabled Then M_LTNO.SetFocus
     Exit Sub
  End If
  
  If Not DataEntered Then
     MsgBox "No Data Found To Save Record !! Can Not Save Record !!", vbInformation, "Cancelled !!"
     Exit Sub
  End If
    
  'check for repeatation of lot
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND DVCD='" & DIVCODE & "' AND LTNO='" & M_LTNO & "' AND RECSTAT='A'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF And SAVEFLAG Then
    MsgBox "Lot No. Already Exist"
    If M_LTNO.Enabled Then M_LTNO.SetFocus
    Exit Sub
  End If
  '-----------------------------------
  
  For INDEX = 1 To FLEX.Rows - 1
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT CODE FROM ITMMST WHERE NAME='" & FLEX.TextMatrix(INDEX, 1) & "' ", CN, adOpenKeyset, adLockPessimistic
     If RS.EOF Then
        MsgBox "Item Not Defined Properly", vbCritical
        FLEX.COL = 1
        FLEX.ROW = INDEX
        FLEX.SetFocus
        Exit Sub
     End If
  Next INDEX
   
           
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT * FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
  "' AND DVCD='" & DIVCODE & "' AND NAME='" & M_FINM & "'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
    M_FICD = RS!CODE
   Else
    M_FINM.SetFocus
    Exit Sub
  End If
  
  If RS.State = 1 Then RS.Close
     RS.Open "SELECT * FROM SUBPKGNGMST WHERE NAME ='" & Trim(SUBPKGNG) & "'", CN, adOpenKeyset, adLockPessimistic
     If Not RS.EOF Then
        SUBPKGCODE = RS!CODE & ""
     Else
        If SUBPKGNG.Enabled = True Then
           SUBPKGNG.SetFocus
           Exit Sub
        End If
  End If

  M_SHCD = 0
  If LBLSHCD.Visible Then
     If RS.State = 1 Then RS.Close
     RS.Open "SELECT CODE FROM GRDMST WHERE GRAD='" & TXTSHCD & "'", CN, adOpenKeyset, adLockPessimistic
     If Not RS.EOF Then
        M_SHCD = RS!CODE & ""
     Else
        MsgBox "Shade not define", vbCritical
        If TXTSHCD.Visible And TXTSHCD.Enabled Then TXTSHCD.SetFocus
        Exit Sub
     End If
  End If

Dim ACT As String
If optActive.Value = True Then
   ACT = "Y"
ElseIf optDeactive.Value = True Then
   ACT = "N"
Else
   ACT = "N"
End If

CN.BeginTrans
CN.Execute "DELETE FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND DVCD='" & DIVCODE & "' AND LTNO='" & M_LTNO & "' AND RECSTAT <>'D'", INDEX

For INDEX = 1 To FLEX.Rows - 1
 
 M_RICD = GetCode("ITMMST", FLEX.TextMatrix(INDEX, 1), "NAME", "CODE")
 
 QUERY = "INSERT INTO TXULOT(COMP,UNIT,DVCD,LTNO,SRCH,FICD,RICD,PERC,ACTIVE,SHCD,SUBPKGCODE,MRGN)VALUES ('" & compPth & _
 "','" & UNCD & "','" & DIVCODE & "','" & M_LTNO & "','" & FLEX.TextMatrix(INDEX, 0) & _
 "','" & M_FICD & "','" & M_RICD & "','" & FLEX.TextMatrix(INDEX, 2) & "','" & ACT & "','" & M_SHCD & "','" & SUBPKGCODE & "','" & FLEX.TextMatrix(INDEX, 3) & "')"

 CN.Execute QUERY

'---------------------------
Next INDEX

'DAILYSTATUS
  If SAVEFLAG = True Then
     Call DAILYSTATUS("LOT", M_LTNO, "", 0, "", 0, cUName, "N", Now, Now)
  Else
     Call DAILYSTATUS("LOT", M_LTNO, "", 0, "", 0, cUName, "M", Now, Now)
  End If
  
CN.CommitTrans

If SAVEFLAG = True Then
   MsgBox "Assigned LotNo. " & M_LTNO
Else
   MsgBox "Lotno. " & M_LTNO & " edit Successfully"
End If
   
  Call RESETALL
  Call btn_sts(False)
  Call cmdCancel_Click
  M_SRCH.Enabled = True
  cmdAdd.SetFocus
  Exit Sub
LAST:
  MsgBox ERR.Description
  RS.CancelUpdate
End Sub

Private Sub M_RINM_LostFocus()
M_RINM.BackColor = vbWhite
End Sub

Public Sub btn_sts(bool As Boolean)
    cmdSave.Enabled = bool
    cmdCancel.Enabled = bool
    cmdAdd.Enabled = Not bool
    cmdEdit.Enabled = Not bool
    cmdDelete.Enabled = Not bool
    M_FINM.Enabled = bool
    frameMain.Enabled = bool
End Sub

Private Sub M_SRCH_GotFocus()
  M_SRCH.BackColor = RGB(BRED, BGREEN, BBLUE)
  
  M_SRCH.SelStart = 0
  M_SRCH.SelLength = Len(M_SRCH)
  
  
  If Not DataEntered Then
    M_SRCH = "1"
    M_SRCH.Enabled = False
  End If
End Sub

Private Sub M_SRCH_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub M_SRCH_LostFocus()
  M_SRCH.BackColor = vbWhite
   
  Dim NO As Double
    
  NO = Val(M_SRCH)
      
  If NO <= 9 Then
    M_SRCH = "0" + nstr(NO, 1, 0)
   Else
    M_SRCH = nstr(NO, 2, 0)
  End If
End Sub

Private Function DataEntered() As Boolean
Dim i As Integer
  DataEntered = False
    
  If M_FINM = Empty And M_FINM.Enabled = True Then M_FINM.SetFocus: DataEntered = True: Exit Function
  If M_LTNO = Empty And M_LTNO.Enabled = True Then M_LTNO.SetFocus: DataEntered = True:  Exit Function
  If SUBPKGNG = Empty And SUBPKGNG.Enabled = True Then SUBPKGNG.SetFocus: DataEntered = True: Exit Function
  
    For i = 1 To FLEX.Rows - 1
        If Trim(FLEX.TextMatrix(i, 0)) = "" Then
            DataEntered = False
            Exit Function
        End If
        
        If Trim(FLEX.TextMatrix(i, 1)) <> "" Then
            If FLEX.TextMatrix(FLEX.Rows - 1, 1) = Empty Then FLEX.Rows = FLEX.Rows - 1
            DataEntered = True
            Exit Function
        End If
        
        
    Next
End Function

Private Sub CLEARDATA()
        M_RINM = Empty
        TXTPER = Empty
        TXTMRGN = Empty
End Sub

Private Function CheckData(RNO As Long) As Boolean
Dim INDEX As Long
    If Trim(M_RINM) = Empty Then
        MsgBox "Please Select Items From List !!", vbInformation
        M_RINM.SetFocus
        CheckData = True
        Exit Function
    End If
    
     If Not IsNumeric(TXTPER) Then
        MsgBox "Please Enter Valid Percentage !!", vbInformation, "Invalid Percentage !"
        TXTPER.SetFocus
        CheckData = True
        Exit Function
     End If
            
     If Val(TXTPER) <= 0 Then
        MsgBox "Please Enter Valid Percentage !!", vbInformation, "Percentage Missing !!"
        TXTPER.SetFocus
        CheckData = True
        Exit Function
    End If
    
      If BOXREQ = "Y" Then
      If TXTMRGN = Empty Then
         MsgBox "Please Enter Merge No.", vbOKOnly
         TXTMRGN.SetFocus
         CheckData = True
        Exit Function
      End If
      End If
      
      
    For INDEX = 1 To FLEX.Rows - 1
        If Trim(FLEX.TextMatrix(INDEX, 1)) = M_RINM And (Not SWITCH Or (SWITCH And INDEX <> RNO)) Then
           MsgBox "Raw Material Already Exist"
           M_RINM.SetFocus
           CheckData = True
           Exit Function
        End If
    Next INDEX
End Function

Private Sub FILLFLEX()
On Error GoTo LAST
Dim i As Double:   i = 0
Dim RECSTAT  As ADODB.Recordset
Set RECSTAT = New ADODB.Recordset

If RECSTAT.State = 1 Then RECSTAT.Close
RECSTAT.Open "SELECT * FROM TXULOT WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DVCD='" & DIVCODE & "' AND LTNO = '" & M_LTNO & "'", CN, adOpenDynamic, adLockOptimistic
If Not RECSTAT.EOF Then
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM FINITMMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
          "' AND DVCD='" & DIVCODE & "' AND CODE='" & RECSTAT!FICD & "'", CN, adOpenKeyset, adLockPessimistic
  If Not RS.EOF Then
    M_FINM = RS!NAME
  Else
    If M_FINM.Enabled Then M_FINM.SetFocus
    Exit Sub
  End If
  
  M_SRCH = Trim(RECSTAT!SRCH)
  M_RINM = GetCode("ITMMST", Trim(RECSTAT!RICD & ""), "CODE", "NAME")
  SUBPKGNG = GetCode("SUBPKGNGMST", Trim(RECSTAT!SUBPKGCODE & ""), "CODE", "NAME")
  If Not IsNull(RECSTAT!SHCD) Then
     M_SHCD = GetCode("GRDMST", Trim(RECSTAT!SHCD), "CODE", "GRAD")
     TXTSHCD = M_SHCD
  End If
  
  If RECSTAT!ACTIVE = "Y" Then
     optActive.Value = True
  Else
     optDeactive.Value = True
  End If
  
  TXTPER = nstr(RECSTAT!PERC, 4, 2)
  Do While Not RECSTAT.EOF
     i = i + 1
     If i >= FLEX.Rows - 1 Then
       FLEX.Rows = FLEX.Rows + 1
     End If
       
     FLEX.TextMatrix(i, 0) = Trim(RECSTAT!SRCH)
     FLEX.TextMatrix(i, 1) = GetCode("ITMMST", Trim(RECSTAT!RICD), "CODE", "NAME")
     FLEX.TextMatrix(i, 2) = Format(STR(RECSTAT!PERC), "00.00")
     FLEX.TextMatrix(i, 3) = Trim(RECSTAT!MRGN & "")
     
     RECSTAT.MoveNext
     If FLEX.Rows > 6 Then FLEX.TopRow = FLEX.TopRow + 2
    Loop
    
    Call SetPercentage
    FLEX.Rows = FLEX.Rows - 1
    btn_sts (True)
    SWITCH = False
    ROWNO = 1
    Call Flex_Click
End If

Exit Sub
LAST:
MsgBox ERR.Description
End Sub

Private Sub SetPercentage()
Dim i As Long
Dim PER As Double: PER = 0

For i = 1 To FLEX.Rows - 1
  PER = PER + Val(FLEX.TextMatrix(i, 2))
Next i

lblPer.Caption = CStr(PER) + " %"

End Sub

Private Function ChkPercentage(SWITCH As Boolean) As Boolean
Dim i As Long
Dim PER As Double: PER = 0

For i = 1 To FLEX.Rows - 1
  PER = PER + Val(FLEX.TextMatrix(i, 2))
Next i

If Not SWITCH Then
  PER = PER + Val(TXTPER)
Else
  PER = PER + Val(TXTPER) - Val(FLEX.TextMatrix(ROWNO, 2))
End If

If PER > 100 Then
   ChkPercentage = True
End If

End Function

Private Sub SUBPKGNG_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(SUBPKGNG.Text) = Empty Or KeyCode = vbKeyF2 Then
    SUBPKGNG.Text = SearchList1("select TOP 20  CODE,NAME from SUBPKGNGMST WHERE RECSTAT <> 'D'", 0, SUBPKGNG.Text, "SELECT SUB PACKAGING TYPE FROM MASTER")
    End If
  If KeyCode = vbKeyDelete Then SUBPKGNG.Text = Empty

End Sub

Private Sub TXTMRGN_GotFocus()
TXTMRGN.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTMRGN_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        TXTMRGN = Empty
    ElseIf KeyCode = vbKeyF2 Or TXTMRGN = Empty Then
        M_DESC = Empty
        Key = Empty
        NEW_VISIBLE = False
        TXTMRGN = SearchList1("Select DISTINCT MRGN,MRGN  From MRGMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND ICOD = '" & GetCode("ITMMST", M_RINM, "NAME", "CODE") & "'", 0, Empty, "Select MERGE FROM MASTER")
        'Me.Tag = Key
        'MERGE = Key
    End If

End Sub

Private Sub TXTMRGN_LostFocus()
TXTMRGN.BackColor = vbWhite
End Sub

Private Sub txtPer_GotFocus()
 TXTPER.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPER_KeyPress(KeyAscii As Integer)
 If CheckNumericKey(KeyAscii, TXTPER, Me) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPer_LostFocus()
 TXTPER.BackColor = vbWhite
End Sub

Private Sub TXTSHCD_GotFocus()
TXTSHCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSHCD_KeyDown(KeyCode As Integer, Shift As Integer)
  If (TXTSHCD = Empty And KeyCode = vbKeyReturn) Or KeyCode = vbKeyF2 Then
    NEW_VISIBLE = True
    TXTSHCD = SearchList1("SELECT DISTINCT GRAD AS GRD,GRAD FROM GRDMST", 0, TXTSHCD, "SELECT " & LBLSHCD.Caption)
      If key_PressNew = True Then
          M_DESC = ""
          TXTSHCD = Empty
          FRM_GRDMST.Show
      End If
  End If
End Sub

Private Sub TXTSHCD_LostFocus()
TXTSHCD.BackColor = vbWhite
End Sub
