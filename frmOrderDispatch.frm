VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmOrderDispatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dispatch Against Order"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   10260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6131.778
   ScaleMode       =   0  'User
   ScaleWidth      =   27391.85
   Begin FramePlusCtl.FramePlus FrmAutoConsumption 
      Height          =   6195
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10927
      BackgroundPictureAlignment=   5
      BorderStyle     =   10
      BackColorGradient=   12632319
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
      Begin VB.ComboBox cmbPackingType 
         BackColor       =   &H0080C0FF&
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
         Height          =   315
         ItemData        =   "frmOrderDispatch.frx":0000
         Left            =   2040
         List            =   "frmOrderDispatch.frx":0002
         TabIndex        =   5
         Tag             =   "0"
         Text            =   "Select Type of Dispatch"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtCONSINEE 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox txtPCOD 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   5175
      End
      Begin VB.TextBox TXTADDRESS 
         Height          =   525
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3120
         Width           =   5175
      End
      Begin VB.TextBox TXTRMRK 
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
         Left            =   240
         MaxLength       =   50
         TabIndex        =   17
         Top             =   4560
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker TXTVBDT 
         Height          =   405
         Left            =   8280
         TabIndex        =   7
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
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
         Format          =   54591489
         CurrentDate     =   39347
      End
      Begin MSComctlLib.ListView lstOrdn 
         Height          =   3855
         Left            =   5760
         TabIndex        =   15
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   6800
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order No."
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pallet No."
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qnty."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DBCD"
            Object.Width           =   0
         EndProperty
      End
      Begin WelchButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2880
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
         Image           =   "frmOrderDispatch.frx":0004
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4320
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
         Image           =   "frmOrderDispatch.frx":0D8E
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   5760
         TabIndex        =   3
         Top             =   5400
         Width           =   975
         _ExtentX        =   1720
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
         Image           =   "frmOrderDispatch.frx":11E0
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdAdd 
         Height          =   495
         Left            =   1440
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
         Image           =   "frmOrderDispatch.frx":1632
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdSearch 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Search"
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
         Image           =   "frmOrderDispatch.frx":19CC
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   735
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   9135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Dispatch :"
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
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   5640
         X2              =   5640
         Y1              =   1200
         Y2              =   5280
      End
      Begin VB.Shape BORDER3 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   6000
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks "
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
         TabIndex        =   16
         Tag             =   "S"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   4095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   9855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C Party"
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
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Address"
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
         TabIndex        =   12
         Tag             =   "S"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Consignee Name"
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
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carton Name"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Tag             =   "S"
         Top             =   -2040
         Width           =   1335
      End
      Begin VB.Label LBLHEADING1 
         BackStyle       =   0  'Transparent
         Caption         =   "Division :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape BORDER1 
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label LBLDIV 
         BackStyle       =   0  'Transparent
         Caption         =   "XXXXXXXXX"
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
         Left            =   1320
         TabIndex        =   21
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label LBLCHDT 
         BackStyle       =   0  'Transparent
         Caption         =   "Export Challan Date :"
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
         Left            =   6120
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label LBLCHLN 
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
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LBLHEAD 
         BackStyle       =   0  'Transparent
         Caption         =   "Export Challan No ."
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
         Left            =   6120
         TabIndex        =   19
         Tag             =   "0"
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmOrderDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CHLNTYP As String
Dim ALLOWEDITDEL As Boolean
Dim DIVCODE As String
Dim SAVEFLAG As Boolean
Dim ORDREQ As Boolean
Dim TTLCOPS As Long
Dim TTLBOXES As Long
Dim TTLQTY As Double, TTLGRSWGT As Double, TTLTAREWGT As Double
Dim SPARTY As String, SCONSINEE As String, SADD As String, grad As String, SUBGRD As String, SGRD As String, SITEM As String
Public VTCD As String
Public ORDN As String
Public DONO As String
Public M_DBCD As String
Public chln As String


Private Sub cmbPackingType_Click()
If InStr(1, UCase(cmbPackingType.Text), "EXPORT") <> 0 Then
   Me.Caption = "Box Dispatch (Export Challan) "
   LBLHEAD = "Export Challan No ."
   LBLCHDT = "Export Challan Date :"
End If

   TXTVBDT = Now
   Call SetGlobal
   LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
End Sub

Private Sub cmbPackingType_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmdadd_Click()
 zoomflag = False
 btn_sts (False)
 SAVEFLAG = True
 If cmbPackingType.Enabled Then cmbPackingType.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Call ClsData(Me)
    Call btn_sts(True)
    lstOrdn.ListItems.Clear
    If zoomflag = True Then
        Call cmdExit_Click
        Exit Sub
    End If
    TXTVBDT = Now
    TXTVBDT.Enabled = True
    LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)
    If cmbPackingType.ListCount > 1 Then cmbPackingType.ListIndex = 0
    If cmdExit.Enabled Then cmdExit.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo LAST
Dim DONO As String
Dim Index As Long
Dim FLAG As Boolean  'OK
Dim SLIP As String
Dim COPS As Double
Dim PCS As Double
Dim SQL As String
Dim TEMPRS As ADODB.Recordset
Set TEMPRS = New ADODB.Recordset

'TO FIND ATLEAST ONE SELECTION
FLAG = False
For Index = 1 To lstOrdn.ListItems.COUNT
  If lstOrdn.ListItems(Index).Checked = True Then: FLAG = True: Exit For
Next

If FLAG = False Then Exit Sub

Dim SALMANCOD As String
For Index = 1 To lstOrdn.ListItems.COUNT
If lstOrdn.ListItems(Index).Checked = True Then
     If SALMANCOD = Empty Then
        SALMANCOD = Trim(lstOrdn.ListItems(Index).SubItems(3))
     Else
        If SALMANCOD <> Trim(lstOrdn.ListItems(Index).SubItems(3)) Then
           MsgBox "Multiple Sales Man including in this dispatch.", vbCritical
           lstOrdn.SetFocus
           Exit Sub
        End If
     End If
End If
Next

'TO FIND NOT BLAND DATA
If Not CHKSAVEDATA Then Exit Sub

'TO FIND CODE
Call SetGlobal

CN.BeginTrans

If SAVEFLAG Then ''INSERT MODE

LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE) 'GENERATE CHALLAN

If TEMPRS.State = 1 Then TEMPRS.Close
TEMPRS.Open "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
            "' AND DVCD='" & DIVCODE & "" & "' AND DBCD='" & VTCD & _
            "' AND VTYP='DPF' AND VBNO='" & LBLCHLN & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
If Not TEMPRS.EOF Then
   MsgBox "CHALLAN ALREADY EXIST : CHECK UNIT CONFIGURATION FOR LAST CHALLAN NO."
   CN.RollbackTrans
   Exit Sub
End If

Dim ORDN As String, PALLETNO As String, CHK As Long
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset

'FIRST SET ALL BOX AS DISPATCH ON BASIS OF PALLET NO WHICH IS SELECTED BY USER
For Index = 1 To lstOrdn.ListItems.COUNT
  If lstOrdn.ListItems(Index).Checked = True Then
     PALLETNO = lstOrdn.ListItems(Index).SubItems(1)
          
     SQL = "UPDATE BOXREGISTER SET VTYP='DPF',RVBNO='" & LBLCHLN & "',RVBDT= '" & Format(TXTVBDT, "MM/DD/YYYY") & _
     "',RDBC = '" & VTCD & "',RVTYP='DPF' WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
     "' AND (VTYP='PPF' OR VTYP='OPN') AND PLTNO ='" & PALLETNO & "' AND RECSTAT<>'D' AND ORDN IS NOT NULL"
     
     CN.Execute SQL, CHK
 
  End If
Next

'FIND FROM UNIQUE COMBINATION OF 3 : ORDN,LOTNO,GRAD BECAUSE OF SINGLE TIME RAW MATERIAL CONSUMPTION
     If TEMPRS.State = 1 Then TEMPRS.Close
     TEMPRS.Open "SELECT ORDN,LOTNO,ICOD,GRAD,COUNT(PLTNO) AS PALLET,SUM(GRSWGT) AS GRSWGT," & _
                 "SUM(TRWGT) AS TRWGT,SUM(NTWGT) AS NTWGT,ISNULL(SUM(ISNULL(COPS,0)),0) AS COPS FROM BOXREGISTER WHERE COMP='" & compPth & _
                 "' AND UNIT='" & UNCD & "' AND VTYP='DPF' AND RECSTAT<>'D' AND RVBNO='" & LBLCHLN & _
                 "' AND RDBC = '" & VTCD & "' AND ORDN IS NOT NULL GROUP BY ORDN,LOTNO,ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
     
     Dim SRCH As Long: SRCH = 0
     Do While Not TEMPRS.EOF
         
        'TO FIND EXCHANGE RATE FOR EXPORT ORDER
         Dim EXRATE As Double: EXRATE = 1
         SQL = "SELECT EXRAT FROM EXPORD WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
               "' AND ORDN='" & TEMPRS!ORDN & "' "
                     
         If RS.State = 1 Then RS.Close
         RS.Open SQL, CN, adOpenDynamic, adLockOptimistic
         If Not RS.EOF Then
            If Val(RS!EXRAT) <> 0 Then
               EXRATE = nstr(RS!EXRAT, 7, 2)
            End If
         End If
         RS.Close
     
        Dim SUBRS As ADODB.Recordset
        Set SUBRS = New ADODB.Recordset
        Dim ORDRATE As Double: ORDRATE = 0
        Dim DollarRate As Double: DollarRate = 0
        Dim BRCD As String, TXCD As String, RTCD As String
        
        If SUBRS.State = 1 Then SUBRS.Close
        SUBRS.Open "SELECT BRCD,RATE,TXCD,RTCD FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                   "' AND ORDN='" & TEMPRS!ORDN & "" & "' AND ICOD='" & TEMPRS!ICOD & "" & _
                   "' AND TRCD='" & TEMPRS!grad & "' AND RECSTAT<>'D'", CN, adOpenDynamic, adLockOptimistic
        If Not SUBRS.EOF Then
           DollarRate = Val(SUBRS!RATE)
           ORDRATE = Val(SUBRS!RATE) * EXRATE
           BRCD = Trim(SUBRS!BRCD & "")
           TXCD = Trim(SUBRS!TXCD & "")
           RTCD = Trim(SUBRS!RTCD & "")
        End If
                    
        SRCH = SRCH + 1
     
        SALMANCOD = FindSalesManCode(Trim(TEMPRS!ORDN & ""))
        DONO = GenDONO(Trim(TEMPRS!ORDN & ""))
        
        SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD," & _
        "QNTY,RATE,ARAT,ORDN,OSRC,BRMK,PRDL,ICOD,RTCD,TXRT,TXCD,SLIP,SLIPDATE,RDBC,DELQNTY,DFLG,DOAPRVBY) VALUES ('" & compPth & _
        "','" & UNCD & "','" & DIVCODE & "','DOS','" & SALMANCOD & "','" & DONO & _
        "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SPARTY & "','" & SCONSINEE & _
        "','" & SRCH & "','" & BRCD & "','" & Trim(TEMPRS!LOTNO & "") & _
        "','" & Trim(TEMPRS!grad & "") & "','0','" & Val(TEMPRS!NTWGT & "") & "'," & ORDRATE & "," & ORDRATE & _
        ",'" & Trim(TEMPRS!ORDN & "") & "','" & SRCH & "','" & Trim(TXTRMRK) & "','','" & Trim(TEMPRS!ICOD & "") & _
        "','" & RTCD & "','" & RTCD & "','" & TXCD & _
        "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & VTCD & "','" & Val(TEMPRS!NTWGT & "") & "','Y','ADMIN')"
        
        CN.Execute SQL
        
        SQL = "INSERT INTO ORDTRN (COMP,UNIT, DVCD,VTYP,DBCD,DONO,DODT,PCOD,DCOD,SRCH,BRCD,LTNO,GRAD,SUBGRD," & _
        "QNTY,RATE,ARAT,ORDN,OSRC,BRMK,PRDL,ICOD,RTCD,TXRT,TXCD,SLIP,SLIPDATE,RDBC,DELQNTY,DFLG) VALUES ('" & compPth & _
        "','" & UNCD & "','" & DIVCODE & "','DPF','" & SALMANCOD & "','" & DONO & _
        "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & SPARTY & "','" & SCONSINEE & _
        "','" & SRCH & "','" & BRCD & "','" & Trim(TEMPRS!LOTNO & "") & _
        "','" & Trim(TEMPRS!grad & "") & "','0','" & Val(TEMPRS!NTWGT & "") & "'," & ORDRATE & "," & ORDRATE & _
        ",'" & Trim(TEMPRS!ORDN & "") & "','" & SRCH & "','" & Trim(TXTRMRK) & "','','" & Trim(TEMPRS!ICOD & "") & _
        "','" & RTCD & "','" & RTCD & "','" & TXCD & _
        "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','" & VTCD & "','" & Val(TEMPRS!NTWGT & "") & "','Y')"
        
        CN.Execute SQL
                                    
        SQL = "INSERT INTO SPTRAN(COMP,UNIT,DVCD,VTYP,DBCD,VBNO,SRCH,CHLN,CHDT,DATE,CRAC,DRAC,PCOD,BRCD,"
        SQL = SQL & "DCOD,ADDRESS,TXCD,RTCD,TXRT,LTNO,ICOD,GRAD,SUBGRD,PCES,QNTY,GWGT,TWGT,EXRATE,RATE,AMNT,QORP,[SYSR],[USER],OPER,"
        SQL = SQL & "COPS,RECSTAT,EXTRA1,EXTRA4)VALUES('" & compPth & "','" & UNCD & "','" & DIVCODE & _
        "','DPF','" & VTCD & "','" & LBLCHLN & "','" & SRCH & "','" & LBLCHLN & "','" & Format(TXTVBDT, "YYYY/MM/DD") & _
        "','" & Format(TXTVBDT, "YYYY/MM/DD") & "','XXXXXX','" & SPARTY & "','" & SPARTY & "','" & BRCD & "','" & SCONSINEE & _
        "','" & SADD & "','" & TXCD & "','" & RTCD & "','RETAIL INVOICE','" & Trim(TEMPRS!LOTNO & "") & "','" & Trim(TEMPRS!ICOD & "") & _
        "','" & Trim(TEMPRS!grad & "") & "','0','" & Val(TEMPRS!PALLET & "") & _
        "','" & Val(TEMPRS!NTWGT & "") & "','" & Val(TEMPRS!GRSWGT & "") & _
        "','" & Val(TEMPRS!TRWGT & "") & "'," & DollarRate & "," & ORDRATE & "," & (Val(TEMPRS!NTWGT & "") * ORDRATE) & _
        ",'Q','N','" & cUName & "','-','" & Val(TEMPRS!COPS & "") & "','A','" & Trim(TEMPRS!ORDN & "") & _
        "','" & Trim(TXTRMRK) & "')"
        
        CN.Execute SQL
                
        'UPDATE DISPATCH QTY AT ORDER TABLE
        SQL = "UPDATE ORDMAN SET DISPATCHQTY = DISPATCHQTY + " & Val(TEMPRS!NTWGT & "") & _
              ",DOQTY = DOQTY - " & Val(TEMPRS!NTWGT & "") & _
              " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DCOD='" & DIVCODE & _
              "' AND ORDN = '" & Trim(TEMPRS!ORDN & "") & "' AND ICOD = '" & Trim(TEMPRS!ICOD & "") & _
              "' AND TRCD='" & Trim(TEMPRS!grad & "") & "'"
            
      CN.Execute SQL
     
     TEMPRS.MoveNext
     Loop
     TEMPRS.Close
     
     Dim M_ORDN As String
     'FOR PRINTING PURPOSE
     If TEMPRS.State = 1 Then TEMPRS.Close
     TEMPRS.Open "SELECT DISTINCT ORDN FROM BOXREGISTER WHERE COMP='" & compPth & _
                 "' AND UNIT='" & UNCD & "' AND VTYP='DPF' AND RECSTAT<>'D' AND RVBNO='" & LBLCHLN & _
                 "' AND RDBC = '" & VTCD & "' AND ORDN IS NOT NULL GROUP BY ORDN,LOTNO,ICOD,GRAD", CN, adOpenDynamic, adLockOptimistic
                 
     CN.Execute "UPDATE SPTRAN SET ORDN = '' WHERE ORDN IS NULL"
        
     Do While Not TEMPRS.EOF
        If M_ORDN <> Empty Then M_ORDN = M_ORDN & ","
        M_ORDN = M_ORDN & Trim(TEMPRS!ORDN & "")
     TEMPRS.MoveNext
     Loop
     TEMPRS.Close
     
     CN.Execute "UPDATE SPTRAN SET ORDN = LTRIM(RTRIM(ORDN)) + '" & M_ORDN & "' WHERE COMP='" & compPth & _
                 "' AND UNIT='" & UNCD & "' AND VTYP='DPF' AND DBCD = '" & VTCD & "' AND VBNO='" & LBLCHLN & _
                 "' AND RECSTAT<>'D' ", NOR
     
'-------------------
'UPDATE SERIAL
Dim UPSQL As String
UPSQL = "UPDATE SERIALMASTER SET [SRNO]='" & LBLCHLN & "' WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & _
"' AND VTYP='DPF' AND CODE='" & VTCD & "' AND FYCD='" & FYCD & "' "

If UNT_DIVSERIES_REQ = "Y" Then
   UPSQL = UPSQL & " AND DVCD='" & DIVCODE & "' "
End If
CN.Execute UPSQL
'=========================

Call DAILYSTATUS("DPF", GetCode("ACCMST", txtPCOD, "NAME", "CODE"), VTCD, Val(txtNTWT), LBLCHLN, 0, cUName, "N", Now, TXTVBDT)

CN.CommitTrans

End If

If SAVEFLAG Then
   MsgBox "Your Challan No. is : " & LBLCHLN
End If

lstOrdn.ListItems.Clear
Call CLEARDATA
Call cmdCancel_Click

TXTVBDT = Now

Exit Sub
LAST:
MsgBox ERR.Description
CN.RollbackTrans
Exit Sub
End Sub

Private Sub cmdSearch_Click()

Dim PKGRS As ADODB.Recordset
Set PKGRS = New ADODB.Recordset

If Trim(txtPCOD) <> Empty Then
   txtPCOD.Tag = GetCode("ACCMST", txtPCOD, "NAME", "CODE")
End If

Dim SRCRS As ADODB.Recordset
Set SRCRS = New ADODB.Recordset
If SRCRS.State = 1 Then SRCRS.Close
SRCRS.Open "SELECT DISTINCT ORDN,DBCD FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
           "' AND RECSTAT<>'D' AND PCOD='" & txtPCOD.Tag & "' AND OFLG<>'Y' AND FIN_APRV = 'O' AND DOQTY > 0", CN, adOpenDynamic, adLockOptimistic
           
lstOrdn.ListItems.Clear
Do While Not SRCRS.EOF
   
   If PKGRS.State = 1 Then PKGRS.Close
   PKGRS.Open "SELECT ORDN,PLTNO,SUM(NTWGT) AS QTY FROM BOXREGISTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
              "' AND (VTYP='PPF' OR VTYP='OPN') AND DVCD = '" & DIVCODE & "' AND ORDN='" & SRCRS!ORDN & _
              "' AND VBDT <= '" & Format(TXTVBDT.Value, "MM/DD/YYYY") & "' AND RECSTAT<>'D' GROUP BY ORDN,PLTNO", CN, adOpenDynamic, adLockOptimistic
   Do While Not PKGRS.EOF
        Set Item = lstOrdn.ListItems.ADD
        Item.Text = Trim(PKGRS!ORDN & "")
        Item.SubItems(1) = Trim(PKGRS!PLTNO & "")
        Item.SubItems(2) = nstr(Val(PKGRS!QTY), 8, 3)
        Item.SubItems(3) = Trim(SRCRS!dbcd & "")
   PKGRS.MoveNext
   Loop
   PKGRS.Close
   
SRCRS.MoveNext
Loop
SRCRS.Close

If lstOrdn.ListItems.COUNT > 0 Then
  lstOrdn.SetFocus
End If

End Sub

Private Sub Form_Activate()
If DIVCODE = Empty Or LBLDIV = Empty Then
  Unload Me
End If

Me.BackColor = RGB(RED, GREEN, BLUE)
If zoomflag = True Then
   btn_sts (False)
   SAVEFLAG = False
Else
   btn_sts (True)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call CenterChild(frm_Main, Me)
Call ColorComponent(Me)
TXTADDRESS.FontBold = False
Me.Left = 50: Me.KeyPreview = True
SAVEFLAG = True
  M_DESC = Empty: Key = Empty: NEW_VISIBLE = False
  DIVCODE = Empty
  If DIVCODE = Empty Then
    LBLDIV.Caption = SearchList1("SELECT  TOP 20 CODE,NAME FROM DIVMST WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND CODE<>'000001' AND RECSTAT='A'", 0, "", "SELECT DIVISION MASTER")
    DIVCODE = Key
  End If
    
  Call SetPackingType
  TXTVBDT = Now
  LBLCHLN = GenDPFVNO("DPF", VTCD, DIVCODE)

End Sub

Private Sub SetPackingType()
Dim PKTYPRS As ADODB.Recordset
Set PKTYPRS = New ADODB.Recordset
Dim SQL As String
SQL = "SELECT DISTINCT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME NOT LIKE '%CAPTIVE%' AND NAME NOT LIKE '%WASTAGE%' AND FYCD='" & FYCD & "'  AND NAME<>''"

If PKTYPRS.State = 1 Then PKTYPRS.Close
PKTYPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic

If Not PKTYPRS.EOF Then VTCD = Trim(PKTYPRS!CODE)

Do While Not PKTYPRS.EOF
 cmbPackingType.AddItem Trim(PKTYPRS!NAME)
 PKTYPRS.MoveNext
Loop
If cmbPackingType.ListCount > 0 Then cmbPackingType.ListIndex = 1
End Sub

Private Sub txtCONSINEE_KeyDown(KeyCode As Integer, Shift As Integer)
 Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
     txtCONSINEE = Empty
  ElseIf KeyCode = vbKeyF2 Or txtCONSINEE = Empty Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtCONSINEE = SearchList1("Select DISTINCT CODE,NAME From PADDMST", 0, Empty, "Select Consinee Name ")
  End If
  
 Me.KeyPreview = True

End Sub

Private Sub txtPCOD_GotFocus()
  txtPCOD.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub txtConsinee_GotFocus()
  txtCONSINEE.BackColor = RGB(BRED, BGREEN, BBLUE)
  SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTADDRESS_GotFocus()
  TXTADDRESS.BackColor = RGB(BRED, BGREEN, BBLUE)
   SendKeys "{HOME}+{END}"
End Sub

Private Sub TXTADDRESS_KeyDown(KeyCode As Integer, Shift As Integer)
   TXTADDRESS.FontSize = 8
   If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
    TXTADDRESS = Empty
   ElseIf KeyCode = vbKeyF2 Or (TXTADDRESS = Empty And KeyCode = vbKeyReturn) Then
    TXTADDRESS = SearchAddress("Select SRNO,ADDR From PADDMST WHERE NAME='" & txtCONSINEE & "'", 0, Empty, "Select Consignee Address from List")
   End If
End Sub

Private Sub txtPCOD_KeyDown(KeyCode As Integer, Shift As Integer)
  Me.KeyPreview = False
    
  If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
      txtPCOD = Empty
  ElseIf KeyCode = vbKeyF2 Or (KeyCode = vbKeyReturn And txtPCOD = Empty) Then
     M_DESC = Empty:   NEW_VISIBLE = False
     txtPCOD = SearchList1("Select TOP 20 Code,Name From ACCMST", 0, Empty, "Select A/c Party ")
  End If
  
  Me.KeyPreview = True

End Sub


Private Sub TXTRMRK_GotFocus(): TXTRMRK.BackColor = RGB(BRED, BGREEN, BBLUE): End Sub

Private Sub txtPCOD_LostFocus()
 txtPCOD.BackColor = vbWhite
End Sub
Private Sub txtConsinee_LostFocus(): txtCONSINEE.BackColor = vbWhite: End Sub
Private Sub TXTADDRESS_LostFocus(): TXTADDRESS.BackColor = vbWhite: End Sub
Private Sub TXTRMRK_LostFocus(): TXTRMRK.BackColor = vbWhite: End Sub

Private Sub CLEARDATA()
 txtPCOD = Empty: txtCONSINEE = Empty: TXTADDRESS = Empty: txtItem = Empty
 TXTRMRK = Empty
End Sub
Private Sub cmbPackingType_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub
Private Sub TXTVBDT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Public Sub btn_sts(Yes As Boolean)
 cmdSave.Enabled = Not Yes: cmdCancel.Enabled = Not Yes: cmdAdd.Enabled = Yes
 txtPCOD.Enabled = Not Yes: txtCONSINEE.Enabled = Not Yes: TXTADDRESS.Enabled = Not Yes
 TXTRMRK.Enabled = Not Yes
End Sub

Private Sub SetGlobal()
Dim Index As Long
TTLQTY = 0: TTLBOXES = 0: TTLCOPS = 0
TTLGRSWGT = 0: TTLTAREWGT = 0
Dim DBCDRS As ADODB.Recordset
Set DBCDRS = New ADODB.Recordset
If DBCDRS.State = 1 Then DBCDRS.Close
DBCDRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
"' AND VTYP='DPF' AND NAME = '" & cmbPackingType.Text & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
If Not DBCDRS.EOF Then
   VTCD = Trim(DBCDRS!CODE & "")
Else
   VTCD = Empty
End If
DBCDRS.Close

SPARTY = GetCode("ACCMST", txtPCOD, "NAME", "CODE")
SCONSINEE = GetCode("PADDMST", txtCONSINEE, "NAME", "CODE")

Dim RSDATA As ADODB.Recordset
Set RSDATA = New ADODB.Recordset

If RSDATA.State = 1 Then RSDATA.Close
RSDATA.Open "SELECT SRNO FROM PADDMST WHERE CODE='" & SCONSINEE & "' AND ADDR='" & TXTADDRESS & "'", CN, adOpenDynamic, adLockOptimistic
If Not RSDATA.EOF Then
  SADD = RSDATA!SRNO
Else
  SADD = Empty
End If
RSDATA.Close

End Sub

Private Function CHKSAVEDATA() As Boolean
  CHKSAVEDATA = True
  If SAVEFLAG And (Trim(txtPCOD) = Empty Or Trim(txtCONSINEE) = Empty Or Trim(TXTADDRESS) = Empty) Then
     If txtPCOD = Empty Then txtPCOD.Enabled = True: txtPCOD.SetFocus: CHKSAVEDATA = False: Exit Function
     If txtCONSINEE = Empty Then txtCONSINEE.Enabled = True: txtCONSINEE.SetFocus: CHKSAVEDATA = False: Exit Function
     If TXTADDRESS = Empty Then TXTADDRESS.Enabled = True: TXTADDRESS.SetFocus: CHKSAVEDATA = False: Exit Function
  End If
End Function

Private Function GenDONO(ORDN As String) As String
Dim DORS As ADODB.Recordset
Set DORS = New ADODB.Recordset
Dim NO As Double

If DORS.State = 1 Then DORS.Close
DORS.Open "SELECT ISNULL(MAX(RIGHT(DONO,4)),0) AS DONUM FROM ORDTRN WHERE COMP='" & compPth & _
          "' AND UNIT ='" & UNCD & "' AND ORDN='" & ORDN & "' ", CN, adOpenDynamic, adLockOptimistic

NO = Val(DORS!DONUM)
NO = NO + 1
DORS.Close
        
If NO < 10 Then
   GenDONO = "000" + Trim(nstr(NO, 1, 0))
ElseIf NO < 100 Then
   GenDONO = "00" + Trim(nstr(NO, 1, 0))
ElseIf NO < 1000 Then
   GenDONO = "0" + Trim(nstr(NO, 1, 0))
ElseIf NO < 10000 Then
   GenDONO = Trim(nstr(NO, 1, 0))
End If
      
   GenDONO = Mid$(CStr(ORDN), 1, 6) & GenDONO

End Function

Private Function FindSalesManCode(ORDN As String)
FindSalesManCode = Empty

Dim DORS As ADODB.Recordset
Set DORS = New ADODB.Recordset
Dim NO As Double

If DORS.State = 1 Then DORS.Close
DORS.Open "SELECT DBCD FROM ORDMAN WHERE COMP='" & compPth & "' AND UNIT ='" & UNCD & "' AND ORDN='" & ORDN & "'", CN, adOpenDynamic, adLockOptimistic
If Not DORS.EOF Then
   FindSalesManCode = Trim(DORS!dbcd & "")
End If

End Function

