VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WelchButton.ocx"
Begin VB.Form frmIsValidData 
   Caption         =   "Is Valid Data ???"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000D&
      Caption         =   "Match BillMain And Sptran"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin WelchButton.lvButtons_H cmdSearch 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      Caption         =   "&Search Data Related Problems"
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
      Image           =   "frmIsValidData.frx":0000
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ListView lstBill 
      Height          =   5385
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9499
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bill No."
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmIsValidData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
Dim VALIDRS As ADODB.Recordset
Set VALIDRS = New ADODB.Recordset

Dim SQL As String

SQL = "SELECT BILLMAIN.VBNO AS BILL,SPTRAN.VBNO FROM BILLMAIN " & _
      "LEFT JOIN SPTRAN ON SPTRAN.COMP=BILLMAIN.COMP AND SPTRAN.UNIT=BILLMAIN.UNIT AND " & _
      "SPTRAN.VTYP = BILLMAIN.VTYP And SPTRAN.dbcd = BILLMAIN.dbcd And " & _
      "SPTRAN.VBNO = BILLMAIN.VBNO WHERE BILLMAIN.UNIT='000001' AND BILLMAIN.VTYP='SAL' " & _
      "AND BILLMAIN.RECSTAT<>'D' AND SPTRAN.VBNO IS NULL  "

If VALIDRS.State = 1 Then VALIDRS.Close
VALIDRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
Do While Not VALIDRS.EOF
   
   Set lstItem = lstBill.ListItems.ADD
   lstItem.Text = Trim(VALIDRS![BILL])
   
VALIDRS.MoveNext
Loop
VALIDRS.Close


End Sub
