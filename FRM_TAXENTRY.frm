VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form FRM_TAXENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAX ENTRY FORM"
   ClientHeight    =   7425
   ClientLeft      =   1815
   ClientTop       =   1230
   ClientWidth     =   8325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8325
   Begin VB.Frame Frame5 
      Caption         =   "Tax Form Information Related to Party"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   60
      TabIndex        =   21
      Top             =   5595
      Width           =   8205
      Begin VB.TextBox TXTFORM 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   720
         Width           =   3705
      End
      Begin VB.TextBox TXTPTYSHOW 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   6645
      End
      Begin MSComCtl2.DTPicker txtTaxDT 
         Height          =   330
         Left            =   1380
         TabIndex        =   25
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53215233
         CurrentDate     =   38429
      End
      Begin VB.Label Label6 
         Caption         =   "Form No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3285
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Rec. Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   24
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Party Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   330
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView LSTVW 
      Height          =   2550
      Left            =   90
      TabIndex        =   20
      Top             =   3030
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4498
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bill No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Challan No."
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Form Rec. Date"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "SRNO"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PCOD"
         Object.Width           =   18
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Party Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2310
      TabIndex        =   14
      Top             =   1440
      Width           =   5925
      Begin VB.TextBox TXTPTYNAME 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   555
         Width           =   4095
      End
      Begin VB.OptionButton OPTPTYPART 
         Alignment       =   1  'Right Justify
         Caption         =   "Particular"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   16
         Top             =   555
         Width           =   1365
      End
      Begin VB.OptionButton OPTPTYALL 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tax Form Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   60
      TabIndex        =   10
      Top             =   1425
      Width           =   2190
      Begin VB.OptionButton OPTTAXCLR 
         Alignment       =   1  'Right Justify
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   1605
      End
      Begin VB.OptionButton OPTTAXPEND 
         Alignment       =   1  'Right Justify
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   690
         Width           =   1605
      End
      Begin VB.OptionButton OPTTAXALL 
         Alignment       =   1  'Right Justify
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sales Tax Form Collection Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   8145
      Begin VB.TextBox txtTXCD 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   285
         Width           =   7185
      End
      Begin VB.TextBox TXTDBNAME 
         Height          =   285
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3645
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Form :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   285
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tax Form Date Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   8175
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   330
         Left            =   5385
         TabIndex        =   9
         Top             =   255
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53215233
         CurrentDate     =   38429
      End
      Begin MSComCtl2.DTPicker txtFrDate 
         Height          =   330
         Left            =   2535
         TabIndex        =   7
         Top             =   255
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53215233
         CurrentDate     =   38429
      End
      Begin VB.Label Label3 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4335
         TabIndex        =   8
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         Top             =   255
         Width           =   1065
      End
   End
   Begin WelchButton.lvButtons_H cmdSave 
      Height          =   495
      Left            =   6000
      TabIndex        =   28
      Top             =   6840
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
      Image           =   "FRM_TAXENTRY.frx":0000
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H cmdExit 
      Height          =   495
      Left            =   7200
      TabIndex        =   29
      Top             =   6840
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
      Image           =   "FRM_TAXENTRY.frx":0D8A
      cBack           =   -2147483633
   End
   Begin WelchButton.lvButtons_H btnview 
      Height          =   495
      Left            =   2280
      TabIndex        =   30
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
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
      Image           =   "FRM_TAXENTRY.frx":11DC
      cBack           =   -2147483633
   End
   Begin VB.Label lblForm 
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   4680
      TabIndex        =   19
      Top             =   2565
      Width           =   3465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Form Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   18
      Top             =   2565
      Width           =   1245
   End
End
Attribute VB_Name = "FRM_TAXENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pCode As String
Dim DBCOD As String
Dim NEWSSQL As String
Dim SQL As String
Dim SSQL As String
Dim PARTYLIST As String
Dim lstItm As ListItem
Dim spara As String

Private Sub btnView_Click()
    SSQL = Empty
    If FRMPARA = "SAL" Then
        If OPTTAXALL.Value = True Then
            SSQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND (VTYP='" & FRMPARA & "' OR VTYP='DBN'  OR VTYP='OPC' ) AND RECSTAT<>'D'"
        ElseIf OPTTAXPEND.Value = True Then
            SSQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND (FORM ='' OR FORM IS NULL) AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND RECSTAT<>'D'"
        Else
            SSQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND FORM <> '' AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND RECSTAT<>'D'"
        End If
        If TXTPTYNAME <> "" Then SSQL = SSQL & " AND PCOD='" & pCode & "'"
        If txtTXCD <> "" Then SSQL = SSQL & " AND BILLMAIN.TXCD='" & txtTXCD.Tag & "'"
        SSQL = SSQL & " AND UNIT='" & UNCD & "'"
        SSQL = SSQL & "  ORDER BY COMP, VTYP, DATE, VBNO"
      Else
       If OPTTAXALL.Value = True Then
            SSQL = "SELECT * FROM PURMAN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND RECSTAT<>'D'"
        ElseIf OPTTAXPEND.Value = True Then
            SSQL = "SELECT * FROM PURMAN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND (FORM ='' OR FORM IS NULL) AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND RECSTAT<>'D'"
        Else
            SSQL = "SELECT * FROM PURMAN WHERE COMP='" & compPth & "' AND DATE>='" & Format(txtFrDate, "MM/dd/YYYY") & "' AND DATE<='" & Format(txtToDate, "mm/dd/yyyy") & "' AND FORM <> '' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND RECSTAT<>'D'"
        End If
        If TXTPTYNAME <> "" Then SSQL = SSQL & " AND PCOD='" & pCode & "'"
        If txtTXCD <> "" Then SSQL = SSQL & " AND PURMAN.TXCD='" & txtTXCD.Tag & "'"
        SSQL = SSQL & " AND UNIT='" & UNCD & "'"
        SSQL = SSQL & "  ORDER BY COMP, VTYP, DATE, VBNO"
    End If
    
    TAXLISTGEN (SSQL)
    
End Sub
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  On Error GoTo LAST
  Dim i As Integer
    
    If LSTVW.ListItems.COUNT < 1 Then
        MsgBox "No Bill Detail Found !! Please Select Proper Option and Try Again", vbInformation, "No Bill Found"
        btnview.SetFocus
        Exit Sub
    End If
    If MsgBox("Are You Sure ? Want To Save Changes You Have Made ??", vbQuestion + vbYesNo, "Update Form ") = VBNO Then
        TXTFORM.SetFocus
        Exit Sub
    End If
    If LSTVW.ListItems.COUNT > 0 Then
        If OPTPTYALL.Value = True Then
            SQL = ""
            If FRMPARA = "SAL" Then
              SQL = "UPDATE BILLMAIN SET FORM='" & Trim(TXTFORM) & "',FDAT='" & Format(txtTaxDT, "MM/dd/YYYY") & "' WHERE COMP='" & compPth & "'" & _
              " AND UNIT='" & UNCD & "'  AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND VBNO='" & LSTVW.SelectedItem.SubItems(1) & "' AND DBCD='" & LSTVW.SelectedItem.SubItems(6) & "'"
             Else
              SQL = "UPDATE PURMAN SET FORM='" & Trim(TXTFORM) & "',FDAT='" & Format(txtTaxDT, "MM/dd/YYYY") & "' WHERE COMP='" & compPth & "'" & _
              " AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND PSNO='" & LSTVW.SelectedItem.SubItems(1) & "' AND DBCD='" & LSTVW.SelectedItem.SubItems(6) & "'"
            End If
            CN.Execute SQL
            LSTVW.SelectedItem.Selected = False
        Else
            For i = 1 To LSTVW.ListItems.COUNT
                If LSTVW.ListItems(i).Selected = True Then
                    SQL = ""
                    If TXTFORM <> "" Then
                        If FRMPARA = "SAL" Then
                          SQL = "UPDATE BILLMAIN SET FORM='" & Trim(TXTFORM) & "',FDAT='" & Format(txtTaxDT, "MM/dd/YYYY") & "' WHERE COMP='" & compPth & "'" & _
                             " AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND VBNO='" & LSTVW.ListItems(i).SubItems(1) & "' AND DBCD='" & LSTVW.SelectedItem.SubItems(6) & "'"
                         Else
                          SQL = "UPDATE PURMAN SET FORM='" & Trim(TXTFORM) & "',FDAT='" & Format(txtTaxDT, "MM/dd/YYYY") & "' WHERE COMP='" & compPth & "'" & _
                             " AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND PSNO='" & LSTVW.ListItems(i).SubItems(1) & "' AND DBCD='" & LSTVW.ListItems(i).SubItems(6) & "'"
                        End If
                        LSTVW.ListItems(i).SubItems(5) = txtTaxDT
                    Else
                        If FRMPARA = "SAL" Then
                          SQL = "UPDATE BILLMAIN SET FORM=NULL,FDAT=NULL WHERE COMP='" & compPth & "'" & _
                             " AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND VBNO='" & LSTVW.ListItems(i).SubItems(1) & "' AND DBCD='" & LSTVW.ListItems(i).SubItems(6) & "'"
                         Else
                          SQL = "UPDATE PURMAN SET FORM=NULL,FDAT=NULL WHERE COMP='" & compPth & "'" & _
                             " AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND PSNO='" & LSTVW.ListItems(i).SubItems(1) & "' AND DBCD='" & LSTVW.ListItems(i).SubItems(6) & "'"
                        End If
                        LSTVW.ListItems(i).SubItems(5) = ""
                    End If
                    CN.Execute SQL
                End If
                LSTVW.ListItems(i).Selected = False
            Next
Restart:
            If Not OPTTAXALL.Value = True Then
                For i = 1 To LSTVW.ListItems.COUNT
                    If (OPTTAXPEND.Value = True And LSTVW.ListItems(i).SubItems(5) <> "") Or (OPTTAXCLR.Value = True And LSTVW.ListItems(i).SubItems(5) = "") Then
                        If LSTVW.ListItems.COUNT > 0 Then LSTVW.ListItems.Remove i
                        GoTo Restart
                    End If
                Next
            End If
    End If
    LSTVW.SetFocus
    TXTPTYSHOW.Text = ""
    TXTFORM.Text = ""
End If

Exit Sub

LAST:
  MsgBox ERR.Description
  Resume
End Sub

Private Sub Form_Activate()
  Call ColorComponent(Me)
  Me.BackColor = RGB(RED, GREEN, BLUE)
  FRMPARA = Me.Tag
  If Allow_view_only = "Y" Then
     Unload Me
     Exit Sub
  End If

End Sub

Private Sub Form_Load()
Call ColorComponent(Me)
  Call CenterChild(frm_Main, Me)
  txtFrDate = FSDT
  txtToDate = FEDT
  txtTaxDT = Now
  TXTPTYNAME.Enabled = False
  cmdExit.Cancel = True
  Me.Tag = FRMPARA
End Sub

Private Sub LSTVW_Click()
    If LSTVW.ListItems.COUNT > 0 Then Call PARTYTAXVIEW
End Sub

Private Sub LSTVW_GotFocus()
 LSTVW.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub LSTVW_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If txtTaxDT.Visible = True Then
    txtTaxDT.SetFocus
    Exit Sub
  End If
End If
Call PARTYTAXVIEW
End Sub
Private Sub LSTVW_KeyUp(KeyCode As Integer, Shift As Integer)
Call PARTYTAXVIEW
End Sub

Private Sub LSTVW_LostFocus()
 LSTVW.BackColor = vbWhite
End Sub

Private Sub OPTPTYALL_Click()
    Call OPTPTYALL_KeyPress(0)
End Sub

Private Sub OPTPTYALL_KeyPress(KeyAscii As Integer)
 
 If OPTPTYALL = True Then
    LSTVW.MultiSelect = False
    TXTPTYNAME.Enabled = False
    TXTPTYNAME = Empty
    NEWSSQL = SSQL
    LSTVW.ListItems.Clear
    btnview.SetFocus
    If spara = "N" Then Exit Sub
 End If

End Sub

Private Sub OPTPTYPART_Click()
    Call OPTPTYPART_KeyDown(13, 0)
End Sub

Private Sub OPTPTYPART_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  TXTPTYNAME.Enabled = True
  TXTPTYNAME.SetFocus
End If
End Sub


Private Sub OPTTAXALL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub OPTTAXCLR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub OPTTAXPEND_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
 
End Sub

Private Sub TXTDBNAME_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If TXTDBNAME.Text = "" Or KeyCode = vbKeyF2 Then
    TXTDBNAME.Text = SearchList1("select  TOP 20 DBCD,name from daybok WHERE COMP='" & compPth & "' AND VTYP='" & FRMPARA & "' and unit='" & UNCD & "'", 0, "", "Select Day Book Name")
    DBCOD = Key
    SendKeys "{TAB}"
  End If
  
End Sub

Private Sub txtDBName_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 And TXTDBNAME.Text <> "" Then txtFrDate.SetFocus

End Sub

Private Sub TXTFORM_Change()
    If LSTVW.ListItems.COUNT < 1 Then Exit Sub
    If Len(TXTFORM) <> 0 Then cmdSave.Default = True Else cmdSave.Default = False
End Sub

Private Sub TXTFORM_GotFocus()
 TXTFORM.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTFORM_LostFocus()
 TXTFORM.BackColor = vbWhite
End Sub

Private Sub txtFrDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtToDate.SetFocus
End Sub

Private Sub txtPTYName_GotFocus()
 TXTPTYNAME.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtPTYName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CANCEL_VISIBLE = True
        TXTPTYNAME = Empty
          TXTPTYNAME.Text = SearchList1("select  TOP 20 code,name from accmst", 0, "", "List Of Party")
          pCode = Key
          If TXTPTYNAME <> "" Then
            LSTVW.MultiSelect = True
            LSTVW.ListItems.Clear
          End If
    End If

    If KeyCode = vbKeyReturn Then
        If OPTPTYPART.Value = True And TXTPTYNAME <> Empty Then LSTVW.MultiSelect = True
        btnview.SetFocus
    End If
End Sub

Private Sub TXTPTYNAME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If TXTPTYNAME.Text = "" Then
        txtPTYName_KeyDown vbKeyF2, 15
        btnview.SetFocus
       Else

      End If
    End If
End Sub

Private Sub txtPTYName_LostFocus()
 TXTPTYNAME.BackColor = vbWhite
  If TXTPTYNAME.Text = "" Then Exit Sub
    NEWSSQL = SSQL
      

    If spara = "N" Then Exit Sub


End Sub

Private Sub TXTPTYSHOW_GotFocus()
TXTPTYSHOW.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTPTYSHOW_LostFocus()
 TXTPTYSHOW.BackColor = vbWhite
End Sub

Private Sub TXTTAXDT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then TXTFORM.SetFocus
End Sub

Private Sub txtToDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
  
    If OPTTAXALL = True Then
       OPTTAXALL.SetFocus
    ElseIf OPTTAXPEND = True Then
       OPTTAXPEND.SetFocus
    Else
       OPTTAXCLR.SetFocus
    End If
  End If
End Sub

Public Function TAXLISTGEN(ByVal STR As String) As Boolean
Dim TEMPRS As New ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
LSTVW.ListItems.Clear

With TEMPRS

  If .State = adStateOpen Then .Close
  If STR = Empty Then Exit Function
  .Open STR, CN, adOpenDynamic, adLockOptimistic
  
  If .EOF = True Then
    MsgBox "No Tax Form Received !!", vbOKOnly + vbInformation
    If TXTDBNAME.Visible = True Then
      TXTDBNAME.SetFocus
    End If
    spara = "N"
    TAXLISTGEN = False
    Exit Function
  End If
  
  .MoveFirst
  
  While .EOF = False
  
  Set lstItm = LSTVW.ListItems.ADD
  
    lstItm.Text = Format(!Date, "dd/MM/yyyy")
    If FRMPARA = "SAL" Then
      lstItm.SubItems(1) = !VBNO
     Else
      lstItm.SubItems(1) = !psno
    End If
    If FRMPARA = "PRM" Then
      If Not IsNull(!VBNO) Then lstItm.SubItems(2) = !VBNO
     Else
      If Not IsNull(!chln) Then lstItm.SubItems(2) = !chln
    End If
    lstItm.SubItems(3) = !TQTY
    lstItm.SubItems(4) = !BNET
    If (Not IsNull(!Form)) And (Trim(!Form) <> "") Then lstItm.SubItems(5) = Format(!FDAT, "dd/MM/yyyy")
    lstItm.SubItems(6) = !dbcd
    lstItm.SubItems(7) = !PCOD
  .MoveNext
  Wend

End With
spara = "Y"
If LSTVW.ListItems.COUNT > 0 Then LSTVW.SetFocus
End Function

Public Function PARTYTAXVIEW()

Dim TEMPRS As New ADODB.Recordset
Set TEMPRS = New ADODB.Recordset
If LSTVW.ListItems.COUNT = 0 Then Exit Function

PARTYLIST = ""
If FRMPARA = "SAL" Then
   PARTYLIST = "SELECT BILLMAIN.*,ACCMST.NAME FROM BILLMAIN INNER JOIN ACCMST ON BILLMAIN.PCOD=ACCMST.CODE " & _
   "WHERE BILLMAIN.COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='DBN' OR VTYP='OPC' ) AND DBCD='" & LSTVW.SelectedItem.SubItems(6) & "'  AND VBNO='" & LSTVW.SelectedItem.SubItems(1) & "' AND RECSTAT<>'D'"
 Else
  PARTYLIST = "SELECT PURMAN.*,ACCMST.NAME FROM PURMAN INNER JOIN ACCMST ON PURMAN.PCOD=ACCMST.CODE " & _
   "WHERE PURMAN.COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND (VTYP='" & FRMPARA & "' OR VTYP='CRN') AND DBCD='" & LSTVW.SelectedItem.SubItems(6) & "' AND PSNO='" & LSTVW.SelectedItem.SubItems(1) & "' AND RECSTAT<>'D'"
End If

  TXTFORM.Text = ""
With TEMPRS

  If .State = adStateOpen Then .Close
  
  .Open PARTYLIST, CN, adOpenDynamic, adLockOptimistic
  
  If .EOF = True Then Exit Function

  TXTPTYSHOW = !NAME
  If Not IsNull(!Form) Then TXTFORM = !Form
  If Not IsNull(!FDAT) Then
     If Format(!FDAT, "mm/dd/yyyy") = "12/30/1899" Then
       txtTaxDT.Value = Now
      Else
       txtTaxDT.Value = Format(!FDAT, "mm/dd/yyyy")
     End If
   Else
     txtTaxDT.Value = Now
  End If
  If Not IsNull(!TXCD) Then lblForm.Caption = GETREF(!TXCD)
.Close
End With

End Function

Private Sub txtTXCD_GotFocus()
 txtTXCD.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txttxcd_KeyDown(KeyCode As Integer, Shift As Integer)
    txtTXCD.Tag = Empty
    If KeyCode = vbKeyF2 Then
        txtTXCD = SearchList1("Select  TOP 20 Code,Name From TAXMST WHERE RECSTAT='A'", 0, Empty, "Select Tax Category")
        txtTXCD.Tag = Key
        SendKeys "{TAB}"
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtTXCD_LostFocus()
 txtTXCD.BackColor = vbWhite
End Sub

Public Function Allow_view_only() As String
 Allow_view_only = "N"
 Dim VIEWRS As New ADODB.Recordset
 Set VIEWRS = New ADODB.Recordset
 VIEWRS.Open "SELECT DISTINCT FYCLOSE FROM SERIALMASTER where comp='" & compPth & "' and unit='" & UNCD & "' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
 If Not VIEWRS.EOF Then
   Allow_view_only = VIEWRS!FYCLOSE & ""
 End If
 If Allow_view_only = "Y" Then
   MsgBox "Fy-Close !! All the Trancasation is in View Mode Only "
 End If
End Function
