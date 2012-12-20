VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BA0F0D53-DEAE-44A6-B2FD-31C81438FAF1}#1.0#0"; "WELCHBUTTON.OCX"
Begin VB.Form Frm_invoiceDeletion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deletion of Invoice and Challan."
   ClientHeight    =   4695
   ClientLeft      =   5295
   ClientTop       =   2205
   ClientWidth     =   8070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8070
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   4755
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8387
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
      Begin VB.PictureBox Search 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3720
         Picture         =   "Frm_invoiceDeletion.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   300
      End
      Begin VB.TextBox TXTSALTYP 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   4485
      End
      Begin VB.TextBox TXTVBNO 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtConsignee 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtAccoutParty 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtNetAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3360
         Width           =   1335
      End
      Begin MSMask.MaskEdBox dtDate 
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
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
      Begin WelchButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   4080
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
         Image           =   "Frm_invoiceDeletion.frx":058A
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   4080
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
         Image           =   "Frm_invoiceDeletion.frx":0B24
         cBack           =   -2147483633
      End
      Begin WelchButton.lvButtons_H cmdExit 
         Height          =   495
         Left            =   4800
         TabIndex        =   7
         Top             =   4080
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
         Image           =   "Frm_invoiceDeletion.frx":0F76
         cBack           =   -2147483633
      End
      Begin VB.Label lblAlert 
         BackStyle       =   0  'Transparent
         Caption         =   "Deletion of Invoice && Challan "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale &Bill No."
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
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   7815
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Qnty."
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
         Left            =   2520
         TabIndex        =   18
         Tag             =   "S"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   1920
         Y2              =   2880
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
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
         Left            =   960
         TabIndex        =   17
         Tag             =   "S"
         Top             =   3000
         Width           =   855
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7920
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   16
         Tag             =   "S"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000080&
         X1              =   120
         X2              =   7920
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000080&
         X1              =   2160
         X2              =   2160
         Y1              =   2880
         Y2              =   3840
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   1935
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   7815
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000080&
         X1              =   4080
         X2              =   4080
         Y1              =   2880
         Y2              =   3840
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
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
         Left            =   5040
         TabIndex        =   15
         Tag             =   "S"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice &Type"
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
         Left            =   600
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Left            =   5400
         TabIndex        =   14
         Tag             =   "S"
         Top             =   3000
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frm_invoiceDeletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim CUR_PCOD As String
Dim CUR_CRAC As String
Dim CUR_DLPT As String
Dim CHG_PCOD As String
Dim CHG_CRAC As String
Dim CHG_DLAC As String
Dim txtDVCD As String

Private Sub cmdCancel_Click()
    TXTSALTYP = Empty: TXTSALTYP.Tag = Empty: txtAccoutParty = Empty: txtConsignee = Empty
    dtDate.Text = Format(Now, "DD/MM/YYYY")
    TXTVBNO = Empty: TXTSALTYP = Empty: TXTSALTYP.SetFocus: txtQty = Empty: txtNetAmt = Empty
End Sub

Private Sub CMDEXIT_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo LAST
  Dim IS_WASTAGE_CHALLAN As Boolean
  
  IS_WASTAGE_CHALLAN = False
  
  Dim BILLRS As ADODB.Recordset
  Set BILLRS = New ADODB.Recordset
  
  Dim TMPRS As ADODB.Recordset
  Set TMPRS = New ADODB.Recordset
         
      'FIND SALE TYPE CODE
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open "SELECT CODE FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND NAME='" & TXTSALTYP & "' AND VTYP='SAL' AND FYCD='" & FYCD & "'", CN, adOpenDynamic, adLockOptimistic
      
      If TMPRS.EOF Then
        MsgBox "Invalid Sale Type", vbCritical: TXTSALTYP.SetFocus: Exit Sub
      Else
        TXTSALTYP.Tag = TMPRS!CODE
      End If
      TMPRS.Close
      '--------------------
      
      'CHECK VALLID BILL IS READY TO DELETE
      If InvalidBill Then Exit Sub
      '-------------------------------------
      
      'USE CONFIRMATION
      Dim AYS
      AYS = MsgBox("Are You sure to Delete this Invoice ? ", vbYesNo)
      If AYS = VBNO Then Exit Sub
      '------------------
      'OPERATION BEGIN
      CN.BeginTrans
                        
      'DELETE INVOICE
      
      SQL = "DELETE FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
      
      SQL = "DELETE FROM EGPMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
      "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
      
      SQL = "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
            " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
       
       
      'Delete from storetran on 14/10/2011
      
      SQL = "DELETE FROM STORETRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
            " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
      CN.Execute SQL
      
      'Deletetion from Ledger File
      
      SQL = "DELETE FROM TRNMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
      
      CN.Execute SQL
      
      'Delete from tradingstock
      Dim NOS1
      SQL = "DELETE FROM TRADINGSTOCK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
      
      CN.Execute SQL, NOS1
      
      'Delete from rg23-d
      
      SQL = "DELETE FROM RG23D WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "'"
      
      CN.Execute SQL
      
      '-------------------
                   
      'DELETE BOX USING REFRENCE FROM CHALLAN
      SQL = "SELECT DVCD,DBCD,VBNO,LTNO FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
            " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
       
      If TMPRS.State = 1 Then TMPRS.Close
      TMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
      
      If Not TMPRS.EOF Then  'WASTAGE CHALLAN WILL NEVER DELETE FROM HERE
         If Trim(TMPRS!ltno & "") = "WASTE" Then
            AYS = MsgBox("Current Invoice is of Wastage Material, So Only Invoice will be delete but Challan Exist! Do You Continue ?? ", vbYesNo)
            If AYS = VBNO Then CN.RollbackTrans: Exit Sub
            IS_WASTAGE_CHALLAN = True
         Else
            Do While Not TMPRS.EOF
                SQL = "UPDATE BOXREGISTER SET VTYP='PPF',RDBC = NULL,RVBNO = NULL,RVBDT =NULL,RVTYP=NULL " & _
                      "WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='DPF' AND " & _
                      "DVCD='" & TMPRS!DVCD & "' AND RDBC='" & TMPRS!dbcd & _
                      "' AND RVBNO='" & TMPRS!VBNO & "' AND RECSTAT<>'D'"
                CN.Execute SQL
             
                CN.Execute "DELETE FROM PKGMAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                           "' AND VTYP='DPF' AND DVCD='" & TMPRS!DVCD & _
                           "' AND SLIPNO='" & TMPRS!VBNO & "' AND DBCD='" & TMPRS!dbcd & "'"
            TMPRS.MoveNext
            Loop
            TMPRS.Close
         End If
      End If
      
       '----------------------
       
       'CHALLAN DELETION
       SQL = "SELECT COMP,UNIT,DVCD,DBCD,VBNO,ICOD,GRAD,EXTRA1 AS ORDN,SUM(QNTY) AS QNTY FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
            " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D' GROUP BY COMP,UNIT,DVCD,DBCD,VBNO,ICOD,GRAD,EXTRA1"
       
       If TMPRS.State = 1 Then TMPRS.Close
       TMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
                 
       Do While Not TMPRS.EOF
          CN.Execute "DELETE FROM ORDTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='DPF' " & _
            " AND DVCD='" & TMPRS!DVCD & "' AND RDBC='" & TMPRS!dbcd & "' AND SLIP='" & TMPRS!VBNO & "' AND RECSTAT<>'D'"
            
          CN.Execute "DELETE FROM ORDTRN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='DOS' " & _
            " AND DVCD='" & TMPRS!DVCD & "' AND RDBC='" & TMPRS!dbcd & "' AND SLIP='" & TMPRS!VBNO & "' AND RECSTAT<>'D'"
            
          SQL = "UPDATE ORDMAN SET DISPATCHQTY = DISPATCHQTY - " & Val(TMPRS!QNTY & "") & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DCOD='" & TMPRS!DVCD & _
              "' AND ORDN = '" & Trim(TMPRS!ORDN & "") & "' AND ICOD = '" & Trim(TMPRS!ICOD & "") & _
              "' AND TRCD='" & Trim(TMPRS!grad & "") & "'"
                
          CN.Execute SQL
          
          'BECAUSE CONTROL TO ACCESS PRODUCTION QTY FROM ORDER THROUGH DOQTY FIELD OF ORDMAN IN CASE OF EXPORT PACKING
          If UNT_EXPPKG_REQ = "Y" Then
             SQL = "UPDATE ORDMAN SET DOQTY = DOQTY + " & Val(TMPRS!QNTY & "") & " WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND DCOD='" & TMPRS!DVCD & _
                   "' AND ORDN = '" & Trim(TMPRS!ORDN & "") & "' AND ICOD = '" & Trim(TMPRS!ICOD & "") & _
                   "' AND TRCD='" & Trim(TMPRS!grad & "") & "'"
             CN.Execute SQL
          End If
                 
       TMPRS.MoveNext
       Loop
       TMPRS.Close
       
       If Not IS_WASTAGE_CHALLAN Then
            'PKGSTK : BEFORE CHALLAN DELETION
            SQL = "SELECT * FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
                 " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
            
            If TMPRS.State = 1 Then TMPRS.Close
            TMPRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
            Dim L As Long
            If Not TMPRS.EOF Then
               SQL = "DELETE FROM PKGSTK WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                     "' AND DVCD='" & TMPRS!DVCD & "' AND DBCD='" & TMPRS!dbcd & "' AND VTYP='DPF' AND CHLN='" & TMPRS!VBNO & "' "
             
               CN.Execute SQL, L
            End If
            TMPRS.Close
            '====================================================================================================
            SQL = "DELETE FROM SPTRAN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
            " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D' "
            
            CN.Execute SQL
       Else
            SQL = "UPDATE SPTRAN SET RTYP=NULL,SDBC=NULL,SVBN=NULL WHERE COMP='" & compPth & _
            "' AND UNIT='" & UNCD & "' AND RTYP='SAL' " & _
            " AND SDBC='" & TXTSALTYP.Tag & "' AND SVBN='" & TXTVBNO.Text & "' AND RECSTAT<>'D' "
            
            CN.Execute SQL
       End If
                     
      'OPERATION FINISH SUCCESSFULLY
      If RS.State = 1 Then RS.Close
      RS.Open "SELECT * FROM DAILYSTAT WHERE 1=2", CN, adOpenDynamic, adLockOptimistic
      RS.AddNew
      RS!COMP = compPth
      RS!VTYP = "SAL"
      RS!PCOD = ""
      RS!SRNO = "Delete"
      RS!VBNO = Trim(TXTVBNO)
      RS!AMNT = Val(txtNetAmt)
      RS!CUSR = cUName
      RS!QNTY = Val(txtQty)
      RS!ACTN = "D"
      RS!DTTM = Format(Now, "YYYY/MM/DD HH:MM:SS")
      RS.Update
       CN.CommitTrans
            
      MsgBox "Invoice Deleted Successfuly "
      
  Call cmdCancel_Click
  Exit Sub
  
LAST:
  MsgBox ERR.Description
  Resume
  CN.RollbackTrans
End Sub

Private Sub GenDetails()
  On Error GoTo LAST
  
  Dim BILLRS As New ADODB.Recordset
  Set BILLRS = New ADODB.Recordset
  
  SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & "' AND VTYP='SAL' " & _
        " AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
  
  If BILLRS.State = 1 Then BILLRS.Close
  BILLRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
  
  If BILLRS.EOF Then
    MsgBox "Invalid Bill No."
    TXTVBNO.SetFocus
    Exit Sub
  End If
  
  txtQty = BILLRS!TQTY
  txtNetAmt = BILLRS!BNET
  dtDate.Text = Format(BILLRS!Date, "DD/MM/YYYY")
  
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM ACCMST WHERE CODE='" & BILLRS!PCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then txtAccoutParty = Trim(RS!NAME & "")
    
  If RS.State = 1 Then RS.Close
  RS.Open "SELECT NAME FROM PADDMST WHERE CODE='" & BILLRS!DCOD & "'", CN, adOpenDynamic, adLockOptimistic
  If Not RS.EOF Then txtConsignee = Trim(RS!NAME & "")
  
  RS.Close
  
  cmdDelete.SetFocus
  
Exit Sub
LAST:
  MsgBox ERR.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If ActiveControl.NAME = "TXTVBNO" Then Exit Sub
  If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
  Call CenterChild(frm_Main, Me)
  Call ColorComponent(Me)
  dtDate.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub Search_Click()
   Call GenDetails
End Sub

Private Sub TXTSALTYP_GotFocus()
 TXTSALTYP.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub TXTSALTYP_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn And Trim(TXTSALTYP) = Empty) Or KeyCode = vbKeyF2 Then
    TXTSALTYP.Text = SearchList1("SELECT CODE,NAME FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
                     "' AND VTYP='SAL' AND FYCD='" & FYCD & "' AND ACTIVE='Y'", 0, TXTSALTYP.Text, "SELECT INVOICE TYPE FROM LIST")
    TXTSALTYP.Tag = Key
    Call FindLastBill
  End If
End Sub

Private Sub TXTSALTYP_LostFocus()
 TXTSALTYP.BackColor = vbWhite
End Sub

Private Sub TXTVBNO_GotFocus()
 TXTVBNO.BackColor = RGB(BRED, BGREEN, BBLUE)
End Sub

Private Sub txtVBNO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TXTVBNO <> Empty And Len(TXTVBNO) = 10 Then
    Call GenDetails
    cmdDelete.SetFocus
End If
End Sub

Private Sub TXTVBNO_LostFocus()
 TXTVBNO.BackColor = vbWhite
End Sub


Private Function InvalidBill() As Boolean
InvalidBill = False 'CONSIDER IT IS A VALID BILL

Dim VALIDRS As ADODB.Recordset
 Set VALIDRS = New ADODB.Recordset

 'CASE :1 : EXIST BILL AND NOT HAVING ANY PAYMENT RECEIPT ENTRY
 SQL = "SELECT * FROM BILLMAIN WHERE COMP='" & compPth & "' AND UNIT='" & UNCD & _
 "' AND VTYP='SAL' AND DBCD='" & TXTSALTYP.Tag & "' AND VBNO='" & TXTVBNO.Text & "' AND RECSTAT<>'D'"
 
 If VALIDRS.State = 1 Then VALIDRS.Close
 VALIDRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
 If VALIDRS.EOF Then
    InvalidBill = True  'PROVE INVALID BILL
    MsgBox "Invoice not Created till date", vbCritical
    TXTVBNO.SetFocus
    Exit Function
 Else
    
    If UCase(Trim(VALIDRS!BSTS & "")) = "A" Then
       InvalidBill = True
       MsgBox "Further Entry Exist", vbCritical, "Sale Bill is Audited"
       TXTVBNO.SetFocus
       Exit Function
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT COMP FROM RPTRAN WHERE COMP='" & VALIDRS!COMP & _
            "' AND BSR1='" & VALIDRS!VTYP & "'", CN, adOpenDynamic, adLockOptimistic
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT COMP FROM RPTRAN WHERE COMP='" & VALIDRS!COMP & "' AND BSR1='" & VALIDRS!VTYP & _
            "' AND SVBN='" & VALIDRS!VBNO & "' AND SDBC='" & VALIDRS!dbcd & _
            "' AND UNIT='" & VALIDRS!unit & "'", CN, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       InvalidBill = True
       MsgBox "Further Entry Exist Can not modified"
       TXTVBNO.SetFocus
       Exit Function
    End If
 End If
End Function
 
Private Sub FindLastBill()
On Error GoTo LAST
Dim FINDRS As ADODB.Recordset
Set FINDRS = New ADODB.Recordset

SQL = "SELECT SRNO FROM SERIALMASTER WHERE COMP='" & compPth & "' AND UNIT = '" & UNCD & "' AND VTYP='SAL' " & _
      "AND CODE='" & TXTSALTYP.Tag & "' AND FYCD='" & FYCD & "'"

If FINDRS.State = 1 Then FINDRS.Close
FINDRS.Open SQL, CN, adOpenDynamic, adLockOptimistic
If Not FINDRS.EOF Then
   TXTVBNO = FINDRS!SRNO & ""
Else
   TXTVBNO = Empty
End If
Exit Sub
LAST:
MsgBox ERR.Description
End Sub
