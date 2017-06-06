VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Status_CS 
   Caption         =   "Form Status CS"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "Summary Status Completion Slip / Order"
      Height          =   2775
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   2280
         Width           =   840
      End
      Begin VB.Line Line12 
         X1              =   240
         X2              =   5160
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lbl_grandtotal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   1440
         TabIndex        =   25
         Top             =   2280
         Width           =   3645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "On Going"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pending"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Closed"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   480
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   240
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   5160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   5160
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   5160
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   645
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   5160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "FBI"
         Height          =   195
         Left            =   1425
         TabIndex        =   20
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label lbl_closed_fbi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lbl_pending_fbi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl_on_going_fbi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   1005
      End
      Begin VB.Line Line8 
         X1              =   2520
         X2              =   2520
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Line Line9 
         X1              =   3720
         X2              =   3720
         Y1              =   360
         Y2              =   2160
      End
      Begin VB.Label lbl_on_going_sip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   2640
         TabIndex        =   16
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lbl_pending_sip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   2640
         TabIndex        =   15
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lbl_closed_sip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   2640
         TabIndex        =   14
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "SIP"
         Height          =   195
         Left            =   2610
         TabIndex        =   13
         Top             =   480
         Width           =   990
      End
      Begin VB.Line Line10 
         X1              =   5160
         X2              =   5160
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Label lbl_on_going_export 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   3840
         TabIndex        =   12
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label lbl_pending_export 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label lbl_closed_export 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "EXPORT"
         Height          =   195
         Left            =   3855
         TabIndex        =   9
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lbl_subtotal_fbi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Line Line11 
         X1              =   240
         X2              =   5160
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Sub Total"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label lbl_subtotal_sip 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lbl_subtotal_export 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   3840
         TabIndex        =   5
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   1320
         Y1              =   360
         Y2              =   2520
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   7680
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cstatus 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pilih Status CS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Menu print 
      Caption         =   "Print"
      Begin VB.Menu fajar_benua 
         Caption         =   "Fajar Benua"
         Begin VB.Menu fbi_on_going 
            Caption         =   "On Going"
         End
         Begin VB.Menu fbi_pending 
            Caption         =   "Pending"
         End
         Begin VB.Menu fbi_closed 
            Caption         =   "Closed"
         End
      End
      Begin VB.Menu toko_sip 
         Caption         =   "Toko SIP"
         Begin VB.Menu sip_on_going 
            Caption         =   "On Going"
         End
         Begin VB.Menu sip_pending 
            Caption         =   "Pending"
         End
         Begin VB.Menu sip_closed 
            Caption         =   "Closed"
         End
      End
      Begin VB.Menu export 
         Caption         =   "Export"
         Begin VB.Menu export_on_going 
            Caption         =   "On Going"
         End
         Begin VB.Menu export_pending 
            Caption         =   "Pending"
         End
         Begin VB.Menu export_closed 
            Caption         =   "Closed"
         End
      End
      Begin VB.Menu summary 
         Caption         =   "Summary"
      End
   End
End
Attribute VB_Name = "Form_Status_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
If Label7.Caption = 1 Then
    With CrystalReport2
        .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
        .SQLQuery = "select*from view_pending_fbi"
        .ReportFileName = App.Path & "\list_on_going.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
End If
End Sub

Private Sub cstatus_Click()
intip = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where status='" & cstatus.Text & "'"
Set rscompletion_slip = conn.Execute(intip)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
End Sub

Private Sub export_on_going_Click()
With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=Purchasing"
        .SQLQuery = "select * from completion_slip where left(no_so,1)='E' and status='On Going' order by finish_date asc"
        .ReportFileName = App.Path & "\list_on_going.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub fbi_on_going_Click()
With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
        .SQLQuery = "select * from completion_slip where left(no_so,1)='F' and status='On Going' order by finish_date asc"
        .ReportFileName = App.Path & "\list_on_going.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub Form_Activate()
'Call db
'conn.CursorLocation = adUseClient
If rscompletion_slip.State = 1 Then rscompletion_slip.Close
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where deleted=0", conn
Call tampilgrid

Me.cstatus.AddItem "All"
Me.cstatus.AddItem "On Going"
Me.cstatus.AddItem "Pending"
Me.cstatus.AddItem "Closed"

Label7 = print_id
End Sub

Sub tampilgrid()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

lihat = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip"
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub Form_Load()
qry_on_going_FBI = "select sum(qty) AS On_Going_FBI from completion_slip where left(no_so,1)='F' and status='On Going' "
Set rscompletion_slip = conn.Execute(qry_on_going_FBI)
on_going_fbi = rscompletion_slip.Fields("On_Going_FBI")
Me.lbl_on_going_fbi = on_going_fbi

qry_pending_FBI = "select sum(qty) AS Pending_FBI from completion_slip where left(no_so,1)='F' and status='Pending' "
Set rscompletion_slip = conn.Execute(qry_pending_FBI)
pending_fbi = rscompletion_slip.Fields("Pending_FBI")
Me.lbl_pending_fbi = IIf(IsNull(pending_fbi), "0", pending_fbi)

qry_closed_FBI = "select sum(qty) AS Closed_FBI from completion_slip where left(no_so,1)='F' and status='Closed' "
Set rscompletion_slip = conn.Execute(qry_closed_FBI)
closed_fbi = rscompletion_slip.Fields("Closed_FBI")
Me.lbl_closed_fbi = closed_fbi

qry_subtotal_FBI = "select sum(qty) AS Subtotal_FBI from completion_slip where left(no_so,1)='F'"
Set rscompletion_slip = conn.Execute(qry_subtotal_FBI)
subtotal_fbi = rscompletion_slip.Fields("Subtotal_FBI")
Me.lbl_subtotal_fbi = subtotal_fbi

qry_on_going_SIP = "select sum(qty) AS On_Going_SIP from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='On Going' "
Set rscompletion_slip = conn.Execute(qry_on_going_SIP)
on_going_sip = rscompletion_slip.Fields("On_Going_SIP")
Me.lbl_on_going_sip = on_going_sip

qry_pending_SIP = "select sum(qty) AS Pending_SIP from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='Pending' "
Set rscompletion_slip = conn.Execute(qry_pending_SIP)
pending_sip = rscompletion_slip.Fields("Pending_SIP")
Me.lbl_pending_sip = IIf(IsNull(pending_sip), "0", pending_sip)

qry_closed_SIP = "select sum(qty) AS Closed_SIP from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='Closed' "
Set rscompletion_slip = conn.Execute(qry_closed_SIP)
closed_sip = rscompletion_slip.Fields("Closed_SIP")
Me.lbl_closed_sip = closed_sip

qry_subtotal_SIP = "select sum(qty) AS Subtotal_SIP from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S')"
Set rscompletion_slip = conn.Execute(qry_subtotal_SIP)
subtotal_sip = rscompletion_slip.Fields("Subtotal_SIP")
Me.lbl_subtotal_sip = subtotal_sip

qry_subtotal_SIP = "select sum(qty) AS Subtotal_SIP from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S')"
Set rscompletion_slip = conn.Execute(qry_subtotal_SIP)
subtotal_sip = rscompletion_slip.Fields("Subtotal_SIP")
Me.lbl_subtotal_sip = subtotal_sip

qry_grandtotal = "select sum(qty) AS Grand_Total from completion_slip "
Set rscompletion_slip = conn.Execute(qry_grandtotal)
grandtotal = rscompletion_slip.Fields("Grand_Total")
Me.lbl_grandtotal = grandtotal

qry_on_going_export = "select sum(qty) AS On_Going_Export from completion_slip where left(no_so,1)='E' and status='On Going' "
Set rscompletion_slip = conn.Execute(qry_on_going_export)
on_going_export = rscompletion_slip.Fields("On_Going_Export")
Me.lbl_on_going_export = IIf(IsNull(on_going_export), "0", on_going_export)

qry_pending_export = "select sum(qty) AS Pending_Export from completion_slip where left(no_so,1)='E' and status='Pending' "
Set rscompletion_slip = conn.Execute(qry_pending_export)
pending_export = rscompletion_slip.Fields("Pending_Export")
Me.lbl_pending_export = IIf(IsNull(pending_export), "0", pending_export)

qry_closed_export = "select sum(qty) AS Closed_Export from completion_slip where left(no_so,1)='E' and status='Closed' "
Set rscompletion_slip = conn.Execute(qry_closed_export)
closed_export = rscompletion_slip.Fields("Closed_Export")
Me.lbl_closed_export = IIf(IsNull(closed_export), "0", closed_export)

qry_subtotal_export = "select sum(qty) AS Subtotal_Export from completion_slip where left(no_so,1)='E'"
Set rscompletion_slip = conn.Execute(qry_subtotal_export)
subtotal_export = rscompletion_slip.Fields("Subtotal_Export")
Me.lbl_subtotal_export = IIf(IsNull(subtotal_export), "0", subtotal_export)


End Sub

Private Sub lbl_closed_jfi_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_closed_FBI = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='F' and status='Closed'"
Set rscompletion_slip = conn.Execute(qry_closed_FBI)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
End Sub

Private Sub lbl_closed_export_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_closed_export = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='E' and status='Closed'"
Set rscompletion_slip = conn.Execute(qry_closed_export)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
End Sub

Private Sub lbl_closed_fbi_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_closed_FBI = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='F' and status='Closed'"
Set rscompletion_slip = conn.Execute(qry_closed_FBI)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_closed_sip_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_closed_SIP = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='Closed'"
Set rscompletion_slip = conn.Execute(qry_closed_SIP)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_on_going_export_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_on_going_export = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='E' and status='On Going'"
Set rscompletion_slip = conn.Execute(qry_on_going_export)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_on_going_fbi_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_on_going_FBI = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='F' and status='On Going'"
Set rscompletion_slip = conn.Execute(qry_on_going_FBI)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

print_id = 1

End Sub

Private Sub lbl_on_going_sip_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_on_going_SIP = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='On Going'"
Set rscompletion_slip = conn.Execute(qry_on_going_SIP)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_pending_export_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_pending_export = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='E' and status='Pending'"
Set rscompletion_slip = conn.Execute(qry_pending_export)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_pending_fbi_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_pending_FBI = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='F' and status='Pending'"
Set rscompletion_slip = conn.Execute(qry_pending_FBI)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

print_id = 2

End Sub

Private Sub lbl_pending_sip_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_pending_SIP = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='Pending'"
Set rscompletion_slip = conn.Execute(qry_pending_SIP)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_subtotal_export_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_subtotal_export = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='E' "
Set rscompletion_slip = conn.Execute(qry_subtotal_export)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub lbl_subtotal_fbi_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_subtotal_FBI = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where left(no_so,1)='F'"
Set rscompletion_slip = conn.Execute(qry_subtotal_FBI)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
End Sub

Private Sub lbl_subtotal_sip_Click()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select no_slip,no_so,jic,qty,status from completion_slip", conn
End If

qry_subtotal_SIP = "select no_slip,no_so,jic,qty,status,qty_pending from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S')"
Set rscompletion_slip = conn.Execute(qry_subtotal_SIP)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

End Sub

Private Sub sip_on_going_Click()
With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
        .SQLQuery = "select * from completion_slip where (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S') and status='On Going' order by finish_date asc"
        .ReportFileName = App.Path & "\list_on_going.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub
