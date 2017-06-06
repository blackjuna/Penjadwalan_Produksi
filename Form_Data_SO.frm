VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Data_SO 
   Caption         =   "Form Data Sales Order"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Print Berdasarkan Delivery"
      Height          =   1815
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cprint 
         Caption         =   "Print"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox ccustomer 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dttanggal1 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41794
      End
      Begin MSComCtl2.DTPicker dttanggal2 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41794
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DT PPIC"
         Height          =   195
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Costumer"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sampai"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Berdasarkan Sales Order"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmb_category 
         Height          =   315
         ItemData        =   "Form_Data_SO.frx":0000
         Left            =   1560
         List            =   "Form_Data_SO.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print CS"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox tno_so 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan No SO"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1275
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8880
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   6165
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
End
Attribute VB_Name = "Form_Data_SO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ccustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If ccustomer.Text = "Fajar Benua" Then
        qry_tampil = "select no_slip,no_so,proses_2,jic,qty,customer,finish_date from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and left(no_so,1)='F'"
        Set rscompletion_slip = conn.Execute(qry_tampil)
        Set DataGrid1.DataSource = rscompletion_slip.DataSource
        With DataGrid1
            .Columns(0).Width = 2000
            .Columns(1).Width = 800
            .Columns(2).Width = 1050
            .Columns(3).Width = 2500
            .Columns(4).Width = 400
            .Columns(5).Width = 2000
            .Columns(6).Width = 1200
            .Columns(6).Caption = "DT PPIC"
        End With
    DataGrid1.Refresh
    ElseIf ccustomer.Text = "Toko SIP" Then
        qry_tampil = "select no_slip,no_so,proses_2,jic,qty,customer,finish_date from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S')"
        Set rscompletion_slip = conn.Execute(qry_tampil)
        Set DataGrid1.DataSource = rscompletion_slip.DataSource
        With DataGrid1
            .Columns(0).Width = 2000
            .Columns(1).Width = 800
            .Columns(2).Width = 1050
            .Columns(3).Width = 2500
            .Columns(4).Width = 400
            .Columns(5).Width = 2000
            .Columns(6).Width = 1200
            .Columns(6).Caption = "DT PPIC"
        End With
    DataGrid1.Refresh
    Else
    qry_tampil = "select no_slip,no_so,proses_2,jic,qty,customer,finish_date from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and left(no_so,1)='E'"
        Set rscompletion_slip = conn.Execute(qry_tampil)
        Set DataGrid1.DataSource = rscompletion_slip.DataSource
        With DataGrid1
            .Columns(0).Width = 2000
            .Columns(1).Width = 800
            .Columns(2).Width = 1050
            .Columns(3).Width = 2500
            .Columns(4).Width = 400
            .Columns(5).Width = 2000
            .Columns(6).Width = 1200
            .Columns(6).Caption = "DT PPIC"
        End With
    DataGrid1.Refresh
    End If

End Sub

Private Sub Command1_Click()
Call tampilgrid
tno_so.Text = ""
tno_so.SetFocus
ListView1.ListItems.Clear
End Sub

Private Sub Command2_Click()
    Dim stringformula As String
    
    Select Case cmb_category.Text
        Case "SWG"
            strcategory = "IsNull({completion_slip.category}) and "
        Case "DJG"
            strcategory = "{completion_slip.category} =2 and "
        Case "SMG"
            strcategory = "{completion_slip.category} =3 and "
        Case ""
            MsgBox "Mohon pilih category terlebih dahulu", vbCritical + vbOKOnly, "Warning"
            Exit Sub
    End Select
    
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
               stringformula = stringformula & "{completion_slip.no_so}='" & ListView1.ListItems(i).Text & "' OR "
        End If
    Next i
    
    If Not stringformula = "" Then ' bila ada yang dicentang
        'stringformula = Mid(stringformula, 1, Len(stringformula) - 3)   'menghilangkan 'OR' diakhir stringformula
        stringformula = strcategory & "(" & Mid(stringformula, 1, Len(stringformula) - 3) & ")" 'menghilangkan 'OR' diakhir stringformula
    Else ' tidak ada yang dicentang
        stringformula = "{completion_slip.no_so}=''"
    End If
    CrystalReport1.Reset
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
    
    CrystalReport1.SelectionFormula = stringformula
    Select Case cmb_category.Text
        Case "SWG"
            CrystalReport1.ReportFileName = App.Path + "\list_foreman_custom.rpt"
        Case "DJG"
            CrystalReport1.ReportFileName = App.Path + "\list_foreman_custom_djg.rpt"
        Case "SMG"
            CrystalReport1.ReportFileName = App.Path + "\list_foreman_custom_smg.rpt"
    End Select
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.WindowShowGroupTree = False
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1

End Sub

Private Sub Command3_Click()
Dim stringformula As String

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
           stringformula = stringformula & "{completion_slip.no_so}='" & ListView1.ListItems(i).Text & "' OR "
    End If
Next i

If Not stringformula = "" Then ' bila ada yang dicentang
    stringformula = Mid(stringformula, 1, Len(stringformula) - 3) 'menghilangkan 'OR' diakhir stringformula
Else ' tidak ada yang dicentang
    stringformula = "{completion_slip.no_so}=''"
End If
        CrystalReport1.Reset
        CrystalReport1.Destination = crptToWindow
        CrystalReport1.Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=penjadwalan_produksi"
        CrystalReport1.ReportFileName = App.Path + "\cs.rpt"
        CrystalReport1.SelectionFormula = stringformula
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.WindowShowGroupTree = False
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.Action = 1

End Sub



Private Sub cprint_Click()
If ccustomer.Text = "Fajar Benua" Then
    With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=1234567;DSQ=penjadwalan_produksi"
        .SQLQuery = "select*from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and left(no_so,1)='F'"
        .ReportFileName = App.Path & "\list_foreman_custom.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
ElseIf ccustomer.Text = "Toko SIP" Then
    With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=1234567;DSQ=penjadwalan_produksi"
        .SQLQuery = "select*from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and (left(no_so,1)='J' or left(no_so,1)='M' or left(no_so,1)='C' or left(no_so,1)='S')"
        .ReportFileName = App.Path & "\list_foreman_custom.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
Else
    With CrystalReport1
        .Connect = "DSN=produksi;UID=sa;PWD=1234567;DSQ=penjadwalan_produksi"
        .SQLQuery = "select*from completion_slip where finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' and finish_date<= '" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' and left(no_so,1)='E' order by no_slip asc"
        .ReportFileName = App.Path & "\list_foreman_custom.rpt"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
End If

End Sub

Private Sub dttanggal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ccustomer.SetFocus
End If
End Sub

Private Sub Form_Activate()
Call db
conn.CursorLocation = adUseClient
rscompletion_slip.Open "select*from completion_slip", conn
Call tampilgrid

tno_so.SetFocus

With DataGrid1
    .Columns(0).Width = 2000
    .Columns(1).Width = 800
    .Columns(2).Width = 1050
    .Columns(3).Width = 2500
    .Columns(4).Width = 400
    .Columns(5).Width = 2000
    .Columns(6).Width = 1200
    .Columns(6).Caption = "DT PPIC"
End With

ccustomer.AddItem "Fajar Benua"
ccustomer.AddItem "Toko SIP"
ccustomer.AddItem "Export"

End Sub

Sub tampilgrid()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select*from completion_slip", conn
End If

lihat = "select no_slip,no_so,proses_2,jic,qty,customer,finish_date from completion_slip"
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
With DataGrid1
    .Columns(0).Width = 2000
    .Columns(1).Width = 800
    .Columns(2).Width = 1050
    .Columns(3).Width = 2500
    .Columns(4).Width = 400
    .Columns(5).Width = 2000
    .Columns(6).Width = 1200
    .Columns(6).Caption = "DT PPIC"
End With

DataGrid1.Refresh

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End If
End Sub

Private Sub tno_so_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lst As ListItem


If KeyCode = vbKeyReturn Then
    If rscompletion_slip.State = 0 Then
        rscompletion_slip.Open "select*from completion_slip", conn
    End If

    qry = "select no_slip,no_so,jic,proses_2,size,qty,customer,finish_date from completion_slip where no_so='" & tno_so.Text & "'"
    Set rscompletion_slip = conn.Execute(qry)
    Set DataGrid1.DataSource = rscompletion_slip.DataSource
    With DataGrid1
    .Columns(0).Width = 2000
    .Columns(1).Width = 800
    .Columns(2).Width = 1050
    .Columns(3).Width = 2500
    .Columns(4).Width = 400
    .Columns(5).Width = 2000
    .Columns(6).Width = 1200
    .Columns(7).Width = 1200
    .Columns(6).Caption = "DT PPIC"
End With
    DataGrid1.Refresh
Set lst = ListView1.ListItems.Add(, , tno_so.Text)

End If
End Sub
