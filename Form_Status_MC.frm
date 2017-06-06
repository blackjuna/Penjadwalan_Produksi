VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Status_MC 
   Caption         =   "Form Status MC"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Edit Jadwal Produksi"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   11055
      Begin VB.CommandButton crevisi 
         Caption         =   "REVISI SCHEDULE"
         Height          =   495
         Left            =   9480
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox cnomc 
         Height          =   315
         Left            =   6240
         TabIndex        =   24
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox tcustomer 
         Height          =   405
         Left            =   960
         TabIndex        =   22
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton chapus 
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   9480
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cedit 
         Caption         =   "EDIT"
         Height          =   495
         Left            =   9480
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cbatal 
         Caption         =   "BATAL"
         Height          =   495
         Left            =   9480
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox tjic 
         Height          =   405
         Left            =   960
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox tsize 
         Height          =   405
         Left            =   960
         TabIndex        =   14
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox tqty 
         Height          =   405
         Left            =   960
         TabIndex        =   13
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox tno_slip 
         Height          =   405
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox tno_so 
         Height          =   405
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cshift 
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   1800
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtdate_printed 
         Height          =   375
         Left            =   6240
         TabIndex        =   6
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98172929
         CurrentDate     =   41714
      End
      Begin MSComCtl2.DTPicker dtfinish_date 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98172929
         CurrentDate     =   41714
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No MC"
         Height          =   195
         Left            =   4680
         TabIndex        =   25
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "JIC"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Slip"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Finish Date"
         Height          =   195
         Left            =   4680
         TabIndex        =   11
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date Printed"
         Height          =   195
         Left            =   4680
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. SO"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Shift"
         Height          =   195
         Left            =   4680
         TabIndex        =   8
         Top             =   1800
         Width           =   315
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4683
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   40
      Cols            =   721
      FixedRows       =   3
   End
   Begin MSComCtl2.DTPicker DtStart 
      Height          =   375
      Left            =   1080
      TabIndex        =   27
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98172929
      CurrentDate     =   41900
   End
   Begin MSComCtl2.DTPicker DtEnd 
      Height          =   375
      Left            =   5640
      TabIndex        =   28
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98172929
      CurrentDate     =   41900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Start Date  :"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "End Date  :"
      Height          =   195
      Left            =   4680
      TabIndex        =   29
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "Form_Status_MC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As Recordset
Dim vCount As Long
Dim vrow As Long
Dim vPath As String
Dim vDate As Date
Dim vintv As Integer

Private Sub cbatal_Click()
Call bersih
cedit.Caption = "EDIT"
End Sub

Private Sub cedit_Click()
If cedit.Caption = "EDIT" Then
    cari = "select * from completion_slip where no_slip='" & DataGrid1.Columns("no_slip") & "'"
    Set rscompletion_slip = conn.Execute(cari)
    cedit.Caption = "UPDATE"
    Call tampil
Else
    ubah = "UPDATE completion_slip SET date_printed='" & Format(dtdate_printed.Value, "YYYY/mm/dd") & "', finish_date='" & Format(dtfinish_date.Value, "YYYY/mm/dd") & "',  shift='" & cshift.Text & "',proses_2='" & cnomc.Text & "' where no_slip='" & tno_slip.Text & "'"
    Set rscompletion_slip = conn.Execute(ubah)
    tampilgrid
    Call bersih
    cedit.Caption = "EDIT"
    MSFlexGrid1.Refresh
    Call hitung_kolom
    'Call cek_kapasitas
End If

End Sub

Private Sub chapus_Click()
x = MsgBox("Yakin Mau Dihapus...?", vbYesNo + vbInformation, "Hapus Data")
If x = vbYes Then
    hapus = "delete from completion_slip where no_slip='" & DataGrid1.Columns("no_slip") & "'"
    Set rscompletion_slip = conn.Execute(hapus)
    Call tampilgrid
    Call bersih
    Call hitung_kolom
End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub crevisi_Click()
Form_Revisi_CS.Show
End Sub

Private Sub csplit_Click()

End Sub

Private Sub DtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call FlexStatusMC
        Exit Sub
    End Select
End Sub

Private Sub DtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DtEnd.SetFocus
        Exit Sub
    End Select
End Sub

Private Sub Form_Activate()
    DtStart.SetFocus
End Sub

Private Sub Form_Load()
    Call db
    
    DtStart.Value = Format(Date, "dd-mm-yyyy")
    DtEnd.Value = Format(DateAdd("d", 30, DtStart.Value), "dd-mm-yyyy")
    
    
'    conn.CursorLocation = adUseClient
'rsdata_mesin.Open "select*from data_mesin", conn
'rscompletion_slip.Open "select*from completion_slip", conn
'Call tampilgrid
'Call Shift
'MSFlexGrid1.AllowUserResizing = flexResizeColumns
'MSFlexGrid1.ColWidth(0) = 1000
'Dim rst As Recordset
'Dim vCount As Long
'Dim vrow As Long
'Dim vPath As String
'Dim vDate As Date
'
'
''Coding Buat Kolom Tanggal
'vDate = "2015/01/1"
'For vCount = 1 To 720 Step 2
'    MSFlexGrid1.TextMatrix(1, vCount) = (Format(vDate, "YYYY/mm/dd"))
'    MSFlexGrid1.TextMatrix(1, vCount + 1) = (Format(vDate, "YYYY/mm/dd"))
'
'    vDate = vDate + 1
'
'    MSFlexGrid1.MergeCells = flexMergeRestrictRows
'    MSFlexGrid1.MergeRow(1) = True
'Next
'
'
''Coding Buat Kolom Shift
'For vCount = 1 To 720
'    If (vCount Mod 2) = 0 Then
'        MSFlexGrid1.TextMatrix(2, vCount) = 2
'        Else
'        MSFlexGrid1.TextMatrix(2, vCount) = 1
'    End If
'    MSFlexGrid1.ColAlignment(vCount) = flexAlignCenterCenter
'Next
'
''MSFlexGrid1.Rows = 3
'BarisData = 2
'
'If rsdata_mesin.BOF Then
'    Exit Sub
'Else
'    rsdata_mesin.MoveFirst
'    Do While Not rsdata_mesin.EOF
'        BarisData = BarisData + 1
'        MSFlexGrid1.Rows = BarisData + 1
'        MSFlexGrid1.TextMatrix(BarisData, 0) = rsdata_mesin.Fields("nama_mesin")
'        cnomc.AddItem rsdata_mesin("nama_mesin")
'        rsdata_mesin.MoveNext
'    Loop
'End If
'
'Call hitung_kolom

End Sub
Sub tampilgrid()
lihat = "select no_so,no_slip,jic,size,qty,customer,qty_pending,status,date_printed,delivery_date,finish_date,proses_2,shift from completion_slip where proses_2='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.MouseRow, 0) & "' and finish_date='" & Format(MSFlexGrid1.TextMatrix(1, MSFlexGrid1.MouseCol), "YYYY/mm/dd") & "' and shift='" & MSFlexGrid1.TextMatrix(2, MSFlexGrid1.MouseCol) & "'"
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
End Sub

Private Sub MSFlexGrid1_Click()
Call tampilgrid
End Sub

Sub tampil()
Me.tno_slip = rscompletion_slip.Fields("no_slip")
Me.tno_so = rscompletion_slip.Fields("no_so")
Me.tjic = rscompletion_slip.Fields("jic")
Me.tsize = rscompletion_slip("size")
Me.tcustomer = rscompletion_slip("customer")
Me.tqty = rscompletion_slip("qty")
Me.dtdate_printed.Value = rscompletion_slip("date_printed")
Me.dtfinish_date.Value = rscompletion_slip("finish_date")
Me.cnomc.Text = rscompletion_slip("proses_2")
Me.cshift.Text = rscompletion_slip("shift")

End Sub

Sub Shift()
cshift.AddItem "1"
cshift.AddItem "2"
End Sub

Sub bersih()
For Each A In Me
    If TypeOf A Is TextBox Then A.Text = ""
    If TypeOf A Is ComboBox Then A.Text = ""
Next A
End Sub

Sub hitung_kolom()
    For rowna = 3 To MSFlexGrid1.Rows - 1
        For colna = 1 To MSFlexGrid1.Cols - 1
        
        
        jumlah = "select sum(qty) AS MyTotal from completion_slip where " & _
            "proses_2='" & MSFlexGrid1.TextMatrix(rowna, 0) & "' and " & _
            "finish_date='" & Format(MSFlexGrid1.TextMatrix(1, colna), "YYYY/mm/dd") & "' " & _
            "and shift='" & MSFlexGrid1.TextMatrix(2, colna) & "'"
        Set rscompletion_slip = conn.Execute(jumlah)
        
        
        MSFlexGrid1.TextMatrix(rowna, colna) = IIf(IsNull(rscompletion_slip.Fields("MyTotal")), "-", rscompletion_slip.Fields("MyTotal"))
        
        Next
    Next

End Sub

Sub cek_kapasitas()
Dim sisaqty As Integer


For colna = 1 To MSFlexGrid1.Cols - 1
    For rowna = 3 To MSFlexGrid1.Rows - 1
    
    banding = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & MSFlexGrid1.TextMatrix(rowna, 0) & "'"
    Set rsdata_mesin = conn.Execute(banding)
    
    sisaqty = 0
    If Val(MSFlexGrid1.TextMatrix(rowna, colna)) > Val(rsdata_mesin.Fields("MyCapa")) Then
        
        sisaqty = Val(MSFlexGrid1.TextMatrix(rowna, colna)) - Val(rsdata_mesin.Fields("MyCapa"))
        MSFlexGrid1.TextMatrix(rowna, colna) = rsdata_mesin.Fields("MyCapa")
        MSFlexGrid1.TextMatrix(rowna, colna + 1) = Val(MSFlexGrid1.TextMatrix(rowna, colna + 1)) + Val(sisaqty)
        'MsgBox ("Kapasitas Terpenuhi")
    End If
                
    Next
Next
End Sub

Public Sub FlexStatusMC()
'    'conn.CursorLocation = adUseClient

    vintv = Val((DateDiff("d", DtStart.Value, DtEnd.Value) * 2) + 3)
    MSFlexGrid1.Cols = vintv

    Call Shift
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.ColWidth(0) = 1000


    'Coding Buat Kolom Tanggal
    vDate = DtStart.Value
    For vCount = 1 To vintv - 1
        MSFlexGrid1.TextMatrix(1, vCount) = (Format(vDate, "YYYY/mm/dd"))
        MSFlexGrid1.TextMatrix(1, vCount + 1) = (Format(vDate, "YYYY/mm/dd"))

        vDate = vDate + 1

        MSFlexGrid1.MergeCells = flexMergeRestrictRows
        MSFlexGrid1.MergeRow(1) = True
        vCount = vCount + 1
    Next

    'Coding Buat Kolom Shift
    For vCount = 1 To vintv - 1
        If (vCount Mod 2) = 0 Then
            MSFlexGrid1.TextMatrix(2, vCount) = 2
            Else
            MSFlexGrid1.TextMatrix(2, vCount) = 1
        End If
        MSFlexGrid1.ColAlignment(vCount) = flexAlignCenterCenter
    Next

    'MSFlexGrid1.Rows = 3
    BarisData = 2

    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    rsdata_mesin.Open "select*from data_mesin ", conn
    rscompletion_slip.Open "select*from completion_slip order by id", conn

    If rsdata_mesin.EOF Then
        Exit Sub
    Else
        rsdata_mesin.MoveFirst
        Do While Not rsdata_mesin.EOF
            BarisData = BarisData + 1
            MSFlexGrid1.Rows = BarisData + 1
            MSFlexGrid1.TextMatrix(BarisData, 0) = rsdata_mesin.Fields("nama_mesin")
            cnomc.AddItem rsdata_mesin("nama_mesin")
            rsdata_mesin.MoveNext
        Loop
    End If

'    Call tampilgrid
    Call hitung_kolom
    'Call cek_kapasitas


'=============================================
'conn.CursorLocation = adUseClient
'rsdata_mesin.Open "select*from data_mesin", conn
'rscompletion_slip.Open "select*from completion_slip", conn
'Call tampilgrid
'Call Shift
'MSFlexGrid1.AllowUserResizing = flexResizeColumns
'MSFlexGrid1.ColWidth(0) = 1000
'Dim rst As Recordset
'Dim vCount As Long
'Dim vrow As Long
'Dim vPath As String
'Dim vDate As Date


'Coding Buat Kolom Tanggal
'vDate = "2015/01/1"
'For vCount = 1 To 720 Step 2
'    MSFlexGrid1.TextMatrix(1, vCount) = (Format(vDate, "YYYY/mm/dd"))
'    MSFlexGrid1.TextMatrix(1, vCount + 1) = (Format(vDate, "YYYY/mm/dd"))
'
'    vDate = vDate + 1
'
'    MSFlexGrid1.MergeCells = flexMergeRestrictRows
'    MSFlexGrid1.MergeRow(1) = True
'Next
'
'
''Coding Buat Kolom Shift
'For vCount = 1 To 720
'    If (vCount Mod 2) = 0 Then
'        MSFlexGrid1.TextMatrix(2, vCount) = 2
'        Else
'        MSFlexGrid1.TextMatrix(2, vCount) = 1
'    End If
'    MSFlexGrid1.ColAlignment(vCount) = flexAlignCenterCenter
'Next
'
''MSFlexGrid1.Rows = 3
'BarisData = 2
'
'If rsdata_mesin.BOF Then
'    Exit Sub
'Else
'    rsdata_mesin.MoveFirst
'    Do While Not rsdata_mesin.EOF
'        BarisData = BarisData + 1
'        MSFlexGrid1.Rows = BarisData + 1
'        MSFlexGrid1.TextMatrix(BarisData, 0) = rsdata_mesin.Fields("nama_mesin")
'        cnomc.AddItem rsdata_mesin("nama_mesin")
'        rsdata_mesin.MoveNext
'    Loop
'End If
'
'Call hitung_kolom
'Call cek_kapasitas
'=====================
End Sub
