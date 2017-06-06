VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Data_Kode 
   Caption         =   "DATA KODE PART "
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   12990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tclass 
      Height          =   375
      Left            =   4800
      TabIndex        =   39
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox tsize 
      Height          =   375
      Left            =   1080
      TabIndex        =   37
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ComboBox tthickness 
      Height          =   315
      Left            =   4800
      TabIndex        =   36
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox tmarking_ir_lokal 
      Height          =   375
      Left            =   9120
      TabIndex        =   30
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox tmarking_or_eksport 
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox tmarking_ir_eksport 
      Height          =   375
      Left            =   9120
      TabIndex        =   28
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox tmarking_or_lokal 
      Height          =   375
      Left            =   9120
      TabIndex        =   27
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton ctutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   8400
      TabIndex        =   26
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton chapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   6960
      TabIndex        =   25
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   4080
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton csimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   2640
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukkan No Part Yang Dicari"
      Height          =   855
      Left            =   8040
      TabIndex        =   19
      Top             =   3000
      Width           =   4815
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.ComboBox cmbproses 
      Height          =   315
      ItemData        =   "Form_Data_Kode.frx":0000
      Left            =   1080
      List            =   "Form_Data_Kode.frx":0002
      TabIndex        =   18
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cmbfiller 
      Height          =   315
      Left            =   4800
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox cmbir 
      Height          =   315
      Left            =   4800
      TabIndex        =   16
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox cmbor 
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cmbhoop 
      Height          =   315
      Left            =   4800
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbtype 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtjic 
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtsize 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtnopart 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7223
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   195
      Left            =   3360
      TabIndex        =   40
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   3360
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Thickness"
      Height          =   195
      Left            =   3360
      TabIndex        =   35
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stamp Marking OR Lokal"
      Height          =   195
      Left            =   7080
      TabIndex        =   34
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Stamp Marking IR Eksport"
      Height          =   195
      Left            =   7080
      TabIndex        =   33
      Top             =   2400
      Width           =   1860
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Stamp Marking OR Eksport"
      Height          =   195
      Left            =   7080
      TabIndex        =   32
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Stamp Marking IR Lokal"
      Height          =   195
      Left            =   7080
      TabIndex        =   31
      Top             =   1440
      Width           =   1710
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Proses"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "JIC"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Size / Class"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Inner Ring"
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Outer Ring"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Hoop"
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Filler"
      Height          =   195
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No Part"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "Form_Data_Kode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String

Sub combo()
Me.cmbproses.AddItem "FINISH GOOD"
Me.cmbproses.AddItem "ASSEMBLING"
Me.cmbproses.AddItem "MARKING"
Me.cmbproses.AddItem "WINDING1"
Me.cmbproses.AddItem "WINDING2"
Me.cmbproses.AddItem "WINDING3"
Me.cmbproses.AddItem "WINDING4"
Me.cmbproses.AddItem "WINDING5"
Me.cmbproses.AddItem "WINDING6"
Me.cmbproses.AddItem "TURNING"
Me.cmbproses.AddItem "ELECTROPLATING"
Me.cmbproses.AddItem "GROVE OR"
Me.cmbproses.AddItem "GROVE IR"
Me.cmbtype.AddItem "Complete"
Me.cmbtype.AddItem "Basic"
Me.cmbtype.AddItem "OR + Basic"
Me.cmbtype.AddItem "IR + Basic"
Me.cmbtype.AddItem "Ring OR"
Me.cmbtype.AddItem "Ring IR"
Me.cmbir.AddItem "CS"
Me.cmbir.AddItem "304"
Me.cmbir.AddItem "304L"
Me.cmbir.AddItem "316"
Me.cmbir.AddItem "316L"
Me.cmbir.AddItem "TITANIUM"
Me.cmbir.AddItem "MONEL"
Me.cmbir.AddItem "SPCC"
Me.cmbir.AddItem "CU"
Me.cmbir.AddItem "321"
Me.cmbir.AddItem "410"
Me.cmbor.AddItem "CS"
Me.cmbor.AddItem "304"
Me.cmbor.AddItem "304L"
Me.cmbor.AddItem "316"
Me.cmbor.AddItem "316L"
Me.cmbor.AddItem "TITANIUM"
Me.cmbor.AddItem "MONEL"
Me.cmbor.AddItem "SPCC"
Me.cmbor.AddItem "CU"
Me.cmbor.AddItem "321"
Me.cmbor.AddItem "410"
Me.cmbhoop.AddItem "304"
Me.cmbhoop.AddItem "304L"
Me.cmbhoop.AddItem "316"
Me.cmbhoop.AddItem "316L"
Me.cmbhoop.AddItem "TITANIUM"
Me.cmbhoop.AddItem "MONEL"
Me.cmbfiller.AddItem "GRP"
Me.cmbfiller.AddItem "ASBES"
Me.cmbfiller.AddItem "PTFE / TF"
Me.cmbfiller.AddItem "NA"
Me.cmbfiller.AddItem "CE"
Me.tthickness.AddItem "3.0"
Me.tthickness.AddItem "3.2"
Me.tthickness.AddItem "4.5"
Me.tthickness.AddItem "6.4"

Call bersih

End Sub

Sub bersih()
For Each A In Me
    If TypeOf A Is TextBox Then A.Text = ""
    If TypeOf A Is ComboBox Then A.Text = "- PILIH -"
Next A
End Sub
Sub tampil()
Me.txtnopart = rscode(0)
Me.cmbproses = rscode(1)
Me.txtjic = rscode(2)
Me.txtsize = rscode(3)
Me.cmbtype = rscode(4)
Me.cmbir = rscode(5)
Me.cmbor = rscode(6)
Me.cmbhoop = rscode(7)
Me.cmbfiller = rscode(8)
Me.tmarking_or_lokal = rscode(9)
Me.tmarking_ir_lokal = rscode(10)
Me.tmarking_or_eksport = rscode(11)
Me.tmarking_ir_eksport = rscode(12)
Me.tthickness = rscode(13)
Me.tsize = rscode(14)
Me.tclass = rscode(15)
End Sub

Sub tampilgrid()
lihat = "select*from code"
Set rscode = conn.Execute(lihat)
Set DataGrid1.DataSource = rscode.DataSource
End Sub

Private Sub cbatal_Click()
Call bersih
cedit.Caption = "EDIT"
csimpan.Enabled = True
End Sub

Private Sub cedit_Click()
If cedit.Caption = "EDIT" Then
cedit.Caption = "UPDATE"
csimpan.Enabled = False
Call tampil
Else
ubah = "UPDATE code SET no_part='" & txtnopart.Text & "', proses='" & cmbproses.Text & "',  jic='" & txtjic.Text & "', size='" & txtsize.Text & "', type='" & cmbtype.Text & "', inner_ring='" & cmbir.Text & "', outer_ring='" & cmbor.Text & "', hoop='" & cmbhoop.Text & "', filler='" & cmbfiller.Text & "',marking_stamp_lokal_or='" & tmarking_or_lokal.Text & "',marking_stamp_lokal_ir='" & tmarking_ir_lokal.Text & "',marking_stamp_eksport_or='" & tmarking_or_eksport.Text & "',marking_stamp_eksport_ir='" & tmarking_ir_eksport.Text & "',thickness='" & tthickness.Text & "',size_2='" & tsize.Text & "'  where no_part='" & txtnopart.Text & "'"
Set rscode = conn.Execute(ubah)
tampilgrid
Call bersih
cedit.Caption = "EDIT"
csimpan.Enabled = True
End If

End Sub

Private Sub chapus_Click()
x = MsgBox("Yakin Mau Dihapus...?", vbYesNo + vbInformation, "Hapus Data")
If x = vbYes Then
hapus = "delete from code where no_part='" & DataGrid1.Columns("no_part") & "'"
Set rscode = conn.Execute(hapus)
MsgBox ("DATA SUDAH TERHAPUS")
Call tampilgrid
Call bersih
End If
End Sub

Private Sub cmbfiller_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tthickness.SetFocus
End If

End Sub

Private Sub cmbhoop_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbfiller.SetFocus
End If

End Sub

Private Sub cmbir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbor.SetFocus
End If

End Sub

Private Sub cmbor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbhoop.SetFocus
End If

End Sub

Private Sub cmbproses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtjic.SetFocus
End If
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
tsize.SetFocus
End If

End Sub

Private Sub csearch_Click()
cari = "select * from code where no_part='" & txtcari.Text & "'"
Set rscode = conn.Execute(cari)
If rscode.EOF Then
    MsgBox "Data tidak ditemukan"
    txtcari.Text = ""
    txtcari.SetFocus
Else
    Call tampil
    cedit.Caption = "UPDATE"
    csimpan.Enabled = False
End If

End Sub

Private Sub csimpan_Click()
If rscode.State = 1 Then rscode.Close
strsql = "Select no_part from code where no_part='" & txtnopart.Text & "' and deleted=0"
rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
If rscode.EOF Then
    simpan = "INSERT INTO code (no_part,proses,jic,size,type,inner_ring,outer_ring,hoop,filler,marking_stamp_lokal_or," & _
        "marking_stamp_lokal_ir,marking_stamp_eksport_or,marking_stamp_eksport_ir,thickness," & _
        "size_2,class,deleted) VALUES('" & txtnopart.Text & "','" & cmbproses.Text & "'," & _
        "'" & txtjic.Text & "','" & txtsize.Text & "','" & cmbtype.Text & "'," & _
        "'" & cmbir.Text & "','" & cmbor.Text & "','" & cmbhoop.Text & "'," & _
        "'" & cmbfiller.Text & "','" & tmarking_or_lokal.Text & "'," & _
        "'" & tmarking_ir_lokal.Text & "','" & tmarking_or_eksport.Text & "'," & _
        "'" & tmarking_ir_eksport.Text & "','" & tthickness.Text & "'," & _
        "'" & tsize.Text & "','" & tclass.Text & "',0)"
    Set rscode = conn.Execute(simpan)
    Call bersih
    Call tampilgrid
Else
    
    MsgBox "Maaf, kode part sudah ada.", vbOKOnly + vbCritical, "Informasi"
    txtnopart.SetFocus
End If
End Sub

Private Sub ctutup_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call db
conn.CursorLocation = adUseClient
rscode.Open "select*from code", conn
tampilgrid
Call combo
txtnopart.SetFocus
End Sub


Private Sub Text1_Change()

End Sub

Private Sub tclass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
tmarking_or_lokal.SetFocus
End If

End Sub

Private Sub tmarking_ir_eksport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    csimpan_Click
    txtnopart.SetFocus
End If
End Sub

Private Sub tmarking_ir_lokal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tmarking_or_eksport.SetFocus
End If
End Sub

Private Sub tmarking_or_eksport_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tmarking_ir_eksport.SetFocus
End If
End Sub

Private Sub tmarking_or_lokal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tmarking_ir_lokal.SetFocus
End If
End Sub

Private Sub tsize_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbir.SetFocus
End If

End Sub

Private Sub tthickness_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
tclass.SetFocus

End If

End Sub

Private Sub txtjic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtsize.SetFocus

End If

End Sub

Private Sub txtnopart_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbproses.SetFocus
End If
End Sub

Private Sub txtsize_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmbtype.SetFocus
End If

End Sub
