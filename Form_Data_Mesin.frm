VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Data_Mesin 
   Caption         =   "FORM DATA MESIN"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ctutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtnamamesin 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtkapasitas 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton csimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton chapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nama Mesin"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Kapasitas"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   690
   End
End
Attribute VB_Name = "Form_Data_Mesin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
For Each A In Me
    If TypeOf A Is TextBox Then A.Text = ""
Next A
End Sub
Sub tampil()
If rsdata_mesin.State = 0 Then rsdata_mesin.Open
Me.txtnamamesin.Text = rsdata_mesin.Fields("nama_mesin")
Me.txtkapasitas.Text = rsdata_mesin.Fields("kapasitas")
End Sub

Sub tampilgrid()
lihat = "select*from data_mesin "
Set rsdata_mesin = conn.Execute(lihat)
Set DataGrid1.DataSource = rsdata_mesin.DataSource
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
txtnamamesin.SetFocus
Else
ubah = "UPDATE data_mesin SET nama_mesin='" & txtnamamesin.Text & "', kapasitas='" & txtkapasitas.Text & "' where id_mesin='" & DataGrid1.Columns("id_mesin") & "'"
Set rsdata_mesin = conn.Execute(ubah)
tampilgrid
Call bersih
'Call id_mesin
cedit.Caption = "EDIT"
csimpan.Enabled = True
Me.txtnamamesin.SetFocus
End If
End Sub

Private Sub chapus_Click()
X = MsgBox("Yakin Mau Dihapus...?", vbYesNo + vbInformation, "Hapus Data")
If X = vbYes Then
hapus = "delete from data_mesin where nama_mesin='" & DataGrid1.Columns("nama_mesin") & "'"
Set rsdata_mesin = conn.Execute(hapus)
Call tampilgrid
Call bersih
'Call id_mesin
End If
End Sub

Private Sub csimpan_Click()
simpan = "INSERT INTO data_mesin (nama_mesin,kapasitas) VALUES('" & txtnamamesin.Text & "','" & txtkapasitas.Text & "')"
Set rsdata_mesin = conn.Execute(simpan)
Call bersih
Call tampilgrid
Me.txtnamamesin.SetFocus
'Call id_mesin
End Sub

Private Sub ctutup_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call db
conn.CursorLocation = adUseClient
rsdata_mesin.Open "select*from data_mesin", conn
tampilgrid
Me.txtnamamesin.SetFocus
'Call id_mesin

End Sub


Private Sub txtkapasitas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
csimpan_Click
txtnamamesin.SetFocus
End If

End Sub

Private Sub txtnamamesin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtkapasitas.SetFocus
End If

End Sub
