VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormData_Kode_DJG 
   Caption         =   "DATA KODE PART DJG"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   23
      Top             =   2400
      Width           =   3495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtnopart 
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox txtjic 
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox cmbfiller 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   1920
      Width           =   3495
   End
   Begin VB.ComboBox cmbproses 
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukkan Description atau Part No"
      Height          =   855
      Left            =   5760
      TabIndex        =   6
      Top             =   2520
      Width           =   4935
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton csimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   10920
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   13800
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   12360
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton chapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   15240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton ctutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   16680
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox tsize 
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9340
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
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "FILLER"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "TYPE DJG"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "JIC CODE"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PART NO"
      Height          =   195
      Left            =   5760
      TabIndex        =   18
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PLAT"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "DESCRIPTION"
      Height          =   195
      Left            =   5760
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "PROCESS"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "SIZE"
      Height          =   195
      Left            =   5760
      TabIndex        =   14
      Top             =   960
      Width           =   360
   End
End
Attribute VB_Name = "FormData_Kode_DJG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Form_Load()
    Dim chr_1, chr_2 As String
    Dim asc_1, asc_2 As Integer
    
    chr_1 = Left$(UCase(Text1.Text), 1)
    chr_2 = Right$(UCase(Text1.Text), 1)
    
    If chr_2 = "Z" And chr_1 <> "Z" Then
        asc_1 = Asc(chr_1)
        asc_1 = asc_1 + 1
        chr_2 = "A"
        chr_1 = Chr(asc_1)
        Label3.Caption = chr_1 + chr_2
    ElseIf chr_1 = "Z" And chr_2 = "Z" Then
        MsgBox "Kode Size sudah habis. Mohon segera hubungi admministrator!", vbOKOnly + vbCritical, "WARNING!!!"
    Else
        asc_1 = Asc(chr_1)
        asc_2 = Asc(chr_2)
        asc_2 = asc_2 + 1
        chr_1 = Chr(asc_1)
        chr_2 = Chr(asc_2)
        Label3.Caption = chr_1 + chr_2
    End If
End Sub
