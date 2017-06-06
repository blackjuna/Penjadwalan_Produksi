VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Cetak_LF 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_category 
      Height          =   315
      ItemData        =   "Form_Cetak_LF.frx":0000
      Left            =   1440
      List            =   "Form_Cetak_LF.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cnomachine 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dttanggal1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100139009
      CurrentDate     =   41729
   End
   Begin MSComCtl2.DTPicker dttanggal2 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100139009
      CurrentDate     =   41729
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label5 
      Caption         =   "Category"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "No. Machine"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CETAK BERDASARKAN"
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
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Width           =   120
   End
End
Attribute VB_Name = "Form_Cetak_LF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
    Select Case cmb_category.Text
        Case "SWG"
            strcategory = "and {completion_slip.category} is null"
        Case "DJG"
            strcategory = "and {completion_slip.category} =2"
        Case "SMG"
            strcategory = "and {completion_slip.category} =3"
        Case ""
            MsgBox "Mohon pilih category terlebih dahulu", vbCritical + vbOKOnly, "Warning"
            Exit Sub
    End Select
    
    With CrystalReport2
        .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
        '.SQLQuery = "select * from completion_slip where completion_slip.finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' " & _
            "and completion_slip.finish_date<='" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' " & _
            "and completion_slip.proses_2='" & cnomachine.Text & "' "
        .SelectionFormula = "{completion_slip.proses_2} ='" & cnomachine.Text & "' and " & _
            "{completion_slip.finish_date} >= DateTime('" & Format(dttanggal1.Value, "YYYY/mm/dd hh:mm:ss") & "') and " & _
            "{completion_slip.finish_date} <= DateTime('" & Format(dttanggal2.Value, "YYYY/mm/dd hh:mm:ss") & "') "
        Select Case cmb_category.Text
            Case "SWG"
                .ReportFileName = App.Path & "\list_foreman.rpt"
            Case "DJG"
                .ReportFileName = App.Path & "\list_foreman_djg.rpt"
            Case "SMG"
                .ReportFileName = App.Path & "\list_foreman_smg.rpt"
        End Select
        
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
rsdata_mesin.Open "select*from data_mesin", conn
cnomachine.AddItem "All"
Do While Not rsdata_mesin.EOF
cnomachine.AddItem rsdata_mesin("nama_mesin")
rsdata_mesin.MoveNext
Loop
End Sub

Private Sub Form_Load()
Me.dttanggal1.Value = Date
Me.dttanggal2.Value = Date

End Sub

