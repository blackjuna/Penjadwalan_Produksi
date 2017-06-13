VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Form_Utama 
   Caption         =   "Main Form Completion Slip SWG"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   14880
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txt_coc_no 
      Height          =   285
      Left            =   5040
      TabIndex        =   86
      Text            =   "Text1"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt_coc 
      Height          =   285
      Left            =   6840
      TabIndex        =   85
      Top             =   10440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame6 
      Caption         =   "Certificate"
      Height          =   4335
      Left            =   9120
      TabIndex        =   83
      Top             =   4800
      Width           =   2295
      Begin MSComctlLib.ListView lv_cert 
         Height          =   3975
         Left            =   120
         TabIndex        =   84
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7011
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmd_add_cert 
      Caption         =   "Add Certificate"
      Height          =   495
      Left            =   7200
      TabIndex        =   82
      Top             =   9360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   1020
      Left            =   10440
      Picture         =   "Form_Utama.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   975
      TabIndex        =   66
      Top             =   9240
      Width           =   1035
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   65
      Top             =   4800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton crefresh 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   4440
      TabIndex        =   64
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox tclass 
      Height          =   405
      Left            =   4440
      TabIndex        =   45
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tsize2 
      Height          =   405
      Left            =   3000
      TabIndex        =   44
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tbatchnumber 
      Height          =   405
      Left            =   1560
      TabIndex        =   43
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tlotnumber 
      Height          =   405
      Left            =   120
      TabIndex        =   42
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   3000
      TabIndex        =   39
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Masukkan No Slip Yang Dicari"
      Height          =   855
      Left            =   11520
      TabIndex        =   36
      Top             =   9240
      Width           =   3615
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   2400
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   1560
      TabIndex        =   30
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton chapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   11520
      TabIndex        =   22
      Top             =   7440
      Width           =   3615
      Begin VB.ComboBox cproses2 
         Height          =   315
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox tproses1 
         Height          =   405
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox tproses3 
         Height          =   405
         Left            =   1440
         TabIndex        =   23
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "From Location"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "End Process"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "To Location"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6495
      Left            =   11520
      TabIndex        =   4
      Top             =   840
      Width           =   3615
      Begin VB.ComboBox cmb_hoop 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   4680
         Width           =   2055
      End
      Begin VB.ComboBox cmb_filler 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   5160
         Width           =   2055
      End
      Begin VB.ComboBox cmb_outer 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   4200
         Width           =   2055
      End
      Begin VB.ComboBox cmb_inner 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Txt_cert 
         Height          =   405
         Left            =   1440
         TabIndex        =   74
         Top             =   6000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txt_Heat_cert 
         Height          =   405
         Left            =   1440
         TabIndex        =   72
         Top             =   5520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ComboBox tthickness 
         Height          =   315
         ItemData        =   "Form_Utama.frx":0B9D
         Left            =   1440
         List            =   "Form_Utama.frx":0BAD
         TabIndex        =   67
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox tmarking_ir 
         Height          =   405
         Left            =   1440
         TabIndex        =   33
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox tmarking_or 
         Height          =   405
         Left            =   1440
         TabIndex        =   31
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox tfiller 
         Height          =   405
         Left            =   1440
         TabIndex        =   19
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox thoop 
         Height          =   405
         Left            =   1440
         TabIndex        =   18
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox tor 
         Height          =   405
         Left            =   1440
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox tir 
         Height          =   405
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Cert. No. Mat."
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   6120
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Heat"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   5640
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Filler Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Hoop Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   4680
         Width           =   720
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Outer Ring Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   69
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Inner Ring Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   3720
         Width           =   1065
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Thickness"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Marking IR"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Marking OR"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Outer Ring"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Filler"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Hoop"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inner Ring"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   6375
      Begin VB.ComboBox cshift 
         Height          =   315
         ItemData        =   "Form_Utama.frx":0BC5
         Left            =   1800
         List            =   "Form_Utama.frx":0BC7
         TabIndex        =   55
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cnowinding 
         Height          =   315
         Left            =   1800
         TabIndex        =   54
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox tnote 
         Height          =   645
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox tcustomer 
         Height          =   405
         Left            =   1800
         TabIndex        =   14
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox tqty 
         Height          =   405
         Left            =   1800
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtdate_printed 
         Height          =   375
         Left            =   1800
         TabIndex        =   56
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96927745
         CurrentDate     =   41714
      End
      Begin MSComCtl2.DTPicker dtfinish_date 
         Height          =   375
         Left            =   1800
         TabIndex        =   57
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96927745
         CurrentDate     =   41714
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Date PPIC"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Input Date"
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Shift"
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "No MC Winding"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4695
      Begin VB.ComboBox cstatus 
         Height          =   315
         ItemData        =   "Form_Utama.frx":0BC9
         Left            =   1800
         List            =   "Form_Utama.frx":0BCB
         TabIndex        =   62
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox tjic 
         Height          =   405
         Left            =   1800
         TabIndex        =   49
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox tsize 
         Height          =   405
         Left            =   1800
         TabIndex        =   48
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox ttype 
         Height          =   405
         Left            =   1800
         TabIndex        =   47
         Top             =   2760
         Width           =   2655
      End
      Begin VB.ComboBox cno_part 
         Height          =   315
         Left            =   1800
         TabIndex        =   46
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox tno_so 
         Height          =   405
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox tno_slip 
         Height          =   405
         Left            =   1800
         TabIndex        =   20
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   240
         TabIndex        =   63
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No Part"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "JIC"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. SO"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Slip"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Label T_Showing_Records 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2520
      TabIndex        =   81
      Top             =   9840
      Width           =   90
   End
   Begin VB.Label T_Results 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   840
      TabIndex        =   80
      Top             =   9840
      Width           =   90
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Showing :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1680
      TabIndex        =   79
      Top             =   9840
      Width           =   705
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   78
      Top             =   9840
      Width           =   615
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto page :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   77
      Top             =   9480
      Width           =   840
   End
   Begin VB.Label T_Page 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1080
      MouseIcon       =   "Form_Utama.frx":0BCD
      MousePointer    =   99  'Custom
      TabIndex        =   76
      Top             =   9480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Caption         =   "AVAILABLE QUANTITY "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   41
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblavaqty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   40
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu pop_up_menu 
      Caption         =   "Pop Up menu"
      Visible         =   0   'False
      Begin VB.Menu reschedule 
         Caption         =   "Reschedule"
      End
   End
   Begin VB.Menu mn_master 
      Caption         =   "Master"
      Begin VB.Menu smn_swg 
         Caption         =   "Kode SWG"
      End
   End
   Begin VB.Menu mn_cetak 
      Caption         =   "Cetak"
      Begin VB.Menu smn_cs 
         Caption         =   "Completion Slip"
      End
      Begin VB.Menu smn_listforeman 
         Caption         =   "List Foreman"
      End
   End
End
Attribute VB_Name = "Form_Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strmenu As String
Public Strnoslip As String
Public strsql As String

Private Sub data_mesin()
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "Select * from data_mesin where nama_mesin like '%w%'", conn, adOpenDynamic, adLockOptimistic
    If Not rsdata_mesin.EOF Then rsdata_mesin.MoveFirst
    Do While Not rsdata_mesin.EOF
        cnowinding.AddItem rsdata_mesin("nama_mesin")
        rsdata_mesin.MoveNext
        
    Loop
End Sub

Sub Warna_List()
Dim i As Long

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).SubItems(8) = "Pending" Then 'Field Stok pada kolom 5
    ListView1.ListItems(i).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(1).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(2).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(3).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(4).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(5).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(6).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(7).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(8).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(9).ForeColor = vbRed
    ListView1.ListItems(i).ListSubItems(10).ForeColor = vbRed
    ElseIf ListView1.ListItems(i).SubItems(8) = "Partial" Then
    ListView1.ListItems(i).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(1).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(2).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(3).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(4).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(5).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(6).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(7).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(8).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(9).ForeColor = vbGreen
    ListView1.ListItems(i).ListSubItems(10).ForeColor = vbGreen
    ElseIf ListView1.ListItems(i).SubItems(8) = "Closed" Then
    ListView1.ListItems(i).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(1).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(2).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(3).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(4).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(5).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(6).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(7).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(8).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(9).ForeColor = vbBlue
    ListView1.ListItems(i).ListSubItems(10).ForeColor = vbBlue
    ElseIf ListView1.ListItems(i).SubItems(8) = "Reschedule" Then
    ListView1.ListItems(i).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(1).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(2).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(3).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(4).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(5).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(6).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(7).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(8).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(9).ForeColor = vbMagenta
    ListView1.ListItems(i).ListSubItems(10).ForeColor = vbMagenta
    Else
    ListView1.ListItems(i).ForeColor = vbBlack
    ListView1.ListItems(i).ListSubItems(1).ForeColor = vbBlack
    ListView1.ListItems(i).ListSubItems(2).ForeColor = vbBlack
    ListView1.ListItems(i).ListSubItems(3).ForeColor = vbBlack
    ListView1.ListItems(i).ListSubItems(4).ForeColor = vbBlack
    ListView1.ListItems(i).ListSubItems(5).ForeColor = vbBlack
    End If
Next

End Sub

Public Sub SetLV()
With ListView1
    .View = lvwReport
    .GridLines = True
    .MultiSelect = True
    .FullRowSelect = True
    .HotTracking = True
    .HoverSelection = True
    ' tambahkan kolom2 ke, , Judul,lebar,aligment
    .ColumnHeaders.Add 1, , "No Slip", 0
    .ColumnHeaders.Add 2, , "No Slip", 1700
    .ColumnHeaders.Add 3, , "No Sales Order", 1400
    .ColumnHeaders.Add 4, , "JIC", 2500
    .ColumnHeaders.Add 5, , "Size", 2200
    .ColumnHeaders.Add 6, , "Qty", 750
    .ColumnHeaders.Add 7, , "DT PPIC", 1000
    .ColumnHeaders.Add 8, , "Realisasi DT", 1100
    .ColumnHeaders.Add 9, , "Status", 1100
    .ColumnHeaders.Add 10, , "Qty Pending", 1100
    .ColumnHeaders.Add 11, , "Remarks Produksi", 3000
    .ColumnHeaders.Add 12, , "Proses Winding", 3000
    .ColumnHeaders.Add 13, , "id", 0
    
End With
End Sub
Sub TplGrid()
    Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 0 Then
'        rscompletion_slip.Open "select*from completion_slip order by id asc", conn
        rscompletion_slip.Open "select top 1000 *from completion_slip order by id desc", conn
    End If
'    lihat = "select * from completion_slip order by id asc"
    lihat = "select top 1000 * from completion_slip order by id desc"
    Set rscompletion_slip = conn.Execute(lihat)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!jic
    lst.SubItems(4) = rscompletion_slip!Size
    lst.SubItems(5) = rscompletion_slip!qty
    lst.SubItems(6) = rscompletion_slip!finish_date
    lst.SubItems(7) = rscompletion_slip!delivery_date
    lst.SubItems(8) = rscompletion_slip!status
    lst.SubItems(9) = IIf(IsNull(rscompletion_slip!qty_pending), 0, rscompletion_slip!qty_pending)
    lst.SubItems(10) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(11) = rscompletion_slip!proses_2
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
End Sub
Sub warning_reschedule()
cari = "select * from completion_slip where status='Reschedule'"
Set rscompletion_slip = conn.Execute(cari)

If rscompletion_slip.EOF Then
    Picture1.Visible = False
Else
    Picture1.Visible = True
End If

End Sub
Sub no_slip()
Dim thn As String

thn = Format(Date, "YYYY")
If rscompletion_slip.State = 1 Then rscompletion_slip.Close
rscompletion_slip.Open "select date_printed,lot_number,batch_number,no_slip from completion_slip where deleted =0 and category is null order by id asc", conn, adOpenDynamic, adLockOptimistic

If rscompletion_slip.RecordCount = 0 Then
    Me.tlotnumber.Text = "0001"
Else
    rscompletion_slip.MoveLast
    last_date = Format(rscompletion_slip.Fields("date_printed"), "YYYY/mm/dd")
    date_now = Format(Date, "YYYY/mm/dd")
    If date_now = last_date Then
        qry_lot_number = rscompletion_slip.Fields("lot_number")
        no_lot = qry_lot_number
        lot_number = no_lot
        Me.tlotnumber.Text = lot_number
    Else
        qry_lot_number = rscompletion_slip.Fields("lot_number")
        no_lot = qry_lot_number
        lot_number = (Val(Right(no_lot, 4)) + 1)
        If lot_number < 10 Then
            Me.tlotnumber.Text = "000" & lot_number
        ElseIf lot_number < 100 Then
            Me.tlotnumber.Text = "00" & lot_number
        ElseIf lot_number < 1000 Then
            Me.tlotnumber.Text = "0" & lot_number
        Else
            Me.tlotnumber.Text = lot_number
        End If
    End If
End If

If rscompletion_slip.RecordCount = 0 Then
    Me.tbatchnumber.Text = "001"
Else
    rscompletion_slip.MoveLast
    last_date = Format(rscompletion_slip!date_printed, "YYYY/mm/dd")
    date_now = Format(Date, "YYYY/mm/dd")
    If Weekday(date_now) = vbMonday And date_now = last_date Then
        qry_batch_number = rscompletion_slip!batch_number
        batch_number = qry_batch_number
        Me.tbatchnumber.Text = batch_number
    ElseIf Weekday(date_now) = vbMonday And date_now > last_date Then
        qry_batch_number = rscompletion_slip!batch_number
        no_batch = qry_batch_number
        batch_number = Val(Right(no_batch, 4)) + 1
        If batch_number < 10 Then
            Me.tbatchnumber.Text = "00" & batch_number
        ElseIf batch_number < 100 Then
            Me.tbatchnumber.Text = "0" & batch_number
        Else
            Me.tbatchnumber.Text = batch_number
        End If
    Else
        qry_batch_number = rscompletion_slip.Fields("batch_number")
        batch_number = qry_batch_number
        Me.tbatchnumber.Text = batch_number
    End If
End If

    
If rscompletion_slip.RecordCount = 0 Then
    Me.tno_slip.Text = "2014-001-0000001"
Else
    rscompletion_slip.MoveLast
    Z = Val(Mid(rscompletion_slip!no_slip, 11, 7)) + 1
    If Z < 10 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "000000" & Z
    ElseIf Z < 100 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "00000" & Z
    ElseIf Z < 1000 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "0000" & Z
    ElseIf Z < 10000 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "000" & Z
    ElseIf Z < 100000 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "00" & Z
    ElseIf Z < 1000000 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & "0" & Z
    ElseIf Z < 10000000 Then
        Me.tno_slip.Text = thn & "-" & tlotnumber.Text & "-" & Z
    End If
End If
'rscompletion_slip.Close
End Sub
Sub bersih()
For Each A In Me
    If TypeOf A Is TextBox Then A.Text = ""
    If TypeOf A Is ComboBox Then
        If Not A.Style = 2 Then
            A.Text = ""
        Else
            A.ListIndex = -1
        End If
    End If
Next A
End Sub
Sub tampil()

cari = "select * from completion_slip where no_slip='" & ListView1.SelectedItem.Text & "'"
Set rscompletion_slip = conn.Execute(cari)
    If Not rscompletion_slip.EOF Then
        Me.tno_slip.Text = rscompletion_slip.Fields("no_slip")
        Me.tno_so = rscompletion_slip.Fields("no_so")
        Me.cnowinding = rscompletion_slip.Fields("proses_2")
        Me.dtdate_printed = rscompletion_slip.Fields("date_printed")
        Me.dtfinish_date = rscompletion_slip.Fields("finish_date")
        Me.cshift.Text = rscompletion_slip.Fields("shift")
        Me.cstatus.Text = rscompletion_slip.Fields("status")
        Me.cno_part.Text = rscompletion_slip.Fields("no_part")
        Me.tjic.Text = rscompletion_slip.Fields("jic")
        Me.tsize.Text = rscompletion_slip.Fields("size")
        Me.ttype.Text = rscompletion_slip.Fields("type")
        Me.tqty.Text = rscompletion_slip.Fields("qty")
        Me.tcustomer.Text = rscompletion_slip.Fields("customer")
        Me.tnote.Text = rscompletion_slip.Fields("note")
        Me.tir.Text = rscompletion_slip.Fields("inner_ring")
        Me.tor.Text = rscompletion_slip.Fields("outer_ring")
        Me.thoop.Text = rscompletion_slip.Fields("hoop")
        Me.tfiller.Text = rscompletion_slip.Fields("filler")
        Me.tmarking_or.Text = rscompletion_slip.Fields("marking_stamp_or")
        Me.tmarking_ir.Text = rscompletion_slip.Fields("marking_stamp_ir")
        Me.tthickness.Text = rscompletion_slip.Fields("thickness")
        Me.tproses1.Text = rscompletion_slip.Fields("proses_1")
        Me.cproses2.Text = Me.cnowinding.Text
        Me.tproses3.Text = rscompletion_slip.Fields("proses_3")
        If rscompletion_slip!inner_ring_cert = "" Then
            cmb_inner.ListIndex = -1
        Else
            Me.cmb_inner.Text = rscompletion_slip.Fields("inner_ring_cert")
        End If
        
        If rscompletion_slip!outer_ring_cert = "" Then
            cmb_inner.ListIndex = -1
        Else
            Me.cmb_outer.Text = rscompletion_slip.Fields("outer_ring_cert")
        End If
        
        If rscompletion_slip!hoop_cert = "" Then
            cmb_inner.ListIndex = -1
        Else
            Me.cmb_hoop.Text = rscompletion_slip.Fields("hoop_cert")
        End If
        
        If rscompletion_slip!filler_cert = "" Then
            cmb_inner.ListIndex = -1
        Else
            Me.cmb_filler.Text = rscompletion_slip.Fields("filler_cert")
        End If
    End If

End Sub

Sub tampilgrid()
If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select*from completion_slip", conn
End If

lihat = "select no_slip,no_so,jic,size,qty,date_printed,finish_date,lot_number,batch_number,status,remarks_produksi,proses_2,shift,no_part,type,customer,note,inner_ring,outer_ring,hoop,filler,marking_stamp_or,marking_stamp_ir,thickness,proses_1,proses_3 from completion_slip"
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
With DataGrid1
    .Columns(0).Width = 1500
    .Columns(0).Caption = "No Slip"
    .Columns(1).Width = 1000
    .Columns(1).Caption = "No SO"
    .Columns(2).Width = 2700
    .Columns(2).Caption = "JIC"
    .Columns(3).Width = 3000
    .Columns(3).Caption = "Size"
    .Columns(4).Width = 500
    .Columns(4).Caption = "Qty"
    .Columns(5).Width = 1100
    .Columns(5).Caption = "Date Printed"
    .Columns(6).Width = 1100
    .Columns(6).Caption = "DT PPIC"
    .Columns(7).Width = 0
    .Columns(8).Width = 0
    .Columns(9).Width = 1000
    .Columns(9).Caption = "Status"
    .Columns(10).Width = 2000
    .Columns(10).Caption = "Remarks Produksi"
    .Columns(11).Width = 0
    .Columns(12).Width = 0
    .Columns(13).Width = 0
    .Columns(14).Width = 0
    .Columns(15).Width = 0
    .Columns(16).Width = 0
    .Columns(17).Width = 0
    .Columns(18).Width = 0
    .Columns(19).Width = 0
    .Columns(20).Width = 0
    .Columns(21).Width = 0
    .Columns(22).Width = 0
    .Columns(23).Width = 0
    .Columns(24).Width = 0
    .Columns(25).Width = 0
End With
DataGrid1.Refresh
End Sub



Private Sub cbatal_Click()
cedit.Caption = "EDIT"
Call bersih
Call no_slip
lblavaqty.Caption = 0
Call Build_Results
Call Warna_List
Call warning_reschedule
tno_so.SetFocus
End Sub

Private Sub cedit_Click()
If cedit.Caption = "EDIT" Then
cedit.Caption = "UPDATE"
Call tampil
tno_so.SetFocus
Else
ubah = "UPDATE completion_slip SET no_so='" & tno_so.Text & "'," & _
    "date_printed='" & Format(dtdate_printed.Value, "YYYY/mm/dd") & "'," & _
    "finish_date='" & Format(dtfinish_date.Value, "YYYY/mm/dd") & "'," & _
    "shift='" & cshift.Text & "',status='" & cstatus.Text & "'," & _
    "no_part='" & cno_part.Text & "',jic='" & tjic.Text & "'," & _
    "size='" & tsize.Text & "',type='" & ttype.Text & "',qty='" & tqty.Text & "'," & _
    "customer='" & tcustomer.Text & "',note='" & tnote.Text & "'," & _
    "inner_ring='" & tir.Text & "',outer_ring='" & tor.Text & "'," & _
    "hoop='" & thoop.Text & "',filler='" & tfiller.Text & "'," & _
    "marking_stamp_or='" & tmarking_or.Text & "'," & _
    "marking_stamp_ir='" & tmarking_ir.Text & "'," & _
    "thickness='" & tthickness.Text & "',proses_1='" & tproses1.Text & "'," & _
    "proses_2='" & cproses2.Text & "',proses_3='" & tproses3.Text & "'," & _
    "inner_ring_cert='" & cmb_inner.Text & "',outer_ring_cert='" & cmb_outer.Text & "'," & _
    "hoop_cert='" & cmb_hoop.Text & "',filler_cert='" & cmb_filler.Text & "' " & _
    "where no_slip='" & tno_slip.Text & "'"
Set rscompletion_slip = conn.Execute(ubah)
Call coc_no
Call save_certificate
Call bersih
cedit.Caption = "EDIT"
Call no_slip

Call Build_Results
Call Warna_List
Call warning_reschedule
End If

End Sub

Private Sub chapus_Click()
x = MsgBox("Yakin Mau Dihapus...?", vbYesNo + vbInformation, "Hapus Data")
If x = vbYes Then
hapus = "delete from completion_slip where no_slip='" & ListView1.SelectedItem.Text & "'"
Set rscompletion_slip = conn.Execute(hapus)
Call bersih
Call no_slip
Call Build_Results
Call Warna_List
Call warning_reschedule

End If

End Sub

Private Sub cmb_filler_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            If KeyCode = vbKeyReturn Then
                cproses2.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmb_hoop_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            If KeyCode = vbKeyReturn Then
                cmb_filler.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmb_inner_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            If KeyCode = vbKeyReturn Then
                cmb_outer.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmb_outer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            If KeyCode = vbKeyReturn Then
                cmb_hoop.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmd_add_cert_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Mohon pilih no cs terlebih dahulu !", vbInformation + vbInformation, "Information"
    Else
        Form_add_cer.Show vbModal
    End If
End Sub

Private Sub cno_part_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    cari = "select * from code where no_part='" & cno_part.Text & "' and deleted=0"
    Set rscode = conn.Execute(cari)
        If Not rscode.EOF Then
            tjic.Text = rscode.Fields("jic")
            tsize.Text = rscode.Fields("size")
            ttype.Text = rscode.Fields("type")
            tir.Text = rscode.Fields("inner_ring")
            tor.Text = rscode.Fields("outer_ring")
            thoop.Text = rscode.Fields("hoop")
            tfiller.Text = rscode.Fields("filler")
            tsize2.Text = rscode.Fields("size_2")
            tclass.Text = rscode.Fields("class")
            If Mid(tno_so.Text, 1, 1) = "F" Or Mid(tno_so.Text, 1, 1) = "f" Then
                tmarking_or.Text = "3STAR" & rscode.Fields("marking_stamp_lokal_or")
                tmarking_ir.Text = "3STAR" & rscode.Fields("marking_stamp_lokal_ir")
                
            ElseIf Mid(tno_so.Text, 1, 1) = "E" Or Mid(tno_so.Text, 1, 1) = "e" Then
                tmarking_or.Text = rscode.Fields("marking_stamp_eksport_or")
                tmarking_ir.Text = rscode.Fields("marking_stamp_eksport_ir")
                
            Else
                tmarking_or.Text = rscode.Fields("marking_stamp_lokal_or")
                tmarking_ir.Text = rscode.Fields("marking_stamp_lokal_ir")
                
            End If
            tthickness.Text = rscode.Fields("thickness")
        Else
            x = MsgBox("No Part Belum Diregistrasi, Registrasi Sekarang ...?", vbYesNo + vbInformation, "Warning")
            If x = vbYes Then
                Form_Data_Kode.Show
                Else
                cno_part.SetFocus
            End If
        End If
    tjic.Enabled = False
    tsize.Enabled = False
    ttype.Enabled = False
    tir.Enabled = False
    tor.Enabled = False
    thoop.Enabled = False
    tfiller.Enabled = False
    Call alokasi_mc_winding
    cstatus.SetFocus
End If

End Sub

Private Sub cnowinding_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Me.cproses2.Text = Me.cnowinding.Text
    dtfinish_date.SetFocus
End If
End Sub

Private Sub completion_slip_Click()
Form_Status_CS.Show
End Sub

Private Sub cproses2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If rschkcode.State = 1 Then rschkcode.Close
    Set rschkcode = New ADODB.Recordset
    sql = "Select no_slip from completion_slip where " & _
        "no_slip='" & tno_slip.Text & "' and deleted=0"
    rschkcode.Open sql, conn, adOpenDynamic, adLockOptimistic
    If Not rschkcode.EOF Then
        MsgBox "No Slip Sudah ada"
        cproses2.SetFocus
        rschkcode.Close
        Set rschkcode = Nothing
    Else
        simpan = "INSERT INTO completion_slip (no_slip,no_so,date_printed,delivery_date,finish_date," & _
        "lot_number,batch_number,proses_1,proses_2,proses_3,no_part,shift,jic,size,type,inner_ring,outer_ring," & _
        "hoop,filler," & _
        "marking_stamp_or,marking_stamp_ir,qty,customer,status,note,thickness,deleted,inner_ring_cert," & _
        "outer_ring_cert,hoop_cert,filler_cert) VALUES " & _
        "('" & tno_slip.Text & "'," & _
        "'" & tno_so.Text & "','" & Format(dtdate_printed.Value, "YYYY/mm/dd") & "'," & _
        "'','" & Format(dtfinish_date.Value, "YYYY/mm/dd") & "'," & _
        "'" & tlotnumber.Text & "','" & tbatchnumber.Text & "'," & _
        "'" & tproses1.Text & "','" & cproses2.Text & "','" & tproses3.Text & "'," & _
        "'" & cno_part.Text & "','" & cshift.Text & "','" & tjic.Text & "'," & _
        "'" & tsize.Text & "','" & ttype.Text & "','" & tir.Text & "'," & _
        "'" & tor.Text & "','" & thoop.Text & "','" & tfiller.Text & "'," & _
        "'" & tmarking_or.Text & "','" & tmarking_ir.Text & "'," & _
        "'" & tqty.Text & "','" & tcustomer.Text & "','" & cstatus.Text & "'," & _
        "'" & tnote.Text & "','" & tthickness.Text & "',0," & _
        "'" & cmb_inner.Text & "','" & cmb_outer.Text & "','" & cmb_hoop.Text & "'," & _
        "'" & cmb_filler.Text & "')"
        Set rscompletion_slip = conn.Execute(simpan)
        Call coc_no
        Call save_certificate
        Call bersih
        tno_so.SetFocus
        Call no_slip
        Call Build_Results
        Call Warna_List
        Call warning_reschedule
        rschkcode.Close
        Set rschkcode = Nothing
    End If
End If

'Call hitung_kolom_2
'Call cek_kapasitas_2
lblavaqty.Caption = 0
End Sub

Private Sub crefresh_Click()
Call Build_Results
Call Warna_List
Call warning_reschedule

End Sub

Private Sub cs_Click()
Form_Cetak_CS.Show
End Sub

Private Sub csearch_Click()
cari = "select * from completion_slip where no_slip='" & txtcari.Text & "'"
Set rscompletion_slip = conn.Execute(cari)
If rscompletion_slip.EOF Then
    MsgBox "Data tidak ditemukan"
    txtcari.Text = ""
    txtcari.SetFocus
Else
    Me.tno_slip.Text = rscompletion_slip.Fields("no_slip")
    Me.tno_so = rscompletion_slip.Fields("no_so")
    Me.cnowinding = rscompletion_slip.Fields("proses_2")
    Me.dtdate_printed = rscompletion_slip.Fields("date_printed")
    Me.dtfinish_date = rscompletion_slip.Fields("finish_date")
    Me.cshift.Text = rscompletion_slip.Fields("shift")
    Me.cstatus.Text = rscompletion_slip.Fields("status")
    Me.cno_part.Text = rscompletion_slip.Fields("no_part")
    Me.tjic.Text = rscompletion_slip.Fields("jic")
    Me.tsize.Text = rscompletion_slip.Fields("size")
    Me.ttype.Text = rscompletion_slip.Fields("type")
    Me.tqty.Text = rscompletion_slip.Fields("qty")
    Me.tcustomer.Text = rscompletion_slip.Fields("customer")
    Me.tnote.Text = rscompletion_slip.Fields("note")
    Me.tir.Text = rscompletion_slip.Fields("inner_ring")
    Me.tor.Text = rscompletion_slip.Fields("outer_ring")
    Me.thoop.Text = rscompletion_slip.Fields("hoop")
    Me.tfiller.Text = rscompletion_slip.Fields("filler")
    Me.tmarking_or.Text = rscompletion_slip.Fields("marking_stamp_or")
    Me.tmarking_ir.Text = rscompletion_slip.Fields("marking_stamp_ir")
    Me.tthickness.Text = rscompletion_slip.Fields("thickness")
    Me.tproses1.Text = rscompletion_slip.Fields("proses_1")
    Me.cproses2.Text = Me.cnowinding.Text
    Me.tproses3.Text = rscompletion_slip.Fields("proses_3")
    cedit.Caption = "UPDATE"
    
End If


End Sub

Private Sub cshift_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

    If cnowinding.Text = "-" Then
        tqty.SetFocus
    Else
        Call available_qty
        tqty.SetFocus
    End If
End If
End Sub

Private Sub cstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cnowinding.SetFocus
End If

End Sub

Private Sub ctype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cno_part.SetFocus
End If

End Sub

Private Sub daftarmesin_Click()
Form_Data_Mesin.Show
End Sub

Private Sub dtdate_printed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
dtfinish_date.SetFocus
End If

End Sub

Private Sub dtfinish_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cshift.SetFocus
End If

End Sub

Private Sub Build_Results(Optional Start_From = 0)
    
On Error GoTo Err_1
    
    Dim lst As ListItem   ' ListItem object
    Dim Temp_Counter, nmr As Long
    Dim Last_Page As Long ' Last page in current recordset
    Dim Start_Page As Long ' The page we will start from [ Start from=21 , Start_Page = 20 ]
    Dim x As Long
    
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    rscompletion_slip.Open "select no_slip, no_so, jic, Size, finish_date, qty, delivery_date, status, " & _
        "qty_pending,remarks_produksi,proses_2,id from completion_slip where category is null and deleted =0 order by id desc", conn, adOpenDynamic, adLockOptimistic
'    If rscompletion_slip.RecordCount > 0 Then
'        rscompletion_slip.MoveLast
'        rscompletion_slip.MoveFirst
'    End If
    
    ListView1.ListItems.Clear
    Temp_Counter = 0
    
    With rscompletion_slip
        If .RecordCount > 0 Then
            .Move Start_From * 20, 1
        End If
        
        Do While Not .EOF And Temp_Counter < 20
            Set lst = ListView1.ListItems.Add
            lst.Text = rscompletion_slip!no_slip
            lst.SubItems(1) = rscompletion_slip!no_slip
            lst.SubItems(2) = rscompletion_slip!no_so
            lst.SubItems(3) = rscompletion_slip!jic
            lst.SubItems(4) = rscompletion_slip!Size
            lst.SubItems(5) = rscompletion_slip!qty
            lst.SubItems(6) = rscompletion_slip!finish_date
            lst.SubItems(7) = rscompletion_slip!delivery_date
            lst.SubItems(8) = rscompletion_slip!status
            lst.SubItems(9) = IIf(IsNull(rscompletion_slip!qty_pending), 0, rscompletion_slip!qty_pending)
            lst.SubItems(10) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
            lst.SubItems(11) = rscompletion_slip!proses_2
            lst.SubItems(12) = rscompletion_slip!id
            .MoveNext
            Temp_Counter = Temp_Counter + 1
        Loop
        
        T_Results.Caption = CStr(.RecordCount)
        
        ' Calculating Showing_Records value
        If .RecordCount > 0 Then
            T_Showing_Records.Caption = (Start_From * 20) + 1 & " - "
            If (Start_From * 20) + 1 + 20 >= .RecordCount Then
                T_Showing_Records.Caption = T_Showing_Records.Caption & .RecordCount
            Else
                T_Showing_Records.Caption = T_Showing_Records.Caption & (Start_From * 20) + 20
            End If
        Else
            T_Showing_Records.Caption = "0"
        End If

        ' Removing old page navigators
        For T = 1 To T_Page.Count - 1
            Unload T_Page(T)
        Next
            
        ' Getting last page in current recordset
        If .RecordCount Mod 20 > 0 Then
            Last_Page = Int(.RecordCount / 20) + 1
        Else
            Last_Page = Int(.RecordCount / 20)
        End If
   
        ' Geting first page we will show [ Start_Page ]
        For y = 1 To Last_Page Step 6
            If Start_From + 1 >= y And Start_From + 1 <= y + 5 Then
                Exit For
            End If
        Next
   
        Start_Page = y
        x = 1
            
        ' If we are showing pages not from first 20... <<- [ Previous ]
        If y > 1 Then
            Load T_Page(T_Page.Count)
            T_Page(T_Page.Count - 1).Caption = "<<-"
            T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
            T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
            T_Page(T_Page.Count - 1).Visible = True
        End If
            
        For T = Start_Page To Last_Page
            Load T_Page(T_Page.Count)
            If x > 6 Then ' If there are more pages then we can show... ->> [ Next ]
                T_Page(T_Page.Count - 1).Caption = "->>"
                T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
                T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
                T_Page(T_Page.Count - 1).Visible = True
                Exit For
            Else
                T_Page(T_Page.Count - 1).Caption = CStr(T)
                T_Page(T_Page.Count - 1).Left = T_Page(T_Page.Count - 2).Left + T_Page(T_Page.Count - 2).Width + 90
                T_Page(T_Page.Count - 1).Top = T_Page(T_Page.Count - 2).Top
                If T = Start_From + 1 Then ' If this is a current page
                    T_Page(T_Page.Count - 1).ForeColor = &HFF&
                End If
                T_Page(T_Page.Count - 1).Visible = True
            End If
            x = x + 1
        Next
    End With
    
    
Exit_Sub:
   Exit Sub
    
Err_1:
    MsgBox Err.Description, vbOKOnly + vbCritical + vbApplicationModal, "StaCS : System error # " & Err.Number
    Resume Exit_Sub
    
End Sub

Private Sub Form_Activate()
'    Call db
 '   conn.CursorLocation = adUseClient
    tno_so.SetFocus
End Sub

Private Sub Form_Load()
    dtdate_printed.Value = Date
    dtfinish_date.Value = Date
    
    cshift.AddItem "1"
    cshift.AddItem "2"
    cstatus.AddItem "On Going"
    
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from certificate_files order by file_name asc", conn, adOpenDynamic, adLockOptimistic
    Do While Not rsdata_mesin.EOF
        cmb_inner.AddItem Mid(rsdata_mesin("file_name"), 1, Len(rsdata_mesin("file_name")) - 4)
        cmb_inner.ItemData(cmb_inner.NewIndex) = rsdata_mesin!id
        rsdata_mesin.MoveNext
    Loop
    
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from certificate_files order by file_name asc", conn, adOpenDynamic, adLockOptimistic
    Do While Not rsdata_mesin.EOF
        cmb_outer.AddItem Mid(rsdata_mesin("file_name"), 1, Len(rsdata_mesin("file_name")) - 4)
        cmb_outer.ItemData(cmb_outer.NewIndex) = rsdata_mesin!id
        rsdata_mesin.MoveNext
    Loop
    
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from certificate_files order by file_name asc", conn, adOpenDynamic, adLockOptimistic
    Do While Not rsdata_mesin.EOF
        cmb_hoop.AddItem Mid(rsdata_mesin("file_name"), 1, Len(rsdata_mesin("file_name")) - 4)
        cmb_hoop.ItemData(cmb_hoop.NewIndex) = rsdata_mesin!id
        rsdata_mesin.MoveNext
    Loop
    
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from certificate_files order by file_name asc", conn, adOpenDynamic, adLockOptimistic
    Do While Not rsdata_mesin.EOF
        cmb_filler.AddItem Mid(rsdata_mesin("file_name"), 1, Len(rsdata_mesin("file_name")) - 4)
        cmb_filler.ItemData(cmb_filler.NewIndex) = rsdata_mesin!id
        rsdata_mesin.MoveNext
    Loop

'    If rscode.State = 1 Then rscode.Close
'    rscode.Open "select*from code where deleted=0", conn
'    Do While Not rscode.EOF
'        cno_part.AddItem rscode("no_part")
'        rscode.MoveNext
'    Loop
    
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from data_mesin", conn
    Do While Not rsdata_mesin.EOF
        cproses2.AddItem rsdata_mesin("nama_mesin")
        rsdata_mesin.MoveNext
    Loop
    
    Call SetLV
    Call SetLVCert
    Call no_slip
    
    Call Build_Results

    lblavaqty.Caption = 0
    dtdate_printed.Enabled = False
    Call Warna_List
    Call warning_reschedule


End Sub

Private Sub kapasitasmesin_Click()
Form_Status_MC.Show
End Sub

Private Sub kode_Click()
Form_Data_Kode.Show
End Sub

Private Sub laporan_Click()
Form_Finish_Good.Show
End Sub

Private Sub list_foreman_Click()
Form_Cetak_LF.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Strnoslip = Trim(Item.Text)
    strsql = "select * from cs_files where id_cs='" & ListView1.SelectedItem.SubItems(12) & "' "
    Call LoadListView(strsql)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then PopupMenu pop_up_menu
End Sub

Private Sub Picture1_Click()
Dim lst As ListItem, nmr As Integer
    If rscompletion_slip.State = 0 Then
        rscompletion_slip.Open "select*from completion_slip", conn
    End If
    lihat = "select * from completion_slip where status='Reschedule'"
    Set rscompletion_slip = conn.Execute(lihat)
    With rscompletion_slip
    ListView1.ListItems.Clear
    Do While Not rscompletion_slip.EOF
    Set lst = ListView1.ListItems.Add
    lst.Text = rscompletion_slip!no_slip
    lst.SubItems(1) = rscompletion_slip!no_slip
    lst.SubItems(2) = rscompletion_slip!no_so
    lst.SubItems(3) = rscompletion_slip!jic
    lst.SubItems(4) = rscompletion_slip!Size
    lst.SubItems(5) = rscompletion_slip!qty
    lst.SubItems(6) = rscompletion_slip!finish_date
    lst.SubItems(7) = rscompletion_slip!delivery_date
    lst.SubItems(8) = rscompletion_slip!status
    lst.SubItems(9) = rscompletion_slip!qty_pending
    lst.SubItems(10) = IIf(IsNull(rscompletion_slip.Fields("remarks_produksi")), "", rscompletion_slip.Fields("remarks_produksi"))
    lst.SubItems(11) = rscompletion_slip!proses_2
    lst.SubItems(12) = rscompletion_slip!id
    rscompletion_slip.MoveNext
    Loop
    End With
Call Warna_List
End Sub

Private Sub reschedule_Click()
Dim s As Integer

kapasitas = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & ListView1.SelectedItem.SubItems(11) & "'"
Set rsdata_mesin = conn.Execute(kapasitas)
kapasitas = rsdata_mesin!MyCapa

qty_pending = "select sum(qty) AS MyPending from completion_slip where proses_2='" & ListView1.SelectedItem.SubItems(11) & "' and finish_date='" & Format(ListView1.SelectedItem.SubItems(6), "yyyy/mm/dd") & "' and status='Reschedule'"
Set rscompletion_slip = conn.Execute(qty_pending)
pending = rscompletion_slip!MyPending

tanggal = ListView1.SelectedItem.SubItems(6)
tanggal2 = Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd")

qty_last_slip = "SELECT qty From completion_slip WHERE finish_date = '" & tanggal2 & "' AND proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' and no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' AND finish_date='" & tanggal2 & "' and shift='1')"
Set rscompletion_slip = conn.Execute(qty_last_slip)
If rscompletion_slip.EOF Then
    myqty = 0
Else
    myqty = rscompletion_slip!qty
End If
    
sisapending = pending

tglrevisi = Format(tanggal, "YYYY/mm/dd")


ubah = "update completion_slip set status='On Going', finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "', shift='1' where status='Reschedule'"
Set rscompletion_slip = conn.Execute(ubah)

Do While sisapending > 0

    For s = 1 To 2
            qty_nextdate_1 = "select sum(qty) AS MyTotal from completion_slip where proses_2='" & ListView1.SelectedItem.SubItems(11) & "' and finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "'"
            Set rscompletion_slip = conn.Execute(qty_nextdate_1)
            strmytotal = rscompletion_slip!mytotal
            totalqty = strmytotal
        
        If totalqty <= kapasitas Then
            sisapending = 0
            Exit For
            
        Else
            
            sisapending = totalqty - kapasitas
            pindahqty = 0
            Do Until pindahqty >= sisapending
                qty_last_slip = "SELECT qty From completion_slip WHERE finish_date = '" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' AND proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' and no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' AND finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "')"
                Set rscompletion_slip = conn.Execute(qty_last_slip)
                qty_slip_max = rscompletion_slip!qty
                
                If s = 2 Then
                    proses_ubah_2 = "update completion_slip set status='On Going', finish_date='" & Format(DateAdd("d", 2, tanggal), "YYYY/mm/dd") & "', shift='1' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' AND finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "')"
                    Set rscompletion_slip = conn.Execute(proses_ubah_2)
                Else
                    proses_ubah_2 = "update completion_slip set status='On Going', shift='" & 2 & "' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' AND finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "')"
                    Set rscompletion_slip = conn.Execute(proses_ubah_2)
                    
                End If
                    pindahqty = pindahqty + qty_slip_max
            Loop
        End If
       
    Next
    tglrevisi = Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd")
    
Loop
Call Build_Results
Call Warna_List
Call warning_reschedule
End Sub

Private Sub salesorder_Click()
Form_Data_SO.Show
End Sub

Private Sub summary_order_Click()
Form_Summary_Order.Show
End Sub

Private Sub smn_cs_Click()
    Form_Cetak_CS.Show vbModal
End Sub

Private Sub smn_listforeman_Click()
    Form_Cetak_LF.Show vbModal
End Sub

Private Sub smn_swg_Click()
    Form_Data_Kode.Show vbModal
End Sub

Private Sub T_Page_Click(Index As Integer)
    
On Error GoTo Err_1
    
    Me.MousePointer = 11
    Me.AutoRedraw = False
    
    If T_Page(Index).Caption = "->>" Then
        Call Build_Results(Val(T_Page(Index - 1).Caption))
    ElseIf T_Page(Index).Caption = "<<-" Then
        Call Build_Results(Val(T_Page(Index + 1).Caption) - 2)
    Else
        Call Build_Results(Val(T_Page(Index).Caption) - 1)
    End If
    
    Me.AutoRedraw = True
    Me.MousePointer = 0
Exit_Sub:
   Exit Sub
    
Err_1:
    Resume Exit_Sub
End Sub

Private Sub tcustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
tnote.SetFocus
End If

End Sub

Private Sub Timer1_Timer()


If rscompletion_slip.State = 0 Then
rscompletion_slip.Open "select*from completion_slip", conn
End If

lihat = "select*from completion_slip"
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource
DataGrid1.Refresh

With DataGrid1
    .Columns(0).Width = 2000
    .Columns(1).Width = 1000
    .Columns(2).Width = 1050
    .Columns(3).Width = 0
    .Columns(4).Width = 1100
    .Columns(5).Width = 1100
    .Columns(6).Width = 1100
    .Columns(7).Width = 0
    .Columns(8).Width = 1000
    .Columns(9).Width = 0
    .Columns(10).Width = 1400
    .Columns(11).Width = 500
    .Columns(12).Width = 2500
    .Columns(13).Width = 1600
    .Columns(14).Width = 0
    .Columns(15).Width = 0
    .Columns(16).Width = 0
    .Columns(17).Width = 0
    .Columns(18).Width = 0
    .Columns(19).Width = 0
    .Columns(20).Width = 0
    .Columns(23).Width = 0
    .Columns(24).Width = 0
    .Columns(25).Width = 0
    .Columns(26).Width = 0
    .Columns(27).Width = 0
    .Columns(28).Width = 0
End With

End Sub

Private Sub tno_so_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cno_part.SetFocus
End If
End Sub

Private Sub tnote_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            tthickness.SetFocus
        End If
        Me.tproses1.Enabled = False
        Me.tproses3.Enabled = False
        
        Me.cproses2.Text = Me.cnowinding.Text
        Me.tproses3.Text = "FINISH GOOD"
    End If
End If

End Sub

Private Sub tproses1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cproses2.SetFocus
End If

End Sub

Private Sub tqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cnowinding.Text = "-" Then
        tcustomer.SetFocus
    Else
        If Val(Me.tqty.Text) > Val(Me.lblavaqty.Caption) Then
            MsgBox ("Kapasitas Sudah Terpenuhi")
            tqty.SetFocus
        Else
            tcustomer.SetFocus
        End If
    End If
End If

If ttype.Text = "Basic" Then
tproses1.Text = "MATERIAL"
tmarking_or.Text = "-"
tmarking_ir.Text = "-"

End If

If ttype.Text = "IR + Basic" Then
tproses1.Text = "MARKING"
tmarking_or.Text = "-"
End If

If ttype.Text = "OR + Basic" Then
tproses1.Text = "MARKING"
tmarking_ir.Text = "-"
End If

If ttype.Text = "Complete" Then
tproses1.Text = "MARKING"
tmarking_ir.Text = "-"
End If

If ttype.Text = "Ring OR" Then
tproses1.Text = "MATERIAL"
tmarking_ir.Text = "-"
End If

If ttype.Text = "Ring IR" Then
tproses1.Text = "MATERIAL"
tmarking_ir.Text = "-"
End If

End Sub

Sub cek_kapasitas_2()
Dim sisaqty As Integer


For colna = 1 To Form_Status_MC.MSFlexGrid1.Cols - 1
    For rowna = 3 To Form_Status_MC.MSFlexGrid1.Rows - 1
        
        banding = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, 0) & "'"
        jumlah = "select sum(qty) AS MyTotal from completion_slip where proses_2='" & Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, 0) & "' and date_printed='" & Format(Form_Status_MC.MSFlexGrid1.TextMatrix(1, colna), "YYYY/mm/dd") & "' and shift='" & Form_Status_MC.MSFlexGrid1.TextMatrix(2, colna) & "'"
        Set rscompletion_slip = conn.Execute(jumlah)
        Set rsdata_mesin = conn.Execute(banding)
        sisaqty = 0
        If Val(Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna)) > Val(rsdata_mesin.Fields("MyCapa")) Then
            
            x = MsgBox(" Kapasitas MC '" & Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, 0) & "' Tanggal '" & Form_Status_MC.MSFlexGrid1.TextMatrix(1, colna) & "' Shift '" & Form_Status_MC.MSFlexGrid1.TextMatrix(2, colna) & "' Sudah Terpenuhi, Akan dimasukkan ke Tanggal '" & Form_Status_MC.MSFlexGrid1.TextMatrix(1, colna) & "' Shift '" & Form_Status_MC.MSFlexGrid1.TextMatrix(2, colna) & "' dan Akan Dibuat No CS Baru..!!!!!", vbYesNo + vbCritical, "Warning")
            If x = vbYes Then
                sisaqty = Val(Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna)) - Val(rsdata_mesin.Fields("MyCapa"))
                sisa = Val(rscompletion_slip.Fields("MyTotal")) - Val(sisaqty)
                Call no_slip
                MsgBox (sisa)
                Me.tqty = sisaqty
                Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna) = rsdata_mesin.Fields("MyCapa")
                Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna + 1) = Val(Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna + 1)) + Val(sisaqty)
                Exit Sub
            Else
                cari = "select * from completion_slip where no_slip='" & DataGrid1.Columns("no_slip") & "'"
                Set rscompletion_slip = conn.Execute(cari)
                Call tampil
                cedit.Caption = "UPDATE"
                sisaqty = Val(Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna))
            End If
            
        End If
    Next
Next
End Sub


Sub hitung_kolom_2()
For rowna = 3 To Form_Status_MC.MSFlexGrid1.Rows - 1
For colna = 1 To Form_Status_MC.MSFlexGrid1.Cols - 1


jumlah = "select sum(qty) AS MyTotal from completion_slip where " & _
    "proses_2='" & Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, 0) & "' and " & _
    "date_printed='" & Format(Form_Status_MC.MSFlexGrid1.TextMatrix(1, colna), "YYYY/mm/dd") & "' " & _
    "and shift='" & Form_Status_MC.MSFlexGrid1.TextMatrix(2, colna) & "'"
Set rscompletion_slip = conn.Execute(jumlah)


Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna) = IIf(IsNull(rscompletion_slip.Fields("MyTotal")), "-", rscompletion_slip.Fields("MyTotal"))

Next
Next

End Sub

Sub available_qty()
jumlah = "select sum(qty+ISNULL(qty_pending,0)) AS MyTotal from completion_slip where proses_2='" & Me.cnowinding.Text & "' " & _
    "and finish_date='" & Format(Me.dtfinish_date.Value, "YYYY/mm/dd") & "' and shift='" & Me.cshift.Text & "'"
banding = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & Me.cnowinding.Text & "'"
Set rscompletion_slip = conn.Execute(jumlah)
Set rsdata_mesin = conn.Execute(banding)

available = Val(rsdata_mesin.Fields("MyCapa")) - Val(IIf(IsNull(rscompletion_slip.Fields("MyTotal")), "-", rscompletion_slip.Fields("MyTotal")))

Me.lblavaqty.Caption = available

End Sub

Sub alokasi_mc_winding()
If tsize2.Text = "1/2" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "3/4" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "1" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "1-1/4" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "1-1/2" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "2" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "2-1/2" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "3" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "3-1/2" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "4" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "5" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "6" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "8" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "10" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "12" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "14" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "16" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "18" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "10A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "15A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "20A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "25A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "32A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "40A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "50A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "65A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "80A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "90A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "100A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "125A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "150A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "175A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "200A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "225A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "250A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "300A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "350A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "400A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "450A" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "DN 10" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 15" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 20" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 25" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 32" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 40" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING1"
    cnowinding.AddItem "WINDING2"
    cnowinding.AddItem "WINDING3"
    cnowinding.AddItem "WINDING4"
ElseIf tsize2.Text = "DN 50" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "DN 65" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "DN 80" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "DN 100" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING5"
    cnowinding.AddItem "WINDING6"
    cnowinding.AddItem "WINDING7"
    cnowinding.AddItem "WINDING8"
ElseIf tsize2.Text = "DN 125" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 150" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 175" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 200" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 250" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 300" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING9"
    cnowinding.AddItem "WINDING10"
ElseIf tsize2.Text = "DN 350" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "DN 400" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
ElseIf tsize2.Text = "DN 450" Then
    cnowinding.Clear
    cnowinding.AddItem "WINDING11"
Else
    cnowinding.Clear
    Call data_mesin
'    cnowinding.AddItem "WINDING12"
'    cnowinding.AddItem "WINDING13"
End If

End Sub

Private Sub tthickness_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            cmb_inner.SetFocus
        End If
    End If
End If
End Sub

'Private Sub Txt_cert_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'       Case vbKeyReturn
'            cproses2.SetFocus
'    End Select
'End Sub

Private Sub txt_filler_cert_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case vbKeyReturn
            cproses2.SetFocus
    End Select
End Sub

'Private Sub txt_Heat_cert_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'       Case vbKeyReturn
'            Txt_cert.SetFocus
'    End Select
'End Sub

Private Sub txt_hoop_cert_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case vbKeyReturn
            txt_filler_cert.SetFocus
    End Select
End Sub

Private Sub txt_inner_cert_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case vbKeyReturn
            txt_outer_cert.SetFocus
    End Select
End Sub

Private Sub txt_outer_cert_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case vbKeyReturn
            txt_hoop_cert.SetFocus
    End Select
End Sub

Public Sub SetLVCert()
    With lv_cert
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "", 0
        .ColumnHeaders.Add 2, , "title", 2500
    End With
End Sub

Public Sub LoadListView(strsql As String)
    Dim lst As ListItem
    lv_cert.ListItems.Clear
    
    If vread.State = 1 Then vread.Close
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        vread.MoveFirst
        Do While Not vread.EOF
            If rscode.State = 1 Then rscode.Close
            strsql = "Select * from certificate_files where id='" & vread!id_certificate_files & "'"
            rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rscode.EOF Then
                Do While Not rscode.EOF
                    Set lst = lv_cert.ListItems.Add
                    lst.Text = Format(IIf(IsNull(rscode!id), "", rscode!id))
                    lst.SubItems(1) = Format(IIf(IsNull(rscode!Title), "", rscode!Title))
                    rscode.MoveNext
                Loop
            End If
            vread.MoveNext
        Loop
        vread.Close
    End If
End Sub

Private Sub coc_no()
    Dim tahun As String
    Dim bulan As String
    Dim tanggal As String
    Dim crxApplication As CRAXDRT.Application
    Dim oReport As CRAXDRT.Report
    Dim crxExportOptions As CRAXDRT.ExportOptions
    
    tahun = Format(Date, "YYYY")
    bulan = Format(Date, "MM")
    tanggal = Format(Date, "yyyy-mm-dd")
                
    Set crxApplication = New CRAXDRT.Application
    Set oReport = crxApplication.OpenReport(App.Path & "\Certificate.rpt", 1)
    oReport.Database.LogOnServer "pdsodbc.dll", "produksi", "PURCHASING", "sa", "admin123"
    oReport.RecordSelectionFormula = "{certificate_files.date} = '" & Format(tanggal, "YYYY-mm-dd") & "' "

    Set crxExportOptions = oReport.ExportOptions
        
    '======= CEK CERTIFICATE FILES =========
    If vread.State = 1 Then vread.Close
    strsql = "select is_coc,title,date as tanggal from certificate_files where " & _
        "deleted =0 and is_coc = 1 order by id asc"
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
       
    If vread.EOF Then
        date_now = Format(Date, "YYYY/mm/dd")
        Me.txt_coc.Text = "0001"
        
        txt_coc_no.Text = txt_coc.Text & "/JFI-QA/" & bulan & "/" & tahun
        txt_filepath = "D:\htdocs\ci\app1\upload\images\certificate"
        txt_note = "COC 2.1"
        
        strsql = "insert into certificate_files(date,file_path,file_name," & _
            "note,title,deleted,is_coc) " & _
            "values ('" & Format(Date, "yyyy-mm-dd") & "'," & _
            "'" & txt_filepath & "','COC_" & Replace(txt_coc_no.Text, "/", "_") & "'," & _
            "'" & txt_note & "','" & Replace(txt_coc_no.Text, "/", "_") & "',0,1) "
        conn.Execute (strsql)
        
        If rschkcode.State = 1 Then rschkcode.Close
        strsql = "select id from certificate_files " & _
            "where cast(date as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
        rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not rschkcode.EOF Then id_certificate = rschkcode!id
        
        If rschkcode.State = 1 Then rschkcode.Close
        strsql = "Select date_printed,id from completion_slip " & _
            "where cast(date_printed as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
        rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not rschkcode.EOF Then
            rschkcode.MoveFirst
            Do While Not rschkcode.EOF
                If rsso.State = 1 Then rsso.Close
                strsql = "select id from cs_files where " & _
                    "id_cs='" & rschkcode!id & "' and id_certificate_files='" & id_certificate & "' "
                rsso.Open strsql, conn, adOpenDynamic, adLockOptimistic
                If Not rsso.EOF Then
                    
                Else
                    strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                        "values ('" & rschkcode!id & "', '" & id_certificate & "',0)"
                    conn.Execute (strsql)
                End If
                
                rschkcode.MoveNext
            Loop
            
        End If
        
        crxExportOptions.DestinationType = crEDTDiskFile
        crxExportOptions.DiskFileName = App.Path & "\COC_" & Replace(txt_coc_no.Text, "/", "_") & ".pdf"
        crxExportOptions.FormatType = crEFTPortableDocFormat
        crxExportOptions.PDFFirstPageNumber = 1
        crxExportOptions.PDFLastPageNumber = 1
        crxExportOptions.PDFExportAllPages = True
        oReport.export False
    Else
        vread.MoveLast
        last_date = Format(vread!tanggal, "YYYY/mm/dd")
        date_now = Format(Date, "YYYY/mm/dd")
        
        If date_now <> last_date Then
            coc_number = Val(Left(vread!Title, 4)) + 1
            If coc_number < 10 Then
                Me.txt_coc.Text = "000" & coc_number
            ElseIf coc_number < 100 Then
                Me.txt_coc.Text = "00" & coc_number
            ElseIf coc_number < 1000 Then
                Me.txt_coc.Text = "0" & coc_number
            Else
                Me.txt_coc.Text = coc_number
            End If
            txt_coc_no.Text = txt_coc.Text & "/JFI-QA/" & bulan & "/" & tahun
            txt_filepath = "D:\htdocs\ci\app1\upload\images\certificate"
            txt_note = "COC 2.1"
            
            strsql = "insert into certificate_files(date,file_path,file_name," & _
                "note,title,deleted,is_coc) " & _
                "values ('" & Format(Date, "yyyy-mm-dd") & "'," & _
                "'" & txt_filepath & "','COC_" & Replace(txt_coc_no.Text, "/", "_") & "'," & _
                "'" & txt_note & "','" & txt_coc_no.Text & "',0,1) "
            conn.Execute (strsql)
            
            If rschkcode.State = 1 Then rschkcode.Close
            strsql = "select id from certificate_files " & _
                "where cast(date as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
            rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rschkcode.EOF Then id_certificate = rschkcode!id
            
            If rschkcode.State = 1 Then rschkcode.Close
            strsql = "Select date_printed,id from completion_slip " & _
                "where cast(date_printed as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
            rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rschkcode.EOF Then
                rschkcode.MoveFirst
                Do While Not rschkcode.EOF
                    If rsso.State = 1 Then rsso.Close
                    strsql = "select id from cs_files where " & _
                        "id_cs='" & rschkcode!id & "' and id_certificate_files='" & id_certificate & "' "
                    rsso.Open strsql, conn, adOpenDynamic, adLockOptimistic
                    If Not rsso.EOF Then
                        
                    Else
                        strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                            "values ('" & rschkcode!id & "', '" & id_certificate & "',0)"
                        conn.Execute (strsql)
                    End If
                    
                    rschkcode.MoveNext
                Loop
                
            End If
            
            crxExportOptions.DestinationType = crEDTDiskFile
            crxExportOptions.DiskFileName = App.Path & "\COC_" & Replace(txt_coc_no.Text, "/", "_") & ".pdf"
            crxExportOptions.FormatType = crEFTPortableDocFormat
            crxExportOptions.PDFFirstPageNumber = 1
            crxExportOptions.PDFLastPageNumber = 1
            crxExportOptions.PDFExportAllPages = True
            oReport.export False
        Else
            Me.txt_coc.Text = Left(vread!Title, 4)
            txt_coc_no.Text = txt_coc.Text & "/JFI-QA/" & bulan & "/" & tahun
            txt_filepath = "D:\htdocs\ci\app1\upload\images\certificate"
            txt_note = "COC 2.1"
            
            If rschkcode.State = 1 Then rschkcode.Close
            strsql = "select id from certificate_files " & _
                "where cast(date as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
            rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rschkcode.EOF Then id_certificate = rschkcode!id
            
            If rschkcode.State = 1 Then rschkcode.Close
            strsql = "Select date_printed,id from completion_slip " & _
                "where cast(date_printed as date) = '" & Format(Date, "yyyy-mm-dd") & "' "
            rschkcode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            If Not rschkcode.EOF Then
                rschkcode.MoveFirst
                Do While Not rschkcode.EOF
                    If rsso.State = 1 Then rsso.Close
                    strsql = "select id from cs_files where " & _
                        "id_cs='" & rschkcode!id & "' and id_certificate_files='" & id_certificate & "' "
                    rsso.Open strsql, conn, adOpenDynamic, adLockOptimistic
                    If rsso.EOF Then
                        strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                            "values ('" & rschkcode!id & "', '" & id_certificate & "',0)"
                        conn.Execute (strsql)
                    End If
                    
                    rschkcode.MoveNext
                Loop
            End If
            
            crxExportOptions.DestinationType = crEDTDiskFile
            crxExportOptions.DiskFileName = App.Path & "\COC_" & Replace(txt_coc_no.Text, "/", "_") & ".pdf"
            crxExportOptions.FormatType = crEFTPortableDocFormat
            crxExportOptions.PDFFirstPageNumber = 1
            crxExportOptions.PDFLastPageNumber = 1
            crxExportOptions.PDFExportAllPages = True
            oReport.export False
        End If
    End If
End Sub

Private Sub save_certificate()
    Dim idcs As String
    Dim idcertificate As String
    Dim insertidcert As String
    
    If vread.State = 1 Then vread.Close
    strsql = "select id from completion_slip where no_slip ='" & tno_slip.Text & "' "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        idcs = vread!id
    End If
    
    If vread.State = 1 Then vread.Close
    strsql = "select cs_files.id, certificate_files.is_coc, certificate_files.id as certificate_id from cs_files " & _
        "Left join certificate_files on cs_files.id_certificate_files =certificate_files.id " & _
        "where id_cs='" & idcs & "' and certificate_files.is_coc = 1 "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        idcertificate = vread!certificate_id
    End If
    
    strsql = "delete from cs_files where id_cs = '" & idcs & "' and id_certificate_files <> '" & idcertificate & "' "
    conn.Execute (strsql)
    
    If cmb_inner.Text <> "" Then
        If vread.State = 1 Then vread.Close
        strsql = "Select id from certificate_files where substring(file_name, 1, Len(file_name) - 4) ='" & cmb_inner & "' "
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not vread.EOF Then
            insertidcert = vread!id
            strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                "values ('" & idcs & "', '" & insertidcert & "',0)"
            conn.Execute (strsql)
        End If
    End If
    
    If cmb_outer.Text <> "" And cmb_outer.Text <> cmb_inner.Text Then
        If vread.State = 1 Then vread.Close
        strsql = "Select id from certificate_files where substring(file_name, 1, Len(file_name) - 4) ='" & cmb_outer & "' "
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not vread.EOF Then
            insertidcert = vread!id
            strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                "values ('" & idcs & "', '" & insertidcert & "',0)"
            conn.Execute (strsql)
        End If
    End If
    
    If cmb_hoop.Text <> "" Then
        If vread.State = 1 Then vread.Close
        strsql = "Select id from certificate_files where substring(file_name, 1, Len(file_name) - 4) ='" & cmb_hoop & "' "
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not vread.EOF Then
            insertidcert = vread!id
            strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                "values ('" & idcs & "', '" & insertidcert & "',0)"
            conn.Execute (strsql)
        End If
    End If
    
    If cmb_filler.Text <> "" Then
        If vread.State = 1 Then vread.Close
        strsql = "Select id from certificate_files where substring(file_name, 1, Len(file_name) - 4) ='" & cmb_filler & "' "
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        If Not vread.EOF Then
            insertidcert = vread!id
            strsql = "insert into cs_files(id_cs,id_certificate_files,deleted) " & _
                "values ('" & idcs & "', '" & insertidcert & "',0)"
            conn.Execute (strsql)
        End If
    End If
    
End Sub

