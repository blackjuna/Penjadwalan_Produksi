VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Form_Utama_DJG 
   Caption         =   "Main Form Completion Slip DJG"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   15225
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   46
      Top             =   960
      Width           =   4695
      Begin VB.TextBox tno_slip 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   53
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox tno_so 
         Height          =   405
         Left            =   1800
         TabIndex        =   52
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cno_part 
         Height          =   315
         Left            =   1800
         TabIndex        =   51
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox ttype 
         Height          =   405
         Left            =   1800
         TabIndex        =   50
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox tsize 
         Height          =   405
         Left            =   1800
         TabIndex        =   49
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox tjic 
         Height          =   405
         Left            =   1800
         TabIndex        =   48
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cstatus 
         Height          =   315
         ItemData        =   "Form_Utama_DJG.frx":0000
         Left            =   1800
         List            =   "Form_Utama_DJG.frx":0007
         TabIndex        =   47
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "No. Slip"
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. SO"
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   59
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "JIC"
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   1800
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   240
         TabIndex        =   56
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No Part"
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   3240
         Width           =   450
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   5160
      TabIndex        =   31
      Top             =   960
      Width           =   4695
      Begin VB.TextBox tqty 
         Height          =   405
         Left            =   1800
         TabIndex        =   36
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox tcustomer 
         Height          =   405
         Left            =   1800
         TabIndex        =   35
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox tnote 
         Height          =   645
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ComboBox cnowinding 
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cshift 
         Height          =   315
         ItemData        =   "Form_Utama_DJG.frx":0015
         Left            =   1800
         List            =   "Form_Utama_DJG.frx":001F
         TabIndex        =   32
         Top             =   1680
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtdate_printed 
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97910785
         CurrentDate     =   41714
      End
      Begin MSComCtl2.DTPicker dtfinish_date 
         Height          =   375
         Left            =   1800
         TabIndex        =   38
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97910785
         CurrentDate     =   41714
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Note"
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "No Machine"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Shift"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Input Date"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Date PPIC"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   10080
      TabIndex        =   20
      Top             =   960
      Width           =   4695
      Begin VB.TextBox txt_partition2 
         Height          =   405
         Left            =   1440
         TabIndex        =   76
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox txt_metal 
         Height          =   405
         Left            =   1440
         TabIndex        =   74
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txt_Metal_cert 
         Height          =   405
         Left            =   1440
         TabIndex        =   71
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txt_filler_cert 
         Height          =   405
         Left            =   1440
         TabIndex        =   70
         Top             =   4320
         Width           =   2655
      End
      Begin VB.TextBox tfiller 
         Height          =   405
         Left            =   1440
         TabIndex        =   25
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox tmarking_or 
         Height          =   405
         Left            =   1440
         TabIndex        =   24
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txt_radius 
         Height          =   405
         Left            =   1440
         TabIndex        =   23
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txt_width 
         Height          =   405
         Left            =   1440
         TabIndex        =   22
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txt_partition 
         Height          =   405
         Left            =   1440
         TabIndex        =   21
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Partition 2"
         Height          =   195
         Left            =   240
         TabIndex        =   77
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Metal"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Metal Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Filler Cert"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   4320
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Filler"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Marking "
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Radius"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Partition 1"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2640
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   10080
      TabIndex        =   13
      Top             =   6000
      Width           =   4695
      Begin VB.TextBox tproses3 
         Height          =   405
         Left            =   1440
         TabIndex        =   16
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tproses1 
         Height          =   405
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cproses2 
         Height          =   315
         ItemData        =   "Form_Utama_DJG.frx":0029
         Left            =   1440
         List            =   "Form_Utama_DJG.frx":002B
         TabIndex        =   14
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "To Location"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "End Process"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "From Location"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.CommandButton chapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Masukkan No Slip Yang Dicari"
      Height          =   855
      Left            =   10080
      TabIndex        =   8
      Top             =   7800
      Width           =   3495
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox tlotnumber 
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tbatchnumber 
      Height          =   405
      Left            =   1680
      TabIndex        =   5
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tsize2 
      Height          =   405
      Left            =   3120
      TabIndex        =   4
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tmetal 
      Height          =   405
      Left            =   4560
      TabIndex        =   3
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton crefresh 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1020
      Left            =   13680
      Picture         =   "Form_Utama_DJG.frx":002D
      ScaleHeight     =   960
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   7800
      Width           =   1035
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8070
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Left            =   10560
      TabIndex        =   69
      Top             =   240
      Width           =   4215
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
      Left            =   6120
      TabIndex        =   68
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label T_Page 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1200
      MouseIcon       =   "Form_Utama_DJG.frx":0BCA
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   9600
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Goto page :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   66
      Top             =   9600
      Width           =   840
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   65
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Showing :"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1800
      TabIndex        =   64
      Top             =   9960
      Width           =   705
   End
   Begin VB.Label T_Results 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   960
      TabIndex        =   63
      Top             =   9960
      Width           =   90
   End
   Begin VB.Label T_Showing_Records 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   2640
      TabIndex        =   62
      Top             =   9960
      Width           =   90
   End
   Begin VB.Menu pop_up_menu 
      Caption         =   "Pop Up Menu"
      Visible         =   0   'False
      Begin VB.Menu reschedule 
         Caption         =   "Reschedule"
      End
   End
   Begin VB.Menu mn_master 
      Caption         =   "Master"
      Begin VB.Menu smn_djd 
         Caption         =   "Kode DJG"
      End
   End
   Begin VB.Menu mn_cetak 
      Caption         =   "Cetak"
      Begin VB.Menu smn_cs 
         Caption         =   "Completion Slip"
      End
      Begin VB.Menu smn_list_foreman 
         Caption         =   "List Foreman"
      End
   End
End
Attribute VB_Name = "Form_Utama_DJG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strsql, Strnoslip As String

Private Sub cbatal_Click()
    cedit.Caption = "EDIT"
    Call ClearText
    Call Create_No_Slip
    lblavaqty.Caption = 0
    Call Build_Results
    Call Warna_List
    Call warning_reschedule
    tno_so.SetFocus
End Sub

Private Sub cedit_Click()
    If cedit.Caption = "EDIT" Then
        cedit.Caption = "UPDATE"
        If rscompletion_slip.State = 1 Then rscompletion_slip.Close
        cari = "select no_slip,no_so,proses_2,date_printed,finish_date,shift,status,completion_slip.no_part," & _
            "completion_slip.jic,completion_slip.size,completion_slip.type,completion_slip.qty, " & _
            "customer,note,completion_slip.filler,marking_stamp_or,proses_1,proses_3,radius,width,partition,partition2,metal_cert,metal, " & _
            "filler_cert from completion_slip left join " & _
            "code on completion_slip.no_part=code.no_part " & _
            "where completion_slip.id='" & Strnoslip & "'"
        rscompletion_slip.Open cari, conn, adOpenDynamic, adLockOptimistic
        If Not rscompletion_slip.EOF Then
            Call SetEdit
        End If
        tjic.Enabled = False
        tsize.Enabled = False
        txt_radius.Enabled = False
        txt_width.Enabled = False
        txt_partition.Enabled = False
        txt_partition2.Enabled = False
        ttype.Enabled = False
        tmarking_or.Enabled = False
        tfiller.Enabled = False
        txt_metal.Enabled = False
        tproses3.Enabled = False
        tno_so.SetFocus
    Else
        ubah = "UPDATE completion_slip SET no_so='" & tno_so.Text & "'," & _
            "date_printed='" & Format(dtdate_printed.Value, "YYYY/mm/dd") & "'," & _
            "finish_date='" & Format(dtfinish_date.Value, "YYYY/mm/dd") & "'," & _
            "shift='" & cshift.Text & "',status='" & cstatus.Text & "'," & _
            "no_part='" & cno_part.Text & "',jic='" & tjic.Text & "'," & _
            "size='" & tsize.Text & "',type='" & ttype.Text & "',qty='" & tqty.Text & "'," & _
            "customer='" & tcustomer.Text & "',note='" & tnote.Text & "'," & _
            "filler='" & tfiller.Text & "'," & _
            "marking_stamp_or='" & tmarking_or.Text & "'," & _
            "proses_1='" & tproses1.Text & "'," & _
            "proses_2='" & cproses2.Text & "',proses_3='" & tproses3.Text & "',metal_cert='" & txt_Metal_cert.Text & "'," & _
            "filler_cert='" & txt_filler_cert.Text & "' where no_slip='" & tno_slip.Text & "'"
        Set rscompletion_slip = conn.Execute(ubah)
        Call ClearText
        cedit.Caption = "EDIT"
        Call Create_No_Slip
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
        Call ClearText
        Call Create_No_Slip
        Call Build_Results
        Call Warna_List
        Call warning_reschedule
    End If
End Sub

Private Sub cproses2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call save
        lblavaqty.Caption = 0
    End If
End Sub

Private Sub crefresh_Click()
    Call Build_Results
    Call Warna_List
    Call warning_reschedule
    tno_so.SetFocus
End Sub

Private Sub csearch_Click()
    cari = "select no_slip,no_so,proses_2,date_printed,finish_date,shift,status,completion_slip.no_part," & _
        "completion_slip.jic,completion_slip.size,completion_slip.type,completion_slip.qty, " & _
        "customer,note,completion_slip.filler,marking_stamp_or,proses_1,proses_3,radius,width,partition,partition2 from completion_slip left join " & _
        "code on completion_slip.no_part=code.no_part " & _
        "where no_slip='" & txtcari.Text & "'"
    Set rscompletion_slip = conn.Execute(cari)
    If rscompletion_slip.EOF Then
        MsgBox "Data tidak ditemukan"
        txtcari.Text = ""
        txtcari.SetFocus
    Else
        Call SetEdit
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Strnoslip = Trim(Item.Text)
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu pop_up_menu
End Sub

Private Sub reschedule_Click()
    Dim s As Integer
    
    If ListView1.SelectedItem.SubItems(12) = "A" Then
        kapasitas = 15
    Else
        kapasitas = 12
    End If

'    kapasitas = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & ListView1.SelectedItem.SubItems(12) & "'"
'    Set rsdata_mesin = conn.Execute(kapasitas)
'    kapasitas = rsdata_mesin!MyCapa
    
    If rscompletion_slip.State Then rscompletion_slip.Close
    strsql = "select sum(qty) AS MyPending from completion_slip where " & _
        "proses_2='" & ListView1.SelectedItem.SubItems(11) & "' and " & _
        "finish_date='" & Format(ListView1.SelectedItem.SubItems(6), "yyyy/mm/dd") & "' and status='Reschedule' and category=2"
    rscompletion_slip.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not rscompletion_slip.EOF Then
        pending = IIf(IsNull(rscompletion_slip!MyPending), 0, rscompletion_slip!MyPending)
    Else
        pending = 0
    End If
    
    tanggal = ListView1.SelectedItem.SubItems(6)
    tanggal2 = Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd")
    
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    qty_last_slip = "SELECT qty From completion_slip WHERE finish_date = '" & tanggal2 & "' " & _
        "AND proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' and no_slip=(SELECT MAX(no_slip) " & _
        "FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' " & _
        "AND finish_date='" & tanggal2 & "' and shift='1' and category=2)"
    rscompletion_slip.Open qty_last_slip, conn, adOpenDynamic, adLockOptimistic
    
    If rscompletion_slip.EOF Then
        myqty = 0
    Else
        myqty = rscompletion_slip!qty
    End If
    
    tglrevisi = Format(tanggal, "YYYY/mm/dd")
    
    
    ubah = "update completion_slip set status='On Going', " & _
        "finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "', " & _
        "shift='1' where status='Reschedule'"
    Set rscompletion_slip = conn.Execute(ubah)
    
    Do While pending > 0
    
        For s = 1 To 2
            qty_nextdate_1 = "select sum(qty) AS MyTotal from completion_slip where " & _
                "proses_2='" & ListView1.SelectedItem.SubItems(11) & "' and " & _
                "finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "' and category=2"
            Set rscompletion_slip = conn.Execute(qty_nextdate_1)
            strmytotal = rscompletion_slip!mytotal
            'totalqty = strmytotal
            If strmytotal <= kapasitas Then
                sisapending = 0
                Exit For
            Else
                sisapending = Val(strmytotal) - Val(kapasitas)
                pindahqty = 0
                Do Until pindahqty >= sisapending
                    qty_last_slip = "SELECT qty From completion_slip WHERE " & _
                        "finish_date = '" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' " & _
                        "AND proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' " & _
                        "and no_slip=(SELECT MAX(no_slip) FROM completion_slip " & _
                        "where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' " & _
                        "AND finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "' and category =2)"
                    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
                    rscompletion_slip.Open qty_last_slip, conn, adOpenDynamic, adLockOptimistic
                    If Not rscompletion_slip.EOF Then
                        qty_slip_max = rscompletion_slip!qty
                    Else
                        qty_slip_max = 0
                    End If
                    
                    If s = 2 Then
                        proses_ubah_2 = "update completion_slip set status='On Going', " & _
                            "finish_date='" & Format(DateAdd("d", 2, tanggal), "YYYY/mm/dd") & "', " & _
                            "shift='1' where no_slip=(SELECT MAX(no_slip) " & _
                            "FROM completion_slip where proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' " & _
                            "AND finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "')"
                        conn.Execute (proses_ubah_2)
                    Else
                        proses_ubah_2 = "update completion_slip set status='On Going', shift='" & 2 & "' " & _
                            "where no_slip=(SELECT MAX(no_slip) FROM completion_slip where " & _
                            "proses_2 = '" & ListView1.SelectedItem.SubItems(11) & "' AND " & _
                            "finish_date='" & Format(DateAdd("d", 1, tanggal), "YYYY/mm/dd") & "' and shift='" & s & "')"
                        conn.Execute (proses_ubah_2)
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

Private Sub smn_cs_Click()
    Form_Cetak_CS.Show vbModal
End Sub

Private Sub smn_djd_Click()
    Form_Data_DJG.Show vbModal
End Sub

Private Sub smn_list_foreman_Click()
    Form_Cetak_LF.Show vbModal
End Sub

Private Sub tcustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        tnote.SetFocus
    End If
End Sub

Private Sub tnote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If KeyCode = vbKeyReturn Then
            If KeyCode = vbKeyReturn Then
                txt_Metal_cert.SetFocus
            End If
            
            Me.cproses2.Text = Me.cnowinding.Text
        End If
    End If
End Sub

Private Sub tnote_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tqty_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        If cnowinding.Text = "-" Then
            tproses1.Text = "DRAWING"
            tproses1.Enabled = False
            tcustomer.SetFocus
        ElseIf Val(tqty.Text) > Val(Me.lblavaqty.Caption) Then
            MsgBox ("Kapasitas Sudah Terpenuhi")
            tqty.SetFocus
        Else
            tproses1.Text = "DRAWING"
            tproses1.Enabled = False
            tcustomer.SetFocus
        End If
    End If
End Sub

Private Sub cno_part_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        Call SearchProduct
    End If
End Sub

Private Sub cnowinding_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cnowinding.Text = "-" Then
            lblavaqty.Caption = "Unlimited"
            dtfinish_date.SetFocus
        Else
            dtfinish_date.SetFocus
        End If
    End If
End Sub

Private Sub cshift_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cnowinding.Text = "-" Then
            tqty.SetFocus
        Else
            Call available_qty
            tqty.SetFocus
        End If
        
    End If
End Sub

Private Sub cstatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cnowinding.SetFocus
    End If
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

Private Sub tno_so_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        cno_part.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    tno_so.SetFocus
End Sub

Private Sub Form_Load()
    Call code_part
    dtdate_printed.Value = Date
    dtfinish_date.Value = Date
    Call Create_Machine
    dtdate_printed.Enabled = False
    
    Call DrawLv
    Call Build_Results
    Call Create_No_Slip
    
    lblavaqty.Caption = 0
    dtdate_printed.Enabled = False
    Call Warna_List
    Call warning_reschedule
End Sub

Public Sub code_part()
    If rscode.State = 1 Then rscode.Close
    strsql = "select no_part from code where deleted=0 and category=2"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not rscode.EOF Then
        rscode.MoveFirst
        cno_part.Clear
        Do While Not rscode.EOF
            cno_part.AddItem rscode!no_part
            rscode.MoveNext
        Loop
    End If
End Sub

Public Sub DrawLv()
    With ListView1
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        .ColumnHeaders.Add 1, , "id", 0
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
        .ColumnHeaders.Add 13, , "Type", 300
    End With
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
        "qty_pending,remarks_produksi,proses_2,id,type from completion_slip where category =2 and deleted =0 order by id desc", conn, adOpenDynamic, adLockOptimistic
'    If rscompletion_slip.RecordCount > 0 Then
'        rscompletion_slip.MoveLast
'        rscompletion_slip.MoveFirst
'    End If
    
    ListView1.ListItems.Clear
    Temp_Counter = 0
    
    With rscompletion_slip
        If .RecordCount > 0 Then
            .Move Start_From * 100, 1
        End If
        
        Do While Not .EOF And Temp_Counter < 100
            Set lst = ListView1.ListItems.Add
            lst.Text = rscompletion_slip!id
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
            lst.SubItems(12) = rscompletion_slip!Type
            .MoveNext
            Temp_Counter = Temp_Counter + 1
        Loop
        
        T_Results.Caption = CStr(.RecordCount)
        
        ' Calculating Showing_Records value
        If .RecordCount > 0 Then
            T_Showing_Records.Caption = (Start_From * 100) + 1 & " - "
            If (Start_From * 100) + 1 + 100 >= .RecordCount Then
                T_Showing_Records.Caption = T_Showing_Records.Caption & .RecordCount
            Else
                T_Showing_Records.Caption = T_Showing_Records.Caption & (Start_From * 100) + 100
            End If
        Else
            T_Showing_Records.Caption = "0"
        End If

        ' Removing old page navigators
        For T = 1 To T_Page.Count - 1
            Unload T_Page(T)
        Next
            
        ' Getting last page in current recordset
        If .RecordCount Mod 100 > 0 Then
            Last_Page = Int(.RecordCount / 100) + 1
        Else
            Last_Page = Int(.RecordCount / 100)
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

Public Sub Create_No_Slip()
    Dim thn As String

    thn = Format(Date, "YYYY")
    If rscompletion_slip.State = 1 Then rscompletion_slip.Close
    rscompletion_slip.Open "select date_printed,lot_number,batch_number,no_slip from completion_slip where deleted =0 and " & _
        "category=2 order by id asc", conn, adOpenDynamic, adLockOptimistic
    
    If rscompletion_slip.RecordCount = 0 Then
        Me.tlotnumber.Text = "0001"
    Else
        rscompletion_slip.MoveLast
        last_date = Format(rscompletion_slip.Fields("date_printed"), "YYYY/mm/dd")
        date_now = Format(Date, "YYYY/mm/dd")
        If date_now = last_date Then
            qry_lot_number = rscompletion_slip.Fields("lot_number")
'            no_lot = qry_lot_number
'            lot_number = no_lot
            Me.tlotnumber.Text = qry_lot_number
        Else
            qry_lot_number = rscompletion_slip.Fields("lot_number")
'            no_lot = qry_lot_number
            lot_number = (Val(Right(qry_lot_number, 4)) + 1)
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
        Me.tbatchnumber.Text = "0001"
    Else
        rscompletion_slip.MoveLast
        last_date = Format(rscompletion_slip!date_printed, "YYYY/mm/dd")
        date_now = Format(Date, "YYYY/mm/dd")
        If Weekday(date_now) = vbMonday And date_now = last_date Then
            qry_batch_number = rscompletion_slip!batch_number
'            batch_number = qry_batch_number
            Me.tbatchnumber.Text = qry_batch_number
        ElseIf Weekday(date_now) = vbMonday And date_now > last_date Then
            qry_batch_number = rscompletion_slip!batch_number
'            no_batch = qry_batch_number
            batch_number = Val(Right(qry_batch_number, 4)) + 1
            If batch_number < 10 Then
                Me.tbatchnumber.Text = "000" & batch_number
            ElseIf batch_number < 100 Then
                Me.tbatchnumber.Text = "00" & batch_number
            ElseIf batch_number < 1000 Then
                Me.tbatchnumber.Text = "0" & batch_number
            Else
                Me.tbatchnumber.Text = batch_number
            End If
        Else
            qry_batch_number = rscompletion_slip.Fields("batch_number")
'            batch_number = qry_batch_number
            Me.tbatchnumber.Text = qry_batch_number
            
        End If
    End If
    
        
    If rscompletion_slip.RecordCount = 0 Then
        Me.tno_slip.Text = "2017-0001-D0000001"
    Else
        rscompletion_slip.MoveLast
        Z = Val(Mid(rscompletion_slip!no_slip, 12, 10)) + 1
'        Y = Mid(rscompletion_slip!no_slip, 12, 10)
        If Z < 10 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D000000" & Z
        ElseIf Z < 100 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D00000" & Z
        ElseIf Z < 1000 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D0000" & Z
        ElseIf Z < 10000 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D000" & Z
        ElseIf Z < 100000 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D00" & Z
        ElseIf Z < 1000000 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D0" & Z
        ElseIf Z < 10000000 Then
            Me.tno_slip.Text = thn & "-" & tbatchnumber.Text & "-" & "D" & Z
        End If
    End If
End Sub

Public Sub Create_Machine()
    If rsdata_mesin.State = 1 Then rsdata_mesin.Close
    rsdata_mesin.Open "select*from data_mesin", conn
    If Not rsdata_mesin.EOF Then
        rsdata_mesin.MoveFirst
        cproses2.Clear
        cproses2.AddItem "ALL"
        cnowinding.AddItem "ALL"
        Do While Not rsdata_mesin.EOF
            cproses2.AddItem rsdata_mesin("nama_mesin")
            cnowinding.AddItem rsdata_mesin("nama_mesin")
            rsdata_mesin.MoveNext
        Loop
    End If
End Sub

Public Sub SearchProduct()
    If rscode.State = 1 Then rscode.Close
    strsql = "select jic,proses,size,type,filler,size_2,filler,metal,marking_stamp_lokal_or,radius,width,partition,partition2 from code where " & _
        "no_part='" & cno_part.Text & "' and deleted=0 and category=2"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not rscode.EOF Then
        tjic.Text = rscode.Fields("jic")
        tsize.Text = rscode.Fields("size")
        ttype.Text = rscode.Fields("type")
        txt_radius.Text = rscode.Fields("radius")
        txt_width.Text = rscode.Fields("width")
        txt_partition.Text = rscode.Fields("partition")
        txt_partition2.Text = IIf(IsNull(rscode.Fields("partition2")), "", rscode.Fields("partition2"))
        tfiller.Text = IIf(IsNull(rscode.Fields("filler")), "", rscode.Fields("filler"))
        txt_metal.Text = rscode.Fields("metal")
        tsize2.Text = rscode.Fields("size_2")
        tmetal.Text = rscode.Fields("metal")
        tmarking_or.Text = rscode.Fields("marking_stamp_lokal_or")
        tproses3.Text = rscode.Fields("proses")
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
    txt_radius.Enabled = False
    txt_width.Enabled = False
    txt_partition.Enabled = False
    txt_partition2.Enabled = False
    ttype.Enabled = False
    tmarking_or.Enabled = False
    tfiller.Enabled = False
    txt_metal.Enabled = False
    tproses3.Enabled = False
    cstatus.SetFocus
End Sub

Public Sub available_qty()
    strsql = "select sum(qty+qty_pending) AS MyTotal from completion_slip where left(type,1)='" & Left(ttype, 1) & "' " & _
        "and finish_date='" & Format(Me.dtfinish_date.Value, "YYYY/mm/dd") & "' and shift='" & Me.cshift.Text & "' and category=2"
    Set rscompletion_slip = conn.Execute(strsql)
    
    If cnowinding.Text = "-" Then
        lblavaqty.Caption = "Unlimited"
    Else
        available = IIf(Left(ttype, 1) = "A", 15, 12) - Val(IIf(IsNull(rscompletion_slip.Fields("MyTotal")), "0", rscompletion_slip.Fields("MyTotal")))
        Me.lblavaqty.Caption = available
    End If
End Sub

Public Sub save()
    If rscode.State = 1 Then rscode.Close
    strsql = "Select no_slip from completion_slip where " & _
        "no_slip='" & tno_slip.Text & "' and deleted=0 and category=2"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not rscode.EOF Then
        MsgBox "No Slip Sudah ada"
        cproses2.SetFocus
    Else
        strsql = "INSERT INTO completion_slip (no_slip,no_so,date_printed,delivery_date,finish_date," & _
        "lot_number,batch_number,proses_1,proses_2,proses_3,no_part,shift,jic,size,type,filler," & _
        "marking_stamp_or,qty,customer,status,note,deleted,category,qty_pending,metal_cert,filler_cert) VALUES('" & tno_slip.Text & "'," & _
        "'" & tno_so.Text & "','" & Format(dtdate_printed.Value, "YYYY/mm/dd") & "',''," & _
        "'" & Format(dtfinish_date.Value, "YYYY/mm/dd") & "'," & _
        "'" & tlotnumber.Text & "','" & tbatchnumber.Text & "'," & _
        "'" & tproses1.Text & "','" & cproses2.Text & "','" & tproses3.Text & "'," & _
        "'" & cno_part.Text & "','" & cshift.Text & "','" & tjic.Text & "'," & _
        "'" & tsize.Text & "','" & ttype.Text & "','" & tfiller.Text & "'," & _
        "'" & tmarking_or.Text & "','" & tqty.Text & "','" & tcustomer.Text & "','" & cstatus.Text & "'," & _
        "'" & tnote.Text & "',0,2,0,'" & txt_Metal_cert.Text & "','" & txt_filler_cert.Text & "')"
        conn.Execute (strsql)
        
        MsgBox "Data sudah tersimpan", vbOKOnly + vbInformation, "Informasi"
        
        Call ClearText
        tno_so.SetFocus
        Call Create_No_Slip
        Call Build_Results
        Call Warna_List
        Call warning_reschedule
    End If
End Sub


Public Sub ClearText()
    For Each A In Me
        If TypeOf A Is TextBox Then A.Text = ""
        If TypeOf A Is ComboBox Then A.Text = ""
    Next A
End Sub

Sub Warna_List()
    Dim i As Long
    Dim y As Integer
        
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).SubItems(8) = "Pending" Then 'Field Stok pada kolom 5
            ListView1.ListItems(i).ForeColor = vbRed
            For y = 1 To 11
                ListView1.ListItems(i).ListSubItems(y).ForeColor = vbRed
            Next
        ElseIf ListView1.ListItems(i).SubItems(8) = "Partial" Then
            ListView1.ListItems(i).ForeColor = vbGreen
            For y = 1 To 11
                ListView1.ListItems(i).ListSubItems(y).ForeColor = vbGreen
            Next
        ElseIf ListView1.ListItems(i).SubItems(8) = "Closed" Then
            ListView1.ListItems(i).ForeColor = vbBlue
            For y = 1 To 11
                ListView1.ListItems(i).ListSubItems(y).ForeColor = vbBlue
            Next
        ElseIf ListView1.ListItems(i).SubItems(8) = "Reschedule" Then
            ListView1.ListItems(i).ForeColor = vbMagenta
            For y = 1 To 11
                ListView1.ListItems(i).ListSubItems(y).ForeColor = vbMagenta
            Next
        Else
            ListView1.ListItems(i).ForeColor = vbBlack
            For y = 1 To 11
                ListView1.ListItems(i).ListSubItems(y).ForeColor = vbBlack
            Next
        End If
    Next

End Sub

Public Sub warning_reschedule()
    cari = "select status from completion_slip where status='Reschedule'"
    Set rscompletion_slip = conn.Execute(cari)
    
    If rscompletion_slip.EOF Then
        Picture1.Visible = False
    Else
        Picture1.Visible = True
    End If
End Sub

Public Sub SetEdit()
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
    Me.tfiller.Text = rscompletion_slip.Fields("filler")
    Me.txt_metal.Text = rscompletion_slip.Fields("metal")
    Me.tmarking_or.Text = rscompletion_slip.Fields("marking_stamp_or")
    txt_radius.Text = rscompletion_slip.Fields("radius")
    txt_width.Text = rscompletion_slip.Fields("width")
    txt_partition.Text = rscompletion_slip.Fields("partition")
    txt_partition2.Text = IIf(IsNull(rscompletion_slip.Fields("partition2")), "", rscompletion_slip.Fields("partition2"))
    Me.tproses1.Text = rscompletion_slip.Fields("proses_1")
    Me.cproses2.Text = Me.cnowinding.Text
    Me.tproses3.Text = rscompletion_slip.Fields("proses_3")
    Me.txt_Metal_cert.Text = IIf(IsNull(rscompletion_slip.Fields("metal_cert")), "", rscompletion_slip.Fields("metal_cert"))
    Me.txt_filler_cert.Text = IIf(IsNull(rscompletion_slip.Fields("filler_cert")), "", rscompletion_slip.Fields("filler_cert"))
    cedit.Caption = "UPDATE"
End Sub

Private Sub txt_filler_cert_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cproses2.SetFocus
    End If
End Sub

Private Sub txt_Metal_cert_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt_filler_cert.SetFocus
    End If
End Sub
