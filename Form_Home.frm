VERSION 5.00
Begin VB.Form Form_Home 
   Caption         =   "Home"
   ClientHeight    =   8625
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mn_pp 
      Caption         =   "Penjadwalan Produksi"
      Begin VB.Menu smn_swg 
         Caption         =   "SWG"
      End
      Begin VB.Menu smn_djg 
         Caption         =   "DJG"
      End
      Begin VB.Menu smn_smg 
         Caption         =   "SMG"
      End
   End
   Begin VB.Menu mn_master 
      Caption         =   "Master"
      Begin VB.Menu smn_kodepart 
         Caption         =   "Kode Part"
         Begin VB.Menu smn_code_swg 
            Caption         =   "SWG"
         End
         Begin VB.Menu smn_code_djg 
            Caption         =   "DJG"
         End
         Begin VB.Menu smn_code_smg 
            Caption         =   "SMG"
         End
      End
      Begin VB.Menu smn_mesin 
         Caption         =   "Daftar Mesin"
      End
      Begin VB.Menu smn_so 
         Caption         =   "Daftar SO"
      End
      Begin VB.Menu smn_certificate 
         Caption         =   "Certificate"
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
   Begin VB.Menu mn_status 
      Caption         =   "Status"
      Begin VB.Menu smn_status_cs 
         Caption         =   "Completion Slip"
      End
      Begin VB.Menu smn_status_mesin 
         Caption         =   "Kapasitas Mesin"
      End
   End
   Begin VB.Menu mn_laporan 
      Caption         =   "Laporan Finish Goods"
   End
End
Attribute VB_Name = "Form_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call db
End Sub

Private Sub mn_laporan_Click()
    Form_Finish_Good.Show vbModal
End Sub

Private Sub smn_certificate_Click()
    Form_Cert.Show vbModal
End Sub

Private Sub smn_code_djg_Click()
    Form_Data_DJG.Show vbModal
End Sub

Private Sub smn_code_smg_Click()
    Form_Data_SMG.Show vbModal
End Sub

Private Sub smn_code_swg_Click()
    Form_Data_Kode.Show vbModal
End Sub

Private Sub smn_cs_Click()
    Form_Cetak_CS.Show vbModal
End Sub

Private Sub smn_djg_Click()
    Form_Utama_DJG.Show vbModal
End Sub

Private Sub smn_list_foreman_Click()
    Form_Cetak_LF.Show vbModal
End Sub

Private Sub smn_mesin_Click()
    Form_Data_Mesin.Show vbModal
End Sub

Private Sub smn_smg_Click()
    Form_Utama_SMG.Show vbModal
End Sub

Private Sub smn_so_Click()
    Form_Data_SO.Show vbModal
End Sub

Private Sub smn_status_cs_Click()
    Form_Status_CS.Show vbModal
End Sub

Private Sub smn_status_mesin_Click()
    Form_Status_MC.Show vbModal
End Sub

Private Sub smn_swg_Click()
    Form_Utama.Show vbModal
End Sub
