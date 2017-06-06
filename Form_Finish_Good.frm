VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Finish_Good 
   Caption         =   "Laporan Finish Good"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7335
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12938
      _Version        =   393216
      Rows            =   33
      Cols            =   3
      AllowUserResizing=   2
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98238465
      CurrentDate     =   41801
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Tanggal Yang Akan Ditampilkan"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox cmb_category 
         Height          =   315
         ItemData        =   "Form_Finish_Good.frx":0000
         Left            =   120
         List            =   "Form_Finish_Good.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98238465
         CurrentDate     =   41801
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "-"
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   480
         Width           =   45
      End
   End
End
Attribute VB_Name = "Form_Finish_Good"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intdiff As Integer

Private Sub cmb_category_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DTPicker1.SetFocus
    End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        DTPicker2.SetFocus
    End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim vCount As Long
        Dim start_date As Date
        Dim end_date As Date
        
        start_date = Format(DTPicker1.Value, "YYYY/mm/dd")
        end_date = Format(DTPicker2.Value, "YYYY/mm/dd")
        intdiff = DateDiff("d", start_date, end_date)
        MSFlexGrid1.Rows = Val(intdiff) + 3
        Call SetFlexGrid
        
        For vCount = 1 To Val(intdiff) + 1
            MSFlexGrid1.TextMatrix(vCount, 0) = (Format(start_date, "YYYY/mm/dd"))
            
            If start_date = end_date Then
                Exit For
            End If
            start_date = start_date + 1
        Next
        Call hitung
        
'        If MSFlexGrid1.TextMatrix(31, 0) = "" Then
'            MSFlexGrid1.RowHeight(31) = 0
'        End If
    End If

End Sub

Private Sub Form_Load()

    DTPicker1.Value = Date
    DTPicker2.Value = Date

End Sub

Sub hitung()
    Select Case cmb_category.Text
        Case "SWG"
            strcategory = "and category is null and deleted=0"
        Case "DJG"
            strcategory = "and category =2 and deleted=0"
        Case "SMG"
            strcategory = "and category =3 and deleted=0"
        Case ""
            MsgBox "Mohon pilih category terlebih dahulu", vbCritical + vbOKOnly, "Warning"
            Exit Sub
    End Select
    
    For rowna = 1 To MSFlexGrid1.Rows - 2
        qty_ppic = "select sum(qty) AS MyTotalPPIC from completion_slip where " & _
            "finish_date='" & Format(MSFlexGrid1.TextMatrix(rowna, 0), "YYYY/mm/dd") & "' " & strcategory & ""
        Set rscompletion_slip = conn.Execute(qty_ppic)
        
        MSFlexGrid1.TextMatrix(rowna, 1) = IIf(IsNull(rscompletion_slip.Fields("MyTotalPPIC")), "0", rscompletion_slip.Fields("MyTotalPPIC"))
        
        qty_produksi = "select sum(qty) AS MyProd from completion_slip where " & _
            "delivery_date='" & Format(MSFlexGrid1.TextMatrix(rowna, 0), "YYYY/mm/dd") & "' " & strcategory & ""
        Set rscompletion_slip = conn.Execute(qty_produksi)
        
        MSFlexGrid1.TextMatrix(rowna, 2) = IIf(IsNull(rscompletion_slip.Fields("MyProd")), "0", rscompletion_slip.Fields("MyProd"))
        
'        If Val(MSFlexGrid1.TextMatrix(rowna, 1)) = 0 Then
'            If Val(MSFlexGrid1.TextMatrix(rowna, 2)) = 0 Then
'                MSFlexGrid1.RowHeight(rowna) = 0
'            End If
'        End If
    Next
    
    bulan = Month(Format(DTPicker1.Value, "YYYY/mm/dd"))
    tahun = Year(Format(DTPicker1.Value, "YYYY/mm/dd"))
    
    qty_all_ppic = "select sum(qty) AS MyTotalAllPPIC from completion_slip where " & _
        "month(finish_date)='" & bulan & "' and year(finish_date)='" & tahun & "'" & strcategory & ""
    Set rscompletion_slip = conn.Execute(qty_all_ppic)
    ppic = IIf(IsNull(rscompletion_slip!MyTotalAllPPIC), "0", rscompletion_slip!MyTotalAllPPIC)
    
    qty_all_produksi = "select sum(qty) AS MyTotalAllProd from completion_slip where " & _
        "month(delivery_date)='" & bulan & "' and year(delivery_date)='" & tahun & "'" & strcategory & ""
    Set rscompletion_slip = conn.Execute(qty_all_produksi)
    produksi = IIf(IsNull(rscompletion_slip!MyTotalAllProd), "0", rscompletion_slip!MyTotalAllProd)
    
    MSFlexGrid1.TextMatrix(intdiff + 2, 1) = ppic

    MSFlexGrid1.TextMatrix(intdiff + 2, 2) = IIf(IsNull(produksi), "0", produksi)
End Sub

Public Sub SetFlexGrid()
    MSFlexGrid1.TextMatrix(0, 0) = "Tanggal"
    MSFlexGrid1.TextMatrix(0, 1) = "Plan Order PPIC"
    MSFlexGrid1.TextMatrix(0, 2) = "Production Qty"
    'MSFlexGrid1.TextMatrix(0, 3) = "Percentage ( % )"
    MSFlexGrid1.TextMatrix(intdiff + 2, 0) = "Grand Total"
    MSFlexGrid1.ColWidth(1) = 1500
    MSFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(2) = 1500
    MSFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    'MSFlexGrid1.ColWidth(3) = 1500
    'MSFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
End Sub
