VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Revisi_CS 
   Caption         =   "Form Revisi CS"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cnomesin 
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cexecute 
      Caption         =   "Execute"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2143
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
   Begin MSComCtl2.DTPicker dtrevisi 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   35848193
      CurrentDate     =   41758
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Masukkan Nomor Mesin"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Masukkan Tanggal Referensi"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "Form_Revisi_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cexecute_Click()
Dim s As Integer

kapasitas = "select sum(kapasitas) AS MyCapa from data_mesin where nama_mesin='" & cnomesin.Text & "'"
Set rsdata_mesin = conn.Execute(kapasitas)
kapasitas = rsdata_mesin!MyCapa

qty_pending = "select sum(qty_pending) AS MyPending from completion_slip where proses_2='" & cnomesin.Text & "' and finish_date='" & Format(dtrevisi.Value, "YYYY/mm/dd") & "'"
Set rscompletion_slip = conn.Execute(qty_pending)
pending = rscompletion_slip!MyPending

qty_last_slip = "SELECT qty From completion_slip WHERE finish_date = '" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "' AND proses_2 = '" & cnomesin.Text & "' and no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "' and shift='1')"
Set rscompletion_slip = conn.Execute(qty_last_slip)
myqty = rscompletion_slip!qty
    
sisapending = pending

tglrevisi = Format(dtrevisi.Value, "YYYY/mm/dd")


ubah = "update completion_slip set finish_date='" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "', shift='1' where status='Pending'"
Set rscompletion_slip = conn.Execute(ubah)

Do While sisapending > 0

    For s = 1 To 2
            qty_nextdate_1 = "select sum(qty) AS MyTotal from completion_slip where proses_2='" & cnomesin.Text & "' and finish_date='" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "' and shift='" & s & "'"
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
                qty_last_slip = "SELECT qty From completion_slip WHERE finish_date = '" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "' AND proses_2 = '" & cnomesin.Text & "' and no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "' and shift='" & s & "')"
                Set rscompletion_slip = conn.Execute(qty_last_slip)
                qty_slip_max = rscompletion_slip!qty
                
                If s = 2 Then
                    proses_ubah_2 = "update completion_slip set finish_date='" & Format(DateAdd("d", 2, tglrevisi), "YYYY/mm/dd") & "', shift='1' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "' and shift='" & s & "')"
                    Set rscompletion_slip = conn.Execute(proses_ubah_2)
                Else
                    proses_ubah_2 = "update completion_slip set shift='" & 2 & "' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd") & "' and shift='" & s & "')"
                    Set rscompletion_slip = conn.Execute(proses_ubah_2)
                    
                End If
                    pindahqty = pindahqty + qty_slip_max
            Loop
        End If
       
    Next
    tglrevisi = Format(DateAdd("d", 1, tglrevisi), "YYYY/mm/dd")
    
Loop

'If Val(pending) < Val(myqty) Then
'    ubah = "update completion_slip set finish_date='" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "', shift='1' where status='Pending'"
'    Set rscompletion_slip = conn.Execute(ubah)
    
    
    
'    proses_ubah_1 = "update completion_slip set finish_date='" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "' and shift='1')"
'    Set rscompletion_slip = conn.Execute(proses_ubah_1)
'Else
'    pending_last_slip = "select qty_pending from completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(dtrevisi.Value, "YYYY/mm/dd") & "' and no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(dtrevisi.Value, "YYYY/mm/dd") & "')"
'    Set rscompletion_slip = conn.Execute(pending_last_slip)
'    last_pending = rscompletion_slip!qty_pending
'    sisa_pending = Val(pending)
'    Do Until Val(sisa_pending) <= Val(myqty)
'        sisa_pending = Val(sisa_pending) - Val(myqty)
'        proses_ubah_2 = "update completion_slip set finish_date='" & Format(dtrevisi.Value + 2, "YYYY/mm/dd") & "' where no_slip=(SELECT MAX(no_slip) FROM completion_slip where proses_2 = '" & cnomesin.Text & "' AND finish_date='" & Format(dtrevisi.Value + 1, "YYYY/mm/dd") & "' and shift='1')"
'        Set rscompletion_slip = conn.Execute(proses_ubah_2)
'    Loop
'End If




End Sub

Private Sub cnomesin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call tampilgrid
End If
End Sub

Private Sub dtrevisi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cnomesin.SetFocus
End If
End Sub

Private Sub Form_Activate()
Call db
conn.CursorLocation = adUseClient

rscompletion_slip.Open "select*from completion_slip", conn
rsdata_mesin.Open "select*from data_mesin", conn

Do While Not rsdata_mesin.EOF
    cnomesin.AddItem rsdata_mesin("nama_mesin")
    rsdata_mesin.MoveNext
Loop

dtrevisi.Value = Date



End Sub

Sub tampilgrid()
lihat = "select no_slip,no_so,jic,customer,qty_pending from completion_slip where status='Pending' and finish_date='" & Format(dtrevisi.Value, "YYYY/mm/dd") & "' AND proses_2='" & cnomesin.Text & "' "
Set rscompletion_slip = conn.Execute(lihat)
Set DataGrid1.DataSource = rscompletion_slip.DataSource

With DataGrid1
    .Columns(0).Width = 1000
    .Columns(1).Width = 800
    .Columns(2).Width = 2500
End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
For rowna = 3 To Form_Status_MC.MSFlexGrid1.Rows - 1
For colna = 1 To Form_Status_MC.MSFlexGrid1.Cols - 1


jumlah = "select sum(qty) AS MyTotal from completion_slip where proses_2='" & Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, 0) & "' and finish_date='" & Format(Form_Status_MC.MSFlexGrid1.TextMatrix(1, colna), "YYYY/mm/dd") & "' and shift='" & Form_Status_MC.MSFlexGrid1.TextMatrix(2, colna) & "'"
Set rscompletion_slip = conn.Execute(jumlah)


Form_Status_MC.MSFlexGrid1.TextMatrix(rowna, colna) = IIf(IsNull(rscompletion_slip.Fields("MyTotal")), "-", rscompletion_slip.Fields("MyTotal"))

Next
Next

End Sub

