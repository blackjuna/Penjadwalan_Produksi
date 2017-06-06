VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_add_cer 
   Caption         =   "Form Add Cert"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_select 
      Caption         =   "Select All"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_deselect 
      Caption         =   "Unselect All"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin MSComctlLib.ListView lv_add_cert 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form_add_cer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strid As String
Public strsql As String


Private Sub cmd_deselect_Click()
    For i = 1 To lv_add_cert.ListItems.Count
        lv_add_cert.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmd_save_Click()
    If vread.State = 1 Then vread.Close
    strsql = "select id from completion_slip where no_slip='" & strid & "' "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        strsql = "delete from cs_files where id_cs='" & vread!id & "'"
        conn.Execute (strsql)
        For i = 1 To lv_add_cert.ListItems.Count
            If lv_add_cert.ListItems(i).Checked = True Then
                strsql = "insert into cs_files (id_cs,id_certificate_files,deleted) values " & _
                    "('" & vread!id & "','" & lv_add_cert.ListItems(i) & "',0)"
                conn.Execute (strsql)
            End If
        Next
        MsgBox "Certificate Sudah tersimpan", vbOKOnly + vbInformation, "Informasi"
    End If
End Sub

Private Sub cmd_select_Click()
    For i = 1 To lv_add_cert.ListItems.Count
        lv_add_cert.ListItems(i).Checked = True
    Next
End Sub

Private Sub Form_Load()
Call SetLVCert
strid = Form_Utama.Strnoslip
strsql = "Select * from certificate_files where deleted=0"
Call LoadListView(strsql)

End Sub

Public Sub LoadListView(strsql As String)
    Dim lst As ListItem
    lv_add_cert.ListItems.Clear
    If vread.State = 1 Then vread.Close
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not vread.EOF Then
        vread.MoveFirst
        i = 1
        Do While Not vread.EOF
            Set lst = lv_add_cert.ListItems.Add
            
            If rscode.State = 1 Then rscode.Close
            strsql = "select no_slip,no_so,no_part,jic,size,marking_stamp_or,marking_stamp_ir,qty,completion_slip.note as cs_note,customer,status," & _
                "certificate_files.file_name,certificate_files.note as cert_note,certificate_files.title,certificate_files.id as id_cert from completion_slip " & _
                "LEFT JOIN cs_files ON cs_files.id_cs  =completion_slip.id " & _
                "LEFT JOIN certificate_files on cs_files.id_certificate_files=certificate_files.id " & _
                "where completion_slip.no_slip='" & strid & "' and " & _
                "certificate_files.id='" & Format(IIf(IsNull(vread!id), "", vread!id)) & "'"
            rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
            
            lst.Text = Format(IIf(IsNull(vread!id), "", vread!id))
            
            If Not rscode.EOF Then lv_add_cert.ListItems.Item(i).Checked = True
            lst.SubItems(1) = Format(IIf(IsNull(vread!Date), "", Format(vread!Date, "dd-mm-yyyy")))
            lst.SubItems(2) = Format(IIf(IsNull(vread!file_name), "", vread!file_name))
            lst.SubItems(3) = Format(IIf(IsNull(vread!Title), "", vread!Title))
            lst.SubItems(4) = Format(IIf(IsNull(vread!note), "", vread!note))
            i = i + 1
            vread.MoveNext
        Loop
        vread.Close
    End If
End Sub

Public Sub SetLVCert()
    With lv_add_cert
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "", 300
        .ColumnHeaders.Add 2, , "Incoming Date", 1300
        .ColumnHeaders.Add 3, , "File Name", 1800
        .ColumnHeaders.Add 4, , "title", 2500
        .ColumnHeaders.Add 5, , "Note", 4500
    End With
End Sub
