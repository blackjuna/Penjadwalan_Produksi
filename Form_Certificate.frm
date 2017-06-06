VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_Cert 
   Caption         =   "Form Certificate "
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Format          =   100597761
      CurrentDate     =   42788
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukkan Judul Yang Dicari"
      Height          =   855
      Left            =   6720
      TabIndex        =   13
      Top             =   4440
      Width           =   4455
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.TextBox txt_title 
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1680
      Width           =   5175
   End
   Begin MSComctlLib.ListView lv_certificate 
      Height          =   3975
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txt_note 
      Height          =   1935
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Form_Certificate.frx":0000
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox txt_filename 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox txt_filepath 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton cmd_refresh 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_new 
      Caption         =   "New"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Incoming Date  :"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "File Name         :"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Note                 :"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Title                  :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Path           :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "Form_Cert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strid As String
Public strsql As String
Public strfilepath As String

Private Sub cmd_delete_Click()
    Dim intResponse As Integer
    Dim strfilename As String
    
    intResponse = MsgBox("Anda yakin data ini akan dihapus?", vbYesNo + vbCritical, "Warning")
    If intResponse = vbYes Then
        strsql = "update certificate_files set deleted=1 where id='" & strid & "' "
        conn.Execute (strsql)
        
        If vread.State = 1 Then vread.Close
        strsql = "select file_name from certificate_files where id='" & strid & "' "
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        
        If Not vread.EOF Then
            strfilename = vread!file_name
        End If
        Kill "\\192.168.0.7\d\htdocs\ci\app1\upload\images\certificate\" & strfilename & ""
        
        strsql = "Select * from certificate_files where deleted=0"
        Call LoadListView(strsql)
    End If
End Sub

Private Sub cmd_refresh_Click()
    Call ClearText
    cmd_delete.Enabled = False
    cmd_save.Enabled = False
    cmd_save.Caption = "Save"
    cmd_new.Enabled = True
    
    strsql = "Select * from certificate_files where deleted=0"
    Call LoadListView(strsql)
End Sub

Private Sub csearch_Click()
    strsql = "select * from certificate_files where (title like '%" & txtcari & "%' or note like '%" & txtcari & "%') and deleted=0"
    Call LoadListView(strsql)
End Sub

Private Sub lv_certificate_DblClick()
    Call EditData(strid)
    
    cmd_new.Enabled = False
    cmd_save.Caption = "Update"
    cmd_delete.Enabled = False
    cmd_save.Enabled = True
End Sub

Private Sub lv_certificate_GotFocus()
    cmd_delete.Enabled = True
End Sub

Private Sub lv_certificate_ItemClick(ByVal Item As MSComctlLib.ListItem)
    strid = Trim(Item.Text)
End Sub

Private Sub lv_certificate_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lv_certificate.SortOrder = 0
    Else
        lv_certificate.SortOrder = Abs(lv_certificate.SortOrder - 1)
    End If
    
    lv_certificate.SortKey = ColumnHeader.Index - 1
    lv_certificate.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub cmd_new_Click()
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    
    Dim myFile As String
    strfilepath = CommonDialog1.FileName
    myFile = VBA.Mid$(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\") + 1)
    'myFile = VBA.Left$(myFile, InStrRev(myFile, ".") - 1)

    txt_filepath.Text = "D:\htdocs\ci\app1\upload\images\certificate"
    txt_filename.Text = myFile
    
    cmd_save.Enabled = True
    cmd_delete.Enabled = False
End Sub

Private Sub cmd_save_Click()
    If cmd_save.Caption = "Save" Then
        Call save
    Else
        Call update
    End If
End Sub

Private Sub Form_Load()
    Call ClearText
    Call DrawListview
    strsql = "Select * from certificate_files where deleted=0"
    Call LoadListView(strsql)
    DTPicker1.Value = Format(Now, "dd-mm-yyyy")
End Sub

Public Sub ClearText()
    For Each A In Me
        If TypeOf A Is TextBox Then A.Text = ""
    Next A
End Sub

Public Sub DrawListview()
    With lv_certificate
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "ID", 0
        .ColumnHeaders.Add 2, , "Incoming Date", 1300
        .ColumnHeaders.Add 3, , "File Name", 2000
        .ColumnHeaders.Add 4, , "title", 2500
        .ColumnHeaders.Add 5, , "Note", 4500
        .ColumnHeaders.Add 6, , "File Path", 3000
    End With
End Sub

Public Sub LoadListView(strsql As String)
    Dim lst As ListItem
    lv_certificate.ListItems.Clear
    If vread.State = 1 Then vread.Close
    
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not vread.EOF Then
        vread.MoveFirst
        Do While Not vread.EOF
            Set lst = lv_certificate.ListItems.Add
            lst.Text = Format(IIf(IsNull(vread!id), "", vread!id))
            lst.SubItems(1) = Format(IIf(IsNull(vread!Date), "", Format(vread!Date, "dd-mm-yyyy")))
            lst.SubItems(2) = Format(IIf(IsNull(vread!file_name), "", vread!file_name))
            lst.SubItems(3) = Format(IIf(IsNull(vread!Title), "", vread!Title))
            lst.SubItems(4) = Format(IIf(IsNull(vread!note), "", vread!note))
            lst.SubItems(5) = Format(IIf(IsNull(vread!file_path), "", vread!file_path))
            vread.MoveNext
        Loop
        vread.Close
    End If

End Sub

Public Sub save()
    If Dir("\\192.168.0.7\d\htdocs\ci\app2\upload\images\certificate\" & txt_filename & "") <> "" Then
        MsgBox "Sorry, File is exists. Please change file name!"
    Else
        FileCopy strfilepath, "\\192.168.0.7\d\htdocs\ci\app1\upload\images\certificate\" & txt_filename & ""

        strsql = "insert into certificate_files(date,file_path,file_name,note,title,deleted) " & _
            "values ('" & Format(DTPicker1.Value, "yyyy-mm-dd") & "','" & txt_filepath & "','" & txt_filename & "','" & txt_note & "','" & txt_title & "',0) "
        conn.Execute (strsql)
        
        MsgBox "Data Sudah tersimpan", vbOKOnly + vbInformation, "Informasi"
        
        Call ClearText
        
        strsql = "Select * from certificate_files where deleted=0"
        Call LoadListView(strsql)
        
        cmd_save.Enabled = False
        cmd_delete.Enabled = False
    End If
End Sub

Public Sub update()
    Dim FileName As String
    Dim NewFileName As String
    
    If vread.State = 1 Then vread.Close
    strsql = "select file_name from certificate_files where id='" & strid & "' "
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not vread.EOF Then
        strfilename = vread!file_name
    End If
    
    FileName = "\\192.168.0.7\d\htdocs\ci\app1\upload\images\certificate\" & strfilename & ""
    NewFileName = "\\192.168.0.7\d\htdocs\ci\app1\upload\images\certificate\" & txt_filename & ""
    Name FileName As NewFileName
    
    strsql = "update certificate_files set date='" & DTPicker1.Value & "',file_name='" & txt_filename & "',title='" & txt_title.Text & "', note='" & txt_note.Text & "' where id='" & strid & "'"
    conn.Execute (strsql)
    
    MsgBox "Data sudah diperbaharui.", vbOKOnly + vbInformation, "Informasi"
    
    Call ClearText
    cmd_save.Caption = "Save"
    cmd_new.Enabled = True
    cmd_delete.Enabled = False
    cmd_save.Enabled = False
    
    strsql = "Select * from certificate_files where deleted=0"
    Call LoadListView(strsql)
    
End Sub

Public Sub EditData(strcode As String)
    If vread.State = 1 Then vread.Close
    strsql = "Select * from certificate_files where deleted=0 and id='" & strcode & "'"
    vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not vread.EOF Then
        DTPicker1.Value = vread!Date
        txt_filepath.Text = IIf(IsNull(vread!file_path), "", vread!file_path)
        txt_filename.Text = IIf(IsNull(vread!file_name), "", vread!file_name)
        txt_title.Text = IIf(IsNull(vread!Title), "", vread!Title)
        txt_note.Text = IIf(IsNull(vread!note), "", vread!note)
    End If
End Sub



