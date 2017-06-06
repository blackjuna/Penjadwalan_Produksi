VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Data_SMG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product SMG"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox tmetal 
      Height          =   315
      Left            =   4920
      TabIndex        =   33
      Top             =   1080
      Width           =   3615
   End
   Begin VB.ComboBox txtradius 
      Height          =   315
      Left            =   4920
      TabIndex        =   32
      Top             =   1560
      Width           =   3615
   End
   Begin VB.ComboBox txtwidth 
      Height          =   315
      Left            =   4920
      TabIndex        =   31
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtpartitionbar2 
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox tsize 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox tmarking_or_lokal 
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton chapus 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   8040
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cbatal 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton csimpan 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukkan No Part Yang Dicari"
      Height          =   855
      Left            =   8640
      TabIndex        =   9
      Top             =   3120
      Width           =   4455
      Begin VB.CommandButton csearch 
         Caption         =   "SEARCH"
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.ComboBox cmbproses 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cmbtype 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtjic 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtsize 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtnopart 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtpartitionbar 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton cbaru 
      Caption         =   "NEW"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton crefresh 
      Caption         =   "REFRESH"
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Partition Bar"
      Height          =   195
      Left            =   3480
      TabIndex        =   30
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Metal"
      Height          =   195
      Left            =   3480
      TabIndex        =   28
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stamp Marking"
      Height          =   195
      Left            =   3480
      TabIndex        =   26
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Proses"
      Height          =   195
      Left            =   240
      TabIndex        =   25
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "JIC"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   2040
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Size / Class"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Radius"
      Height          =   195
      Left            =   3480
      TabIndex        =   21
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Width"
      Height          =   195
      Left            =   3480
      TabIndex        =   20
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Partition Bar"
      Height          =   195
      Left            =   3480
      TabIndex        =   19
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No Part"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   540
   End
End
Attribute VB_Name = "Form_Data_SMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strsql As String
Public strid As String

Private Sub cbatal_Click()
    Call ClearText
    Call EnabledAll(False)
    cbaru.Enabled = True
    csimpan.Enabled = False
    cedit.Caption = "EDIT"
    chapus.Enabled = True
    cedit.Enabled = False
End Sub

Private Sub cedit_Click()
    If cedit.Caption = "UPDATE" Then
        Call update
    Else
        Call EditData(strid)
        Call EnabledAll(True)
        txtnopart.Enabled = False
    End If
End Sub

Private Sub chapus_Click()
    x = MsgBox("Yakin Mau Dihapus...?", vbYesNo + vbInformation, "Hapus Data")
    If x = vbYes Then
        hapus = "update code set deleted=1 where id='" & strid & "' "
        Set rscode = conn.Execute(hapus)
        MsgBox ("DATA SUDAH TERHAPUS")
    
        Call AddComboBox
        Call ClearText
        Call EnabledAll(False)
        csimpan.Enabled = False
        cedit.Enabled = False
        cedit.Caption = "EDIT"
        cbaru.Enabled = True
        chapus.Enabled = True
        txtcari.Enabled = True
        
        strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
            "from code where deleted=0 and category=3"
        Call FillListview(strsql)
    End If
End Sub

Private Sub cmbproses_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtjic.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            tsize.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub crefresh_Click()
    Call search(Empty)
    Call cbatal_Click
    txtcari.Enabled = True
End Sub

Private Sub csearch_Click()
    Call search(txtcari)
End Sub

Private Sub csimpan_Click()
    Call save
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If ColumnHeader.Index - 1 <> CURR_COL Then
        ListView1.SortOrder = 0
    Else
        ListView1.SortOrder = Abs(ListView1.SortOrder - 1)
    End If
    
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub ListView1_DblClick()
    Call EditData(strid)
    
    cbaru.Enabled = False
    cedit.Caption = "UPDATE"
    chapus.Enabled = False
    
    Call EnabledAll(True)
    txtnopart.Enabled = False
    
End Sub

Private Sub ListView1_GotFocus()
    cedit.Enabled = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    strid = Trim(Item.Text)
End Sub

Private Sub cbaru_Click()
    Call EnabledAll(True)
    csimpan.Enabled = True
    cbaru.Enabled = False
    chapus.Enabled = False
    cedit.Enabled = False
    txtnopart.SetFocus
End Sub

Private Sub Form_Load()
    Call AddComboBox
    Call ClearText
    Call EnabledAll(False)
    txtcari.Enabled = True
    Call DrawListview
    
    strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
        "from code where deleted=0 and category=3"
    Call FillListview(strsql)
    
    csimpan.Enabled = False
    cedit.Enabled = False
End Sub

Public Sub AddComboBox()
    With cmbproses
        .Clear
        .AddItem "FINISH GOODS"
        .AddItem "DRAWING"
        .AddItem "CUTTING"
        .AddItem "WELDING"
        .AddItem "GINTING RIB"
        .AddItem "MOLDING"
        .AddItem "RADIUS"
        .AddItem "ISI FILLER"
        .AddItem "LIPAT"
        .AddItem "FIBRO"
        .AddItem "FINISHING"
    End With
    With cmbtype
        .Clear
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
        .AddItem "D"
        .AddItem "E"
        .AddItem "F"
        .AddItem "G"
        .AddItem "H"
        .AddItem "J"
        .AddItem "K"
        .AddItem "L"
        .AddItem "M"
        .AddItem "N"
        .AddItem "P"
        .AddItem "S"
        .AddItem "T"
        .AddItem "U"
        .AddItem "Y"
        .AddItem "Z"
    End With
    With txtradius
        .Clear
        .AddItem "8"
        .AddItem "10"
        .AddItem "13"
    End With
    With txtwidth
        .Clear
        .AddItem "8"
        .AddItem "10"
        .AddItem "13"
    End With
    With tmetal
        .Clear
        .AddItem "316"
        .AddItem "304"
        .AddItem "SPCC"
        .AddItem "410"
        .AddItem "TITANIUM"
        .AddItem "BRASS"
        .AddItem "COPPER"
        .AddItem "MONEL"
        .AddItem "ALUMUNIUM"
    End With
End Sub

Public Sub ClearText()
    For Each A In Me
        If TypeOf A Is TextBox Then A.Text = ""
        If TypeOf A Is ComboBox Then A.Text = "- PILIH -"
    Next A
End Sub

Public Sub EnabledAll(status As Boolean)
    For Each A In Me
        If TypeOf A Is TextBox Then A.Enabled = status
        If TypeOf A Is ComboBox Then A.Enabled = status
    Next A
End Sub

Public Sub DrawListview()
    With ListView1
        .View = lvwReport
        .GridLines = True
        .MultiSelect = True
        .FullRowSelect = True
        .HotTracking = True
        .HoverSelection = True
        ' tambahkan kolom2 ke, , Judul,lebar,aligment
        .ColumnHeaders.Add 1, , "ID", 0
        .ColumnHeaders.Add 2, , "Stock Code", 2000
        .ColumnHeaders.Add 3, , "Process", 1500
        .ColumnHeaders.Add 4, , "JIC", 2500
        .ColumnHeaders.Add 5, , "Size", 1500
        .ColumnHeaders.Add 6, , "Type", 1500
        .ColumnHeaders.Add 7, , "Size", 1000
        .ColumnHeaders.Add 8, , "Metal", 1000
        .ColumnHeaders.Add 9, , "Radius", 1000
        .ColumnHeaders.Add 10, , "Width", 1000
        .ColumnHeaders.Add 11, , "Partition", 1000
        .ColumnHeaders.Add 12, , "Partition2", 1000
        .ColumnHeaders.Add 13, , "Stamp Marking", 3000
    End With
End Sub

Public Sub FillListview(strsql As String)
    Dim lst As ListItem   ' ListItem object
    If rscode.State = 1 Then rscode.Close
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    ListView1.ListItems.Clear
    
    If Not rscode.EOF Then
        rscode.MoveFirst
        Do While Not rscode.EOF
            Set lst = ListView1.ListItems.Add
            lst.Text = rscode!id
            lst.SubItems(1) = rscode!no_part
            lst.SubItems(2) = rscode!proses
            lst.SubItems(3) = rscode!jic
            lst.SubItems(4) = rscode!Size
            lst.SubItems(5) = rscode!Type
            lst.SubItems(6) = rscode!size_2
            lst.SubItems(7) = IIf(IsNull(rscode!Metal), "", rscode!Metal)
            lst.SubItems(8) = rscode!radius
            lst.SubItems(9) = rscode!Width
            lst.SubItems(10) = rscode!Partition
            lst.SubItems(11) = IIf(IsNull(rscode!Partition2), "", rscode!Partition2)
            lst.SubItems(12) = rscode!marking_stamp_lokal_or
            rscode.MoveNext
        Loop
    End If
End Sub

Public Sub EditData(strcode As String)
    If rscode.State = 1 Then rscode.Close
    strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
        "from code where deleted=0 and category=3 and id='" & strcode & "'"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    If Not rscode.EOF Then
        txtnopart.Text = rscode!no_part
        cmbproses.Text = rscode!proses
        txtjic.Text = rscode!jic
        txtsize.Text = rscode!Size
        cmbtype.Text = rscode!Type
        tsize.Text = rscode!size_2
        tmetal.Text = IIf(IsNull(rscode!Metal), "", rscode!Metal)
        txtradius.Text = rscode!radius
        txtwidth.Text = rscode!Width
        txtpartitionbar.Text = rscode!Partition
        txtpartitionbar2.Text = IIf(IsNull(rscode!Partition2), "", rscode!Partition2)
        tmarking_or_lokal.Text = rscode!marking_stamp_lokal_or
    End If
End Sub

Public Sub save()
    If rscode.State = 1 Then rscode.Close
    strsql = "Select no_part from code where no_part='" & txtnopart.Text & "' and deleted =0"
    rscode.Open strsql, conn, adOpenDynamic, adLockOptimistic
    
    If Not rscode.EOF Then
        MsgBox "Maaf, kode part sudah ada.", vbOKOnly + vbCritical, "Informasi"
        txtnopart.SetFocus
    Else
        strsql = "insert into code(no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or,deleted,category) " & _
            "values ('" & txtnopart & "','" & cmbproses & "','" & txtjic & "','" & txtsize & "','" & cmbtype & "','" & tsize & "','" & tmetal & "' " & _
            ",'" & txtradius & "','" & txtwidth & "','" & txtpartitionbar & "','" & txtpartitionbar2 & "','" & tmarking_or_lokal & "',0,3) "
        conn.Execute (strsql)
        
        MsgBox "Data Sudah tersimpan", vbOKOnly + vbInformation, "Informasi"
        
        Call AddComboBox
        Call ClearText
        Call EnabledAll(False)
        csimpan.Enabled = False
        cedit.Enabled = False
        cbaru.Enabled = True
        txtcari.Enabled = True
        
        strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
            "from code where deleted=0 and category=3"
        Call FillListview(strsql)
    End If
End Sub

Public Sub update()
    strsql = "update code set proses='" & cmbproses & "',jic='" & txtjic & "', size='" & txtsize & "', " & _
        "type='" & cmbtype & "', size_2='" & tsize & "', metal='" & tmetal & "',radius='" & txtradius & "', " & _
        "width='" & txtwidth & "',partition='" & txtpartitionbar & "',partition2='" & txtpartitionbar2 & "', " & _
        "marking_stamp_lokal_or='" & tmarking_or_lokal & "' where no_part='" & txtnopart & "'"
    conn.Execute (strsql)
    MsgBox "Data Berhasil Di Ubah"
    
    Call AddComboBox
    Call ClearText
    Call EnabledAll(False)
    csimpan.Enabled = False
    cedit.Enabled = False
    cedit.Caption = "EDIT"
    cbaru.Enabled = True
    chapus.Enabled = True
    
    strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
        "from code where deleted=0 and category=3"
    Call FillListview(strsql)
    
End Sub

Public Sub search(strcari As String)
    strsql = "Select id,no_part,proses,jic,size,type,size_2,metal,radius,width,partition,partition2,marking_stamp_lokal_or " & _
        "from code where deleted=0 and category=3 and (no_part like '%" & strcari & "%' OR proses like '%" & strcari & "%' " & _
        "or jic like '%" & strcari & " %' or size like '%" & strcari & "%' or type like '%" & strcari & "%' " & _
        "or size_2 like '%" & strcari & "%' or metal like '%" & strcari & "%' or radius like '%" & strcari & "%' " & _
        "or width like '%" & strcari & "%' or partition like '%" & strcari & "%' or partition2 like '%" & strcari & "%'" & _
        "or marking_stamp_lokal_or like '%" & strcari & "%')"
    Call FillListview(strsql)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
End Sub

Private Sub tmetal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtradius.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub tmarking_or_lokal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            csimpan.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub tsize_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            tmetal.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtjic_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtsize.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtnopart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            cmbproses.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtpartitionbar_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtpartitionbar2.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtpartitionbar2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            tmarking_or_lokal.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtradius_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtwidth.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtsize_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            cmbtype.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub txtwidth_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case vbKeyReturn
            txtpartitionbar.SetFocus
            Exit Sub
    End Select
End Sub


