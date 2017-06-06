VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form_Cetak_CS 
   Caption         =   "Form Cetak CS"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicEncode2 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   5520
      ScaleHeight     =   2835
      ScaleWidth      =   4875
      TabIndex        =   14
      Top             =   3840
      Width           =   4935
   End
   Begin VB.PictureBox picEncode 
      AutoRedraw      =   -1  'True
      Height          =   2895
      Left            =   360
      ScaleHeight     =   2835
      ScaleWidth      =   4875
      TabIndex        =   13
      Top             =   3840
      Width           =   4935
   End
   Begin VB.ComboBox cmb_category 
      Height          =   315
      ItemData        =   "Form_Cetak_CS.frx":0000
      Left            =   1800
      List            =   "Form_Cetak_CS.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   720
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6975
      Begin VB.OptionButton optgl 
         Caption         =   "Tanggal"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optnoslip 
         Caption         =   "No Slip"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cnoslip1 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cnoslip2 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dttanggal1 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41729
      End
      Begin MSComCtl2.DTPicker dttanggal2 
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   41729
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
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label3 
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
         Left            =   4080
         TabIndex        =   9
         Top             =   840
         Width           =   120
      End
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   480
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrinterCollation=   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form_Cetak_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Const RESOURCETYPE_DISK = &H1

Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As String
   lpRemoteName As String
   lpComment As String
   lpProvider As String
End Type

Private Sub cmb_category_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case cmb_category.Text
            Case "SWG"
                strcategory = "and category is null"
            Case "DJG"
                strcategory = "and category =2"
            Case "SMG"
                strcategory = "and category =3"
            Case ""
                MsgBox "Mohon pilih category terlebih dahulu", vbCritical + vbOKOnly, "Warning"
                Exit Sub
        End Select
        
        If rscompletion_slip.State = 1 Then rscompletion_slip.Close
        strsql = "select no_slip from completion_slip WHERE deleted=0 " & strcategory & " order by id asc"
        rscompletion_slip.Open strsql, conn, adOpenDynamic, adLockOptimistic
        rscompletion_slip.MoveFirst
        cnoslip1.Clear
        cnoslip2.Clear
        Do While Not rscompletion_slip.EOF
            cnoslip1.AddItem rscompletion_slip("no_slip")
            cnoslip2.AddItem rscompletion_slip("no_slip")
            rscompletion_slip.MoveNext
        Loop
        
        optgl.SetFocus
    End If
    
End Sub

Private Sub cmdprint_Click()
    Dim qrEncoder As New QRCodeEncoder
    Dim networkResource As NETRESOURCE
    With networkResource
        .dwType = RESOURCETYPE_DISK
        .lpLocalName = ""
        .lpRemoteName = "\\192.168.10.250\D$"
        .lpProvider = ""
    End With
    
    lon = WNetAddConnection2(networkResource, "090375", "administrator", 0)
    
    
    qrEncoder.QRCodeEncodeMode = ENCODE_MODE.ENCODE_MODE_BYTE
    qrEncoder.QRCodeScale = 4
    qrEncoder.QRCodeVersion = 7
    qrEncoder.QRCodeErrorCorrect = ERROR_CORRECTION.ERROR_CORRECTION_M
    
    Select Case cmb_category.Text
        Case "SWG"
            strcategory = "and category is null"
        Case "DJG"
            strcategory = "and category =2"
        Case "SMG"
            strcategory = "and category =3"
        Case ""
            MsgBox "Mohon pilih category terlebih dahulu", vbCritical + vbOKOnly, "Warning"
            Exit Sub
    End Select

    If optgl.Value = True Then
        If vread.State = 1 Then vread.Close
        strsql = "select id, marking_stamp_or from completion_slip where " & _
                "finish_date>= '" & Format(dttanggal1.Value, "YYYY/mm/dd") & "' " & _
                "and finish_date<='" & Format(dttanggal2.Value, "YYYY/mm/dd") & "' " & _
                "" & strcategory & ""
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        
        If Not vread.EOF Then
            vread.MoveFirst
            Do While Not vread.EOF
                picEncode.Picture = qrEncoder.EncodeVB6(vread!marking_stamp_or)
                SavePicture picEncode.Picture, "\\192.168.10.250\d$\" & RTrim(vread!id) & ".bmp"
                strsql = "UPDATE completion_slip SET " & _
                    "image_marking = (SELECT BulkColumn from Openrowset(Bulk 'D:\" & Trim(vread!id) & ".bmp', Single_Blob) as Image) where id = '" & vread!id & "'"
                conn.Execute (strsql)
                Kill "\\192.168.10.250\d$\" & RTrim(vread!id) & ".bmp"
                vread.MoveNext
            Loop
        End If
    
        With CrystalReport2
            .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
            '.SelectionFormula = "{completion_slip.finish_date} = datetime(" & Format(dttanggal1.Value, "YYYY-mm-dd") & ") " & _
                " and {completion_slip.finish_date}<=datetime(" & Format(dttanggal2.Value, "YYYY-mm-dd") & ") "
            .SelectionFormula = "{completion_slip.finish_date} >= '" & Format(dttanggal1.Value, "YYYY-mm-dd") & "' " & _
                "and {completion_slip.finish_date} <= '" & Format(dttanggal2.Value, "YYYY-mm-dd") & "'"
                
            If cmb_category.Text = "SWG" Then
                .ReportFileName = App.Path & "\cs.rpt"
            ElseIf cmb_category.Text = "DJG" Then
                .ReportFileName = App.Path & "\cs_DJG.rpt"
            Else
                .ReportFileName = App.Path & "\cs_SMG.rpt"
            End If

            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .RetrieveDataFiles
            .WindowState = crptMaximized
            .Action = 1
        End With
    ElseIf optnoslip.Value = True Then
        If vread.State = 1 Then vread.Close
        strsql = "Select id,marking_stamp_or from completion_slip where no_slip >='" & Trim(cnoslip1.Text) & "' and " & _
            "no_slip <='" & Trim(cnoslip2.Text) & "'"
        vread.Open strsql, conn, adOpenDynamic, adLockOptimistic
        
        If Not vread.EOF Then
            vread.MoveFirst
            Do While Not vread.EOF
                picEncode.Picture = qrEncoder.EncodeVB6(vread!marking_stamp_or)
                SavePicture picEncode.Picture, "\\192.168.10.250\d$\" & RTrim(vread!id) & ".bmp"
                strsql = "UPDATE completion_slip SET " & _
                    "image_marking = (SELECT BulkColumn from Openrowset(Bulk 'D:\" & Trim(vread!id) & ".bmp', Single_Blob) as Image) where id = '" & vread!id & "'"
                conn.Execute (strsql)
                Kill "\\192.168.10.250\d$\" & RTrim(vread!id) & ".bmp"
                vread.MoveNext
            Loop
        End If
    
        With CrystalReport2
            .Connect = "DSN=produksi;UID=sa;PWD=admin123;DSQ=purchasing"
            .SelectionFormula = "{completion_slip.no_slip} >='" & Trim(cnoslip1.Text) & "' and {completion_slip.no_slip} <='" & Trim(cnoslip2.Text) & "' "
            If cmb_category.Text = "SWG" Then
                .ReportFileName = App.Path & "\cs.rpt"
            ElseIf cmb_category.Text = "DJG" Then
                .ReportFileName = App.Path & "\cs_djg.rpt"
            Else
                .ReportFileName = App.Path & "\cs_SMG.rpt"
            End If
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .RetrieveDataFiles
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
    lon = WNetCancelConnection2(UNC, 0, True)
End Sub

Private Sub Form_Activate()
'Call db
    cmb_category.SetFocus
End Sub

Private Sub Form_Load()
Me.dttanggal1.Value = Date
Me.dttanggal2.Value = Date
End Sub

