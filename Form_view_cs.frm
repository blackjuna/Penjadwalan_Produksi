VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form view_cs 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "view_cs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As CRAXDRT.Database
Dim rpt As CRAXDRT.Report
Dim appl As CRAXDRT.Application
Dim WithEvents sect As CRAXDRT.Section
Attribute sect.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    
    Set appl = New CRAXDRT.Application
    Set rpt = appl.OpenReport(App.Path & "\cs.rpt")
    rpt.Database.LogOnServer "p2sodbc.dll", "produksi", "", "sa", "admin123"
    Set db = rpt.Database
    Set sect = rpt.Sections("Section5")
    
    rscode.Open "SELECT * FROM completion_slip where no_slip='2017-002-0000208'", conn, adOpenDynamic, adLockOptimistic
    rpt.Database.SetDataSource rscode, 3, 1
    
    CRViewer1.ReportSource = rpt
    CRViewer1.ViewReport
    CRViewer1.Zoom 1
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cn.Close
End Sub

Private Sub sect_Format(ByVal pFormattingInfo As Object)
Dim bmp As StdPicture

    With sect.ReportObjects
        'Check picture file exist or not using
        'FileSystemObject.FileExists
        'Set bmp = LoadPicture(App.Path & "\cs\" & .Item("Field3").Value)
        Set bmp = LoadPicture(App.Path & "\cs\warning.jpg")
        Set .Item("Picture1").FormattedPicture = bmp
    End With
End Sub

