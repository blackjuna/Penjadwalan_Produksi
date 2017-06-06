VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Form_preview_cs 
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
Attribute VB_Name = "Form_preview_cs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rpt As CRAXDRT.Report
Dim db As CRAXDRT.Database
Dim rs As New ADODB.Recordset
Dim WithEvents sect As CRAXDRT.Section
Attribute sect.VB_VarHelpID = -1

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Set rpt = crx.OpenReport(App.Path & "\cs.rpt")
Set db = rpt.Database
Set sect = rpt.Sections("Section5")
'rs.Open "SELECT * FROM Cake", cn, 1, 1
'rpt.Database.SetDataSource rs, 3, 1
rpt.Database.LogOnServer "p2sodbc.dll", "produksi", "", "sa", "admin123"
CRViewer1.ReportSource = rpt
rpt.RecordSelectionFormula = "{completion_slip.no_slip} ='2017-002-0000314' "
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
conn.Close
End Sub

Private Sub sect_Format(ByVal pFormattingInfo As Object)
Dim bmp As StdPicture
On Error Resume Next
With sect.ReportObjects
    Set .Item("picture2").FormattedPicture = LoadPicture(App.Path & "\321.bmp") 'default
'    If .Item("adoFileName").Value <> "" Then
'        Set bmp = LoadPicture(App.Path & "\Cake\" & .Item("adoFileName").Value)
'        Set .Item("picCake").FormattedPicture = bmp
'    End If
End With

End Sub



