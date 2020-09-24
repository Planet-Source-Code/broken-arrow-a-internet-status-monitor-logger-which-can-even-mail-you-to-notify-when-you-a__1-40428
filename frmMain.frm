VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "iMon"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckCheck 
      Left            =   7440
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.Data dbTemp 
      Caption         =   "dbTemp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "[ Log ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5055
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   7935
      Begin VB.CommandButton cmdDelData 
         Caption         =   "Delete"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdClearDatabase 
         Caption         =   "Clear Record"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Data dbLog 
         Caption         =   "dbLog"
         Connect         =   "Access"
         DatabaseName    =   "E:\Joy\Joy\Project\Visual Basic\Utility\iMon\iMon.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblLog"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLogDisplay 
         Caption         =   "dbLogDisplay"
         Connect         =   "Access"
         DatabaseName    =   "E:\Joy\Joy\Project\Visual Basic\Utility\iMon\iMon.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   $"frmMain.frx":0442
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "frmMain.frx":0578
         Height          =   3135
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   5530
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin Crystal.CrystalReport crpReport 
         Left            =   6720
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "Report"
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdShowData 
         Caption         =   "Show"
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtToDate 
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   225
         Width           =   975
      End
      Begin VB.TextBox txtFromDate 
         Height          =   285
         Left            =   3720
         TabIndex        =   25
         Text            =   "01/01/1901"
         Top             =   225
         Width           =   975
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFC0&
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   4080
         Width           =   7695
      End
      Begin VB.Label lblSummery 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connection gone down"
         ForeColor       =   &H00C0C0FF&
         Height          =   280
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Width           =   7695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "to"
         Height          =   195
         Left            =   4800
         TabIndex        =   26
         Top             =   270
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Show log from"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
         Top             =   270
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Setting ]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtTolerance 
         Height          =   285
         Left            =   6120
         TabIndex        =   43
         Text            =   ".03"
         Top             =   315
         Width           =   375
      End
      Begin MSWinsockLib.Winsock sckMail 
         Left            =   6960
         Top             =   1920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog dlgSound 
         Left            =   2280
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "WAV Files|*.wav|All Files|*.*"
      End
      Begin VB.CommandButton cmdResumeSound 
         Caption         =   "..."
         Height          =   285
         Left            =   3600
         TabIndex        =   39
         Top             =   3120
         Width           =   375
      End
      Begin VB.CommandButton cmdDownSound 
         Caption         =   "..."
         Height          =   285
         Left            =   3600
         TabIndex        =   38
         Top             =   2800
         Width           =   375
      End
      Begin VB.TextBox txtResumeSound 
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   3120
         Width           =   2400
      End
      Begin VB.TextBox txtDownSound 
         Height          =   285
         Left            =   1200
         TabIndex        =   36
         Top             =   2800
         Width           =   2400
      End
      Begin VB.PictureBox picTray 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   7320
         Picture         =   "frmMain.frx":0593
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "Minimize"
         Height          =   255
         Left            =   6045
         TabIndex        =   22
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   255
         Left            =   5120
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "[ Scan Status ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   660
         Left            =   4200
         TabIndex        =   17
         Top             =   2760
         Width           =   3615
         Begin MSComctlLib.ProgressBar prgTotal 
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   375
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   318
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblHost 
            AutoSize        =   -1  'True
            Caption         =   "Host..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   30
            Top             =   195
            Width           =   420
         End
         Begin VB.Image imgOnline 
            Height          =   360
            Left            =   3120
            Picture         =   "frmMain.frx":09D5
            Stretch         =   -1  'True
            Top             =   195
            Width           =   360
         End
         Begin VB.Image imgOffline 
            Height          =   360
            Left            =   3120
            Picture         =   "frmMain.frx":0E17
            Stretch         =   -1  'True
            Top             =   202
            Width           =   360
         End
      End
      Begin VB.TextBox txtSMTPServer 
         Height          =   285
         Left            =   4200
         TabIndex        =   15
         Text            =   "Mail.Hope-Tech.Net"
         Top             =   1995
         Width           =   3615
      End
      Begin VB.TextBox txtHostDescription 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   315
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog dlgTextFile 
         Left            =   6240
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Chose the text file to save log..."
         Filter          =   "Text File|*.txt|Log File|*.log|Data File|*.dat|All Files|*.*"
      End
      Begin VB.TextBox txtMailResume 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1395
         Width           =   3615
      End
      Begin VB.CommandButton cmdTextFile 
         Caption         =   "..."
         Height          =   285
         Left            =   7440
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtTextFile 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton cmdHostRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   865
         Width           =   735
      End
      Begin VB.CommandButton cmdHostUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   865
         Width           =   735
      End
      Begin VB.CommandButton cmdHostAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   865
         Width           =   735
      End
      Begin MSComctlLib.ListView lstHost 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Host IP/Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "sec(s)"
         Height          =   195
         Left            =   6600
         TabIndex        =   44
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Minimum offline tolerance"
         Height          =   195
         Left            =   4200
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Resume sound"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   3165
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Down sound"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2845
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SMTP Mail Server:"
         Height          =   195
         Left            =   4200
         TabIndex        =   14
         Top             =   1755
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Send mail on resume..."
         Height          =   195
         Left            =   4200
         TabIndex        =   9
         Top             =   1155
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Log text file..."
         Height          =   195
         Left            =   4200
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Host IP/Name"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Response As String, CancelScan As Boolean, sck_Error As Boolean

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 4
End Type

Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_RBUTTONDOWN = &H204

Sub CompactDatabase(Database_Name As String)
If Dir(Database_Name & ".tmp") <> "" Then Kill Database_Name & ".tmp"
DBEngine.CompactDatabase Database_Name, Database_Name & ".tmp"
Kill Database_Name
Name Database_Name & ".tmp" As Database_Name
End Sub

Private Sub cmdClearDatabase_Click()
If MsgBox("Are you sure that you want to clear the records?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm clear") = vbNo Then Exit Sub

dbTemp.Database.QueryDefs("qryDeleteAll").Execute
dbLogDisplay.Refresh
End Sub

Private Sub cmdDelData_Click()
dbLogDisplay.Recordset.MoveFirst
dbLogDisplay.Recordset.Move MSFlexGrid1.Row - 1
dbLogDisplay.Recordset.Delete
dbLogDisplay.Refresh
End Sub

Private Sub cmdDownSound_Click()
On Error GoTo DownSound_Cancelled

dlgSound.ShowOpen
txtDownSound = dlgSound.FileName
ExitDownSound_Sub:
Exit Sub

DownSound_Cancelled:
Resume ExitDownSound_Sub
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHostAdd_Click()
Dim Found As Boolean, a As Long
For a = 1 To lstHost.ListItems.Count
    If lstHost.ListItems(a).Text = txtHost.Text Then Found = True
    If Found Then Exit For
Next
If Found Then Exit Sub

With lstHost.ListItems
    .Add , , txtHost
    .Item(.Count).SubItems(1) = txtHostDescription
End With

dbTemp.RecordSource = "tblHost"
dbTemp.Refresh

With dbTemp.Recordset
    .AddNew
    .Fields("Host") = txtHost
    .Fields("Description") = txtHostDescription
    .Update
End With

If lstHost.ListItems.Count > 1 And cmdHostRemove.Enabled = False Then cmdHostRemove.Enabled = True
End Sub

Private Sub cmdHostRemove_Click()
If MsgBox("Are you sure that you want to remove the Host '" & txtHost & "' from the scan list?", vbYesNo + vbQuestion + vbDefaultButton2, "Remove Host...") = vbNo Then Exit Sub
lstHost.ListItems.Remove (lstHost.SelectedItem.Index)

dbTemp.RecordSource = "SELECT * FROM tblHost WHERE tblHost.Host = '" & txtHost & "'"
dbTemp.Refresh
dbTemp.Recordset.Delete

If lstHost.ListItems.Count = 1 Then cmdHostRemove.Enabled = False
End Sub

Private Sub cmdHostUpdate_Click()
cmdHostRemove_Click
cmdHostAdd_Click
End Sub

Private Sub cmdMinimize_Click()
Me.Hide
End Sub

Private Sub cmdReport_Click()
crpReport.ReportFileName = App.Path & "\Report.rpt"
crpReport.DataFiles(0) = App.Path & "\iMon.mdb"
crpReport.Action = 1
End Sub

Private Sub cmdResumeSound_Click()
On Error GoTo ResumeSound_Cancelled

dlgSound.ShowOpen
txtResumeSound = dlgSound.FileName
ExitResumeSound_Sub:
Exit Sub

ResumeSound_Cancelled:
Resume ExitResumeSound_Sub
End Sub

Private Sub cmdShowData_Click()
'On Error Resume Next
With dbLogDisplay
    .RecordSource = "SELECT FORMAT(DownDate,'dd/mm/yy') AS [Down Date], FORMAT(DownDate,'hh:nn:ss') AS [Down Time], FORMAT(ResumeDate,'dd/mm/yy') AS [Resume Date], FORMAT(ResumeDate,'hh:nn:ss') AS [Resume Time], FORMAT(DATEDIFF('s',DownDate,ResumeDate)/60, '0.00') AS [Duration Min(s)] FROM tblLog WHERE DownDate >= #" & txtFromDate & "# AND ResumeDate < #" & CDate(txtToDate) + 1 & "# ORDER BY tblLog.DownDate DESC"
    .Refresh
End With

dbTemp.RecordSource = "SELECT SUM(DATEDIFF('s',DownDate,ResumeDate)/60) AS TotalDownTime FROM tblLog WHERE DownDate >= #" & txtFromDate & "# AND ResumeDate < #" & CDate(txtToDate) + 1 & "#"
dbTemp.Refresh
lblSummery = "Connection gone down " & dbLogDisplay.Recordset.RecordCount & " times and total downtime is " & Format(dbTemp.Recordset.Fields(0), "0.00") & " min(s)"
End Sub

Private Sub cmdTextFile_Click()
On Error GoTo Save_Cancelled

dlgTextFile.ShowSave
If Dir(dlgTextFile.FileName) = "" Then Open dlgTextFile.FileName For Binary As #1
txtTextFile = dlgTextFile.FileName
Exit_Sub:
Exit Sub

Save_Cancelled:
Resume Exit_Sub
End Sub

Private Sub Form_Load()
ShowStatus "Session started on " & Format(Date, "Long Date") & " at " & Time & " ..."

dbTemp.DatabaseName = App.Path & "\iMon.mdb"

dbLogDisplay.DatabaseName = App.Path & "\iMon.mdb"
'dbLogDisplay.RecordSource = "SELECT FORMAT(DownDate,'dd/mm/yy') AS [Down Date], FORMAT(DownDate,'hh:nn:ss') AS [Down Time], FORMAT(ResumeDate,'dd/mm/yy') AS [Resume Date], FORMAT(ResumeDate,'hh:nn:ss') AS [Resume Time], FORMAT(DATEDIFF('s',DownDate,ResumeDate)/60, '0.00') AS [Duration Min(s)] From tblLog ORDER BY tblLog.DownDate DESC"
dbLogDisplay.Refresh
MSFlexGrid1.ColWidth(0) = MSFlexGrid1.Width / 5 - 70
MSFlexGrid1.ColWidth(1) = MSFlexGrid1.Width / 5 - 70
MSFlexGrid1.ColWidth(2) = MSFlexGrid1.Width / 5 - 70
MSFlexGrid1.ColWidth(3) = MSFlexGrid1.Width / 5 - 70
MSFlexGrid1.ColWidth(4) = MSFlexGrid1.Width / 5 - 70

dbTemp.RecordSource = "SELECT SUM(DATEDIFF('s',DownDate,ResumeDate)/60) AS TotalDownTime FROM tblLog"
dbTemp.Refresh
lblSummery = "Connection gone down " & dbLogDisplay.Recordset.RecordCount & " times and total downtime is " & Format(dbTemp.Recordset.Fields(0), "0.00") & " min(s)"

dbLog.DatabaseName = App.Path & "\iMon.mdb"
'dbLog.RecordSource = "tblLog"
dbLog.Refresh

ShowStatus "Database loaded"

Load_Setting
ShowStatus "Settings loaded"

LoadHostList
ShowStatus "Host list loaded"

Me.Show
Me.Refresh

If FileExists(txtTextFile) Then
    Open txtTextFile For Binary As #1
    Seek #1, LOF(1) + 1
End If

txtToDate = Date

CreateIcon

ShowStatus "Starting scan..."
If lstHost.ListItems.Count > 0 Then KeepScanning
End Sub

Sub Load_Setting()
txtTextFile = GetSetting(App.Title, "Setting", "Log Text File", App.Path & "\Log.txt")
txtMailResume = GetSetting(App.Title, "Setting", "Mail on Resume", "sKabir@Hope-Tech.Net")
txtSMTPServer = GetSetting(App.Title, "Setting", "SMTP Server", "Mail.Hope-Tech.Net")
txtDownSound = GetSetting(App.Title, "Setting", "Down Sound", "")
txtResumeSound = GetSetting(App.Title, "Setting", "Resume Sound", "")
txtTolerance = GetSetting(App.Title, "Setting", "Offtime Tolerance", ".04")
End Sub

Sub Save_Setting()
SaveSetting App.Title, "Setting", "Log Text File", txtTextFile
SaveSetting App.Title, "Setting", "Mail on Resume", txtMailResume
SaveSetting App.Title, "Setting", "SMTP Server", txtSMTPServer
SaveSetting App.Title, "Setting", "Down Sound", txtDownSound
SaveSetting App.Title, "Setting", "Resume Sound", txtResumeSound
SaveSetting App.Title, "Setting", "Offtime Tolerance", txtTolerance
End Sub

Private Sub Form_Unload(Cancel As Integer)
CancelScan = True

Save_Setting
Close #1
DeleteIcon

dbTemp.Database.Close
dbLog.Database.Close
dbLogDisplay.Database.Close
CompactDatabase App.Path & "\iMon.mdb"

End
End Sub

Private Sub lstHost_Click()
If lstHost.ListItems.Count < 1 Then Exit Sub
txtHost = lstHost.SelectedItem.Text
txtHostDescription = lstHost.SelectedItem.SubItems(1)
End Sub

Sub LoadHostList()
dbTemp.RecordSource = "tblHost"
dbTemp.Refresh

If dbTemp.Recordset.RecordCount < 1 Then Exit Sub

dbTemp.Recordset.MoveLast
dbTemp.Recordset.MoveFirst

Dim a As Long
For a = 1 To dbTemp.Recordset.RecordCount
    If a > 1 Then dbTemp.Recordset.MoveNext
    
    With lstHost.ListItems
        .Add , , dbTemp.Recordset.Fields("Host")
        If dbTemp.Recordset.Fields("Description") <> vbNull Then .Item(.Count).SubItems(1) = dbTemp.Recordset.Fields("Description")
    End With
Next
End Sub

Sub ShowStatus(StatusText As String)
If Len(txtStatus) > 80 * 100 Then txtStatus = ""
txtStatus.SelStart = Len(txtStatus) + 1
txtStatus.SelText = StatusText & vbCrLf
End Sub

Sub KeepScanning()
If lstHost.ListItems.Count = 0 Then Exit Sub

Dim LastState As Boolean, CurHost As Long, LastDown As Date, LastResume As Date, LogTxt As String
LastState = True

While Not CancelScan
    DoEvents
    
    If CurHost = lstHost.ListItems.Count Then prgTotal.Value = 0
    If CurHost < lstHost.ListItems.Count Then CurHost = CurHost + 1 Else CurHost = 1
    prgTotal.Max = lstHost.ListItems.Count
    
    sck_Error = False
    lblHost = lstHost.ListItems(CurHost).Text
    sckCheck.RemoteHost = lstHost.ListItems(CurHost).Text
    sckCheck.Connect
    While sckCheck.State <> sckConnected And sck_Error = False And Not CancelScan
        DoEvents
    Wend
    prgTotal.Value = CurHost
    If sck_Error And LastState Then 'Down
        LastDown = Now
        LastState = False
        imgOffline.ZOrder
        ShowStatus ":(   Internet gone DOWN on " & Date & " at " & Time
        LogTxt = "Internet gone DOWN on " & Date & " at " & Time
        If Dir(txtTextFile) <> "" Then Put #1, , LogTxt
        PlayWAV txtDownSound
        picTray.Picture = imgOffline.Picture
        ModifyIcon
    End If
    If Not sck_Error And Not LastState Then 'Resumed
        LastResume = Now
        LastState = True
        imgOnline.ZOrder
        ShowStatus ":)   Internet RESUMED on " & Date & " at " & Time
        LogTxt = " and RESUMED on " & Date & " at " & Time & "    Downtime = " & Format(DateDiff("s", LastDown, LastResume) / 60, "0.##") & " min(s)" & vbCrLf
        If FileExists(txtTextFile) Then Put #1, , LogTxt
        PlayWAV txtResumeSound
        picTray.Picture = imgOnline.Picture
        ModifyIcon
        If DateDiff("s", LastDown, LastResume) > Val(txtTolerance) Then
            With dbLog.Recordset
                .AddNew
                .Fields("DownDate") = LastDown
                .Fields("ResumeDate") = LastResume
                .Update
            End With
            dbLogDisplay.Refresh
            
            dbTemp.RecordSource = "SELECT SUM(DATEDIFF('s',DownDate,ResumeDate)/60) AS TotalDownTime FROM tblLog"
            dbTemp.Refresh
            lblSummery = "Connection gone down " & dbLogDisplay.Recordset.RecordCount & " times and total downtime is " & Format(dbTemp.Recordset.Fields(0), "0.00") & " min(s)"
            If txtMailResume > "" Then SMTPSendEmail txtSMTPServer, "iMon", txtMailResume, "", txtMailResume, txtMailResume, "Internet Resume Report", "Internet resumed after a down state." & vbCrLf & vbCrLf & "Down              :     " & LastDown & vbCrLf & "Resume            :     " & LastResume & vbCrLf & "Down dureation    :     " & Format(DateDiff("s", LastDown, LastResume) / 60, "0.00") & " min(s)", sckMail, False
        End If
    End If
    sckCheck.Close
    While sckCheck.State <> sckClosed And Not CancelScan
        DoEvents
    Wend
Wend
End Sub



Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X / Screen.TwipsPerPixelX = WM_LBUTTONDOWN Then
    Me.Show
End If
End Sub

Private Sub sckCheck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sck_Error = True
ShowStatus "ERROR#" & Number & "> " & Description
End Sub

Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = frmMain.picTray.Picture
Tic.szTip = "iMon"
erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub ModifyIcon()
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = frmMain.picTray.Picture
Tic.szTip = "iMon"
erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub

Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA, erg As Long
Tic.cbSize = Len(Tic)
Tic.hwnd = frmMain.picTray.hwnd
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Sub PlayWAV(WAV_File As String)
On Error Resume Next

If Dir(WAV_File) = "" Then Exit Sub
PlaySound WAV_File, 0, 0
End Sub

Function FileExists(strPath As String) As Boolean
    strPath = Trim(strPath)
    If strPath = "" Then
        FileExists = False
        Exit Function
    End If
  FileExists = Len(Dir(strPath)) <> 0
End Function

Sub SaveErrorLog(Error_Description As String)
On Error Resume Next

Open App.Path & "\Error.Log" For Binary As #2
Seek #2, LOF(2) + 1
Put #2, , Error_Description
Close #2
End Sub

Sub SMTPSendEmail(SMTPServer As String, From As String, FromEmailAddress As String, ReceiptEmailAddress As String, ToName As String, ToEmailAddress As String, Subject As String, msgBody As String, sckObject As Winsock, Optional Show_Error As Boolean = True)
On Error Resume Next

'If InStr(FromEmailAddress, ";") > 0 Then FromEmailAddress = Left(FromEmailAddress, InStr(FromEmailAddress, ";") - 1)
If ReceiptEmailAddress = "" Then ReceiptEmailAddress = FromEmailAddress

If sckObject.State <> sckClosed Then sckObject.Close          'Make sure that the socket is closed
While sckObject.State <> sckClosed                           'Wait for the socket to be closed
    DoEvents
Wend
sckObject.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start

sckObject.Protocol = sckTCPProtocol                          'Set protocol for sending
sckObject.RemoteHost = SMTPServer                            'Set the SMTP server address
sckObject.RemotePort = 25                                    'Set the SMTP Port
sckObject.Connect                                            'Start connection

WaitFor ("220"): sckObject.SendData ("HELO worldcomputers.com" + vbCrLf) 'Connect
WaitFor ("250"): sckObject.SendData ("mail from:" + Chr(32) + FromEmailAddress + vbCrLf)
WaitFor ("250"): sckObject.SendData ("rcpt to:" + Chr(32) + ReceiptEmailAddress + vbCrLf)
WaitFor ("250"): sckObject.SendData ("data" + vbCrLf)
WaitFor ("354")
sckObject.SendData ("From:" + Chr(32) + From + vbCrLf + "Date:" + Chr(32) + Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600" + vbCrLf + "X-Mailer: iMon" + vbCrLf + "To:" + Chr(32) + ToName + vbCrLf + "Subject:" + Chr(32) + Subject + vbCrLf + vbCrLf)
sckObject.SendData (msgBody + vbCrLf + vbCrLf)           'Send mail body
sckObject.SendData ("." + vbCrLf)
WaitFor ("250"): sckObject.SendData ("quit" + vbCrLf)    'Disconnect
WaitFor ("221"): sckObject.Close                         'Disconnected
End Sub

Sub WaitFor(ResponseCode As String)
Dim Start As Single

Start = Timer ' Time event so won't get stuck in loop
While Len(Response) = 0
    DoEvents ' Let System keep checking for incoming response **IMPORTANT**
    If Timer - Start > 50 Then Exit Sub
Wend
While Left(Response, 3) <> ResponseCode
    DoEvents
    If Timer - Start > 50 Then Exit Sub
Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)
sckMail.GetData Response ' Check for incoming response *IMPORTANT*
End Sub
