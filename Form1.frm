VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Countdown"
   ClientHeight    =   3810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   105
      TabIndex        =   0
      Top             =   15
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   5054
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time Remaining"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Reminder"
         Object.Width           =   132292
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   2625
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add Date"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Timer1.Enabled = True
    LoadSettings
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ListView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
End Sub

Private Sub mnuFileAdd_Click()
    On Error Resume Next
    Dim ReminderText As String, i As Long
    Dim CountdownDate As Date, s As String
    Dim li As ListItem
    ReminderText = InputBox("Enter the reminder text:")
    If Len(ReminderText) = 0 Then Exit Sub
    s = InputBox("Enter the countdown date:")
    If Len(s) = 0 Then Exit Sub
    CountdownDate = CDate(s)
    AddReminder ReminderText, CountdownDate
End Sub

Private Sub AddReminder(s As String, d As Date)
    Dim li As ListItem
    Set li = ListView1.ListItems.Add(, , " ")
    li.ListSubItems.Add , , d
    li.ListSubItems.Add , , s
End Sub

Private Sub mnuFileExit_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim i As Long, s As String, s2() As String
    Dim SecondsToTarget As Long 'total secs to target date
    Dim dd As Long 'days to target date
    Dim hh As Long 'hrs to target date
    Dim mm As Long 'mins to target date
    Dim ss As Long 'secs to target date

    With ListView1
        For i = 1 To .ListItems.Count
            s = .ListItems(i).ListSubItems(1).Text
            SecondsToTarget = DateDiff("s", Now, s)
            GetRemaining SecondsToTarget, dd, hh, mm, ss
            s = dd & " days " & hh & " hours " & mm & " minutes " & ss & " seconds"
            .ListItems(i).Text = s
        Next i
        .Refresh
    End With
End Sub

Private Sub GetRemaining(ByVal i As Long, DaysLeft As Long, HoursLeft As Long, MinsLeft As Long, SecsLeft As Long)
    
    'extract number of days
    DaysLeft = i \ 86400
    i = i Mod 86400
    
    'extract number of hrs
    HoursLeft = i \ 3600
    i = i Mod 3600
    
    'extract number of mins
    MinsLeft = i \ 60
    i = i Mod 60
    
    'extract number of secs
    SecsLeft = i
    
End Sub

Private Sub SaveSettings()
    Dim s As String, d As Date, i As Long, t As String
    For i = 1 To ListView1.ListItems.Count
        t = t & ListView1.ListItems(i).ListSubItems(1) & vbTab & ListView1.ListItems(i).ListSubItems(2) & vbCrLf
    Next i
    SaveSetting App.Title, "Preferences", "Reminders", t
End Sub

Private Sub LoadSettings()
    Dim s As String, s2() As String, d As Date, t() As String, i As Long
    s = GetSetting(App.Title, "Preferences", "Reminders", "")
    If Len(s) = 0 Then
        Exit Sub
    End If
    s2 = Split(s, vbCrLf)
    For i = 0 To UBound(s2, 1)
        s = s2(i)
        t = Split(s, vbTab)
        AddReminder t(1), CDate(t(0))
    Next i
End Sub
