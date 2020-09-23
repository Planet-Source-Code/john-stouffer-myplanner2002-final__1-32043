VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyPlanner2002"
   ClientHeight    =   4425
   ClientLeft      =   2700
   ClientTop       =   2250
   ClientWidth     =   7185
   Icon            =   "Planner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7185
   Begin TabDlg.SSTab SSTab1 
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   7805
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Calendar"
      TabPicture(0)   =   "Planner.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MonthView1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdRemove2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAdd2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdExit2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSave2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "List2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "General Tasks"
      TabPicture(1)   =   "Planner.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CommonDialog1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdExit"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "List1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdLoad"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSave"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdRemove"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdAdd"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Text1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Settings"
      TabPicture(2)   =   "Planner.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Option7"
      Tab(2).Control(1)=   "Option6"
      Tab(2).Control(2)=   "Text3"
      Tab(2).Control(3)=   "Option5"
      Tab(2).Control(4)=   "Option4"
      Tab(2).Control(5)=   "Option1"
      Tab(2).Control(6)=   "Option3"
      Tab(2).Control(7)=   "Option2"
      Tab(2).Control(8)=   "Check1"
      Tab(2).Control(9)=   "Check2"
      Tab(2).Control(10)=   "Image2"
      Tab(2).Control(11)=   "Label7"
      Tab(2).ControlCount=   12
      Begin VB.OptionButton Option7 
         Caption         =   "Nice Ice"
         Height          =   225
         Left            =   -74790
         TabIndex        =   32
         Top             =   2835
         Width           =   1170
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Salmon"
         Height          =   225
         Left            =   -74790
         TabIndex        =   31
         Top             =   3150
         Width           =   1170
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFC0C0&
         Height          =   3900
         Left            =   -71850
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "Planner.frx":035E
         Top             =   420
         Width           =   3900
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Custom"
         Height          =   225
         Left            =   -74790
         TabIndex        =   29
         Top             =   3780
         Width           =   1065
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Windows Standard"
         Height          =   225
         Left            =   -74790
         TabIndex        =   28
         Top             =   3465
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "The Blues"
         Height          =   225
         Left            =   -74790
         TabIndex        =   26
         Top             =   2520
         Width           =   1170
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Harmony"
         Height          =   225
         Left            =   -74790
         TabIndex        =   24
         Top             =   2205
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Neon Purple"
         Height          =   225
         Left            =   -74790
         TabIndex        =   22
         Top             =   1890
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Automatically Save Calendar Tasks"
         Height          =   225
         Left            =   -74790
         TabIndex        =   20
         ToolTipText     =   "When checked, automatically saves your tasks when Add or Remove is selected."
         Top             =   840
         Value           =   1  'Checked
         Width           =   2850
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Automatically Save General Tasks"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -74790
         TabIndex        =   18
         Top             =   525
         Value           =   1  'Checked
         Width           =   2850
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -74895
         TabIndex        =   13
         Top             =   630
         Width           =   2745
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Task"
         Default         =   -1  'True
         DownPicture     =   "Planner.frx":056B
         Height          =   960
         Left            =   -74895
         MaskColor       =   &H8000000F&
         Picture         =   "Planner.frx":0875
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add the above task to your list of general tasks."
         Top             =   1575
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove Task"
         Height          =   960
         Left            =   -73425
         Picture         =   "Planner.frx":0B7F
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Remove selected task from the list."
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save List"
         Height          =   960
         Left            =   -74895
         Picture         =   "Planner.frx":0E89
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "No need to use this if AutoSave is enabled."
         Top             =   2625
         Width           =   1275
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load List"
         Height          =   960
         Left            =   -73425
         Picture         =   "Planner.frx":1753
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Use this if you accidently press ""Clear List""."
         Top             =   2625
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFC0C0&
         Height          =   330
         Left            =   -74895
         TabIndex        =   15
         Top             =   1155
         Width           =   2745
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFC0C0&
         Height          =   3570
         ItemData        =   "Planner.frx":201D
         Left            =   -72060
         List            =   "Planner.frx":201F
         TabIndex        =   27
         ToolTipText     =   "Shows the general tasks to be done."
         Top             =   630
         Width           =   4005
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Planner"
         Height          =   435
         Left            =   -74895
         TabIndex        =   25
         ToolTipText     =   "Godbye!"
         Top             =   3780
         Width           =   2745
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFC0C0&
         Height          =   1620
         ItemData        =   "Planner.frx":2021
         Left            =   3045
         List            =   "Planner.frx":2023
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2625
         Width           =   4005
      End
      Begin VB.CommandButton cmdSave2 
         Caption         =   "Save Tasks"
         Height          =   750
         Left            =   210
         Picture         =   "Planner.frx":2025
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Manually save tasks for this date"
         Top             =   3045
         Width           =   2745
      End
      Begin VB.CommandButton cmdExit2 
         Caption         =   "Exit Planner"
         Height          =   435
         Left            =   210
         TabIndex        =   6
         ToolTipText     =   "Goodbye!"
         Top             =   3885
         Width           =   1275
      End
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "Add Item"
         Height          =   855
         Left            =   3045
         Picture         =   "Planner.frx":28EF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Adds the above text to the tasks list."
         Top             =   1680
         Width           =   1905
      End
      Begin VB.CommandButton cmdRemove2 
         Caption         =   "Remove Item"
         Height          =   855
         Left            =   5145
         Picture         =   "Planner.frx":2BF9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Removes the selected item from the task list."
         Top             =   1680
         Width           =   1905
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   3045
         TabIndex        =   2
         ToolTipText     =   "What should be done."
         Top             =   1260
         Width           =   4005
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "Planner.frx":2F03
         Left            =   3045
         List            =   "Planner.frx":2F5E
         TabIndex        =   1
         ToolTipText     =   "You can put in your own time by deleting the standard text and typing in your own."
         Top             =   735
         Width           =   4005
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -68805
         Top             =   1785
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   210
         TabIndex        =   8
         ToolTipText     =   "Select the date to review your tasks or edit them."
         Top             =   630
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483648
         Appearance      =   1
         MonthBackColor  =   16761024
         StartOfWeek     =   24444929
         TitleBackColor  =   0
         TitleForeColor  =   16761024
         TrailingForeColor=   4210752
         CurrentDate     =   37310
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   -73110
         Picture         =   "Planner.frx":3072
         Top             =   1785
         Width           =   1110
      End
      Begin VB.Label Label7 
         Caption         =   "Color Schemes:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   33
         Top             =   1575
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000000&
         Caption         =   "Enter Task Here"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -74895
         TabIndex        =   16
         Top             =   945
         Width           =   2745
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000000&
         Caption         =   "Enter Due Date or Time Here"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -74895
         TabIndex        =   14
         Top             =   420
         Width           =   2745
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1575
         TabIndex        =   12
         ToolTipText     =   "Shows the date selected."
         Top             =   3990
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Enter Task Here"
         Height          =   225
         Left            =   3045
         TabIndex        =   11
         Top             =   1050
         Width           =   4005
      End
      Begin VB.Label Label5 
         Caption         =   "Select Time Here or Type Your Own Time"
         Height          =   225
         Left            =   3045
         TabIndex        =   10
         Top             =   525
         Width           =   4005
      End
      Begin VB.Label Label4 
         Caption         =   "Date Selected:"
         Height          =   225
         Left            =   1575
         TabIndex        =   9
         Top             =   3780
         Width           =   1380
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear List"
      End
      Begin VB.Menu asdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove2 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuClear2 
         Caption         =   "Clear List"
      End
      Begin VB.Menu ffff 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave2 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuLoad2 
         Caption         =   "Load"
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
   Begin VB.Menu ToTray 
      Caption         =   "Send To Tray"
   End
   Begin VB.Menu systray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu systrayRestore 
         Caption         =   "Restore MyPlanner2002"
      End
      Begin VB.Menu asdfasdf 
         Caption         =   "-"
      End
      Begin VB.Menu systrayExit 
         Caption         =   "Exit MyPlanner2002"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public directory$
Public directory2$
Public AutoSave1 As Boolean
Public AutoSave2 As Boolean
Public MyColor As String
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Sub About_Click()
    Message = MsgBox("Created by John Stouffer, Stouffer Enterprises.  Freeware, open-source, feel free to edit/update/fix anything you wish.", vbOKOnly, "About")
End Sub

Private Sub Check1_Click()
    AutoSave1 = Check1.Value
End Sub

Private Sub Check2_Click()
    AutoSave2 = Check2.Value
End Sub

Private Sub cmdAdd_Click()
    List1.AddItem (Text1.Text & " -- " & Text2.Text)
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    If AutoSave2 = True Then
        Save1
    End If
End Sub

Private Sub cmdAdd2_Click()
    List2.AddItem (Combo1.Text & " -- " & Text4.Text)
    Combo1.Text = ""
    Text4.Text = ""
    Combo1.SetFocus
    If AutoSave1 = True Then
        Save2
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdExit2_Click()
    Unload Me
    End
End Sub

Private Sub cmdLoad_Click()
    Load1
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
    If AutoSave2 = True Then
        Save1
    End If
End Sub

Private Sub cmdRemove2_Click()
    On Error Resume Next
    List2.RemoveItem List2.ListIndex
    If AutoSave1 = True Then
        Save2
    End If
    If List2.ListCount = 0 Then
        Kill2
    End If
End Sub

Private Sub cmdSave_Click()
    Save1
End Sub

Private Sub cmdSave2_Click()
    Save2
End Sub

Private Sub Combo1_GotFocus()
    cmdAdd2.Default = True
End Sub

Private Sub Exit_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        Message = MsgBox("I'm already running!!!  Check your taskbar for me!", vbOKOnly, "Previous Instance Detected")
        Unload Me
        End
    End If
    directory$ = App.Path & "\MyTasks.tsk"
    directory2$ = App.Path & "\"
    AutoSave1 = True
    AutoSave2 = True
    Load1
    Dim MyTime As SYSTEMTIME
    GetLocalTime MyTime
    MonthView1.Value = MyTime.wDay & "/" & MyTime.wMonth & "/" & MyTime.wYear
    DateClicked = MyTime.wDay & "/" & MyTime.wMonth & "/" & MyTime.wYear
    Label3.Caption = DateClicked
    Load2
    On Error Resume Next
    Open (App.Path & "\MyColor.tsk") For Input As #1
        Input #1, MyColor
    Close #1
    Update
End Sub

Private Sub Help_Click()
    Form2.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Sys As Long
    Sys = X / Screen.TwipsPerPixelX
    Select Case Sys
    Case WM_LBUTTONDOWN:
    Me.PopupMenu systray
    End Select
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
    Me.Hide
    Me.Refresh
    With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = Me.Caption & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    Else
    Shell_NotifyIcon NIM_DELETE, nid
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub

Private Sub Image1_Click()
    Message = MsgBox("Created by John Stouffer, Stouffer Enterprises.  Freeware, open-source, feel free to edit/update/fix anything you wish.", vbOKOnly, "About")
End Sub

Private Sub Image2_Click()
    Message = MsgBox("Created by John Stouffer, Stouffer Enterprises.  Freeware, open-source, feel free to edit/update/fix anything you wish.", vbOKOnly, "About")
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If
    
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopUp2
    End If
End Sub

Private Sub mnuClear_Click()
    List1.Clear
End Sub

Private Sub mnuClear2_Click()
    List2.Clear
End Sub

Private Sub mnuLoad_Click()
    Load1
End Sub

Private Sub mnuLoad2_Click()
    Load2
End Sub

Private Sub mnuRemove_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
    If AutoSave2 = True Then
        Save1
    End If
End Sub

Private Sub mnuRemove2_Click()
    On Error Resume Next
    List2.RemoveItem List2.ListIndex
    If AutoSave1 = True Then
        Save2
    End If
    If List2.ListCount = 0 Then
        Kill2
    End If
End Sub

Private Sub mnuSave_Click()
    Save1
End Sub

Private Sub mnuSave2_Click()
    Save2
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Label3.Caption = DateClicked
    List2.Clear
    Load2
End Sub

Private Sub Option1_Click()
    MyColor = &HFF8080
    Update
End Sub

Private Sub Option2_Click()
    MyColor = &HFFC0C0
    Update
End Sub

Private Sub Option3_Click()
    MyColor = &HC000&
    Update
End Sub

Private Sub Option4_Click()
    MyColor = &HFFFFFF
    Update
End Sub

Private Sub Option5_Click()
    CommonDialog1.ShowColor
    MyColor = CommonDialog1.Color
    Update
End Sub

Private Sub Option6_Click()
    MyColor = &HC0C0FF
    Update
End Sub

Private Sub Option7_Click()
    MyColor = &HFFFFC0
    Update
End Sub

Private Sub systrayExit_Click()
    Unload Me
    End
End Sub

Private Sub systrayRestore_Click()
    WindowState = vbNormal
    Me.Show
End Sub

Private Sub Text1_Click()
    cmdAdd.Default = True
End Sub

Private Sub Text3_Click()
    cmdAdd2.Default = True
End Sub
Private Function Update()
    MonthView1.MonthBackColor = MyColor
    MonthView1.TitleForeColor = MyColor
    Text1.BackColor = MyColor
    Text2.BackColor = MyColor
    Text3.BackColor = MyColor
    Text4.BackColor = MyColor
    List1.BackColor = MyColor
    List2.BackColor = MyColor
    Label3.BackColor = MyColor
    Combo1.BackColor = MyColor
    On Error Resume Next
    Open (App.Path & "\MyColor.tsk") For Output As #1
        Print #1, MyColor
    Close #1
End Function

Public Function Load1()

    Dim CheckThisFile As String
    Dim FileExists As Boolean
    CheckThisFile = Dir(directory$)
    If CheckThisFile <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
    
    If FileExists = False Then
        Exit Function
    End If
    
    List1.Clear
    On Error Resume Next
    Dim MyString$
    Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        List1.AddItem MyString$
    Wend
    Close #1
End Function
Public Function Save1()
    On Error Resume Next
    Open directory$ For Output As #1
        For X = 0 To (List1.ListCount - 1)
            Print #1, List1.List(X)
        Next X
    Close #1
End Function

Public Function Load2()
    Dim CheckThisFile As String
    Dim FileExists As Boolean
    CheckThisFile = Dir(directory2$ & MonthView1.Day & "-" & MonthView1.Month & "-" & MonthView1.Year & ".tsk")
    If CheckThisFile <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
    
    On Error Resume Next
    If FileExists = False Then
        Exit Function
    End If
    
    Dim MyCalendarInput$
    Open (directory2$ & MonthView1.Day & "-" & MonthView1.Month & "-" & MonthView1.Year & ".tsk") For Input As #1
    While Not EOF(1)
        Input #1, MyCalendarInput$
        List2.AddItem MyCalendarInput$
    Wend
    Close #1
End Function

Public Function Save2()
    On Error Resume Next
    Open (directory2$ & MonthView1.Day & "-" & MonthView1.Month & "-" & MonthView1.Year & ".tsk") For Output As #1
        For Y = 0 To (List2.ListCount - 1)
            Print #1, List2.List(Y)
        Next Y
    Close #1
End Function

Public Function Kill2()
    On Error Resume Next
    Close All
    Kill (directory2$ & MonthView1.Day & "-" & MonthView1.Month & "-" & MonthView1.Year & ".tsk")
End Function

Private Sub ToTray_Click()
    WindowState = vbMinimized
End Sub
