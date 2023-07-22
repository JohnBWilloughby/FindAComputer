VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Find A Computer"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "workstation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8281
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Search For :"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mainmenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuProxy 
         Caption         =   "Proxy"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
  ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()

ADFilter = "computer"
ListView1.ListItems.Clear

With ListView1
    .ColumnHeaders.Add , , "Workstation Name", TextWidth("What is longest name of ")
    .ColumnHeaders.Add , , "IP Address", TextWidth("xxx.xxx.xxx.xxx.xxx")
    .ColumnHeaders.Add , , "Username", TextWidth("This one could even ")
    .ColumnHeaders.Add , , "Location", TextWidth("Location NYO")
    .View = lvwReport
    .FullRowSelect = True
    .GridLines = True
    .AllowColumnReorder = False
End With




End Sub


Private Sub Command1_Click()

ListView1.ListItems.Clear


If Form1.txtSearch <> "" Then
    strSearch = LCase(Form1.txtSearch.Text)
    Call CollectPC
Else
    Form1.txtSearch.Text = ""
    Form1.Refresh
End If

'Set itmx = Form1.ListView1.ListItems.Add(, , "Test1")
'    itmx.SubItems(1) = "this is a test"
'    itmx.SubItems(2) = "This is a second"
'    itmx.SubItems(3) = "NY"
'Form1.Refresh

End Sub

Private Sub Command2_Click()
    Unload Form1
    End
    
End Sub

Private Sub ListView1_MouseUP(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim li As ListItem


If Button = vbRightButton Then
    If Not (ListView1.SelectedItem Is Nothing) Then
      If Not (ListView1.HitTest(x, y) Is Nothing) Then
        Me.PopupMenu mainmenu
        Debug.Print "right-click over item, popup now"
      Else
        Debug.Print "not over item"
      End If
    End If
  End If

End Sub

Private Sub mnuProxy_Click()
    ProxyPC = ListView1.SelectedItem.Text
    ' Debug.Print ListView1.SelectedItem.Text
    ' MsgBox ListView1.SelectedItem.Text
    ' for Windows 10 and Proxy 10
    ShellResp = Shell("C:\Program Files (x86)\Proxy Networks\Master\Proxy.exe" & " /S" & Chr(34) & ProxyPC & Chr(34))
    ' for Windows 7 and Proxy 8
    ' ShellResp = Shell("C:\Program Files (x86)\Proxy Networks\PROXY Pro Master\Proxy.exe" & " /S" & Chr(34) & ProxyPC & Chr(34))
    
    'C:\Program Files (x86)\Proxy Networks\PROXY Pro Master\Proxy.exe /S <nameofPC>
End Sub
