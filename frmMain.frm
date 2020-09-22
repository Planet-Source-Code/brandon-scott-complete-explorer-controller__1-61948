VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Explorer Controller"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgUsers 
      Left            =   360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstExplorerWindows 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgUsers"
      SmallIcons      =   "imgUsers"
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "-"
         Object.Width           =   601
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Location"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Menu mnu 
      Caption         =   "Global Control"
      Begin VB.Menu mnuRefreshList 
         Caption         =   "Refresh List..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllGoBack 
         Caption         =   "All Back..."
      End
      Begin VB.Menu mnuAllGoForward 
         Caption         =   "All Forward..."
      End
      Begin VB.Menu mnuAllGoHome 
         Caption         =   "All Home..."
      End
      Begin VB.Menu mnuAllGoSearch 
         Caption         =   "All Search..."
      End
      Begin VB.Menu mnuAllRefresh 
         Caption         =   "All Refresh..."
      End
      Begin VB.Menu mnuAllStop 
         Caption         =   "All Stop..."
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitAllExplorers 
         Caption         =   "Quit All Explorer Windows..."
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Explorer Controller..."
      End
   End
   Begin VB.Menu mnuRemote 
      Caption         =   "Remote Control"
      Begin VB.Menu mnuRemoteGotoURL 
         Caption         =   "Goto URL"
      End
      Begin VB.Menu mnuSetBodyHTML 
         Caption         =   "Set Body HTML..."
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavigation 
         Caption         =   "Navigation"
         Begin VB.Menu mnuRemoteBack 
            Caption         =   "Back..."
         End
         Begin VB.Menu mnuRemoteForward 
            Caption         =   "Forward..."
         End
         Begin VB.Menu mnuRemoteHome 
            Caption         =   "Home..."
         End
         Begin VB.Menu mnuRemoteSearch 
            Caption         =   "Search..."
         End
         Begin VB.Menu mnuRemoteRefresh 
            Caption         =   "Refresh"
         End
         Begin VB.Menu mnuRemoteStop 
            Caption         =   "Stop..."
         End
      End
      Begin VB.Menu mnuEventSink 
         Caption         =   "Event Sink..."
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoteClose 
         Caption         =   "Close..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IE As InternetExplorer
Dim WIN As New ShellWindows

Public Sub AddExplorerWindow(ByRef IEWindow As InternetExplorer)
    Dim AddList As ListItem
    Set AddList = lstExplorerWindows.ListItems.Add(, "h" & IEWindow.hWnd, , 1, 1)
    AddList.SubItems(1) = IEWindow.LocationName
    AddList.SubItems(2) = IEWindow.LocationURL
End Sub

Public Function GetExplorerWindowList()
    lstExplorerWindows.ListItems.Clear
    For Each IE In WIN
        AddExplorerWindow IE
    Next
End Function

Public Function ExplorerActionToAll(Action As Integer)
    On Error Resume Next
    For Each IE In WIN
        If IE.readyState = READYSTATE_COMPLETE Then
            Select Case Action
                Case 0: IE.GoBack
                Case 1: IE.GoForward
                Case 2: IE.GoHome
                Case 3: IE.GoSearch
                Case 4: IE.Refresh
                Case 5: IE.stop
                Case 6: IE.Quit
            End Select
        End If
    Next
End Function

Private Sub Form_Load()
    Call GetExplorerWindowList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstExplorerWindows.Width = Me.Width - 120
    lstExplorerWindows.Height = Me.Height - 720
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAllGoBack_Click()
    Call ExplorerActionToAll(0)
End Sub

Private Sub mnuAllGoForward_Click()
    Call ExplorerActionToAll(1)
End Sub

Private Sub mnuAllGoHome_Click()
    Call ExplorerActionToAll(2)
End Sub

Private Sub mnuAllGoSearch_Click()
    Call ExplorerActionToAll(3)
End Sub

Private Sub mnuAllRefresh_Click()
    Call ExplorerActionToAll(4)
End Sub

Private Sub mnuAllStop_Click()
    Call ExplorerActionToAll(5)
End Sub

Private Sub mnuEventSink_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            Dim NewSink As New frmEventSink
            NewSink.BrowserHandle = "h" & IE.hWnd
            NewSink.Show
        End If
    Next
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuQuitAllExplorers_Click()
    Call ExplorerActionToAll(6)
End Sub

Private Sub mnuRefreshList_Click()
    Call GetExplorerWindowList
End Sub

Private Sub mnuRemoteBack_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.GoBack
        End If
    Next
End Sub

Private Sub mnuRemoteClose_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.Quit
        End If
    Next
End Sub

Private Sub mnuRemoteForward_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.GoForward
        End If
    Next
End Sub

Private Sub mnuRemoteGotoURL_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            Dim tmp As String
            tmp = InputBox("Enter url to navigate to?", "Navigate")
            IE.navigate tmp
        End If
    Next
End Sub

Private Sub mnuRemoteHome_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.GoHome
        End If
    Next
End Sub

Private Sub mnuRemoteRefresh_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.Refresh
        End If
    Next
End Sub

Private Sub mnuRemoteSearch_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.GoSearch
        End If
    Next
End Sub

Private Sub mnuRemoteStop_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            IE.stop
        End If
    Next
End Sub

Private Sub mnuSetBodyHTML_Click()
    For Each IE In WIN
        If lstExplorerWindows.SelectedItem.Key = "h" & IE.hWnd Then
            frmSetBodyHTML.BrowserHandle = "h" & IE.hWnd
            frmSetBodyHTML.Show
        End If
    Next
End Sub
