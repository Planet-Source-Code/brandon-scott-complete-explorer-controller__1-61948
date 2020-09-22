VERSION 5.00
Begin VB.Form frmEventSink 
   Caption         =   "Event Sink - {}"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6375
   Icon            =   "frmEventSink.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IE As InternetExplorer
Dim WIN As New ShellWindows
Private WithEvents EvtSink As InternetExplorer
Attribute EvtSink.VB_VarHelpID = -1
Public BrowserHandle As String

Private Sub EvtSink_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub EvtSink_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    'List1.AddItem "DocumentComplete: " & URL
End Sub

Private Sub EvtSink_DownloadBegin()

End Sub

Private Sub EvtSink_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

End Sub

Private Sub EvtSink_StatusTextChange(ByVal Text As String)
    If Text = "" Then Exit Sub
    List1.AddItem "StatusTextChange: " & Text
    List1.Selected(List1.ListCount - 1) = True
End Sub

Private Sub Form_Load()
    For Each IE In WIN
        If BrowserHandle = "h" & IE.hWnd Then
            Me.Caption = "Event Sink - {" & IE.LocationName & "}"
            Set EvtSink = IE
        End If
    Next
End Sub
