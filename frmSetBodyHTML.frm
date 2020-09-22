VERSION 5.00
Begin VB.Form frmSetBodyHTML 
   Caption         =   "Set Body HTML"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSet 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmSetBodyHTML.frx":0000
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdSetBodyHTML 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetBodyHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IE As InternetExplorer
Dim WIN As New ShellWindows
Public BrowserHandle As String

Private Sub cmdSetBodyHTML_Click()
    On Error Resume Next
    For Each IE In WIN
        If BrowserHandle = "h" & IE.hWnd Then
            'IE.navigate "about:blank"
            IE.document.Close
            IE.document.open
            IE.document.write txtSet.Text
            'IE.document.body.innerHTML = ""
            'IE.document.body.innerHTML = txtSet.Text
        End If
    Next
End Sub
