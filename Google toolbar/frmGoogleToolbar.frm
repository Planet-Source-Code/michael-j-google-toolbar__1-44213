VERSION 5.00
Begin VB.Form frmToolbar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   300
   ClientLeft      =   465
   ClientTop       =   660
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   285
      Left            =   7370
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Google With:"
      Default         =   -1  'True
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtSEarch 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
 '   SET ALWAYS ON TOP(TRUE)
 MakeAlwaysOnTop Me, True
 Me.Top = 0
 Me.Left = Screen.Width / 3
 Me.Width = Screen.Width / 2 - 50
End Sub

Private Sub cmdSearch_Click()
    Dim textSearch
    textSearch = txtSEarch.Text
    '*** Space + Support ***
    If InStr(textSearch, " ") Then
        textSearch = Replace(textSearch, " ", "+")
    End If
    '***
    'Format is "http://www.google.com/search?hl=en&lr=&ie=UTF-8&oe=UTF-8&q=" & textSearch & "&btnG=Google+Search"
    textSearch = "http://www.google.com/search?hl=en&lr=&ie=UTF-8&oe=UTF-8&q=" & textSearch & "&btnG=Google+Search"
    Dim frma As New frmBrowser
    frma.Show
    frma.cboAddress.Text = textSearch
    If mbDontNavigateNow Then Exit Sub
    frma.timTimer.Enabled = True
    frma.brwWebBrowser.Navigate frma.cboAddress.Text
End Sub
