VERSION 5.00
Begin VB.Form frmRevEx 
   Caption         =   "RevEX"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRevEx 
      Caption         =   "RevEx"
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmRevEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Display
End Sub

Private Function Display()
    txtResult = ""
    
    txtResult = txtResult & "Set (lower case only): " & _
        modRevEx.RevEx("[abcdefghijklmnopqrstuvwxyz]")
    txtResult = txtResult & vbCrLf & "Negative Set (upper case only): " & _
        modRevEx.RevEx("[!0123456789abcdefghijklmnopqrstuvwxyz]")
    txtResult = txtResult & vbCrLf & "Range (lower case only): " & _
        modRevEx.RevEx("[a-z]")
    txtResult = txtResult & vbCrLf & "Negative Range (upper case only): " & _
        modRevEx.RevEx("[!0-z]")
    txtResult = txtResult & vbCrLf & "Number (Decimal): " & _
        modRevEx.RevEx("[#]")
    txtResult = txtResult & vbCrLf & "Boolean: " & _
        modRevEx.RevEx("[01]")
    txtResult = txtResult & vbCrLf & "Hex (lower case): " & _
        modRevEx.RevEx("[0-f]")
    txtResult = txtResult & vbCrLf & "Hex (upper case): " & _
        modRevEx.RevEx("[0123456789ABCDEF]")
    txtResult = txtResult & vbCrLf & "Octal: " & _
        modRevEx.RevEx("[0-8]")
    txtResult = txtResult & vbCrLf & "Vowels: " & _
        modRevEx.RevEx("[aeiouAEIOU]")
    txtResult = txtResult & vbCrLf & "Date: " & _
        modRevEx.RevEx("[0-1][0-2][/][0-2][0-9][/][2][0][0][0-2]") ' <<< days only 00 - 29! Month may return "00"! Not always valid results, needs postprocessing!
    txtResult = txtResult & vbCrLf & "Time12: " & _
        modRevEx.RevEx("[0-1][0-2][:][0-5][0-9][:][0-5][0-9][ ][ap][.][m]")
    txtResult = txtResult & vbCrLf & "Time24: " & _
        modRevEx.RevEx("[0-2][0-3][:][0-5][0-9][:][0-5][0-9]")
    txtResult = txtResult & vbCrLf & "Slotmaschine: " & _
        modRevEx.RevEx("[123B][123A][123R]")
    txtResult = txtResult & vbCrLf & "Intl. Phone Number (US): " & _
        modRevEx.RevEx("[+][1][-][#][#][#][-][5][5][5][-][#][#][#][#]")
    txtResult = txtResult & vbCrLf & "ZipCode: " & _
        modRevEx.RevEx("[1-9][1-9][1-9][1-9][1-9]")
    txtResult = txtResult & vbCrLf & "GUID: " & _
        modRevEx.RevEx("[0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][_][{][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][-][0-f][0-f][0-f][0-f][0-f][-][0-f][0-f][0-f][-][0-f][0-f][0-f][0-f][-][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][0-f][}]")
    txtResult = txtResult & vbCrLf & "Network Card: " & _
        modRevEx.RevEx("[0-f][0-f][.][0-f][0-f][.][0-f][0-f][.][0-f][0-f][.][0-f][0-f][.][0-f][0-f]")
    txtResult = txtResult & vbCrLf & "IP: " & _
        modRevEx.RevEx("[0-2][0-5][0-5][.][0-2][0-5][0-5][.][0-2][0-5][0-5][.][0-2][0-5][0-5]") ' <<< needs clean up to remove trailing 0s ...
    txtResult = txtResult & vbCrLf & "Generic Serial: " & _
        modRevEx.RevEx("[A][B][C][-][0-Z][0-Z][0-Z][0-Z][0-Z][-][0-Z][0-Z][0-Z][0-Z][0-Z]")
    txtResult = txtResult & vbCrLf & "Social Security: " & _
        modRevEx.RevEx("[#][#][#][-][#][#][-][#][#][#][#]") ' <<< not valid, unless by chance ;)
    txtResult = txtResult & vbCrLf & "Credit Card: " & _
        modRevEx.RevEx("[#][#][#][#][#][#][#][#][#][#][#][#][#][#][#][#][#][#][#]") ' <<< not valid, unless by chance ;) Not even sure, if the format is correct...
End Function

Private Sub cmdRevEx_Click()
    Display
End Sub

