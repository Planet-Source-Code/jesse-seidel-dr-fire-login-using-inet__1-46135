VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Example - By SpitFire"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2640
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "¤"
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Status: Not logged-in"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By SpitFire
Private Sub Command1_Click()
'Checks Text3 (users.txt) to see if the user that was
'supplied is there, else it dosnt give them the message box
If InStr(1, Text3.Text, Text1.Text) And Text2.Text = "pppassw0rd" Then
'Note: You can put any action here
MsgBox ("You are now Logged-in")
Label3.Caption = "Status: Logged-in"
Else
'Note: You can put any action here
MsgBox ("Login invalid, please try again")
Label3.Caption = "Status: Not logged-in"
End If
End Sub

Private Sub Command2_Click()
'Cancel
End
End Sub

Private Sub Form_Load()
On Error GoTo err:
'You put the URL location of the users.txt so that Inet
'can download it and people can start logging in
Text3.Text = Inet1.OpenURL("http://members.lycos.co.uk/spitfiremsn/login/users.txt")
Exit Sub
err:
MsgBox ("There was a problem loading the login example" & vbCrLf & vbCrLf & "Probably because the Inet URL dosnt exist, or there was some other Inet related problem")
End Sub
