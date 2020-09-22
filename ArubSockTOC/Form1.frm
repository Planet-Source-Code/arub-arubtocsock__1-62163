VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "unblock user"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Block User"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "set profile"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "go away"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "send chat"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "send im"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "chat invite"
      Height          =   435
      Left            =   4320
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "password"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "username"
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "join chat"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "login"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin Project1.ArubTOCSock t 
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
   End
   Begin VB.Label Label3 
      Caption         =   "Online Buddies"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Buddies"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long

'**NOTE**
'THIS IS JUST A SMALL EXAMPLE TO SHOW HOW TO GET THE DATA FROM THE BUDDYLIST
'AND SOME OTHER BASIC STUFF
'figure out the rest ;)
Private Sub Command1_Click()
    If Command1.Caption = "login" Then
        t.Login Text1.Text, Text2.Text, "toc.oscar.aol.com", 5190
        Command1.Caption = "???"
    Else
        t.LogOff
        Command1.Caption = "login"
    End If
End Sub

Private Sub Command2_Click()
    t.JoinChat "arub", 4
End Sub

Private Sub Command3_Click()
    t.ChatInvite Text3.Text, "hi kthnx  bye", "SMARTERCHILD"
End Sub

Private Sub Command4_Click()
    Dim asdf As String
    asdf = InputBox("sn")
    
    t.SendIm asdf, "I'm from the streets biatch >:o", False
End Sub

Private Sub Command5_Click()
        t.SendChat Text3, "HI KTHNX BYE"
End Sub

Private Sub Command6_Click()
    t.SetAway "<FONT FACE = TAHOMA SIZE = 1 COLOR = BLUE>http://www.arubs.net</FONT>"
End Sub

Private Sub Command7_Click()
    t.SetProfile "asdf"
End Sub

Private Sub Command8_Click()
    Dim strFaggot As String
    strFaggot = InputBox("who=")
    t.BlockUser strFaggot
End Sub

Private Sub Command9_Click()
    Dim unblockwho As String
    unblockwho = InputBox("username=", "", "")
    t.UnBlockUser unblockwho
End Sub

Private Sub t_BuddyOffline(strUserName As String)
    For i = 0 To List2.ListCount - 1
        If Minimal(List2.List(i)) = Minimal(strUserName) Then _
        List2.RemoveItem i
    Next i
    List1.AddItem strUserName
End Sub

Private Sub t_BuddyOnline(strUserName As String)
    'will add oncoming buddies to the buddylist
    For i = 0 To List1.ListCount - 1 'remove the buddy from offline list
        If Minimal(List1.List(i)) = Minimal(strUserName) Then List1.RemoveItem i
    Next i
    
    List2.AddItem strUserName
        'online buddies when you sign on are sent like this
End Sub

Private Sub t_Error(intErrorCode As Integer)
    If intErrorCode = 901 Then Exit Sub
    MsgBox intErrorCode
    
End Sub

Private Sub t_IncomingMessage(strUserName As String, strMessage As String)
    MsgBox strMessage
End Sub

Private Sub t_JoinedChat(strRoomName As String, lngRoomID As Long)
    Text3.Text = lngRoomID
End Sub

Private Sub t_LoggedIn(strUserName As String, strBuddyList As String, strBlocked As String)
    Command1.Caption = "LogOut"
    Dim X As Variant
    X = Split(strBuddyList, vbCrLf)
        Dim i As Integer
        For i = 0 To UBound(X)
            List1.AddItem X(i)
        Next i
        
        'to only add buddies and not groups
        'for i = 0 to ubound(X)
          '  if not mid(x(i),1,5) = "GROUP" then list1.additem x(i)
         'next i
End Sub

