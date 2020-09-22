VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ArubTOCSock 
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ScaleHeight     =   1275
   ScaleWidth      =   3285
   Begin VB.Timer tmrAntiIdle 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2400
      Top             =   840
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   2760
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   0
      Picture         =   "ArubTOCSock.ctx":0000
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "ArubTOCSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'most of this stuff can be found in the toc protocol documentation
'www.arubs.net/TOC_Protocol.txt
'----------------------
'www.arubs.net
Option Explicit
Public Event LoggedIn(strUserName As String, strBuddyList As String, strBlocked As String)
Public Event LoggedOff(strUserName As String)
Public Event Error(intErrorCode As Integer)
Public Event IncomingBuddyCapabilities(strUserName As String, strCapUUID As String)
Public Event IncomingPacket(intIndex As Integer, strData As String)
Public Event IncomingMessage(strUserName As String, strMessage As String)
Public Event BuddyOnline(strUserName As String)
Public Event BuddyOffline(strUserName As String)
Public Event Warned(strWarner As String, intNewWarningLevel As Integer)
Public Event JoinedChat(strRoomName As String, lngRoomID As Long)
Public Event IncomingChatMessage(lngRoomID As Long, strUserName As String, blnWhisper As Boolean, strMessage As String)
Public Event ChatUserUpdate(lngRoomID As Long, blnInside As Boolean, strUser As String, StrUser2 As String)
Public Event IncomingChatInvite(strRoomName As String, lngRoomID As Long, strUserName As String, strMessage As String)
Public Event LeftChat(lngRoomID As Long)
Public Event PassWordUpdate(blnSuccess As Boolean)
Public Event ScreenNameFormatUpdate(blnSuccess As Boolean)
Public Event IncomingProfile(strProfile As String)
Public Event RendevousRequest(strUser As String, strUUID As String, strCookie As String, strRendevousIP As String)
Public Event IncomingBartData(strUserName As String, strBartData As String)

Dim A1, A2, A3, A4, A5, A6, A7, A8, A9, A10
Dim UserName As String, Password As String, LocalSeq As Long, RemoteID As Long, lngChatIDS As Long
Dim Server As String, Port As Integer, strServerNick As String, strMsg As String
Private Type StrChats
    lngRoomID As Long
    strRoomName As String
End Type
Dim ChatRooms(1000) As StrChats


Public Function Login(strUserName As String, strPassword As String, strServer As String, intPort As Integer)
    UserName = Minimal(strUserName)
    Password = EncryptPW(strPassword)
    Server = strServer
    Port = intPort
    Sock.Close
    Sock.Connect strServer, intPort
End Function
Public Function LogOff()
    Sock.Close
    Server = vbNullString: Password = vbNullString: Server = vbNullString: Port = 0: LocalSeq = 0
        Dim i As Long
        For i = 0 To 1000
            ChatRooms(i).lngRoomID = 0
            ChatRooms(i).strRoomName = ""
        Next i
    lngChatIDS = 0: RemoteID = 0
    RaiseEvent LoggedOff(UserName)
End Function
Private Function ParseData(intIndex As Integer, strData As String)
    On Error Resume Next
    RaiseEvent IncomingPacket(intIndex, strData)
    Select Case intIndex 'Data Frame
        Case 1 'SIGNON
            Select Case strData
                Case Chr(0) & Chr(0) & Chr(0) & Chr(1)
                        SendPacket 1, Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Word(Len(UserName)) & UserName
                        SendPacket 2, "toc2_login login.oscar.aol.com 29999 " & UserName & " " & Password & " English  " & Qt("ArubSockTOC2.0") & " 160 US " & Qt("") & " " & Qt("") & " 3 0 30303 -kentucky -utf8 74651200" & Chr(0)
    
            End Select
        Case 2 'DATA
            A1 = Split(strData, ":")

            Select Case LCase(A1(0))
                
                Case "sign_on"
                       'SendPacket2 "toc_init_done"
                Case "config2" 'Buddylist
                        
                        SendPacket2 "toc_add_buddy " & "arub PIMP" & Int(7 * Rnd)
                        'gotta send toc_add_buddy so they won't put you on some weird privacy setting
                
                        Dim strBuddies As String, strBlocked As String
                        Combo1.Clear
                        Dim tocsplit
                        tocsplit = Split(strData, Chr(10))
                        
                         Dim i As Integer
                         For i = 0 To UBound(tocsplit) 'coulda just parsed the tocsplit, but you might need this later if you're setting the config again
                             Combo1.AddItem tocsplit(i)
                             DoEvents
                         Next i
                         
X:
                             For i = 0 To Combo1.ListCount - 1
                                 Dim strTemp
                                 strTemp = Split(Combo1.List(i), ":", 2)
                                     Select Case CStr(strTemp(0)) 'get the first letter
                                         Case "d" 'blocked
                                             
                                             strBlocked = strBlocked & CStr(strTemp(1)) & vbCrLf
                                             Combo1.RemoveItem i
                                             GoTo X
                                             
                                         Case "b" 'buddy
                                             
                                             strBuddies = strBuddies & CStr(strTemp(1)) & vbCrLf
                                             Combo1.RemoveItem i
                                             GoTo X
                                         
                                         Case "g" 'group
                                             strBuddies = strBuddies & "GROUP [" & CStr(strTemp(1)) & "]" & vbCrLf
                                             Combo1.RemoveItem i
                                             GoTo X
                                             
                                       End Select
                                       
                             Next i
                                            
                             SendPacket2 "toc_init_done"
                             RaiseEvent LoggedIn(strServerNick, strBuddies, strBlocked)

                    Case "nick" 'formatted nick back from server
                        strServerNick = CStr(A1(1))
                    Case "im_in_enc2" 'incoming im
                        
                        If UBound(A1) = 9 Then
                            strMsg = CStr(A1(9))
                        Else
                            strMsg = ""
                            For i = 9 To UBound(A1)
                                strMsg = strMsg & ":" & A1(i)
                            Next i
                        End If
                        
                        RaiseEvent IncomingMessage(CStr(A1(1)), strMsg)
                    
                    Case "update_buddy2"
                    
                        'update_buddy:<Buddy User>:<Online? T/F>:<Evil Amount>:<Signon Time>:<IdleTime>:<UC>
                        ' just got the buddy's sn
                        If UCase(A1(2)) = "T" Then
                            RaiseEvent BuddyOnline(CStr(A1(1)))
                        Else
                            RaiseEvent BuddyOffline(CStr(A1(1)))
                        End If
                        
                    Case "buddy_caps2" 'buddy capabilities
                        RaiseEvent IncomingBuddyCapabilities(CStr(A1(1)), CStr(A1(2)))
                    Case "bart2"
                        RaiseEvent IncomingBartData(CStr(A1(1)), CStr(A1(2)))
                    Case "error" 'error
                        RaiseEvent Error(CInt(A1(1)))
                    Case "eviled" 'warned
                        RaiseEvent Warned(CStr(A1(2)), CInt(A1(1)))
                    Case "chat_join" 'joined chat
                    
                        RemoteID = CLng(A1(1))
                        lngChatIDS = lngChatIDS + 1
                        If lngChatIDS >= 1000 Then lngChatIDS = 0
                        ChatRooms(lngChatIDS).lngRoomID = CLng(A1(1))
                        ChatRooms(lngChatIDS).strRoomName = CStr(A1(2))
                        RaiseEvent JoinedChat(CStr(A1(2)), RemoteID)
                    
                    Case "chat_in_enc" 'message in chat
                        
                        If UBound(A1) = 6 Then
                            strMsg = A1(6)
                        Else
                            strMsg = ""
                            For i = 6 To UBound(A1)
                                strMsg = strMsg & ":" & A1(i)
                            Next i
                        End If
                        RaiseEvent IncomingChatMessage(CLng(A1(1)), CStr(A1(2)), CBool(A1(3) = "T"), strMsg)
                    
                    Case "chat_update_buddy" 'someone left or joined chat
                        RaiseEvent ChatUserUpdate(CLng(A1(1)), CBool(A1(2) = "T"), CStr(A1(3)), CStr(A1(4)))
                    Case "chat_invite" 'invited to chat
                        RaiseEvent IncomingChatInvite(CStr(A1(1)), CStr(A1(2)), CStr(A1(3)), CStr(A1(4)))
                    Case "chat_left" 'left chat
                        RaiseEvent LeftChat(CStr(A1(1)))
                    Case "goto_url" 'profile
                        RaiseEvent IncomingProfile(CStr(A1(2)))
                    Case "admin_nick_status" 'if the return code is 0, it's a success, if not; it failed.
                        RaiseEvent ScreenNameFormatUpdate((CBool(CInt(A1(1)) = 0)))
                    Case "admin_passwd_status" 'same as format changing
                        RaiseEvent PassWordUpdate(CBool(CInt(A1(1)) = 0))
                    Case "pause" 'pause, not really parsed
                        DoEvents
                    Case "rvous_propose" 'rendevous request
                        RaiseEvent RendevousRequest(CStr(A1(1)), CStr(A1(2)), CStr(A1(3)), CStr(A1(5)))
                End Select
            Case 3, 4, 5
                '3,4 - not used in toc
                '5 - not used here (will use for anti-idle though)
                
            Case Else
                    'RUH ROH =X
                        
        End Select
                            
End Function

Public Function SendPacket(intIndex As Integer, strData As String)
    If Not Sock.State = sckConnected Then Exit Function
    LocalSeq = LocalSeq + 1
    If LocalSeq >= 65535 Then LocalSeq = 0
    A2 = Chr(intIndex)
    A3 = Word(LocalSeq) & Word(Len(strData)) & strData
    Sock.SendData "*" & A2 & A3
End Function
Public Function SendPacket2(strDatas As String)
    SendPacket 2, strDatas & Chr(0)
End Function

Private Sub Sock_Close()
   ' MsgBox "SOCKET CLOSED"
    RaiseEvent LoggedOff(UserName)
End Sub

Private Sub Sock_Connect()
    Sock.SendData "FLAPON" & vbCrLf & vbCrLf
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
'By Xeon
    Dim strData As String
    Dim lngLength As Long
Split:
    Sock.PeekData strData, vbString
    lngLength = GetWord(Mid(strData, 5, 2))
    If bytesTotal >= lngLength + 6 Then
        Sock.GetData strData, vbString, lngLength + 6
        Call ParseData(Asc(Mid(strData, 2, 1)), Mid(strData, 7, Len(strData) - 6))
        bytesTotal = bytesTotal - (lngLength + 6)
        If bytesTotal > 0 Then GoTo Split
    End If
End Sub
Public Function SendIm(strUserName As String, strMessage As String, blnAutoResponse As Boolean)
    Dim strAuto As String
        IIf blnAutoResponse, strAuto = "auto", strAuto = ""
    If strAuto = "" Then
        SendPacket2 "toc_send_im " & Minimal(Qt(strUserName)) & " " & Qt(Normalize(strMessage))
    Else
        SendPacket2 "toc_send_im " & Minimal(Qt(strUserName)) & " " & Qt(Normalize(strMessage)) & " " & strAuto
    End If
End Function
Public Function AddBuddy(strBuddy As String, blnPermanent As Boolean)
    If blnPermanent = False Then 'buddy will be gone after client goes offline
        SendPacket2 "toc_add_buddy " & Qt(strBuddy)
    Else
        Dim strConfigString As String
        Combo1.AddItem "b:" & strBuddy
            Dim i As Long
            For i = 0 To Combo1.ListCount - 1
                strConfigString = strConfigString & Combo1.List(i)
            Next i
        SendPacket2 "toc_set_config " & strConfigString
    End If
End Function
Public Function RemoveBuddy(strBuddy As String)
    SendPacket2 "toc_remove_buddy " & Qt(strBuddy)
End Function
Public Function WarnUser(strUserName As String, blnAnonymous As Boolean)
    Dim strBlockstring As String
        If blnAnonymous = True Then
            strBlockstring = "anon"
        Else
            strBlockstring = "norm"
        End If
    SendPacket2 "toc_evil " & Qt(strUserName) & " " & Qt(strBlockstring)
End Function
Public Function JoinChat(strRoomName As String, intExchange As Integer)
    SendPacket2 "toc_chat_join " & Qt(intExchange) & " " & Qt(strRoomName)
End Function
Public Function SendChat(lngID As Long, strMessage As String)
    SendPacket2 "toc_chat_send " & Qt(lngID) & " " & Qt(Normalize(strMessage))
End Function
Public Function ChatWhisper(lngID As Long, strUserName As String, strMessage As String)
    SendPacket2 "toc_chat_whisper " & Qt(lngID) & " " & Qt(Minimal(strUserName)) & " " & Qt(Normalize(strMessage))
End Function
Public Function ChatInvite(lngID As Long, strInviteMessage As String, strUserName As String)
    SendPacket2 "toc_chat_invite " & Qt(lngID) & " " & Qt(Normalize(strInviteMessage)) & " " & Qt(strUserName)
End Function
Public Function ChatLeave(lngID As Long)
    SendPacket2 "toc_chat_leave " & Qt(lngID)
End Function
Public Function AcceptChatInvite(lngID As Long)
    SendPacket2 "toc_chat_accept " & Qt(lngID)
End Function
Public Function GetProfile(strUserName As String)
    SendPacket2 "toc_get_info " & Qt(Minimal(strUserName))
End Function
Public Function GetStatus(strUserName As String)
    SendPacket2 "toc_get_status " & Qt(Minimal(strUserName))
End Function
Public Function SetProfile(strProfile As String)
    SendPacket2 "toc_set_info " & Qt(strProfile)
End Function
Public Function SetAway(strMessage As String)
    SendPacket2 "toc_set_away " & Qt(Normalize(strMessage))
End Function
Public Function SetIdle(lngMinutes As Long)
    SendPacket2 "toc_set_idle " & Qt(Int(lngMinutes * 60))
End Function
Public Function SetCapabilities(strCapUUID As String)
    SendPacket2 "toc_set_caps " & Qt(strCapUUID)
    'seperated with a "," for each capability
    'UUIDS AS PROVIDED IN DOCUMENTATION:
    'TALK             - 09461341-4C7F-11D1-8222-444553540000
    'SEND FILE        - 09461343-4C7F-11D1-8222-444553540000
    'IM IMAGE         - 09461345-4C7F-11D1-8222-444553540000
    'BUDDYICON        - 09461346-4C7F-11D1-8222-444553540000
    'ADD INS          - 09461347-4C7F-11D1-8222-444553540000
    'GET FILE         - 09461348-4C7F-11D1-8222-444553540000
    'AIM EXPRESSIONS  - 0946134A-4C7F-11D1-8222-444553540000
    'SEND BUDDYLIST   - 0946134B-4C7F-11D1-8222-444553540000
End Function
Public Function FormatUserName(strNewFormattedUserName As String)
    SendPacket2 "toc_format_nick " & Qt(strNewFormattedUserName)
End Function
Public Function ChangePassword(strOldPassword As String, strNewPassword As String)
    SendPacket2 "toc_change_passwd " & Qt(strOldPassword) & " " & Qt(strNewPassword)
End Function
Public Function GetChatID(strRoomName As String) As Long
'get the room ID for a specific chat
        Dim i As Integer
        For i = 0 To 1000
            If Minimal(ChatRooms(i).strRoomName) = Minimal(strRoomName) Then
                GetChatID = ChatRooms(i).lngRoomID
                Exit For
            End If
        Next i
End Function
Public Function LastChatID() As Long
    LastChatID = RemoteID
End Function

Private Sub tmrAntiIdle_Timer()
    SendPacket 5, vbNullString
End Sub

Private Sub UserControl_Initialize()
    UserControl.Height = imgLogo.Height: UserControl.Width = imgLogo.Width
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = imgLogo.Height: UserControl.Width = imgLogo.Width
End Sub
Public Function AntiIdle(blnEnabled As Boolean)
    tmrAntiIdle.Enabled = blnEnabled
End Function
Public Function BlockUser(strUserName As String)
    SendPacket2 "toc2_add_deny " & Qt(strUserName)
End Function
Public Function UnBlockUser(strUserName As String)
    SendPacket2 "toc2_remove_deny " & Qt(strUserName)
End Function
