Attribute VB_Name = "mXaudioPlayer"
Option Explicit


'
'Public Function WindowProc(ByVal lHwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    'this will take care of sending appropriate command to player or execute
'    '   default Windows procedure
'    Dim ret As Long
'
'    'Debug.Print hw, iMsg, wParam, lParam, lPlayerHandle
'
'    If GetWindowLong(hWnd, GWL_USERDATA) = lPlayerHandle Then
'        'Go ahead - execute command
'        If iMsg = eCurrentCommand Then
'            'clear tcurrentmessage structure
'            tCurrentMessage.mAck = 0
'            tCurrentMessage.mNack.mCode = 0
'            tCurrentMessage.mNack.mCommand = 0
'            'execute command here
'            Select Case iMsg
'            Case XA_MSG_COMMAND_INPUT_OPEN 'OK
'                'we can send an address of the string - but i am just too lazy
'                ret = control_message_send_S(lPlayerHandle, iMsg, sCurrentShortFile)
'            Case XA_MSG_COMMAND_INPUT_CLOSE 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_OUTPUT_OPEN 'OK
'                'we can send an address of the string - but i am just too lazy
'                ret = control_message_send_S(lPlayerHandle, iMsg, sCurrentOutputShortFile)
'            Case XA_MSG_COMMAND_OUTPUT_CLOSE 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_OUTPUT_RESET
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_INPUT_SEND_MESSAGE
'                'do not use it - EVER!!!!!
'            Case XA_MSG_COMMAND_PLAY ' start playing - if applicable. OK
'                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PLAYING Then Exit Function
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_STOP ' stop playing. OK
'                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_EOS Or tCurrentMessage.mPlayerState = XA_PLAYER_STATE_STOPPED Then Exit Function
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_EXIT ' close input and output and exit.
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_PAUSE 'OK
'                'execute only if playing
'                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PLAYING Then
'                    ret = control_message_send_N(lPlayerHandle, iMsg)
'                Else
'                    Exit Function
'                End If
'            Case XA_MSG_COMMAND_SEEK
'                'wParam is an offset, lParam is a range
'                ret = control_message_send_II(lPlayerHandle, iMsg, wParam, lParam)
'            Case XA_MSG_COMMAND_SYNC 'DON'T KNOW HOW TO USE
'                'do not use
'            Case XA_MSG_SET_INPUT_POSITION_RANGE 'OK
'                'by default range=400
'                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
'            Case XA_MSG_GET_INPUT_POSITION_RANGE 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_OUTPUT_MUTE 'OK
'                'it does not really mute the output
'                '   it just sets the volume of wav to 0
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_COMMAND_OUTPUT_UNMUTE 'OK
'                'recall value of volume, when it was muted
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_SET_INPUT_TIMECODE_GRANULARITY 'DOES NOT WORK
'                'DOES NOT WORK, BUT DOES NOT HURT EITHER IF GREATER THAN 100
'                'by default granularity is 100. It suppose to take care of
'                '   the speed of sending update messages
'                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
'            Case XA_MSG_GET_INPUT_TIMECODE_GRANULARITY 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_SET_OUTPUT_VOLUME 'OK
'                'wparam is divided in three parts:
'                '   upper 16 bits - master volume 0-100
'                '   bits 8-15 - wav volume 0-100
'                '   bits 0-7 - balance left/right 0-100
'                Dim mast As Long
'                Dim wavvol As Long
'                Dim bal As Long
'                mast = (wParam) \ (256& * 256& * 256&)
'                wavvol = (wParam - mast * 256& * 256& * 256&) \ (256& * 256)
'                bal = (wParam - mast * 256& * 256& * 256& - wavvol * 256& * 256&) \ 256&
'                If mast > 100 Then mast = 100
'                If wavvol > 100 Then wavvol = 100
'                If bal > 100 Then bal = 100
'                If mast < 0 Then mast = 0
'                If wavvol < 0 Then wavvol = 0
'                If bal < 0 Then bal = 0
'                ret = control_message_send_III(lPlayerHandle, iMsg, bal, wavvol, mast)
'            Case XA_MSG_GET_OUTPUT_VOLUME 'OK
'                'send back separate data for master, wav and balance
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_SET_OUTPUT_CHANNELS 'OK
'                'accepts only values from 0 to 4
'                'no needs to use this command
'                If (wParam >= 0) And (wParam < 5) Then
'                    ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
'                End If
'            Case XA_MSG_GET_OUTPUT_CHANNELS 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            Case XA_MSG_SET_CODEC_EQUALIZER 'OK
'                'wParam is an address of XA_equalizer structure
'                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
'                Debug.Print "Equalizer Set"
'            Case XA_MSG_GET_CODEC_EQUALIZER 'OK
'                ret = control_message_send_N(lPlayerHandle, iMsg)
'            End Select
'        ElseIf iMsg > 36864 Then
'            ProcessReturn iMsg - 36864, wParam, lParam
'
'        'EXECUTE DEFAULT PROCEDURE FOR THE WINDOW
'        Else
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
'        End If
'    Else
'        'message was not send by player redirect it to default Windows procedure
'        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
'        Exit Function
'    End If
'
'End Function
'
'
Public Function twipsX( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsX = PixelsIn * Screen.TwipsPerPixelX
End Function
Public Function twipsY( _
                        ByVal PixelsIn As Variant) _
                As Long
    twipsY = PixelsIn * Screen.TwipsPerPixelY
End Function

Public Function pixelsX(ByVal TwipsIn As Long) As Integer
    pixelsX = TwipsIn / Screen.TwipsPerPixelX
End Function

Public Function pixelsY(ByVal TwipsIn As Long) As Integer
    pixelsY = TwipsIn / Screen.TwipsPerPixelY
End Function


