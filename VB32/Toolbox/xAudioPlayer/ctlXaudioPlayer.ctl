VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlXaudioPlayer 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   630
   ScaleWidth      =   2475
   Begin MSComctlLib.ImageList il16 
      Left            =   4470
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   32768
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXaudioPlayer.ctx":0000
            Key             =   "PAUSE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXaudioPlayer.ctx":0552
            Key             =   "PLAY"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXaudioPlayer.ctx":0664
            Key             =   "SETTINGS"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXaudioPlayer.ctx":0BB6
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXaudioPlayer.ctx":1108
            Key             =   "VOLUME"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider slidePosition 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Position"
      Top             =   0
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   397
      _Version        =   393216
      OLEDropMode     =   1
      Max             =   50
      TickStyle       =   3
   End
   Begin MSComctlLib.Toolbar tbPlayer 
      Height          =   330
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "il16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PLAY"
            Object.ToolTipText     =   "Play"
            ImageKey        =   "PLAY"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PAUSE"
            Object.ToolTipText     =   "Pause"
            ImageKey        =   "PAUSE"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "STOP"
            Object.ToolTipText     =   "Stop"
            ImageKey        =   "STOP"
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "VOLUME"
            ImageKey        =   "VOLUME"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SETTINGS"
            ImageKey        =   "SETTINGS"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Slider slideVolume 
      Height          =   570
      Left            =   2190
      TabIndex        =   2
      ToolTipText     =   "Volume"
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   1005
      _Version        =   393216
      Enabled         =   0   'False
      OLEDropMode     =   1
      Orientation     =   1
      Max             =   100
      TickFrequency   =   20
   End
End
Attribute VB_Name = "ctlXaudioPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private bStarting As Boolean
Private bSliding As Boolean
Private bPaused As Boolean
Private bProcessingReturn As Boolean
Private bPushingButton As Boolean
Private bHidden As Boolean
Private bStopping As Boolean
Private bPinging As Boolean
Private Const lNotifyOffset = 36864
Private eResponse As EMsgResponse

Private lPlayerHandle As Long
Private lPreviousWindowProc As Long

Private tCurrentMessage As XA_Message
Private eCurrentCommand As XA_MessageCode

Private sCurrentFile As String
Private sCurrentShortFile As String

Private sCurrentOutputFile As String
Private sCurrentOutputShortFile As String


Public Enum enumButtons
    btnStart = 1
    btnPause = 2
    btnStop = 3
End Enum

Public Enum XA_MISC
    XA_DECODER_EQUALIZER_NB_BANDS = 31
    OT_COMMAND = 5000&
End Enum

Public Enum enumGetWindowLongOptions
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum



Public Enum XA_MessageCode
    XA_MSG_UNKNOWN = 0
    
    'commands to decoder
    XA_MSG_COMMAND_EXIT = 1
    XA_MSG_COMMAND_SYNC = 2
    XA_MSG_COMMAND_PING = 3
    XA_MSG_COMMAND_PLAY = 4
    XA_MSG_COMMAND_PAUSE = 5
    XA_MSG_COMMAND_STOP = 6
    XA_MSG_COMMAND_SEEK = 7
    XA_MSG_COMMAND_INPUT_OPEN = 8
    XA_MSG_COMMAND_INPUT_CLOSE = 9
    XA_MSG_COMMAND_INPUT_SEND_MESSAGE = 10
    XA_MSG_COMMAND_INPUT_ADD_FILTER = 11
    XA_MSG_COMMAND_INPUT_REMOVE_FILTER = 12
    XA_MSG_COMMAND_INPUT_FILTERS_LIST = 13
    XA_MSG_COMMAND_INPUT_MODULE_REGISTER = 14
    XA_MSG_COMMAND_INPUT_MODULE_QUERY = 15
    XA_MSG_COMMAND_INPUT_MODULES_LIST = 16
    XA_MSG_COMMAND_OUTPUT_OPEN = 17
    XA_MSG_COMMAND_OUTPUT_CLOSE = 18
    XA_MSG_COMMAND_OUTPUT_SEND_MESSAGE = 19
    XA_MSG_COMMAND_OUTPUT_MUTE = 20
    XA_MSG_COMMAND_OUTPUT_UNMUTE = 21
    XA_MSG_COMMAND_OUTPUT_RESET = 22
    XA_MSG_COMMAND_OUTPUT_DRAIN = 23
    XA_MSG_COMMAND_OUTPUT_ADD_FILTER = 24
    XA_MSG_COMMAND_OUTPUT_REMOVE_FILTER = 25
    XA_MSG_COMMAND_OUTPUT_FILTERS_LIST = 26
    XA_MSG_COMMAND_OUTPUT_MODULE_REGISTER = 27
    XA_MSG_COMMAND_OUTPUT_MODULE_QUERY = 28
    XA_MSG_COMMAND_OUTPUT_MODULES_LIST = 29
    XA_MSG_COMMAND_CODEC_SEND_MESSAGE = 30
    XA_MSG_COMMAND_CODEC_MODULE_REGISTER = 31
    XA_MSG_COMMAND_CODEC_MODULE_QUERY = 32
    XA_MSG_COMMAND_CODEC_MODULES_LIST = 33
    XA_MSG_SET_PLAYER_MODE = 34
    XA_MSG_GET_PLAYER_MODE = 35
    XA_MSG_SET_PLAYER_ENVIRONMENT_INTEGER = 36
    XA_MSG_GET_PLAYER_ENVIRONMENT_INTEGER = 37
    XA_MSG_SET_PLAYER_ENVIRONMENT_STRING = 38
    XA_MSG_GET_PLAYER_ENVIRONMENT_STRING = 39
    XA_MSG_UNSET_PLAYER_ENVIRONMENT = 40
    XA_MSG_SET_INPUT_NAME = 41
    XA_MSG_GET_INPUT_NAME = 42
    XA_MSG_SET_INPUT_MODULE = 43
    XA_MSG_GET_INPUT_MODULE = 44
    XA_MSG_SET_INPUT_POSITION_RANGE = 45
    XA_MSG_GET_INPUT_POSITION_RANGE = 46
    XA_MSG_SET_INPUT_TIMECODE_GRANULARITY = 47
    XA_MSG_GET_INPUT_TIMECODE_GRANULARITY = 48
    XA_MSG_SET_OUTPUT_NAME = 49
    XA_MSG_GET_OUTPUT_NAME = 50
    XA_MSG_SET_OUTPUT_MODULE = 51
    XA_MSG_GET_OUTPUT_MODULE = 52
    XA_MSG_SET_OUTPUT_POSITION_RANGE = 53
    XA_MSG_GET_OUTPUT_POSITION_RANGE = 54
    XA_MSG_SET_OUTPUT_TIMECODE_GRANULARITY = 55
    XA_MSG_GET_OUTPUT_TIMECODE_GRANULARITY = 56
    XA_MSG_SET_OUTPUT_VOLUME = 57
    XA_MSG_GET_OUTPUT_VOLUME = 58
    XA_MSG_SET_OUTPUT_CHANNELS = 59
    XA_MSG_GET_OUTPUT_CHANNELS = 60
    XA_MSG_SET_OUTPUT_PORTS = 61
    XA_MSG_GET_OUTPUT_PORTS = 62
    XA_MSG_SET_CODEC_QUALITY = 63
    XA_MSG_GET_CODEC_QUALITY = 64
    XA_MSG_SET_CODEC_EQUALIZER = 65
    XA_MSG_GET_CODEC_EQUALIZER = 66
    XA_MSG_SET_CODEC_MODULE = 67
    XA_MSG_GET_CODEC_MODULE = 68
    XA_MSG_SET_NOTIFICATION_MASK = 69
    XA_MSG_GET_NOTIFICATION_MASK = 70
    XA_MSG_SET_DEBUG_LEVEL = 71
    XA_MSG_GET_DEBUG_LEVEL = 72
    
    'notifications from decoder
    XA_MSG_NOTIFY_READY = 73
    XA_MSG_NOTIFY_ACK = 74
    XA_MSG_NOTIFY_NACK = 75
    XA_MSG_NOTIFY_PONG = 76
    XA_MSG_NOTIFY_EXITED = 77
    XA_MSG_NOTIFY_PLAYER_STATE = 78
    XA_MSG_NOTIFY_PLAYER_MODE = 79
    XA_MSG_NOTIFY_PLAYER_ENVIRONMENT_INTEGER = 80
    XA_MSG_NOTIFY_PLAYER_ENVIRONMENT_STRING = 81
    XA_MSG_NOTIFY_INPUT_STATE = 82
    XA_MSG_NOTIFY_INPUT_NAME = 83
    XA_MSG_NOTIFY_INPUT_CAPS = 84
    XA_MSG_NOTIFY_INPUT_POSITION = 85
    XA_MSG_NOTIFY_INPUT_POSITION_RANGE = 86
    XA_MSG_NOTIFY_INPUT_TIMECODE = 87
    XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY = 88
    XA_MSG_NOTIFY_INPUT_MODULE = 89
    XA_MSG_NOTIFY_INPUT_MODULE_INFO = 90
    XA_MSG_NOTIFY_INPUT_DEVICE_INFO = 91
    XA_MSG_NOTIFY_INPUT_FILTER_INFO = 92
    XA_MSG_NOTIFY_OUTPUT_STATE = 93
    XA_MSG_NOTIFY_OUTPUT_NAME = 94
    XA_MSG_NOTIFY_OUTPUT_CAPS = 95
    XA_MSG_NOTIFY_OUTPUT_POSITION = 96
    XA_MSG_NOTIFY_OUTPUT_POSITION_RANGE = 97
    XA_MSG_NOTIFY_OUTPUT_TIMECODE = 98
    XA_MSG_NOTIFY_OUTPUT_TIMECODE_GRANULARITY = 99
    XA_MSG_NOTIFY_OUTPUT_VOLUME = 100
    XA_MSG_NOTIFY_OUTPUT_BALANCE = 101
    XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL = 102
    XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL = 103
    XA_MSG_NOTIFY_OUTPUT_CHANNELS = 104
    XA_MSG_NOTIFY_OUTPUT_PORTS = 105
    XA_MSG_NOTIFY_OUTPUT_MODULE = 106
    XA_MSG_NOTIFY_OUTPUT_MODULE_INFO = 107
    XA_MSG_NOTIFY_OUTPUT_DEVICE_INFO = 108
    XA_MSG_NOTIFY_OUTPUT_FILTER_INFO = 109
    XA_MSG_NOTIFY_STREAM_MIME_TYPE = 110
    XA_MSG_NOTIFY_STREAM_DURATION = 111
    XA_MSG_NOTIFY_STREAM_PARAMETERS = 112
    XA_MSG_NOTIFY_STREAM_PROPERTIES = 113
    XA_MSG_NOTIFY_CODEC_QUALITY = 114
    XA_MSG_NOTIFY_CODEC_EQUALIZER = 115
    XA_MSG_NOTIFY_CODEC_MODULE = 116
    XA_MSG_NOTIFY_CODEC_MODULE_INFO = 117
    XA_MSG_NOTIFY_CODEC_DEVICE_INFO = 118
    XA_MSG_NOTIFY_NOTIFICATION_MASK = 119
    XA_MSG_NOTIFY_DEBUG_LEVEL = 120
    XA_MSG_NOTIFY_PROGRESS = 121
    XA_MSG_NOTIFY_DEBUG = 122
    XA_MSG_NOTIFY_ERROR = 123
    XA_MSG_NOTIFY_PRIVATE_DATA = 124
    
    'commands to timesync
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULE_REGISTER = 125
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULE_QUERY = 126
    XA_MSG_COMMAND_FEEDBACK_HANDLER_MODULES_LIST = 127
    XA_MSG_COMMAND_FEEDBACK_HANDLER_EXIT = 128
    XA_MSG_COMMAND_FEEDBACK_HANDLER_START = 129
    XA_MSG_COMMAND_FEEDBACK_HANDLER_STOP = 130
    XA_MSG_COMMAND_FEEDBACK_HANDLER_PAUSE = 131
    XA_MSG_COMMAND_FEEDBACK_HANDLER_RESTART = 132
    XA_MSG_COMMAND_FEEDBACK_HANDLER_FLUSH = 133
    XA_MSG_COMMAND_FEEDBACK_HANDLER_SEND_MESSAGE = 134
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_AUDIO_EVENT = 135
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_TAG_EVENT = 136
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_TIMECODE_EVENT = 137
    XA_MSG_COMMAND_FEEDBACK_HANDLER_QUEUE_POSITION_EVENT = 138
    XA_MSG_SET_FEEDBACK_AUDIO_EVENT_RATE = 139
    XA_MSG_GET_FEEDBACK_AUDIO_EVENT_RATE = 140
    XA_MSG_SET_FEEDBACK_HANDLER_NAME = 141
    XA_MSG_GET_FEEDBACK_HANDLER_NAME = 142
    XA_MSG_SET_FEEDBACK_HANDLER_MODULE = 143
    XA_MSG_GET_FEEDBACK_HANDLER_MODULE = 144
    XA_MSG_SET_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 145
    XA_MSG_GET_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 146
    XA_MSG_SET_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 147
    XA_MSG_GET_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 148
    XA_MSG_UNSET_FEEDBACK_HANDLER_ENVIRONMENT = 149
    
    'notifications from timesync
    XA_MSG_NOTIFY_FEEDBACK_AUDIO_EVENT_RATE = 150
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_STATE = 151
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_MODULE = 152
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_MODULE_INFO = 153
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_NAME = 154
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_INFO = 155
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_ENVIRONMENT_INTEGER = 156
    XA_MSG_NOTIFY_FEEDBACK_HANDLER_ENVIRONMENT_STRING = 157
    XA_MSG_NOTIFY_FEEDBACK_AUDIO_EVENT = 158
    XA_MSG_NOTIFY_FEEDBACK_TAG_EVENT = 159
    
    'sentinel
    XA_MSG_LAST = 160
End Enum

'input state
Public Enum InputState
    XA_INPUT_STATE_OPEN = 0
    XA_INPUT_STATE_CLOSED = 1
End Enum

'output state
Public Enum OutputState
    XA_OUTPUT_STATE_OPEN = 0
    XA_OUTPUT_STATE_CLOSED = 1
End Enum


'OUTPUT CHANNELS
Public Enum XA_OutputChannels
    XA_OUTPUT_CHANNELS_STEREO = 0
    XA_OUTPUT_CHANNELS_MONO_LEFT = 1
    XA_OUTPUT_CHANNELS_MONO_RIGHT = 2
    XA_OUTPUT_CHANNELS_MONO_MIX = 3
End Enum


Public Type XA_InputStreamInfo
    mChanged As Long             '0 if the stream information has not changed since the last decoded frame, or non zero if it has
    mLevel As Long               'MPEG syntax level (1 for MPEG1, 2 for MPEG2, 0 for MPEG2.5)
    mLayer As Long               'MPEG layer (1, 2 or 3)
    mBitrate As Long             'MPEG bitrate (in bits per second)
    mFrequency As Long           'MPEG sampling frequency (in Hz)
    mMode As Long                'MPEG mode (0 for stereo, 1 for joint-stereo, 2 for dual-channel, 3 for mono)
    mDuration As Long            'estimated stream duration (in milliseconds)
End Type

Public Type XA_TimeCode
    mH As Long       'hours
    mM As Long       'minutes
    mS As Long       'seconds
    mF As Long       'fractures in 100th of second
End Type

Public Type XA_StatusInfo
    mFrame As Long               'current frame number
    mPosition As Single          'Value between 0.0 and 1.0 giving the relative position in the stream
    mInfo As XA_InputStreamInfo  'input stream structure
    mTimecode As XA_TimeCode     'time code structure
End Type

Public Type XA_AbsoluteTime
    mSeconds As Long
    mMicroseconds As Long
End Type

Public Type XA_EnvironmentInfo
    mName As String
    mInteger As Long
    mString As String
End Type

Public Type XA_TimecodeInfo
    mH As Byte
    mM As Byte
    mS As Byte
    mF As Byte
End Type

Public Type XA_NackInfo
    mCommand As Byte
    mCode As Integer
End Type

Public Type XA_VolumeInfo
    mMasterLevel As Byte
    mPCMLevel As Byte
    mBalance As Byte
End Type

Public Type XA_PositionInfo
    mOffset As Long
    mRange As Long
End Type

Public Type XA_ModuleInfo
    nID As Byte
    mNBDevices As Byte
    mName As String
    mDescription As String
End Type

Public Type XA_FilterInfo
    nID As Integer
    mName As String
End Type

Public Type XA_DeviceInfo
    mModuleID As Byte
    mIndex As Byte
    mFlags As Byte
    mName As String
    mDescription As String
End Type

Public Type XA_StreamParameters
    mFrequency As Long
    mBitrate As Integer
    mNBChannels As Byte
End Type

Public Type XA_ModuleMessage
    mType As Integer
    mSize As Long
    mData As Long
End Type

Public Type XA_TagEvent
    mWhen As XA_AbsoluteTime
    mTag As Long
End Type

Public Type XA_AudioEvent
    mWhen As XA_AbsoluteTime
    mSamplingFrequency As Long
    mNBChannels As Integer
    mNBSamples As Integer
    mSamples As String
End Type

Public Type XA_TimecodeEvent
    mWhen As XA_AbsoluteTime
    mTimecode As XA_TimecodeInfo
End Type

Public Type XA_PositionEvent
    mWhen As XA_AbsoluteTime
    mPosition As XA_PositionInfo
End Type

Public Type XA_ProgressInfo
    mSource As Byte
    mCode As Byte
    mValue As Integer
    mMessage As String
End Type

Public Type XA_DebugInfo
    mSource As Byte
    mLevel As Byte
    mMessage As String
End Type

Public Type XA_ErrorInfo
    mSource As Byte
    mCode As Integer
    mMessage As String
End Type

Public Type XA_PrivateData
    mSource As Byte
    mType As Integer
    mData As Long
    mSize As Long
End Type

Public Type XA_EqualizerInfo
    eleft(XA_DECODER_EQUALIZER_NB_BANDS) As Byte
    eright(XA_DECODER_EQUALIZER_NB_BANDS) As Byte
End Type

Public Enum XA_PropertyType
    XA_PROPERTY_TYPE_STRING
    XA_PROPERTY_TYPE_INTEGER
End Enum

Public Type XA_PropertyValue
    mInteger As Long
    mString As String
End Type

Public Type XA_Property
    mName As String
    mType As XA_PropertyType
    mValue As XA_PropertyValue
End Type

Public Type XA_PropertyList
    mNBProperties As Long
    mProperties As XA_Property
End Type

'INPUT/OUTPUT STATE
Public Enum XA_InputOutputState
    XA_STATE_OPEN = 0
    XA_STATE_CLOSED = 1
End Enum

'PLAYER STATE
Public Enum XA_PlayerState
    XA_PLAYER_STATE_STOPPED = 0
    XA_PLAYER_STATE_PLAYING = 1
    XA_PLAYER_STATE_PAUSED = 2
    XA_PLAYER_STATE_EOS = 3
End Enum

Public Type XA_Message
    mCode As XA_MessageCode
    'data structure follows
    mBuffer As String
    mName As String
    mString As String
    mMimeType As String
    mModuleID As Integer
    mMode As Long
    mChannels As Byte
    mQuality As Byte
    mDuration As Long
    mRange As Long
    mGranularity As Long
    mCaps As Long
    mPorts As Byte
    mAck As Byte
    mTag As Long
    mDebugLevel As Byte
    mNotificationMask As Long
    mRate As Byte
    mNack As XA_NackInfo
    mVolume As XA_VolumeInfo
    mPosition As XA_PositionInfo
    mEqualizer As XA_EqualizerInfo
    mModuleinfo As XA_ModuleInfo
    mFilterInfo As XA_FilterInfo
    mDeviceInfo As XA_DeviceInfo
    mStreamParameters As XA_StreamParameters
    mEnvironmentInfo As XA_EnvironmentInfo
    mTimecode As XA_TimecodeInfo
    mModuleMessage As XA_ModuleMessage
    mTagEvent As XA_TagEvent
    mAudioEvent As XA_AudioEvent
    mTimecodeEvent As XA_TimecodeEvent
    mPositionEvent As XA_PositionEvent
    mProperties As XA_PropertyList
    mWhen As XA_AbsoluteTime
    mProgress As XA_ProgressInfo
    mDebug As XA_DebugInfo
    mError As XA_ErrorInfo
    mPrivateData As XA_PrivateData
    mInputState As XA_InputOutputState
    mOutputState As XA_InputOutputState
    mPlayerState As XA_PlayerState
    mInputStreamInfo As XA_InputStreamInfo
    'end of data structure
End Type


'open new player if par=0 then msges will be sent to player, otherwise to specified window handle hWnd
Private Declare Function player_new Lib "xaudio.dll" (hPlayer As Long, par As Long) As Long
'delete specified player
Private Declare Function player_delete Lib "xaudio.dll" (ByVal hPlayer As Long) As Long
Private Declare Function player_set_priority Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal par As XA_CONTROL_PRIORITY) As Long
Private Declare Function player_get_priority Lib "xaudio.dll" (ByVal hPlayer As Long) As XA_CONTROL_PRIORITY
Private Declare Function control_message_send_S Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal msg_str As String) As Long
Private Declare Function control_message_send_N Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long) As Long
Private Declare Function control_message_send_I Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat As Long) As Long
Private Declare Function control_message_get Lib "xaudio.dll" (hPlayer As Long, status As XA_Message) As Long
'Private Declare Function control_message_wait Lib "xaudio.dll" (ByVal hPlayer As Long, status As XA_Message, ByVal TimeOut As Long) As Long
Private Declare Function control_message_wait Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal status As Long, ByVal TimeOut As Long) As Long

'Public Declare Function control_message_sprint Lib "xaudio.dll" (strBuff As String, status As XA_Message) As Long
Private Declare Function xaudio_get_version Lib "xaudio.dll" (ByVal cc As Long) As Long
Private Declare Function control_message_send_IPI Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ptr As Any, ByVal dat2 As Long) As Long
Private Declare Function control_message_send_II Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ByVal dat2 As Long) As Long
Private Declare Function control_message_send_P Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, dat As Any) As Long
Private Declare Function control_message_send_III Lib "xaudio.dll" (ByVal hPlayer As Long, ByVal msg_code As Long, ByVal dat1 As Long, ByVal dat2 As Long, ByVal dat3 As Long) As Long


Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessagePointer Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function agGetStringFromPointer Lib "apigid32.dll" Alias "agGetStringFromLPSTR" (ByVal ptr As Long) As String
Private Declare Sub agCopyData Lib "apigid32.dll" (ByVal source As Long, dest As Any, ByVal nCount As Long)
Private Declare Function agGetAddressForObject Lib "apigid32.dll" (obj As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long



'FEEDBACK HANDLER STATE
Public Enum XA_FEEDBACK_HANDLER_STATES
    XA_FEEDBACK_HANDLER_STATE_STARTED = 0
    XA_FEEDBACK_HANDLER_STATE_STOPPED = 1
End Enum

'ERROR CODES
Public Enum XA_ERRORFLAGS
    XA_SUCCESS = 0
    XA_FAILURE = -1
End Enum

'Priorities
Public Enum XA_CONTROL_PRIORITY
    XA_CONTROL_PRIORITY_LOWEST = 0
    XA_CONTROL_PRIORITY_LOW = 1
    XA_CONTROL_PRIORITY_NORMAL = 2
    XA_CONTROL_PRIORITY_HIGH = 3
    XA_CONTROL_PRIORITY_HIGHEST = 4
End Enum

'general error codes
Public Enum XA_ERRORS
    XA_ERROR_BASE_GENERAL = -100
    XA_ERROR_OUT_OF_MEMORY = XA_ERROR_BASE_GENERAL - 0
    XA_ERROR_OUT_OF_RESOURCES = XA_ERROR_BASE_GENERAL - 1
    XA_ERROR_INVALID_PARAMETERS = XA_ERROR_BASE_GENERAL - 2
    XA_ERROR_INTERNAL = XA_ERROR_BASE_GENERAL - 3
    XA_ERROR_TIMEOUT = XA_ERROR_BASE_GENERAL - 4
    XA_ERROR_VERSION_EXPIRED = XA_ERROR_BASE_GENERAL - 5
    XA_ERROR_VERSION_MISMATCH = XA_ERROR_BASE_GENERAL - 6
    
    'network error codes
    XA_ERROR_BASE_NETWORK = -200
    XA_ERROR_CONNECT_TIMEOUT = XA_ERROR_BASE_NETWORK - 0
    XA_ERROR_CONNECT_FAILED = XA_ERROR_BASE_NETWORK - 1
    XA_ERROR_CONNECTION_REFUSED = XA_ERROR_BASE_NETWORK - 2
    XA_ERROR_ACCEPT_FAILED = XA_ERROR_BASE_NETWORK - 3
    XA_ERROR_LISTEN_FAILED = XA_ERROR_BASE_NETWORK - 4
    XA_ERROR_SOCKET_FAILED = XA_ERROR_BASE_NETWORK - 5
    XA_ERROR_SOCKET_CLOSED = XA_ERROR_BASE_NETWORK - 6
    XA_ERROR_BIND_FAILED = XA_ERROR_BASE_NETWORK - 7
    XA_ERROR_HOST_UNKNOWN = XA_ERROR_BASE_NETWORK - 8
    XA_ERROR_HTTP_INVALID_REPLY = XA_ERROR_BASE_NETWORK - 9
    XA_ERROR_HTTP_ERROR_REPLY = XA_ERROR_BASE_NETWORK - 10
    XA_ERROR_HTTP_FAILURE = XA_ERROR_BASE_NETWORK - 11
    XA_ERROR_FTP_INVALID_REPLY = XA_ERROR_BASE_NETWORK - 12
    XA_ERROR_FTP_ERROR_REPLY = XA_ERROR_BASE_NETWORK - 13
    XA_ERROR_FTP_FAILURE = XA_ERROR_BASE_NETWORK - 14
    
    'control error codes
    XA_ERROR_BASE_CONTROL = -300
    XA_ERROR_PIPE_FAILED = XA_ERROR_BASE_CONTROL - 0
    XA_ERROR_FORK_FAILED = XA_ERROR_BASE_CONTROL - 1
    XA_ERROR_SELECT_FAILED = XA_ERROR_BASE_CONTROL - 2
    XA_ERROR_PIPE_CLOSED = XA_ERROR_BASE_CONTROL - 3
    XA_ERROR_PIPE_READ_FAILED = XA_ERROR_BASE_CONTROL - 4
    XA_ERROR_PIPE_WRITE_FAILED = XA_ERROR_BASE_CONTROL - 5
    XA_ERROR_INVALID_MESSAGE = XA_ERROR_BASE_CONTROL - 6
    XA_ERROR_CIRQ_FULL = XA_ERROR_BASE_CONTROL - 7
    XA_ERROR_POST_FAILED = XA_ERROR_BASE_CONTROL - 8
    
    'url error codes
    XA_ERROR_BASE_URL = -400
    XA_ERROR_URL_UNSUPPORTED_SCHEME = XA_ERROR_BASE_URL - 0
    XA_ERROR_URL_INVALID_SYNTAX = XA_ERROR_BASE_URL - 1
    
    'i/o error codes
    XA_ERROR_BASE_IO = -500
    XA_ERROR_OPEN_FAILED = XA_ERROR_BASE_IO - 0
    XA_ERROR_CLOSE_FAILED = XA_ERROR_BASE_IO - 1
    XA_ERROR_READ_FAILED = XA_ERROR_BASE_IO - 2
    XA_ERROR_WRITE_FAILED = XA_ERROR_BASE_IO - 3
    XA_ERROR_PERMISSION_DENIED = XA_ERROR_BASE_IO - 4
    XA_ERROR_NO_DEVICE = XA_ERROR_BASE_IO - 5
    XA_ERROR_IOCTL_FAILED = XA_ERROR_BASE_IO - 6
    XA_ERROR_MODULE_NOT_FOUND = XA_ERROR_BASE_IO - 7
    XA_ERROR_UNSUPPORTED_INPUT = XA_ERROR_BASE_IO - 8
    XA_ERROR_UNSUPPORTED_OUTPUT = XA_ERROR_BASE_IO - 9
    XA_ERROR_UNSUPPORTED_FORMAT = XA_ERROR_BASE_IO - 10
    XA_ERROR_DEVICE_BUSY = XA_ERROR_BASE_IO - 11
    XA_ERROR_NO_SUCH_DEVICE = XA_ERROR_BASE_IO - 12
    XA_ERROR_NO_SUCH_FILE = XA_ERROR_BASE_IO - 13
    XA_ERROR_INPUT_EOS = XA_ERROR_BASE_IO - 14
        
    'codec error codes
    XA_ERROR_BASE_CODEC = -600
    XA_ERROR_NO_CODEC = XA_ERROR_BASE_CODEC - 0
    
    'bitstream error codes
    XA_ERROR_BASE_BITSTREAM = -700
    XA_ERROR_INVALID_FRAME = XA_ERROR_BASE_BITSTREAM - 0
    
    'dynamic linking error codes
    XA_ERROR_BASE_DYNLINK = -800
    XA_ERROR_DLL_NOT_FOUND = XA_ERROR_BASE_DYNLINK - 0
    XA_ERROR_SYMBOL_NOT_FOUND = XA_ERROR_BASE_DYNLINK - 1
        
    'environment variables / porperties  error codes
    XA_ERROR_BASE_ENVIRONMENT = -900
    XA_ERROR_NO_SUCH_ENVIRONMENT = XA_ERROR_BASE_ENVIRONMENT - 0
    XA_ERROR_NO_SUCH_PROPERTY = XA_ERROR_BASE_ENVIRONMENT - 0
    XA_ERROR_ENVIRONMENT_TYPE_MISMATCH = XA_ERROR_BASE_ENVIRONMENT - 1
    XA_ERROR_PROPERTY_TYPE_MISMATCH = XA_ERROR_BASE_ENVIRONMENT - 1
    
    'modules
    XA_ERROR_BASE_MODULES = -1000
    XA_ERROR_NO_SUCH_INTERFACE = XA_ERROR_BASE_MODULES - 0
End Enum


Public Event EndOfStream()
Public Event Error(tErr As XA_ErrorInfo)
Public Event DebugError(tErr As XA_DebugInfo)
Public Event PlayerStop()
Public Event Position(ByVal lValue As Long, ByVal lMax As Long)
Public Event PlayerStateChanged(ByVal eNewState As XA_PlayerState)


Private Function PushButton(ByVal eButton As enumButtons)
    bPushingButton = True
    On Error Resume Next
    UserControl.tbPlayer.Buttons(eButton).Value = tbrPressed
    bPushingButton = False

End Function




'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' SendCommand
'    The purpose of this function is to send a message to the player,
'    and then wait until the response message has been received,
'    indicating with a boolean value whether it was successful or not.
'
'    This is only place in the code that should have a doevents.
'    In the future, this function could be re-written to use the
'    control_message_wait()
'    function.
Private Function SendCommand( _
                            ByVal cmnd As XA_MessageCode, _
                            Optional ByVal wpar As Long, _
                            Optional ByVal lpar As Long, _
                            Optional ByVal bDoEvents As Boolean = True) _
                As Boolean
    On Error Resume Next
    
    Dim StartTime As Long
    StartTime = GetTickCount
    
    eCurrentCommand = cmnd
    tCurrentMessage.mAck = 0
    tCurrentMessage.mNack.mCode = 0
    tCurrentMessage.mNack.mCommand = 0
    tCurrentMessage.mDebug.mMessage = ""
    Dim lRet As Long
    Dim yRet As Long
    Dim bTimeout As Boolean
    
    lRet = SendMessage(UserControl.hWnd, cmnd, wpar, lpar)
    If bDoEvents Then
        bTimeout = True
        Do While lRet = 0 And ((GetTickCount - StartTime) < OT_COMMAND)
            DoEvents
            Select Case True
                Case tCurrentMessage.mAck = eCurrentCommand
'                    If tCurrentMessage.mDebug.mMessage <> "" Then
'                        MsgBox "Pause here."
'                    End If
                        SendCommand = True
                        bTimeout = False
                    Exit Do
                Case tCurrentMessage.mNack.mCommand = eCurrentCommand
                    bTimeout = False
                    Exit Do
            End Select
        Loop
    Else
    
        If lRet = 0 Then
            lRet = control_message_wait(lPlayerHandle, VarPtr(tCurrentMessage), OT_COMMAND)
            If tCurrentMessage.mAck = eCurrentCommand Then
                lRet = 0
                SendCommand = True
            End If
        End If
    End If
    
    If lRet = XA_ERROR_TIMEOUT Then
        bTimeout = True
    End If
    
    If bTimeout Then
        MsgBox "Command: " & eCurrentCommand & " timed out."
    End If
    
End Function
Private Function QuickCommand( _
                            ByVal eCmd As XA_MessageCode, _
                            Optional ByVal wpar As Long, _
                            Optional ByVal lpar As Long)
                            
    SendMessage UserControl.hWnd, eCmd, 0, 0
End Function

Public Function Pause()
    SendCommand XA_MSG_COMMAND_PAUSE
    'control_message_send_N lPlayerHandle, XA_MSG_COMMAND_PAUSE
    
End Function


Private Property Let ISubclass_MsgResponse(ByVal RHS As jeffStaticDLL.EMsgResponse)
    eResponse = RHS
End Property

Private Property Get ISubclass_MsgResponse() As jeffStaticDLL.EMsgResponse
    ISubclass_MsgResponse = eResponse
End Property


Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'this will take care of sending appropriate command to player or execute
    '   default Windows procedure
    Dim ret As Long
    On Error Resume Next
    'Debug.Print hw, iMsg, wParam, lParam, lPlayerHandle

    If GetWindowLong(hWnd, GWL_USERDATA) = lPlayerHandle Then
        'Go ahead - execute command
        If iMsg = eCurrentCommand Then
            'clear tcurrentmessage structure
            tCurrentMessage.mAck = 0
            tCurrentMessage.mNack.mCode = 0
            tCurrentMessage.mNack.mCommand = 0
            'execute command here
            Select Case iMsg
            Case XA_MSG_COMMAND_INPUT_OPEN 'OK
                'we can send an address of the string - but i am just too lazy
                ret = control_message_send_S(lPlayerHandle, iMsg, sCurrentShortFile)
            Case XA_MSG_COMMAND_INPUT_CLOSE 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_OUTPUT_OPEN 'OK
                'we can send an address of the string - but i am just too lazy
                ret = control_message_send_S(lPlayerHandle, iMsg, sCurrentOutputShortFile)
            Case XA_MSG_COMMAND_OUTPUT_CLOSE 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_OUTPUT_RESET
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_INPUT_SEND_MESSAGE
                'do not use it - EVER!!!!!
            Case XA_MSG_COMMAND_PLAY ' start playing - if applicable. OK
                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PLAYING Then Exit Function
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_STOP ' stop playing. OK
                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_EOS Or tCurrentMessage.mPlayerState = XA_PLAYER_STATE_STOPPED Then Exit Function
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_EXIT ' close input and output and exit.
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_PAUSE 'OK
                'execute only if playing
                If tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PLAYING Then
                    ret = control_message_send_N(lPlayerHandle, iMsg)
                Else
                    Exit Function
                End If
            Case XA_MSG_COMMAND_SEEK
                'wParam is an offset, lParam is a range
                ret = control_message_send_II(lPlayerHandle, iMsg, wParam, lParam)
            Case XA_MSG_COMMAND_SYNC 'DON'T KNOW HOW TO USE
                'do not use
            Case XA_MSG_SET_INPUT_POSITION_RANGE 'OK
                'by default range=400
                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
            Case XA_MSG_GET_INPUT_POSITION_RANGE 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_OUTPUT_MUTE 'OK
                'it does not really mute the output
                '   it just sets the volume of wav to 0
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_COMMAND_OUTPUT_UNMUTE 'OK
                'recall value of volume, when it was muted
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_SET_INPUT_TIMECODE_GRANULARITY 'DOES NOT WORK
                'DOES NOT WORK, BUT DOES NOT HURT EITHER IF GREATER THAN 100
                'by default granularity is 100. It suppose to take care of
                '   the speed of sending update messages
                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
            Case XA_MSG_GET_INPUT_TIMECODE_GRANULARITY 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_SET_OUTPUT_VOLUME 'OK
                'wparam is divided in three parts:
                '   upper 16 bits - master volume 0-100
                '   bits 8-15 - wav volume 0-100
                '   bits 0-7 - balance left/right 0-100
                Dim mast As Long
                Dim wavvol As Long
                Dim bal As Long
                mast = (wParam) \ (256& * 256& * 256&)
                wavvol = (wParam - mast * 256& * 256& * 256&) \ (256& * 256)
                bal = (wParam - mast * 256& * 256& * 256& - wavvol * 256& * 256&) \ 256&
                If mast > 100 Then mast = 100
                If wavvol > 100 Then wavvol = 100
                If bal > 100 Then bal = 100
                If mast < 0 Then mast = 0
                If wavvol < 0 Then wavvol = 0
                If bal < 0 Then bal = 0
                ret = control_message_send_III(lPlayerHandle, iMsg, bal, wavvol, mast)
            Case XA_MSG_GET_OUTPUT_VOLUME 'OK
                'send back separate data for master, wav and balance
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_SET_OUTPUT_CHANNELS 'OK
                'accepts only values from 0 to 4
                'no needs to use this command
                If (wParam >= 0) And (wParam < 5) Then
                    ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
                End If
            Case XA_MSG_GET_OUTPUT_CHANNELS 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            Case XA_MSG_SET_CODEC_EQUALIZER 'OK
                'wParam is an address of XA_equalizer structure
                ret = control_message_send_I(lPlayerHandle, iMsg, wParam)
                Debug.Print "Equalizer Set"
            Case XA_MSG_GET_CODEC_EQUALIZER 'OK
                ret = control_message_send_N(lPlayerHandle, iMsg)
            End Select
        
        'Else
        ElseIf iMsg > 36864 Then
            ProcessReturn iMsg - 36864, wParam, lParam
        '    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        'End If
        'EXECUTE DEFAULT PROCEDURE FOR THE WINDOW
        Else
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        End If
    Else
        'message was not send by player redirect it to default Windows procedure
        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Exit Function
    End If
    
End Function

Public Function ProcessReturn(code As Long, wParam As Long, lParam As Long) As Boolean
    On Error Resume Next
    If Not bProcessingReturn Then
        bProcessingReturn = True
    'Debug.Print code, wParam, lParam
    Select Case code
        Case XA_MSG_NOTIFY_ACK 'OK
            'acknowledged - wParam is a code of command that was acknowledged
            tCurrentMessage.mAck = wParam
        '    Debug.Print "Ack "; wParam; lParam
        Case XA_MSG_NOTIFY_NACK 'OK
            'not acknowledged - wParam - code to be not acjknowledged, lParam - error
            tCurrentMessage.mNack.mCode = lParam
            tCurrentMessage.mNack.mCommand = wParam
        '    Debug.Print "Not Ack"; wParam; lParam
        Case XA_MSG_NOTIFY_READY 'OK - notify when player is created
            'player ready
        Case XA_MSG_NOTIFY_PLAYER_STATE 'OK - state like Playing, stopped, paused, E(nd)O(f)S(tream)
            'returns player state - wParam - code of player state
            tCurrentMessage.mPlayerState = wParam
            RaiseEvent PlayerStateChanged(tCurrentMessage.mPlayerState)
            Select Case True
                Case tCurrentMessage.mPlayerState = XA_PLAYER_STATE_EOS
                    slidePosition.Value = 0
                    PushButton btnStop
                    RaiseEvent EndOfStream
                    'SendCommand XA_MSG_COMMAND_STOP, , , True
                Case tCurrentMessage.mPlayerState = XA_PLAYER_STATE_STOPPED
                    slidePosition.Value = 0
                    PushButton btnStop
                    RaiseEvent PlayerStop
                Case tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PAUSED
                    PushButton btnPause
                    RaiseEvent PlayerStop
                Case tCurrentMessage.mPlayerState = XA_PLAYER_STATE_PLAYING
                    PushButton btnStart
            End Select
        Case XA_MSG_NOTIFY_INPUT_TIMECODE 'OK - during play/decode player sends this info every second
        '   does not make sense to use fraction as it is always 0
        '    tCurrentMessage.mTimecode.mF = CSng(wParam And &HFF000000) / (256& * 256& * 256&)
            tCurrentMessage.mTimecode.mS = (wParam And 16711680) / (256& * 256&)
            tCurrentMessage.mTimecode.mM = (wParam And 65280) / 256&
            tCurrentMessage.mTimecode.mH = wParam And 255
        Case XA_MSG_NOTIFY_DEBUG 'OK
            'if something happened you will be notified to debug
            tCurrentMessage.mDebug.mMessage = agGetStringFromPointer(lParam)
            tCurrentMessage.mDebug.mSource = wParam Mod (256& * 256&)
            tCurrentMessage.mDebug.mLevel = (wParam And &HFFFF0000) \ (256& * 256&)
            Debug.Print "Debug Received "; wParam; agGetStringFromPointer(lParam)
            RaiseEvent DebugError(tCurrentMessage.mDebug)
        Case XA_MSG_NOTIFY_INPUT_STATE 'OK
            'sent every time state of input is changed
            tCurrentMessage.mInputState = wParam
        Case XA_MSG_NOTIFY_OUTPUT_STATE
            'sent every time state of output is changed
            tCurrentMessage.mOutputState = wParam
        Case XA_MSG_NOTIFY_ERROR 'OK
            'difficult to debug
            Debug.Print "Error occured"
            tCurrentMessage.mError.mCode = wParam \ (256& * 256&)
            tCurrentMessage.mError.mSource = wParam Mod (256& * 256&)
            tCurrentMessage.mError.mMessage = agGetStringFromPointer(lParam)
            RaiseEvent Error(tCurrentMessage.mError)
        Case XA_MSG_NOTIFY_PONG
            'did not try
            tCurrentMessage.mTag = wParam
            bPinging = False
        Case XA_MSG_NOTIFY_PLAYER_MODE 'OK
            'notifies when mode is changed
            tCurrentMessage.mMode = wParam
        Case XA_MSG_NOTIFY_INPUT_NAME 'OK
            'does not know how to use it
            tCurrentMessage.mName = agGetStringFromPointer(lParam)
        Case XA_MSG_NOTIFY_INPUT_CAPS 'OK
            'Notifies of the player's current input capabilities.
            'This is typically useful to be notified wether it is possible to seek into an input stream (an disabling a seek bar for instance, if it is not).
            'The flag XA_DECODER_INPUT_SEEKABLE indicates wether it is possible to seek or not.
            tCurrentMessage.mCaps = wParam
        Case XA_MSG_NOTIFY_INPUT_POSITION 'OK
            'notifies when position of input is changed
            tCurrentMessage.mPosition.mOffset = wParam
            tCurrentMessage.mPosition.mRange = lParam
            If Not bSliding And Not bStarting Then
                With slidePosition
                    If .Max <> lParam Then
                        .Max = lParam
                    End If
                    If .Value <> wParam Then
                        .Value = wParam
                    End If
                End With
            End If
            
        Case XA_MSG_NOTIFY_INPUT_POSITION_RANGE 'OK
            'notifies when range is changed - by default the range is 400
            tCurrentMessage.mPosition.mRange = lParam
        Case XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY 'OK
            tCurrentMessage.mGranularity = lParam
        Case XA_MSG_NOTIFY_STREAM_DURATION 'OK
            'shows it when stream is VBR - estimated duration of stream
            tCurrentMessage.mDuration = wParam
        Case XA_MSG_NOTIFY_STREAM_PARAMETERS 'OK
            'shows it when stream is VBR after each chunk of data
            '   and on opening the input file
            tCurrentMessage.mStreamParameters.mNBChannels = CByte(wParam \ (256& * 256&))
            tCurrentMessage.mStreamParameters.mBitrate = CInt(wParam And (255& * 256& + 255&)) 'in kB/s
            tCurrentMessage.mStreamParameters.mFrequency = lParam 'in Hertz
        Case XA_MSG_NOTIFY_STREAM_PROPERTIES
            'it fires but don't know how to use
        Case XA_MSG_NOTIFY_STREAM_MIME_TYPE
            tCurrentMessage.mMimeType = agGetStringFromPointer(lParam)
        Case XA_MSG_NOTIFY_INPUT_MODULE
            'did not try and do not know how to use
            tCurrentMessage.mModuleID = wParam
        Case XA_MSG_NOTIFY_INPUT_MODULE_INFO
            'did not try and do not know how to use
            tCurrentMessage.mModuleinfo.nID = CByte(wParam And &HFF)
            tCurrentMessage.mModuleinfo.mNBDevices = CByte((wParam And &HFF00) / 256)
            tCurrentMessage.mModuleinfo.mName = agGetStringFromPointer(lParam)
        Case XA_MSG_NOTIFY_OUTPUT_NAME
            'did not try and do not know how to use
            
            'tCurrentMessage.mName = agGetStringFromPointer(lParam)
        Case XA_MSG_NOTIFY_OUTPUT_CAPS
            'did not try
            tCurrentMessage.mCaps = wParam
        Case XA_MSG_NOTIFY_OUTPUT_VOLUME 'this does not fire ever
            tCurrentMessage.mVolume.mMasterLevel = CByte(wParam And &HFF)
            tCurrentMessage.mVolume.mPCMLevel = CByte((wParam And &HFF00) / 256&)
            tCurrentMessage.mVolume.mBalance = CByte((wParam And &HFF0000) / (256& * 256&))
        Case XA_MSG_NOTIFY_OUTPUT_BALANCE 'OK
            tCurrentMessage.mVolume.mBalance = wParam
        Case XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL 'OK
            tCurrentMessage.mVolume.mPCMLevel = wParam
        Case XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL 'OK
            tCurrentMessage.mVolume.mMasterLevel = wParam
        Case XA_MSG_NOTIFY_OUTPUT_CHANNELS 'OK
            tCurrentMessage.mChannels = CByte(wParam)
        Case XA_MSG_NOTIFY_OUTPUT_MODULE_INFO
            'did not try and do not know how to use
            tCurrentMessage.mModuleinfo.nID = CByte(wParam And &HFF)
            tCurrentMessage.mModuleinfo.mNBDevices = CByte((wParam And &HFF00) / 256)
            tCurrentMessage.mModuleinfo.mName = agGetStringFromPointer(lParam)
        Case XA_MSG_NOTIFY_CODEC_EQUALIZER 'DOES NOT WORK VERY GOOD
            'I have no idea why I should add 4 to the address
            agCopyData lParam + 4, tCurrentMessage.mEqualizer, Len(tCurrentMessage.mEqualizer)
        '    Dim i As Long
        '    For i = 0 To 31
        '        Debug.Print i, tCurrentMessage.mEqualizer.eleft(i), tCurrentMessage.mEqualizer.eright(i)
        '    Next i
        Case XA_MSG_NOTIFY_NOTIFICATION_MASK
            'did not try
            tCurrentMessage.mNotificationMask = wParam
        Case XA_MSG_NOTIFY_PROGRESS
            'did not try
            tCurrentMessage.mProgress.mSource = CByte(wParam And &HFF)
            tCurrentMessage.mProgress.mCode = CByte((wParam And &HFF00) / 256&)
            tCurrentMessage.mProgress.mValue = CInt((wParam And &HFFFF0000) / (256& * 256&))
            tCurrentMessage.mProgress.mMessage = agGetStringFromPointer(lParam)
    End Select
        bProcessingReturn = False
    Else
        'MsgBox "Return within a return."
    End If
    
End Function

Private Sub slidePosition_Change()
    If Not bProcessingReturn And Not bStopping Then
'        DoEvents
        eCurrentCommand = XA_MSG_COMMAND_SEEK
        SendCommand eCurrentCommand, slidePosition.Value, slidePosition.Max
        'start playing here
'        eCurrentCommand = XA_MSG_COMMAND_PLAY
'        SendCommand eCurrentCommand, 0, 0
    End If
    
End Sub

Private Sub slidePosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bSliding = True
    
End Sub

Private Sub slidePosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bSliding = False
    
End Sub

Private Sub slidePosition_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    DroppedFiles Data
End Sub

Private Sub slideVolume_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    DroppedFiles Data
End Sub

Private Sub tbPlayer_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Not bProcessingReturn Then
        Select Case UCase(Button.Key)
            Case "PLAY"
                Me.Play
            Case "PAUSE"
                Me.Pause
            Case "STOP"
                Me.StopPlayer
        End Select
    End If
    
End Sub

Private Sub tbPlayer_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    DroppedFiles Data
End Sub

Private Sub UserControl_Hide()
'    MsgBox "Control hide"
    bHidden = True
    
End Sub

Private Sub UserControl_Initialize()
    
    'have to preinitialize some vars
    tCurrentMessage.mPlayerState = XA_PLAYER_STATE_STOPPED
    tCurrentMessage.mInputState = XA_STATE_CLOSED
    tCurrentMessage.mOutputState = XA_STATE_CLOSED
    
    'Set fPlayer = New frmPlayer
    'Load fPlayer
    Dim lRet As Long
    lRet = player_new(lPlayerHandle, UserControl.hWnd)
    If lRet < 0 Then
        Err.Raise lRet, "XaudioPlayer.ctlXaudioPlayer", "Failed to initialize a new player.  Double check the xAudio.dll version."
    End If
    
    SubClassMessages
    'set user data for the window - don't know why, just in case
    SetWindowLong UserControl.hWnd, GWL_USERDATA, lPlayerHandle
    
    Me.Priority = XA_CONTROL_PRIORITY_NORMAL
    
    tbPlayer.Buttons(btnStart).Enabled = False
    tbPlayer.Width = twipsX(123)
'''    cbMain.Bands(2).MinWidth = tbPlayer.Width
''    Set cbMain.Bands(2).Child = UserControl.tbPlayer
'    Set cbMain.Bands(1).Child = UserControl.slidePosition
    
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    DroppedFiles Data
End Sub

Private Sub UserControl_Resize()
'    MsgBox "player resize"
    On Error Resume Next
    Static bResizing As Boolean
    If Not bResizing And Not bHidden Then
        bResizing = True
'''        With cbMain
'''            .Align = vbAlignNone
'''            .Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
'''            If ((.Width + twipsX(2)) <> UserControl.Width) Then
'''                UserControl.Width = .Width
'''            End If
'''            If ((.Height + twipsY(2)) <> UserControl.Height) Then
'''                UserControl.Height = .Height
'''            End If
'''        End With
        bResizing = False
    End If
End Sub

Private Sub UserControl_Show()
    bHidden = False
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
'    MsgBox "beginning of  terminate"
    
    'SendCommand XA_MSG_COMMAND_EXIT, , , False
    control_message_send_N lPlayerHandle, XA_MSG_COMMAND_EXIT
    player_delete lPlayerHandle
    UnSubClassMessages
    lPlayerHandle = 0
    
End Sub

Private Function QuickStop(Optional ByVal bDoEvents As Boolean = True)
    ResyncPlayer bDoEvents
    If tCurrentMessage.mPlayerState <> XA_PLAYER_STATE_STOPPED Then
        SendCommand XA_MSG_COMMAND_STOP, , , bDoEvents
    End If
    CloseFiles bDoEvents
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property


Public Property Get PlayerState() As XA_PlayerState
    PlayerState = tCurrentMessage.mPlayerState
    
End Property

Public Property Get InputState() As XA_InputOutputState
    InputState = tCurrentMessage.mInputState
    
End Property

Public Property Get OutputState() As XA_InputOutputState
    OutputState = tCurrentMessage.mOutputState
End Property

Public Property Get Priority() As XA_CONTROL_PRIORITY
    Priority = player_get_priority(lPlayerHandle)
    
End Property

Public Property Let Priority(ByVal eNewPriority As XA_CONTROL_PRIORITY)
    player_set_priority lPlayerHandle, eNewPriority
    
End Property

Public Property Get FileName() As String
    FileName = sCurrentFile
End Property

Public Property Let FileName(ByVal sNewFile As String)
    sCurrentFile = sNewFile
    sCurrentShortFile = GetShortName(sNewFile)
    tbPlayer.Buttons(btnStart).Enabled = True
    
End Property

Private Function GetShortName(ByVal sFilename As String)
    On Error Resume Next
    Dim sShort As String
    sShort = String(255, 0)
    Dim lLen As Long
    lLen = GetShortPathName(sFilename, sShort, Len(sShort))
    GetShortName = Left(sShort, lLen)
    
End Function

Public Function Play(Optional ByVal sFilename As String)
    If Not bStarting Then
        bStarting = True
    If sFilename <> "" Then
        Me.FileName = sFilename
    End If
'    If (Me.InputState <> XA_STATE_CLOSED) Or (Me.OutputState <> XA_STATE_CLOSED) Then
'        Me.StopPlayer
'    End If
    Select Case True
        Case Me.PlayerState = XA_PLAYER_STATE_PLAYING
            Me.StopPlayer
        Case Me.PlayerState = XA_PLAYER_STATE_PAUSED And (sCurrentShortFile <> tCurrentMessage.mName)
            Me.StopPlayer
        Case Me.PlayerState = XA_PLAYER_STATE_EOS
            CloseFiles
    End Select
    If Me.InputState <> XA_STATE_OPEN Then
        tCurrentMessage.mPosition.mRange = 0
        eCurrentCommand = XA_MSG_COMMAND_INPUT_OPEN
        If Not SendCommand(eCurrentCommand, 0, 0) Then
            With tCurrentMessage.mError
                .mMessage = "Failed to open input: " & sCurrentFile
                .mCode = -123
            End With
            SendCommand XA_MSG_COMMAND_INPUT_CLOSE
            RaiseEvent Error(tCurrentMessage.mError)
            bStarting = False
            Exit Function
        End If
    End If
'    If tCurrentMessage.mPosition.mRange = 0 Then
'            With tCurrentMessage.mError
'                .mMessage = "Failed to open input (zero duration): " & sCurrentFile
'                .mCode = -123
'            End With
'            SendCommand XA_MSG_COMMAND_INPUT_CLOSE
'            RaiseEvent Error(tCurrentMessage.mError)
'            bStarting = False
'            Exit Function
'    End If
    
    'prevent from pressing start button when player is playing
'    Select Case ReceivedMsg.mPlayerState
'        Case XA_PLAYER_STATE_EOS
'        Case XA_PLAYER_STATE_PAUSED
'            'go to selected position
'            eCurrentCommand = XA_MSG_COMMAND_SEEK
'            SendCommand eCurrentCommand, slPos.Value, slPos.Max
'            eCurrentCommand = XA_MSG_COMMAND_PLAY
'            If SendCommand(eCurrentCommand, 0, 0) = False Then GoTo erh
'            Exit Sub
'        Case XA_PLAYER_STATE_PLAYING
'            Exit Sub
'        Case XA_PLAYER_STATE_STOPPED
'    End Select

    'open output file:
    '   if output file is empty string will play on speakers
    '   if not empty then will encode to specified location
    '   if filename preceeded with "wav:" then will create wav header
    If Me.OutputState <> XA_STATE_OPEN Then
        eCurrentCommand = XA_MSG_COMMAND_OUTPUT_OPEN
        SendCommand eCurrentCommand, 0, 0  ' ) = False Then GoTo erh
    End If
    
    'set output to stereo - not necessary
    eCurrentCommand = XA_MSG_SET_OUTPUT_CHANNELS
    SendCommand eCurrentCommand, XA_OUTPUT_CHANNELS_STEREO, 0  ' ) = False Then GoTo erh

    'go to selected position
    eCurrentCommand = XA_MSG_COMMAND_SEEK
    SendCommand eCurrentCommand, slidePosition.Value, slidePosition.Max
    
    'start playing here
    eCurrentCommand = XA_MSG_COMMAND_PLAY
    If Not SendCommand(eCurrentCommand, 0, 0) Then    ' = False Then GoTo erh
        If tCurrentMessage.mError.mCode <> 0 Then
            Err.Raise tCurrentMessage.mError.mCode, "ctlXaudioPlayer", "Failed to start playing the intended file." & vbCrLf & vbCrLf & tCurrentMessage.mError.mMessage
        End If
    End If
    
'    Me.PushButton btnStart
    
    bPaused = False
'prevState = XA_PLAYER_STATE_PLAYING
'Do 'loop while playing
'    'all messages are accepted by frmPlay
'    'we just show the time and position.
'    lblTimeElapsed = IIf(ReceivedMsg.mTimecode.mH = 0, "", ReceivedMsg.mTimecode.mH & ":") & Format(ReceivedMsg.mTimecode.mM, "00") & ":" & Format(ReceivedMsg.mTimecode.mS, "00") & "." & Format(ReceivedMsg.mTimecode.mF, "0")
'    If ReceivedMsg.mPlayerState <> XA_PLAYER_STATE_PAUSED Then slPos.Value = ReceivedMsg.mPosition.mOffset
'    DoEvents
'Loop While (ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PLAYING) Or (ReceivedMsg.mPlayerState = XA_PLAYER_STATE_PAUSED)
''if here then play was stopped or end of stream was reached
'slPos.Value = 0
'lblTimeElapsed.Caption = "00:00.0"
'lblState = "Stopped"
'Select Case ReceivedMsg.mPlayerState
'Case XA_PLAYER_STATE_PAUSED
'    lblState = "Paused"
'Case XA_PLAYER_STATE_PLAYING
'    lblState = "Playing"
'Case XA_PLAYER_STATE_STOPPED
'    lblState = "Stopped"
'Case XA_PLAYER_STATE_EOS
'    lblState = "End Of File"
'End Select
'
''if were decoding then close the output
'If WAVFile <> "" Then
'    IssuedComm = XA_MSG_COMMAND_OUTPUT_CLOSE
'    SendCommand IssuedComm, 0, 0
'End If

'Exit Sub
        bStarting = False
    End If
    
End Function

Public Function StopPlayer()
    bStopping = True
    
    slidePosition.Value = 0
    
    If Me.PlayerState <> XA_PLAYER_STATE_STOPPED Then
        SendCommand XA_MSG_COMMAND_STOP  ' Then
    End If
'    control_message_send_N lPlayerHandle, XA_MSG_COMMAND_STOP
    CloseFiles
    
    bStopping = False
    
End Function


Private Function CloseFiles(Optional ByVal bDoEvents As Boolean = True)
    ResyncPlayer bDoEvents
    If Me.OutputState <> XA_STATE_CLOSED Then
'        control_message_send_N lPlayerHandle, XA_MSG_COMMAND_OUTPUT_CLOSE
        SendCommand XA_MSG_COMMAND_OUTPUT_CLOSE, , , bDoEvents
    End If
    If Me.InputState <> XA_STATE_CLOSED Then
'        control_message_send_N lPlayerHandle, XA_MSG_COMMAND_INPUT_CLOSE
        SendCommand XA_MSG_COMMAND_INPUT_CLOSE, , , bDoEvents
    End If
    
End Function

Private Function Attach(ByVal lMsg As Long, Optional ByVal bWithOffset As Boolean)
    Dim lHwnd As Long
    lHwnd = UserControl.hWnd
    AttachMessage Me, lHwnd, lMsg + IIf(bWithOffset, lNotifyOffset, 0)
End Function

Private Function Detach(ByVal lMsg As Long, Optional ByVal bWithOffset As Boolean)
    Dim lHwnd As Long
    lHwnd = UserControl.hWnd
    DetachMessage Me, lHwnd, lMsg + IIf(bWithOffset, lNotifyOffset, 0)
    
End Function

Private Function SubClassMessages()
    Attach XA_MSG_COMMAND_INPUT_OPEN
    Attach XA_MSG_COMMAND_INPUT_CLOSE
    Attach XA_MSG_COMMAND_OUTPUT_OPEN
    Attach XA_MSG_COMMAND_OUTPUT_CLOSE
    Attach XA_MSG_COMMAND_OUTPUT_RESET
    Attach XA_MSG_COMMAND_INPUT_SEND_MESSAGE
    Attach XA_MSG_COMMAND_PLAY
    Attach XA_MSG_COMMAND_STOP
    Attach XA_MSG_COMMAND_EXIT
    Attach XA_MSG_COMMAND_PAUSE
    Attach XA_MSG_COMMAND_SEEK
    Attach XA_MSG_COMMAND_SYNC
    Attach XA_MSG_SET_INPUT_POSITION_RANGE
    Attach XA_MSG_GET_INPUT_POSITION_RANGE
    Attach XA_MSG_COMMAND_OUTPUT_MUTE
    Attach XA_MSG_COMMAND_OUTPUT_UNMUTE
    Attach XA_MSG_SET_INPUT_TIMECODE_GRANULARITY
    Attach XA_MSG_GET_INPUT_TIMECODE_GRANULARITY
    Attach XA_MSG_SET_OUTPUT_VOLUME
    Attach XA_MSG_GET_OUTPUT_VOLUME
    Attach XA_MSG_SET_OUTPUT_CHANNELS
    Attach XA_MSG_GET_OUTPUT_CHANNELS
    Attach XA_MSG_SET_CODEC_EQUALIZER
    Attach XA_MSG_GET_CODEC_EQUALIZER
    SubClassAcknowledgements
    
End Function

Private Function UnSubClassMessages()
    
    UnSubClassAcknowledgements
    Detach XA_MSG_COMMAND_INPUT_OPEN
    Detach XA_MSG_COMMAND_INPUT_CLOSE
    Detach XA_MSG_COMMAND_OUTPUT_OPEN
    Detach XA_MSG_COMMAND_OUTPUT_CLOSE
    Detach XA_MSG_COMMAND_OUTPUT_RESET
    Detach XA_MSG_COMMAND_INPUT_SEND_MESSAGE
    Detach XA_MSG_COMMAND_PLAY
    Detach XA_MSG_COMMAND_STOP
    Detach XA_MSG_COMMAND_EXIT
    Detach XA_MSG_COMMAND_PAUSE
    Detach XA_MSG_COMMAND_SEEK
    Detach XA_MSG_COMMAND_SYNC
    Detach XA_MSG_SET_INPUT_POSITION_RANGE
    Detach XA_MSG_GET_INPUT_POSITION_RANGE
    Detach XA_MSG_COMMAND_OUTPUT_MUTE
    Detach XA_MSG_COMMAND_OUTPUT_UNMUTE
    Detach XA_MSG_SET_INPUT_TIMECODE_GRANULARITY
    Detach XA_MSG_GET_INPUT_TIMECODE_GRANULARITY
    Detach XA_MSG_SET_OUTPUT_VOLUME
    Detach XA_MSG_GET_OUTPUT_VOLUME
    Detach XA_MSG_SET_OUTPUT_CHANNELS
    Detach XA_MSG_GET_OUTPUT_CHANNELS
    Detach XA_MSG_SET_CODEC_EQUALIZER
    Detach XA_MSG_GET_CODEC_EQUALIZER
    
End Function

Private Function SubClassAcknowledgements()

    Attach XA_MSG_NOTIFY_ACK, True
    Attach XA_MSG_NOTIFY_NACK, True
    Attach XA_MSG_NOTIFY_READY, True
    Attach XA_MSG_NOTIFY_PLAYER_STATE, True
    Attach XA_MSG_NOTIFY_INPUT_TIMECODE, True
    Attach XA_MSG_NOTIFY_DEBUG, True
    Attach XA_MSG_NOTIFY_INPUT_STATE, True
    Attach XA_MSG_NOTIFY_OUTPUT_STATE, True
    Attach XA_MSG_NOTIFY_ERROR, True
    Attach XA_MSG_NOTIFY_PONG, True
    Attach XA_MSG_NOTIFY_PLAYER_MODE, True
    Attach XA_MSG_NOTIFY_INPUT_NAME, True
    Attach XA_MSG_NOTIFY_INPUT_CAPS, True
    Attach XA_MSG_NOTIFY_INPUT_POSITION, True
    Attach XA_MSG_NOTIFY_INPUT_POSITION_RANGE, True
    Attach XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY, True
    Attach XA_MSG_NOTIFY_STREAM_DURATION, True
    Attach XA_MSG_NOTIFY_STREAM_PARAMETERS, True
    Attach XA_MSG_NOTIFY_STREAM_PROPERTIES, True
    Attach XA_MSG_NOTIFY_STREAM_MIME_TYPE, True
    Attach XA_MSG_NOTIFY_INPUT_MODULE, True
    Attach XA_MSG_NOTIFY_INPUT_MODULE_INFO, True
    Attach XA_MSG_NOTIFY_OUTPUT_NAME, True
    Attach XA_MSG_NOTIFY_OUTPUT_CAPS, True
    Attach XA_MSG_NOTIFY_OUTPUT_VOLUME, True
    Attach XA_MSG_NOTIFY_OUTPUT_BALANCE, True
    Attach XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL, True
    Attach XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL, True
    Attach XA_MSG_NOTIFY_OUTPUT_CHANNELS, True
    Attach XA_MSG_NOTIFY_OUTPUT_MODULE_INFO, True
    Attach XA_MSG_NOTIFY_CODEC_EQUALIZER, True
    Attach XA_MSG_NOTIFY_NOTIFICATION_MASK, True
    Attach XA_MSG_NOTIFY_PROGRESS, True
    
End Function

Private Function UnSubClassAcknowledgements()

    Detach XA_MSG_NOTIFY_ACK, True
    Detach XA_MSG_NOTIFY_NACK, True
    Detach XA_MSG_NOTIFY_READY, True
    Detach XA_MSG_NOTIFY_PLAYER_STATE, True
    Detach XA_MSG_NOTIFY_INPUT_TIMECODE, True
    Detach XA_MSG_NOTIFY_DEBUG, True
    Detach XA_MSG_NOTIFY_INPUT_STATE, True
    Detach XA_MSG_NOTIFY_OUTPUT_STATE, True
    Detach XA_MSG_NOTIFY_ERROR, True
    Detach XA_MSG_NOTIFY_PONG, True
    Detach XA_MSG_NOTIFY_PLAYER_MODE, True
    Detach XA_MSG_NOTIFY_INPUT_NAME, True
    Detach XA_MSG_NOTIFY_INPUT_CAPS, True
    Detach XA_MSG_NOTIFY_INPUT_POSITION, True
    Detach XA_MSG_NOTIFY_INPUT_POSITION_RANGE, True
    Detach XA_MSG_NOTIFY_INPUT_TIMECODE_GRANULARITY, True
    Detach XA_MSG_NOTIFY_STREAM_DURATION, True
    Detach XA_MSG_NOTIFY_STREAM_PARAMETERS, True
    Detach XA_MSG_NOTIFY_STREAM_PROPERTIES, True
    Detach XA_MSG_NOTIFY_STREAM_MIME_TYPE, True
    Detach XA_MSG_NOTIFY_INPUT_MODULE, True
    Detach XA_MSG_NOTIFY_INPUT_MODULE_INFO, True
    Detach XA_MSG_NOTIFY_OUTPUT_NAME, True
    Detach XA_MSG_NOTIFY_OUTPUT_CAPS, True
    Detach XA_MSG_NOTIFY_OUTPUT_VOLUME, True
    Detach XA_MSG_NOTIFY_OUTPUT_BALANCE, True
    Detach XA_MSG_NOTIFY_OUTPUT_PCM_LEVEL, True
    Detach XA_MSG_NOTIFY_OUTPUT_MASTER_LEVEL, True
    Detach XA_MSG_NOTIFY_OUTPUT_CHANNELS, True
    Detach XA_MSG_NOTIFY_OUTPUT_MODULE_INFO, True
    Detach XA_MSG_NOTIFY_CODEC_EQUALIZER, True
    Detach XA_MSG_NOTIFY_NOTIFICATION_MASK, True
    Detach XA_MSG_NOTIFY_PROGRESS, True
    
End Function



Private Function DroppedFiles(ByRef oData)
    If oData.Files.Count > 0 Then
        If oData.Files.Count = 1 Then
            '
            'Me.StopPlayer
            Me.Play oData.Files(1)
        End If
    End If
        
End Function


Private Function ResyncPlayer(Optional ByVal bDoEvents As Boolean = True)
    'QuickCommand XA_MSG_COMMAND_SYNC, bDoEvents
    QuickCommand XA_MSG_COMMAND_PING, bDoEvents

End Function
