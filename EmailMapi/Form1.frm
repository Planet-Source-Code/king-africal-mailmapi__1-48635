VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Browser"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   120
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   7
         Top             =   3960
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   3840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   3240
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ras, res As String
Dim i As Long



Private Sub Command2_Click()
On Error GoTo ema
With CommonDialog1
  .CancelError = True
  .Filter = "All files(*.*)|*.*"
  .DialogTitle = "Open files..."
  .ShowOpen
  If Len(.FileName) = 0 Then Exit Sub
  ras = .FileName
  res = .FileTitle
End With
ema:
Command1.Enabled = True
End Sub


Private Sub Command1_Click()
Dim g As String


With CommonDialog1
  .CancelError = True
  .Filter = "All files(*.*)|*.*"
  .DialogTitle = "Open files..."
  .Filter = "all csv(*.csv)|*.csv|,All txt(*.txt)|*.txt|"
  .ShowOpen
  If Len(.FileName) = 0 Then Exit Sub
  MAPISession1.SignOn
  Open .FileName For Input As #1
  
Do
Line Input #1, g

If InStr(g, "@") <> 0 Then


If MAPISession1.SessionID <> 0 Then
With MAPIMessages1
    .SessionID = MAPISession1.SessionID
       .Action = 6
        .RecipDisplayName = Trim(Text5)
        .RecipAddress = Trim(g)
        .MsgSubject = Trim(Text2)
        .MsgNoteText = Trim(Text3)
        .AttachmentName = res
        .AttachmentPathName = ras
        .AttachmentPosition = 0
        .AttachmentType = 0
        .RecipIndex = 0
        .RecipType = 1
         
         
        .Action = 3
      
    End With
End If
End If


Loop Until EOF(1)

 Close #1
 MAPISession1.SignOff
End With

End Sub

Private Sub Form_Load()
Command1.Enabled = False
End Sub
