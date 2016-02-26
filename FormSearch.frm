VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSearch 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "検索"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "FormSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2280
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "すべてのファイル|*.*"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   615
      Left            =   5040
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "検索"
      Default         =   -1  'True
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame fraDirection 
      Caption         =   "検索方向"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
      Begin VB.OptionButton optFindDown 
         Caption         =   "下へ"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optFindUp 
         Caption         =   "上へ"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraFind 
      Caption         =   "検索内容"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton cmdFindDifference 
         Caption         =   "ファイル"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtFindDifference 
         Height          =   270
         Left            =   1080
         TabIndex        =   14
         Top             =   1560
         Width           =   3255
      End
      Begin VB.OptionButton optFindDifference 
         Caption         =   "差分"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.PictureBox picFindBytes 
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   4635
         TabIndex        =   7
         ToolTipText     =   "16進数表記のみ"
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txtFindText 
         Height          =   855
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.Frame fraEncode 
         Caption         =   "エンコード"
         Height          =   855
         Left            =   4440
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton optEncodeShiftJIS 
            Caption         =   "Shift-JIS"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optEncodeUnicode 
            Caption         =   "Unicode"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton optFindBytes 
         Caption         =   "バイト列"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optFindString 
         Caption         =   "文字列"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "FormSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Buf() As Byte
Dim BufSize As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim Find As Integer
    Find = 0
    If optFindString.Value Then
        If Len(txtFindText.Text) Then
            Select Case True
            Case optEncodeUnicode.Value
                Buf = txtFindText.Text
            Case optEncodeShiftJIS.Value
                Buf = StrConv(txtFindText.Text, vbFromUnicode)
            End Select
            Find = 1
        End If
    ElseIf optFindDifference.Value Then
        If txtFindDifference.Text = "" Or Dir(txtFindDifference.Text) = "" Then
        Else
            Dim N As Integer
            N = FreeFile
            Open txtFindDifference.Text For Binary As #N
                BufSize = LOF(N)
                If BufSize = 0 Then
                Else
                    ReDim Buf(BufSize - 1)
                    Get #N, , Buf
                    Find = 2
                End If
            Close
        End If
    End If
    If Find Then
        FormMain.Find Buf, optFindDown.Value, Find
        Unload Me
    End If
End Sub

Private Sub cmdFindDifference_Click()
    On Error Resume Next
    With dlgFile
        .DialogTitle = "ファイルを開く"
        .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly
        .ShowOpen
        If Err.Number Then Exit Sub

        txtFindDifference.Text = .FileName
    End With
End Sub

Private Sub Form_Activate()
    txtFindText.SetFocus
End Sub

Private Sub optFindBytes_Click()
    txtFindText.Enabled = False
    fraEncode.Enabled = False
    optEncodeShiftJIS.Enabled = False
    optEncodeUnicode.Enabled = False
    picFindBytes.Enabled = True
End Sub

Private Sub optFindString_Click()
    txtFindText.Enabled = True
    fraEncode.Enabled = True
    optEncodeShiftJIS.Enabled = True
    optEncodeUnicode.Enabled = True
    picFindBytes.Enabled = False
End Sub
