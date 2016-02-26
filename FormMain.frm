VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormMain 
   Caption         =   "ByteEdit"
   ClientHeight    =   5115
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7380
   DrawMode        =   6  'Mask Pen Not
   BeginProperty Font 
      Name            =   "FixedSys"
      Size            =   13.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   6360
      Top             =   0
   End
   Begin VB.PictureBox picByteEdit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      DrawMode        =   10  'Mask Pen
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.HScrollBar hscByteEdit 
      Height          =   255
      LargeChange     =   8
      Left            =   0
      Max             =   15
      TabIndex        =   2
      Top             =   3120
      Width           =   6015
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   6360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "すべてのファイル|*.*"
   End
   Begin VB.VScrollBar vscByteEdit 
      Height          =   3135
      LargeChange     =   16
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ファイル(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新規作成(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "開く(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "上書き保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "名前をつけて保存(&A)..."
      End
      Begin VB.Menu mnuFileS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "終了(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "編集(&E)"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "元に戻す(&U)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "切り取り(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "コピー(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "貼り付け(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "削除(&D)"
      End
      Begin VB.Menu mnuEditS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAllSelect 
         Caption         =   "すべて選択(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "検索(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindSel 
         Caption         =   "選択内容を検索"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "次を検索(&N)"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "オプション(&O)"
      Begin VB.Menu mnuOptionDigitsSelect 
         Caption         =   "１行の桁数(&D)"
         Begin VB.Menu mnuOptionDigits 
            Caption         =   "8桁"
            Index           =   8
         End
         Begin VB.Menu mnuOptionDigits 
            Caption         =   "10桁"
            Index           =   10
         End
         Begin VB.Menu mnuOptionDigits 
            Caption         =   "16桁"
            Index           =   16
         End
         Begin VB.Menu mnuOptionDigits 
            Caption         =   "20桁"
            Index           =   20
         End
         Begin VB.Menu mnuOptionDigits 
            Caption         =   "32桁"
            Index           =   32
         End
      End
      Begin VB.Menu mnuOptionPausesSelect 
         Caption         =   "区切り(&P)"
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "4桁"
            Index           =   4
         End
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "5桁"
            Index           =   5
         End
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "8桁"
            Index           =   8
         End
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "10桁"
            Index           =   10
         End
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "16桁"
            Index           =   16
         End
         Begin VB.Menu mnuOptionPauses 
            Caption         =   "なし"
            Index           =   100
         End
      End
      Begin VB.Menu mnuOptionViewTypeSelect 
         Caption         =   "データ表示形式(&T)"
         Begin VB.Menu mnuOptionViewType 
            Caption         =   "2進数"
            Index           =   2
         End
         Begin VB.Menu mnuOptionViewType 
            Caption         =   "8進数"
            Index           =   8
         End
         Begin VB.Menu mnuOptionViewType 
            Caption         =   "10進数"
            Index           =   10
         End
         Begin VB.Menu mnuOptionViewType 
            Caption         =   "16進数"
            Index           =   16
         End
      End
      Begin VB.Menu mnuOptionGuideTypeSelect 
         Caption         =   "案内表示形式(&G)"
         Begin VB.Menu mnuOptionGuideType 
            Caption         =   "8進数"
            Index           =   8
         End
         Begin VB.Menu mnuOptionGuideType 
            Caption         =   "10進数"
            Index           =   10
         End
         Begin VB.Menu mnuOptionGuideType 
            Caption         =   "16進数"
            Index           =   16
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "ヘルプ(&H)"
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "目次(&I)"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "トピックの検索(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpFind 
         Caption         =   "キーワードで検索(&K)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "バージョン情報(&A)"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API宣言
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Byte, Source As Byte, ByVal Length As Long)

'定数もどき
Dim ScrollWidth As Single, ScrollHeight As Single
Dim tWidth As Single, tHeight As Single

'設定とか
Dim Digits As Integer
Dim Pauses As Integer
Dim NumType As Integer
Dim NumLength As Integer
Dim GuideType As Integer
Dim Blinking As Boolean

'データ
Dim Data() As Byte
Dim Size As Long
Private Type tagDataInfo
    ScrollX As Long
    ScrollY As Long
    SelStart As Long
    SelEnd As Long
End Type
Dim DataInfo As tagDataInfo
Dim Editing As String
Dim IsDirty As Boolean

Dim DataFile As String

Dim MousePressed As Boolean

'クリップボードもどき
Dim ClipData() As Byte
Dim ClipSize As Long

'元に戻す機能
Dim UndoData() As Byte
Dim UndoSize As Long
Dim PrevAction As String

'検索用データ
Dim FindData() As Byte
Dim FindSize As Long
Dim FindDown As Boolean     '下に向かって検索するかどうか
Dim FindWay As Integer      '検索方法(1.FindDataに一致、2.FindDataとの差分、他.何もしない

'スクロール
Dim Scroll As New clsLongScroll

'ホイールマウス用
Dim WithEvents Wheel As clsWheelMouse
Attribute Wheel.VB_VarHelpID = -1


Private Sub InsertData(ByVal Pos As Long, Insert() As Byte)
    Dim iSize As Long
    iSize = UBound(Insert) + 1
    ReDim Preserve Data(Size + iSize - 1)
    If Pos < Size Then MoveMemory Data(Pos + iSize), Data(Pos), Size - Pos
    MoveMemory Data(Pos), Insert(0), iSize
    Size = Size + iSize
End Sub

Private Sub DeleteData(ByVal Pos1 As Long, ByVal Pos2 As Long)
    Dim dStart As Long, dSize As Long
    If Pos1 < Pos2 Then
        dStart = Pos1
        dSize = Pos2 - Pos1
    ElseIf Pos1 > Pos2 Then
        dStart = Pos2
        dSize = Pos1 - Pos2
    Else
        Exit Sub
    End If
    If dStart + dSize < Size Then MoveMemory Data(dStart), Data(dStart + dSize), Size - dStart - dSize
    If Size - dSize - 1 >= 0 Then ReDim Preserve Data(Size - dSize - 1)
    Size = Size - dSize
End Sub

Private Sub ReplaceData(ByVal Pos1 As Long, ByVal Pos2 As Long, Insert() As Byte)
    DeleteData Pos1, Pos2
    If Pos1 <= Pos2 Then
        InsertData Pos1, Insert
    Else
        InsertData Pos2, Insert
    End If
End Sub

Private Sub EndEditing()
    If Editing <> "" Then
        With DataInfo
            Data(.SelEnd) = CByte(IIf(Dec(Editing, NumType) < 256, Dec(Editing, NumType), 255))
            Editing = ""
            .SelEnd = .SelEnd + 1
            .SelStart = .SelEnd
        End With
        Scroll.Max = Size \ Digits
    End If
End Sub

Private Function Enc(ByVal Number As Long, ByVal Length As Integer, ByVal Mode As Integer) As String
    Select Case Mode
    Case 2
        Dim i As Integer
        Enc = ""
        For i = 0 To 7
            Enc = CStr(-((Number And 2 ^ i) <> 0)) & Enc
        Next
    Case 8
        Enc = Oct$(Number)
    Case 10
        Enc = CStr(Number)
    Case 16
        Enc = Hex$(Number)
    End Select
    Enc = Right$(String(Length - 1, "0") & Enc, Length)
End Function

Private Function Dec(Number As String, ByVal Mode As Integer) As Long
    Select Case Mode
    Case 2
        Dim i As Integer
        Dec = 0
        For i = 0 To Len(Number) - 1
            Dec = Dec + CLng(Mid$(Number, Len(Number) - i, 1)) * 2 ^ i
        Next
    Case 8
        Dec = CLng("&O" & Number)
    Case 10
        Dec = CLng(Number)
    Case 16
        Dec = CLng("&H" & Number)
    End Select
End Function

Private Sub OnDraw()
    Dim Msg As String
    Dim X As Long, Y As Long
    Dim Lines As Integer
    Dim StartY As Long, EndY As Long
    With picByteEdit
        .Cls
        .CurrentX = 0
        .CurrentY = 0
        Lines = (ScaleHeight - ScrollHeight) \ tHeight
        StartY = DataInfo.ScrollY
        EndY = DataInfo.ScrollY + Lines - 1

        '数値表示
        For Y = StartY To EndY
            Msg = Msg & Enc(Y * Digits, 8, GuideType) & "  "
            For X = DataInfo.ScrollX To Digits - 1
                If Y * Digits + X < Size Then
                    If Editing = "" Or DataInfo.SelEnd <> Y * Digits + X Then
                        Msg = Msg & Enc(Data(Y * Digits + X), NumLength, NumType)
                    Else
                        Msg = Msg & Right$("       " & Editing, NumLength)
                    End If
                    Msg = Msg & IIf(X Mod Pauses = Pauses - 1, "  ", " ")
                End If
            Next
            Msg = Msg & vbNewLine
        Next
        picByteEdit.Print Msg

        'カーソル
        If Blinking Then
            Dim CursorX As Long, CursorY As Long
            CursorX = DataInfo.SelEnd Mod Digits
            CursorY = DataInfo.SelEnd \ Digits
            If CursorX >= DataInfo.ScrollX Then
                If CursorY >= StartY Then
                    If CursorY <= EndY Then
                        Dim CVX As Long, CVY As Long
                        CVX = CursorX - DataInfo.ScrollX
                        CVX = CVX * (NumLength + 1) + (CursorX \ Pauses) - (DataInfo.ScrollX \ Pauses)
                        CVX = CVX + 10
                        If Editing <> "" Then CVX = CVX + NumLength
                        CVY = CursorY - StartY
                        picByteEdit.Line (CVX * tWidth, CVY * tHeight)-(CVX * tWidth, (CVY + 1) * tHeight)
                    End If
                End If
            End If
        End If

        '選択範囲
        Dim Earlier As Long, Latter As Long
        If DataInfo.SelStart < DataInfo.SelEnd Then
            Earlier = DataInfo.SelStart
            Latter = DataInfo.SelEnd
        Else
            Earlier = DataInfo.SelEnd
            Latter = DataInfo.SelStart
        End If
        Dim X1 As Long, X2 As Long  '開始・終了位置
        For Y = 0 To Lines - 1
            X1 = Earlier - (StartY + Y) * Digits: X2 = Latter - (StartY + Y) * Digits
            If X1 < DataInfo.ScrollX Then X1 = DataInfo.ScrollX Else If X1 > Digits Then X1 = Digits
            If X2 < DataInfo.ScrollX Then X2 = DataInfo.ScrollX Else If X2 > Digits Then X2 = Digits
            If X1 < X2 Then     '選択範囲がある場合
                X1 = (X1 - DataInfo.ScrollX) * (NumLength + 1) + (X1 \ Pauses) - (DataInfo.ScrollX \ Pauses) + 10
                'X2 = X2 + 1
                X2 = (X2 - DataInfo.ScrollX) * (NumLength + 1) + (X2 \ Pauses) - (DataInfo.ScrollX \ Pauses) + 10
                picByteEdit.Line (X1 * tWidth, Y * tHeight)-(X2 * tWidth, (Y + 1) * tHeight), , BF
            End If
        Next
    End With
End Sub

Private Sub Form_Load()
    ScrollWidth = vscByteEdit.Width:    ScrollHeight = hscByteEdit.Height
    tWidth = TextWidth(" "):            tHeight = TextHeight(" ")
    Scroll.ScrollBar = vscByteEdit
    FindDown = True
    Digits = 16
    mnuOptionDigits(Digits).Checked = True
    Pauses = 8
    mnuOptionPauses(Pauses).Checked = True
    hscByteEdit.Max = Digits - Pauses
    NumType = 16
    NumLength = 2
    mnuOptionViewType(NumType).Checked = True
    GuideType = 16
    mnuOptionGuideType(GuideType).Checked = True
    mnuHelpIndex.Enabled = Dir(Replace(App.Path & "\ByteEdit.chm", "\\", "\")) <> ""
    If mnuHelpIndex.Enabled Then
        App.HelpFile = Replace(App.Path & "\ByteEdit.chm", "\\", "\")
    End If
    Set Wheel = New clsWheelMouse
    Wheel.Initialize picByteEdit.hWnd

    If Command$ = "" Then
        mnuFileNew_Click
    Else
        Dim N As Integer
        DataFile = Replace(Command$, """", "")
        Caption = App.Title & " - [" & Dir$(DataFile) & "]"
        N = FreeFile
        EndEditing
        If Dir(DataFile) = "" Then
            ReDim Data(0)
            Size = 0
        Else
            Open DataFile For Binary As #N
                Size = LOF(N)
                If Size = 0 Then
                    ReDim Data(0)
                Else
                    ReDim Data(Size - 1)
                    Get #N, , Data
                End If
            Close
        End If
        With DataInfo
            .ScrollX = 0
            .ScrollY = 0
            .SelStart = 0
            .SelEnd = 0
        End With
        Scroll.Max = Size \ Digits
        OnDraw
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        If ScaleHeight >= tHeight + hscByteEdit.Height Then
            vscByteEdit.Left = ScaleWidth - ScrollWidth
            vscByteEdit.Height = ScaleHeight - ScrollHeight
            hscByteEdit.Top = ScaleHeight - ScrollHeight
            hscByteEdit.Width = ScaleWidth - ScrollWidth
            picByteEdit.Move 0, 0, vscByteEdit.Left, hscByteEdit.Top
        Else
            Height = Height - ScaleHeight + tHeight + hscByteEdit.Height
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not CloseData
    If Not Cancel Then
        Wheel.Terminate
        Set Wheel = Nothing
    End If
End Sub

Private Sub hscByteEdit_Change()
    DataInfo.ScrollX = hscByteEdit.Value
    OnDraw
End Sub

Private Sub hscByteEdit_GotFocus()
    picByteEdit.SetFocus
End Sub

Private Sub hscByteEdit_Scroll()
    hscByteEdit_Change
End Sub

Private Sub mnuEditAllSelect_Click()
    DataInfo.SelStart = 0
    DataInfo.SelEnd = Size
    OnDraw
End Sub

Private Sub mnuEditCopy_Click()
    CopyData DataInfo.SelStart, DataInfo.SelEnd
End Sub

Private Sub mnuEditCut_Click()
    mnuEditCopy_Click
    mnuEditDelete_Click
End Sub

Private Sub mnuEditDelete_Click()
    With DataInfo
        UpdateUndoBuf "Delete"
        DeleteData .SelStart, .SelEnd
        If .SelStart < .SelEnd Then .SelEnd = .SelStart Else .SelStart = .SelEnd
        Scroll.Max = Size \ Digits
        IsDirty = True
        If .SelEnd \ Digits < .ScrollY Then Scroll.Value = .SelEnd \ Digits
        If .SelEnd \ Digits >= .ScrollY + (ScaleHeight - hscByteEdit.Height) \ tHeight Then Scroll.Value = .SelEnd \ Digits - (ScaleHeight - hscByteEdit.Height) \ tHeight + 1
        Blinking = True
        tmrBlink.Enabled = False
        tmrBlink.Enabled = True
    End With
    OnDraw
End Sub

Private Sub mnuEditFind_Click()
    FormSearch.Show vbModal, Me
End Sub

Private Sub mnuEditFindNext_Click()
    FindNext
End Sub

Private Sub mnuEditFindSel_Click()
    Dim dStart As Long
    If DataInfo.SelStart < DataInfo.SelEnd Then
        dStart = DataInfo.SelStart
        FindSize = DataInfo.SelEnd - DataInfo.SelStart
    ElseIf DataInfo.SelStart > DataInfo.SelEnd Then
        dStart = DataInfo.SelEnd
        FindSize = DataInfo.SelStart - DataInfo.SelEnd
    Else
        Exit Sub
    End If
    FindWay = 1
    ReDim FindData(FindSize - 1)
    MoveMemory FindData(0), Data(dStart), FindSize
    FindNext
End Sub

Private Sub mnuEditPaste_Click()
    With DataInfo
        If ClipSize Then
            UpdateUndoBuf "Paste"
            Dim NewSel As Long
            If .SelStart <= .SelEnd Then
                NewSel = .SelStart + ClipSize
            Else
                NewSel = .SelEnd + ClipSize
            End If
            ReplaceData .SelStart, .SelEnd, ClipData
            .SelStart = NewSel
            .SelEnd = NewSel
            Scroll.Max = Size \ Digits
            IsDirty = True
            If .SelEnd \ Digits < .ScrollY Then Scroll.Value = .SelEnd \ Digits
            If .SelEnd \ Digits >= .ScrollY + (ScaleHeight - hscByteEdit.Height) \ tHeight Then Scroll.Value = .SelEnd \ Digits - (ScaleHeight - hscByteEdit.Height) \ tHeight + 1
            Blinking = True
            tmrBlink.Enabled = False
            tmrBlink.Enabled = True
        End If
    End With
    OnDraw
End Sub

Private Sub mnuEditUndo_Click()
    If Len(PrevAction) Then
        EndEditing
        Dim Buf() As Byte, BufSize As Long
        Buf = Data
        BufSize = Size
        Data = UndoData
        Size = UndoSize
        UndoData = Buf
        UndoSize = BufSize
        If DataInfo.SelStart >= Size Then DataInfo.SelStart = Size
        If DataInfo.SelEnd >= Size Then DataInfo.SelEnd = Size
        OnDraw
        IsDirty = True
        PrevAction = "Undo"
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    If Not CloseData Then Exit Sub
    EndEditing
    ReDim Data(0)
    Size = 0
    Scroll.Max = 0
    Scroll.Value = 0
    hscByteEdit.Value = 0
    With DataInfo
        .SelStart = 0
        .SelEnd = 0
    End With
    DataFile = ""
    Caption = App.Title & " - [無題]"
    OnDraw
    IsDirty = False
    PrevAction = ""
End Sub

Private Sub mnuFileOpen_Click()
    If Not CloseData Then Exit Sub
    On Error Resume Next
    With dlgFile
        .DialogTitle = "ファイルを開く"
        .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly
        .ShowOpen
        If Err.Number Then Exit Sub

        Dim N As Integer
        DataFile = .FileName
        Caption = App.Title & " - [" & Dir$(DataFile) & "]"
        N = FreeFile
        EndEditing
        If Dir(DataFile) = "" Then
            ReDim Data(0)
            Size = 0
        Else
            Open DataFile For Binary As #N
                Size = LOF(N)
                If Size = 0 Then
                    ReDim Data(0)
                Else
                    ReDim Data(Size - 1)
                    Get #N, , Data
                End If
            Close
        End If
    End With
    With DataInfo
        .ScrollX = 0
        .ScrollY = 0
        .SelStart = 0
        .SelEnd = 0
    End With
    Scroll.Max = Size \ Digits
    OnDraw
    IsDirty = False
    PrevAction = ""
End Sub

Private Sub mnuFileSave_Click()
    If DataFile = "" Then
        mnuFileSaveAs_Click
        Exit Sub
    End If

    Dim N As Integer
    N = FreeFile
    EndEditing
    If Size Then
        Open DataFile For Output As #N
        Close
        Open DataFile For Binary Access Write Lock Write As #N
            Put #N, , Data
        Close
    Else
        Open DataFile For Output As #N
        Close
    End If
    IsDirty = False
    PrevAction = "Save"
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error Resume Next
    With dlgFile
        .DialogTitle = "名前を付けて保存"
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        .ShowSave
        If Err.Number Then Exit Sub

        DataFile = .FileName
        Caption = App.Title & " - [" & Dir$(DataFile) & "]"
        mnuFileSave_Click
    End With
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "ByteEdit ver 0.10" & vbNewLine & vbNewLine & "        Copyright Mifumi", vbInformation
End Sub

Private Sub mnuHelpIndex_Click()
    SendKeys "{F1}"
End Sub

Private Sub mnuOptionDigits_Click(Index As Integer)
    mnuOptionDigits(Digits).Checked = False
    Digits = Index
    mnuOptionDigits(Digits).Checked = True
    hscByteEdit.Max = IIf(Digits - Pauses > 0, Digits - Pauses, 0)
    Scroll.Max = Size \ Digits
    OnDraw
End Sub

Private Sub mnuOptionGuideType_Click(Index As Integer)
    mnuOptionGuideType(GuideType).Checked = False
    GuideType = Index
    mnuOptionGuideType(GuideType).Checked = True
    OnDraw
End Sub

Private Sub mnuOptionPauses_Click(Index As Integer)
    mnuOptionPauses(Pauses).Checked = False
    Pauses = Index
    mnuOptionPauses(Pauses).Checked = True
    hscByteEdit.Max = IIf(Digits - Pauses > 0, Digits - Pauses, 0)
    hscByteEdit.LargeChange = Pauses
    OnDraw
End Sub

Private Sub mnuOptionViewType_Click(Index As Integer)
    mnuOptionViewType(NumType).Checked = False
    NumType = Index
    NumLength = Switch(NumType = 2, 8, NumType = 8, 3, NumType = 10, 3, NumType = 16, 2)
    mnuOptionViewType(NumType).Checked = True
    OnDraw
End Sub

Private Sub picByteEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift And 2 Then Exit Sub        'Ctrl+Cを使うため
    Dim MoveCursor As Boolean
    With DataInfo
        Select Case KeyCode
        Case vbKeyLeft                  '左移動
            EndEditing
            If .SelEnd >= 1 Then
                .SelEnd = .SelEnd - 1
            End If
            MoveCursor = True
        Case vbKeyRight                 '右移動
            EndEditing
            If .SelEnd <= Size - 1 Then
                .SelEnd = .SelEnd + 1
            End If
            MoveCursor = True
        Case vbKeyUp                    '上移動
            EndEditing
            If .SelEnd >= Digits Then
                .SelEnd = .SelEnd - Digits
            End If
            MoveCursor = True
        Case vbKeyDown                  '下移動
            EndEditing
            If .SelEnd <= Size - Digits Then
                .SelEnd = .SelEnd + Digits
            End If
            MoveCursor = True
        Case vbKeyBack                  '左削除
            If .SelStart = .SelEnd Then
                EndEditing
                If .SelEnd > 0 Then
                    UpdateUndoBuf "Delete"
                    DeleteData .SelEnd - 1, .SelEnd
                    .SelEnd = .SelEnd - 1
                End If
            Else
                UpdateUndoBuf "Delete"
                DeleteData .SelStart, .SelEnd
                If .SelStart < .SelEnd Then .SelEnd = .SelStart Else .SelStart = .SelEnd
            End If
            Scroll.Max = Size \ Digits
            MoveCursor = True
            IsDirty = True
        Case vbKeyDelete                '右削除
            If .SelStart = .SelEnd Then
                EndEditing
                If .SelEnd < Size Then
                    UpdateUndoBuf "Delete"
                    DeleteData .SelEnd, .SelEnd + 1
                End If
            Else
                UpdateUndoBuf "Delete"
                DeleteData .SelStart, .SelEnd
                If .SelStart < .SelEnd Then .SelEnd = .SelStart Else .SelStart = .SelEnd
            End If
            Scroll.Max = Size \ Digits
            MoveCursor = True
            IsDirty = True
        Case vbKey0 To vbKey9, vbKeyA To vbKeyF, vbKeyNumpad0 To vbKeyNumpad9
                                        '数値入力
            Dim Num As String
            If vbKey0 <= KeyCode And KeyCode <= vbKey9 Then
                Num = KeyCode - vbKey0
            ElseIf vbKeyA <= KeyCode And KeyCode <= vbKeyF Then
                Num = Chr(KeyCode)
            Else
                Num = KeyCode - vbKeyNumpad0
            End If
            If CLng("&H" & Num) >= NumType Then Exit Sub
            UpdateUndoBuf "Input"
            If Editing = "" Then
                Dim a(0) As Byte
                a(0) = 0
                ReplaceData .SelStart, .SelEnd, a
                If .SelStart < .SelEnd Then .SelEnd = .SelStart Else .SelStart = .SelEnd
            End If
            Editing = Editing & Num
            If Len(Editing) = NumLength Then EndEditing
            MoveCursor = True
            IsDirty = True
        Case vbKeyReturn
            EndEditing
        End Select
        If MoveCursor Then
            If (Shift And 1) = 0 Then .SelStart = .SelEnd
            If .SelEnd \ Digits < .ScrollY Then Scroll.Value = .SelEnd \ Digits
            If .SelEnd \ Digits >= .ScrollY + (ScaleHeight - hscByteEdit.Height) \ tHeight Then Scroll.Value = .SelEnd \ Digits - (ScaleHeight - hscByteEdit.Height) \ tHeight + 1
            .ScrollY = Scroll.Value
            Blinking = True
            tmrBlink.Enabled = False
            tmrBlink.Enabled = True
        End If
        OnDraw
    End With
End Sub

Private Sub picByteEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        DataInfo.SelEnd = MouseSelPos(X, Y)
        If (Shift And 1) = 0 Then DataInfo.SelStart = DataInfo.SelEnd

        MousePressed = True
        Blinking = True
        tmrBlink.Enabled = False
        tmrBlink.Enabled = True
        OnDraw
    ElseIf Button = 2 Then
        PopupMenu mnuEdit
    End If
End Sub

Private Sub picByteEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MousePressed Then
        DataInfo.SelEnd = MouseSelPos(X, Y)

        Blinking = True
        tmrBlink.Enabled = False
        tmrBlink.Enabled = True
        OnDraw
    End If
End Sub

Private Sub picByteEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePressed = False
End Sub

Private Sub picByteEdit_Resize()
    OnDraw
End Sub

Private Sub tmrBlink_Timer()
    Blinking = Not Blinking
    OnDraw
End Sub

Private Sub vscByteEdit_Change()
    Scroll.Update
    DataInfo.ScrollY = Scroll.Value
    OnDraw
End Sub

Private Sub vscByteEdit_GotFocus()
    picByteEdit.SetFocus
End Sub

Private Sub vscByteEdit_Scroll()
    vscByteEdit_Change
End Sub

Public Function CloseData() As Boolean
    Dim Result As VbMsgBoxResult
    If IsDirty Then
        Result = MsgBox("ファイル内容が変更されています。" & vbNewLine & "保存しますか？", vbYesNoCancel Or vbQuestion)
        Select Case Result
        Case vbYes
            mnuFileSave_Click
            CloseData = Not IsDirty
        Case vbNo
            CloseData = True
        Case vbCancel
            CloseData = False
        End Select
    Else
        CloseData = True
    End If
End Function

Private Sub CopyData(ByVal Pos1 As Long, ByVal Pos2 As Long)
    'DeleteDataんとこからコピー
    Dim dStart As Long, dSize As Long
    If Pos1 < Pos2 Then
        dStart = Pos1
        dSize = Pos2 - Pos1
    ElseIf Pos1 > Pos2 Then
        dStart = Pos2
        dSize = Pos1 - Pos2
    Else
        Exit Sub
    End If
    '↑コピーでした！
    ClipSize = dSize
    ReDim ClipData(ClipSize - 1)
    MoveMemory ClipData(0), Data(dStart), ClipSize
End Sub

Private Sub UpdateUndoBuf(Action As String)
    If PrevAction <> Action Then
        UndoData = Data
        UndoSize = Size
        PrevAction = Action
    End If
End Sub

Public Sub Find(Buf() As Byte, FD As Boolean, Optional FW)
    FindData = Buf
    FindSize = UBound(FindData) + 1
    FindDown = FD
    If Not IsMissing(FW) Then FindWay = FW
    FindNext
End Sub

Private Sub FindNext()
    If FindSize Then
        Dim SelEarlier As Long, SelLater As Long, i As Long, J As Long
        Dim Found As Boolean
        Dim FindLast As Long
        If DataInfo.SelStart > DataInfo.SelEnd Then
            SelEarlier = DataInfo.SelEnd
            SelLater = DataInfo.SelStart
        Else
            SelEarlier = DataInfo.SelStart
            SelLater = DataInfo.SelEnd
        End If
        If FindDown Then
            If FindWay = 1 Then FindLast = Size - FindSize
            If FindWay = 2 Then FindLast = Size - 1
            For i = SelLater To FindLast
                If FindWay = 1 Then
                    If Data(i) = FindData(0) Then
                        Found = True
                        If FindSize > 1 Then
                            For J = 1 To FindSize - 1
                                If Data(i + J) <> FindData(J) Then
                                    Found = False
                                    Exit For
                                End If
                            Next
                        End If
                        If Found Then
                            DataInfo.SelStart = i
                            DataInfo.SelEnd = i + FindSize
                            Exit For
                        End If
                    End If
                ElseIf FindWay = 2 Then
                    If i < FindSize Then
                        If Data(i) <> FindData(i) Then
                            Found = True
                            DataInfo.SelStart = i
                            DataInfo.SelEnd = i + 1
                            Exit For
                        End If
                    Else
                        Found = True
                        DataInfo.SelStart = FindSize
                        DataInfo.SelEnd = Size - 1
                        Exit For
                    End If
                End If
            Next
        Else
            If FindWay = 1 Then FindLast = FindSize - 1
            If FindWay = 2 Then FindLast = 0
            For i = SelEarlier - 1 To FindLast Step -1
                If FindWay = 1 Then
                    If Data(i) = FindData(FindSize - 1) Then
                        Found = True
                        If FindSize > 1 Then
                            For J = 1 To FindSize - 1 Step -1
                                If Data(i - J) <> FindData(FindSize - 1 - J) Then
                                    Found = False
                                    Exit For
                                End If
                            Next
                        End If
                        If Found Then
                            DataInfo.SelStart = i - FindSize + 1
                            DataInfo.SelEnd = i + 1
                            Exit For
                        End If
                    End If
                ElseIf FindWay = 2 Then
                    If i < FindSize Then
                        If Data(i) <> FindData(i) Then
                            Found = True
                            DataInfo.SelStart = i
                            DataInfo.SelEnd = i + 1
                            Exit For
                        End If
                    Else
                        Found = True
                        DataInfo.SelStart = FindSize
                        DataInfo.SelEnd = Size - 1
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    If Not Found Then Beep
    OnDraw
End Sub

'選択位置計算
Public Function MouseSelPos(X As Single, Y As Single) As Long
    Dim XX As Long, YY As Long
    XX = X \ tWidth
    XX = XX - 10
    XX = XX - (XX \ (NumLength + 1) + hscByteEdit.Value) \ Pauses + hscByteEdit.Value \ Pauses
    XX = XX \ (NumLength + 1) + hscByteEdit.Value
    If XX < 0 Then XX = 0
    If XX > Digits Then XX = Digits
    YY = Y \ tHeight
    YY = YY + Scroll.Value
    MouseSelPos = YY * Digits + XX
    If MouseSelPos > Size Then MouseSelPos = Size
    If MouseSelPos < 0 Then MouseSelPos = 0
End Function

Private Sub Wheel_MouseWheel(ByVal iniVector As Integer, ByVal inShift As Integer)
    Dim v As Integer
    v = vscByteEdit.Value - iniVector
    If v < 0 Then v = 0
    If v > vscByteEdit.Max Then v = vscByteEdit.Max
    vscByteEdit.Value = v
    
End Sub

