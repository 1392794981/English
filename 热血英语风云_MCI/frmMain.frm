VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   Caption         =   "热血"
   ClientHeight    =   7005
   ClientLeft      =   480
   ClientTop       =   1170
   ClientWidth     =   12525
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   12525
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdClearDividePoint 
      Caption         =   "清空断点"
      Height          =   375
      Left            =   6120
      TabIndex        =   54
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePointTwo 
      Caption         =   "存2"
      Height          =   360
      Left            =   4560
      TabIndex        =   53
      Top             =   960
      Width           =   675
   End
   Begin VB.CommandButton cmdOpenPointTwo 
      Caption         =   "开2"
      Height          =   360
      Left            =   5280
      TabIndex        =   52
      Top             =   960
      Width           =   675
   End
   Begin VB.CommandButton cmdLoadLrc 
      Caption         =   "载入LRC"
      Height          =   375
      Left            =   6120
      TabIndex        =   51
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   10320
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2520
      Top             =   5880
   End
   Begin VB.TextBox txtRate 
      Height          =   375
      Left            =   6720
      TabIndex        =   49
      Text            =   "0.5"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdCarefulListen 
      Caption         =   "细听"
      Height          =   375
      Left            =   6120
      TabIndex        =   48
      Top             =   1440
      Width           =   585
   End
   Begin VB.CheckBox chkAddTime 
      Caption         =   "加时"
      Height          =   375
      Left            =   9720
      TabIndex        =   47
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtAddTime 
      Height          =   270
      Left            =   10080
      TabIndex        =   46
      Text            =   "3"
      Top             =   1080
      Width           =   345
   End
   Begin VB.CommandButton cmdOpenPointThree 
      Caption         =   "开3"
      Height          =   360
      Left            =   5280
      TabIndex        =   44
      Top             =   1440
      Width           =   675
   End
   Begin VB.CommandButton cmdSavePointThree 
      Caption         =   "存3"
      Height          =   360
      Left            =   4560
      TabIndex        =   43
      Top             =   1440
      Width           =   675
   End
   Begin VB.CommandButton cmdOpenPointOne 
      Caption         =   "开1"
      Height          =   360
      Left            =   5280
      TabIndex        =   42
      Top             =   480
      Width           =   675
   End
   Begin VB.CommandButton cmdSavePointOne 
      Caption         =   "存1"
      Height          =   360
      Left            =   4560
      TabIndex        =   41
      Top             =   480
      Width           =   675
   End
   Begin VB.CommandButton cmdPDF 
      Caption         =   "PDF"
      Height          =   360
      Left            =   9720
      TabIndex        =   40
      Top             =   120
      Width           =   705
   End
   Begin VB.CommandButton cmdQuickCompare 
      Caption         =   "速比"
      Height          =   360
      Left            =   8205
      TabIndex        =   39
      Top             =   600
      Width           =   705
   End
   Begin VB.CheckBox chkNote 
      Caption         =   "显示笔记"
      Height          =   375
      Left            =   9240
      TabIndex        =   38
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdCharOffSet 
      Caption         =   "字距"
      Height          =   255
      Left            =   7440
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txtCharOffSet 
      Height          =   270
      Left            =   6960
      TabIndex        =   35
      Text            =   "80"
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton myFont 
      Caption         =   "字体"
      Height          =   255
      Left            =   8160
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   585
   End
   Begin MCI.MMControl mciOnlyPlayRecord 
      Height          =   330
      Left            =   9720
      TabIndex        =   33
      Top             =   4200
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CheckBox chkCompare 
      Caption         =   "对比"
      Height          =   375
      Left            =   8520
      TabIndex        =   32
      Top             =   1440
      Width           =   735
   End
   Begin VB.Timer tmrRecord 
      Interval        =   100
      Left            =   12840
      Top             =   1200
   End
   Begin MCI.MMControl mciRecord 
      Height          =   330
      Left            =   9840
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdRecordStop 
      Caption         =   "录音停"
      Height          =   360
      Left            =   10335
      TabIndex        =   28
      Top             =   3255
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdRecordPlay 
      Caption         =   "录音放"
      Height          =   360
      Left            =   10335
      TabIndex        =   27
      Top             =   3645
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdRecord 
      Caption         =   "录音"
      Height          =   360
      Left            =   10335
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picSetColor 
      BackColor       =   &H8000000E&
      Height          =   200
      Left            =   9600
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000FF&
      Height          =   245
      Left            =   8880
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   200
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   9120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "颜色"
      Height          =   255
      Left            =   9360
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   580
   End
   Begin RichTextLib.RichTextBox txtTRY 
      Height          =   1935
      Left            =   4560
      TabIndex        =   21
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSaveText 
      Caption         =   "保存"
      Height          =   360
      Left            =   8955
      TabIndex        =   20
      Top             =   120
      Width           =   705
   End
   Begin VB.TextBox txtPauseTime 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9360
      TabIndex        =   19
      Text            =   "60"
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkPauseTime 
      Caption         =   "读完停顿"
      Height          =   255
      Left            =   7440
      TabIndex        =   17
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwLesson2 
      Height          =   1695
      Left            =   10560
      TabIndex        =   16
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483647
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDict 
      BackColor       =   &H80000003&
      Caption         =   "字典"
      Height          =   360
      Left            =   8205
      TabIndex        =   15
      Top             =   120
      Width           =   705
   End
   Begin VB.TextBox txtNewWord 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000315BA&
      Height          =   1485
      Left            =   11520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdAllText 
      BackColor       =   &H80000003&
      Caption         =   "Word"
      Height          =   360
      Left            =   8955
      TabIndex        =   13
      Top             =   600
      Width           =   705
   End
   Begin VB.CommandButton cmdNewWord 
      BackColor       =   &H80000003&
      Caption         =   "词汇"
      Height          =   360
      Left            =   9840
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdTiming 
      BackColor       =   &H80000003&
      Caption         =   "计时"
      Height          =   360
      Left            =   7440
      TabIndex        =   11
      Top             =   120
      Width           =   705
   End
   Begin VB.CommandButton cmdCompareA 
      BackColor       =   &H80000005&
      Caption         =   "对比"
      Height          =   360
      Left            =   7440
      MaskColor       =   &H80000003&
      TabIndex        =   10
      Top             =   600
      Width           =   705
   End
   Begin VB.CheckBox chkQuickPress 
      Caption         =   "快捷方式"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox txtTRY2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000315BA&
      Height          =   4215
      Left            =   11280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H80000003&
      Caption         =   "暂停"
      Height          =   360
      Left            =   9555
      TabIndex        =   4
      Top             =   3645
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Timer tmrStatus 
      Interval        =   50
      Left            =   4560
      Top             =   120
   End
   Begin VB.CommandButton cmdPlay2 
      BackColor       =   &H80000003&
      Caption         =   "断点"
      Height          =   360
      Left            =   9555
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSComctlLib.ListView lvwDividePoint 
      Height          =   1575
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer tmrPlay 
      Interval        =   50
      Left            =   13560
      Top             =   600
   End
   Begin MSComctlLib.Slider sldSound 
      Height          =   420
      Left            =   2040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   741
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2880
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReRead 
      BackColor       =   &H80000003&
      Caption         =   "播放"
      Height          =   360
      Left            =   9555
      TabIndex        =   3
      Top             =   3255
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "开始播放"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin MCI.MMControl mciRead 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.ListView lvwLesson 
      Height          =   2775
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483647
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6600
      TabIndex        =   55
      Top             =   0
      Width           =   855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   720
      TabIndex        =   50
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2778
      _cy             =   1085
   End
   Begin VB.Label lblAddTime 
      Caption         =   "=原+"
      Height          =   255
      Left            =   9720
      TabIndex        =   45
      Top             =   1125
      Width           =   495
   End
   Begin VB.Label lblRecord 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6000
      TabIndex        =   31
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   975
      Left            =   2160
      TabIndex        =   30
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "<="
      Height          =   135
      Left            =   9120
      TabIndex        =   25
      Top             =   2325
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblPauseTime 
      Caption         =   "3/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuMain 
      Caption         =   "主菜单"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearItem 
         Caption         =   "清空"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图"
      Begin VB.Menu mnuFullScreen 
         Caption         =   "全屏"
      End
      Begin VB.Menu mnuStandardScreen 
         Caption         =   "标准"
      End
   End
   Begin VB.Menu mnuDel 
      Caption         =   "删除"
      Visible         =   0   'False
      Begin VB.Menu mnuEditDelDividePoint 
         Caption         =   "删除一个"
      End
      Begin VB.Menu mnuEditDelAllDividePoint 
         Caption         =   "删除全部"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''
 Private WithEvents hk       As clsRegHotKeys
Attribute hk.VB_VarHelpID = -1

'''''''''''''''''''''''''''''''''

Private strFileName As String
Private blMoveSlider As Boolean
Private blPB As Boolean

Private blSave As Boolean
Private blSaveFileName As String

Private blTimeOverPause As Boolean

Private lngRecordStart As Long
Private lngRecordEnd As Long
Private blRecording As Boolean
Private blRecordPlaying As Boolean

Private intScreenType As Integer
Private Const FULL_SCREEN = 1
Private Const STANDARD_SCREEN = 2
Private lngPauseTime As Long

Public blShowCarefullListenForm As Boolean


Private showEnglishLearnWindows As Boolean



Private Sub ReadPlay()
    If lvwDividePoint.ListItems.Count < 1 Then Exit Sub
    'If mciRead.Command = "pause" Then Exit Sub
    
    Dim myFSO As New FileSystemObject
    If myFSO.FileExists(lvwDividePoint.SelectedItem.SubItems(2)) = False Then
        MsgBox "所在文件不存在……", , "不存在"
        Exit Sub
    End If
    
    mciOnlyPlayRecord.Command = "close"
    
    strFileName = lvwDividePoint.SelectedItem.SubItems(2)
    mciRead.FileName = strFileName
    mciRead.Notify = False
    mciRead.Wait = True
    mciRead.Command = "close"
    mciRead.Command = "open"
    mciRead.From = GetCurrentDividePoint
    mciRead.To = GetNextDividePoint
    mciRead.Notify = True
    mciRead.Wait = False
    mciRead.Command = "play"
    If chkAddTime.Value = Checked Then
        txtPauseTime.Text = CStr(Val(txtAddTime.Text) + ((mciRead.To - mciRead.From) \ 1000))
    End If
End Sub

Private Sub OnlyPlayRecord()
    If lvwDividePoint.ListItems.Count < 1 Then Exit Sub
    'If mcionlyplayrecord.Command = "pause" Then Exit Sub
    
    Dim myFSO As New FileSystemObject
    If myFSO.FileExists(lvwDividePoint.SelectedItem.SubItems(2)) = False Then
        MsgBox "所在文件不存在……", , "不存在"
        Exit Sub
    End If
    
    mciRead.Command = "close"
    
    mciOnlyPlayRecord.FileName = lvwDividePoint.SelectedItem.SubItems(2)
    mciOnlyPlayRecord.Notify = False
    mciOnlyPlayRecord.Wait = True
    mciOnlyPlayRecord.Command = "close"
    mciOnlyPlayRecord.Command = "open"
    mciOnlyPlayRecord.From = GetCurrentDividePoint
    mciOnlyPlayRecord.To = GetNextDividePoint
    mciOnlyPlayRecord.Notify = True
    mciOnlyPlayRecord.Wait = False
    mciOnlyPlayRecord.Command = "play"
End Sub

Private Sub cmdChTextSelect()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strFileName As String
    Dim strContent As String
    Dim strLine As String
    'strfilename=App.Path + "\nce2\" + GetNum(lvwLesson.SelectedItem.Text + ".txt"
    If fso.FileExists(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt")
        ts.ReadLine
        strContent = Replace(Trim(ts.ReadLine), Chr(9), "") + Chr(13)
        strContent = strContent + Trim(ts.ReadLine) + Chr(13)
        ts.ReadLine
        ts.ReadLine
        ts.ReadLine
        ts.ReadLine
        strContent = strContent + Trim(ts.ReadLine) + Chr(13)
        ts.ReadLine
        Do Until ts.AtEndOfStream
            strLine = Trim(ts.ReadLine)
            If strLine = "New words and expressions 生词和短语" Then
                Exit Do
            Else
                strContent = strContent + strLine
            End If
        Loop

        strContent = "New words and expressions 生词和短语"
        Do Until ts.AtEndOfStream
            strLine = Trim(ts.ReadLine)
            If strLine = "参考译文" Then
                Exit Do
            Else
                strContent = strContent + vbCrLf + strLine
            End If
        Loop
        
        strContent = "参考译文"
        Do Until ts.AtEndOfStream
            strLine = Trim(ts.ReadLine)
            strContent = strContent + vbCrLf + strLine
        Loop
        ts.Close
    End If
    
    txtNewWord.Text = strContent
    frmTip.txtNewWord.Text = strContent
End Sub

Private Sub chkCompare_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTRY.SetFocus
End Sub

Private Sub chkNote_Click()
     Call Form_Resize
End Sub

Private Sub chkQuickPress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTRY.SetFocus
End Sub

Private Sub cmdAllText_Click()
    On Error Resume Next

    
    
    Dim fso As New FileSystemObject
    Dim fil As File
    Set fil = fso.GetFile(lvwLesson.SelectedItem.SubItems(1))
    
    Dim strEnTxtFileName As String
    strEnTxtFileName = fil.ParentFolder + "\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt"

    If fso.FileExists(strEnTxtFileName) = False Then
        MsgBox "没有找到英文文本!"
        Exit Sub
    End If
           
'    Dim wrdApp As Word.Application
'    Dim wrdDoc As Word.Document
'
'    Set wrdApp = CreateObject("Word.Application")
'    Set wrdDoc = wrdApp.Documents.Open(strEnTxtFileName)
'
'    wrdApp.Visible = True
End Sub

Private Sub cmdCarefulListen_Click()
    DoEvents
    
    Dim rate As Double
    rate = Val(txtRate.Text)
    blShowCarefullListenForm = True
    
    Call EnglishPlay(GetCurrentDividePoint / 1000, rate)
    frmCarefullListen.Show (0)
End Sub

Public Sub EnglishPlay(cPosition As Single, rate As Double)
    If blShowCarefullListenForm = True Then
        frmMain.mciRead.Notify = False
        frmMain.mciRead.Wait = True
        frmMain.mciRead.Command = "stop"
        'frmMain.mciRead.Command = "close"
        WindowsMediaPlayer1.URL = lvwDividePoint.SelectedItem.SubItems(2)
        WindowsMediaPlayer1.Controls.currentPosition = cPosition
        WindowsMediaPlayer1.settings.rate = rate
        WindowsMediaPlayer1.Controls.play
    End If
End Sub

Private Sub cmdClearDividePoint_Click()
    lvwDividePoint.ListItems.Clear
    
    Dim myListItem As ListItem
    Set myListItem = lvwDividePoint.ListItems.Add
    myListItem.Text = Format(0, "000000000")
    myListItem.SubItems(1) = GetPlayTime_Ex(0)
    myListItem.SubItems(2) = strFileName
        
    myListItem.Selected = True
    
    lvwDividePoint.Refresh
    DoEvents
    
    SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
    Call cmdReRead_Click
End Sub

Private Sub cmdLoadLrc_Click()
    Dim strFileName_mp3, strFileName As String
    strFileName_mp3 = lvwLesson.SelectedItem.SubItems(1)
    strFileName = left(strFileName_mp3, Len(strFileName_mp3) - 4) + ".lrc"
    'MsgBox strFileName
    If Dir(strFileName) <> "" Then '文件存在
        'MsgBox "存在"
        'MsgBox Dir(strFileName)
        
        Call cmdSavePointOne_Click '默认保存在‘存1’
                
        Dim str As String
        Dim intStart, intColon, intEnd, Hours, Minutes, Seconds As Integer
        Dim Msecond As Long
        Dim dblMinutes, dblSecond As Double
        Open strFileName For Input As #1
        lvwDividePoint.ListItems.Clear
        Do While Not EOF(1)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Line Input #1, str
            intStart = InStr(1, str, "[")
            intColon = InStr(1, str, ":")
            intEnd = InStr(1, str, "]")
            dblMinutes = Val(Mid(str, intStart + 1, intColon - intStart - 1))
            dblSecond = Val(Mid(str, intColon + 1, intEnd - intColon - 1))
            Msecond = CLng((dblMinutes * 60 + dblSecond) * 1000)
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim myListItem As ListItem
            Set myListItem = lvwDividePoint.ListItems.Add
            myListItem.Text = Format(Msecond, "000000000")
                        
            Seconds = (Msecond \ 1000) Mod 60
            Minutes = (Msecond \ 60000) Mod 60
            Hours = (Msecond \ 3600000)
            myListItem.SubItems(1) = CStr(Hours) + "小时" + CStr(Minutes) + "分" + CStr(Seconds) + "秒"
            myListItem.SubItems(2) = lvwLesson.SelectedItem.SubItems(1)

            lvwDividePoint.Refresh
            DoEvents
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Loop
        Close #1
    Else
        MsgBox "不存在"
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub mnuClearItem_Click()
    If MsgBox("是否确定清空？", vbOKCancel, "警告") = vbOK Then
        lvwLesson.ListItems.Clear
    End If
End Sub

Private Sub Timer1_Timer()
    If blShowCarefullListenForm = True Then
        Dim PositionFrom As Long
        Dim PositionTo As Long
        Dim rate As Double
        PositionFrom = GetCurrentDividePoint
        PositionTo = GetNextDividePoint
        rate = Val(txtRate.Text)

        On Error Resume Next
    
        If WindowsMediaPlayer1.Controls.currentPosition * 1000 >= (PositionTo) Then
            WindowsMediaPlayer1.Controls.stop
            Unload frmCarefullListen
            blShowCarefullListenForm = False
        End If
        
        frmCarefullListen.Label1.Caption = "×" + Trim(Format(rate, "0.0"))
        frmCarefullListen.ProgressBar1.Min = 0
        frmCarefullListen.ProgressBar1.Max = Int(PositionTo) - Int(PositionFrom)
        frmCarefullListen.ProgressBar1.Value = Int(WindowsMediaPlayer1.Controls.currentPosition * 1000) - Int(PositionFrom)
    End If
End Sub

Private Sub cmdCharOffSet_Click()
    If Val(txtCharOffSet.Text) > 0 Then
        txtTRY.SelCharOffset = Val(txtCharOffSet.Text)
        
        
        
            'txtTRY.SelAlignment = 0
            
        
        txtTRY.SetFocus
    End If
End Sub

Private Sub cmdChText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        Call cmdChTextSelect
        'txtNewWord.Visible = True
        frmTip.Show 0
End Sub

Private Sub cmdChText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
        txtNewWord.Visible = False
        Unload frmTip
        txtTRY.SetFocus
End Sub

Private Sub cmdColor_Click()
    On Error GoTo ErrHandler
    
    dlgFile.ShowColor
    txtTRY.SelColor = dlgFile.Color
    picColor.BackColor = dlgFile.Color
    txtTRY.SetFocus
ErrHandler:
    
End Sub

Private Sub cmdCompareA_Click()
    'MsgBox GetNum(lvwLesson.SelectedItem.Text)
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strFileName As String
    Dim strContent As String
    Dim strLine As String
    If fso.FileExists(App.Path + "\第四册（自编）\" + Trim(lvwLesson.SelectedItem.Text) + ".txt") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\第四册（自编）\" + Trim(lvwLesson.SelectedItem.Text) + ".txt")
        strContent = ts.ReadAll
        frmMain2.strInitA = strContent
        frmMain2.txtA.Text = strContent
    Else
        'strfilename=App.Path + "\nce2\" + GetNum(lvwLesson.SelectedItem.Text + ".txt"
        If fso.FileExists(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt") = True Then
            Set ts = fso.OpenTextFile(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt")
            strContent = Replace(Trim(ts.ReadLine), Chr(9), "") + vbCrLf
            strContent = strContent + Trim(ts.ReadLine) + vbCrLf
            ts.ReadLine
            ts.ReadLine
            ts.ReadLine
            ts.ReadLine
            strContent = strContent + Trim(ts.ReadLine) + vbCrLf
            ts.ReadLine
            Do Until ts.AtEndOfStream
                strLine = Trim(ts.ReadLine)
                If strLine = "New words and expressions 生词和短语" Then
    
                    Exit Do
                Else
                    strContent = strContent + strLine + vbCrLf
                End If
            Loop
    
            frmMain2.strInitA = strContent
            frmMain2.txtA.Text = strContent
    
            ts.Close
        End If
    End If
    frmMain2.strInitB = txtTRY.Text
    frmMain2.txtB.Text = txtTRY.Text
    
    DoEvents
    strCurrentMediaFileName = Trim(lvwLesson.SelectedItem.Text)
    frmMain2.Show 1
End Sub

Private Sub cmdNewWordSelect()

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strFileName As String
    Dim strContent As String
    Dim strLine As String
    'strfilename=App.Path + "\nce2\" + GetNum(lvwLesson.SelectedItem.Text + ".txt"
    'MsgBox App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt"
    If fso.FileExists(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\第四册\" + GetNum(lvwLesson.SelectedItem.Text) + ".txt")
'        ts.ReadLine
'        strContent = Replace(Trim(ts.ReadLine), Chr(9), "") + Chr(13)
'        strContent = strContent + Trim(ts.ReadLine) + Chr(13)
'        ts.ReadLine
'        ts.ReadLine
'        ts.ReadLine
'        ts.ReadLine
'        strContent = strContent + Trim(ts.ReadLine) + Chr(13)
'        ts.ReadLine
        Do Until ts.AtEndOfStream
            strLine = Trim(ts.ReadLine)
            If strLine = "New words and expressions 生词和短语" Then
                Exit Do
            Else
                strContent = strContent + strLine
            End If
        Loop
        
        strContent = "New words and expressions 生词和短语"
        Do Until ts.AtEndOfStream
            strLine = Trim(ts.ReadLine)
            If strLine = "参考译文" Then
                Exit Do
            Else
                strContent = strContent + vbCrLf + strLine
            End If
        Loop
        
        ts.Close
    Else
        MsgBox "不存在文件!"
    End If
    
    
    'strContent = Replace(strContent, Chr(13) + Chr(10) + Chr(13) + Chr(10), Chr(13) + Chr(10))
    strContent = Replace(strContent, Chr(13) + Chr(10) + Chr(13) + Chr(10), "&&**")
    strContent = Replace(strContent, Chr(13) + Chr(10), "    ")
    strContent = Replace(strContent, Chr(9), "")
    strContent = Replace(strContent, "&&**", Chr(13) + Chr(10))
    
'    Dim s As String
'    Dim i As Integer
'    For i = 1 To Len(strContent)
'        s = s + Str(Asc(Mid(strContent, i, 1))) + " "
'    Next
'    txtTRY.Text = s
'    Stop
    
    txtNewWord.Text = strContent
    frmTip.txtNewWord.Text = strContent
End Sub

Private Sub cmdDict_Click()
'    If frmDic.Height > lvwLesson.Height + 500 Then
'        frmDic.Height = lvwLesson.Height + 500
'    End If
    If IsfrmDicExist = False Then
        frmDic.Show 0
        frmDic.txtWord.SetFocus
    Else
        Unload frmDic
    End If
    
End Sub

Private Sub cmdNewWord_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

        Call cmdNewWordSelect
        'txtNewWord.Visible = True
        frmTip.Show 0
        
End Sub

Private Sub cmdNewWord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

        'txtNewWord.Visible = False
        Unload frmTip
        
        txtTRY.SetFocus
End Sub

Private Sub cmdOpenPointOne_Click()
    Call OpenPoint(lvwLesson.SelectedItem.Text + ".one")
End Sub

Private Sub cmdOpenPointThree_Click()
    Call OpenPoint(lvwLesson.SelectedItem.Text + ".three")
End Sub

Private Sub cmdOpenPointTwo_Click()
    Call OpenPoint(lvwLesson.SelectedItem.Text + ".two")
End Sub

Private Sub cmdPause_Click()
    'If blTimeOverPause = True Then Exit Sub
    
    If mciRead.Command <> "pause" Then
        mciRead.Command = "pause"
    Else
        mciRead.Command = "play"
    End If
End Sub

Private Sub cmdPDF_Click()
    MsgBox CStr(mciRead.Length)
    'Shell ("C:\Program Files\TTKN\CAJViewer 7.0\CAJViewer.exe")
End Sub

Private Sub cmdPlay_Click()
       
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    Dim myFSO As New FileSystemObject
    If myFSO.FileExists(lvwLesson.SelectedItem.SubItems(1)) = False Then
        MsgBox "所在文件不存在……", , "不存在"
        Exit Sub
    End If
    
    strFileName = lvwLesson.SelectedItem.SubItems(1)
    mciRead.FileName = strFileName
    mciRead.Notify = False
    mciRead.Wait = True
    mciRead.Command = "close"
    mciRead.Command = "open"
    mciRead.Notify = True
    mciRead.Wait = False
    mciRead.Command = "play"
    
    sldSound.Min = 0
    sldSound.Max = mciRead.Length
    
    lvwDividePoint.ListItems.Clear
    Call cmdPlay2_Click
End Sub

Private Sub cmdPlay2_Click()
    Dim myListItem As ListItem
    Set myListItem = lvwDividePoint.ListItems.Add
    myListItem.Text = Format(mciRead.position, "000000000")
    myListItem.SubItems(1) = GetPlayTime()
    myListItem.SubItems(2) = strFileName
        
    myListItem.Selected = True
    
    lvwDividePoint.Refresh
    DoEvents
    
'    '让选择条出现在当前选择项
'    Dim i As Integer
'    For i = 1 To 100
'        SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_UP, 0
'        If lvwDividePoint.SelectedItem.Text = myListItem.Text Then
'            Exit Sub
'        End If
'    Next
'    For i = 1 To 100
'        SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_DOWN, 0
'    If lvwDividePoint.SelectedItem.Text = myListItem.Text Then
'            Exit Sub
'        End If
'    Next
End Sub

Private Sub cmdQuickCompare_Click()
    IsQuickCompare = True
    Call cmdCompareA_Click
End Sub

Private Sub cmdRecord_Click()
    
    mciRecord.FileName = App.Path + "\录音.wav"
    
    mciRecord.Command = "open"
    mciRecord.Command = "record"
    
    lngRecordStart = mciRecord.position
    
    blRecording = True
End Sub

Private Sub cmdRecordPlay_Click()
    'mciRecord.FileName = App.Path + "\录音.wav"
    'mciRecord.Command = "open"
    mciRecord.From = lngRecordStart
    mciRecord.To = lngRecordEnd
    mciRecord.Command = "play"
    
    blRecordPlaying = True
End Sub

Private Sub cmdRecordStop_Click()

    DoEvents
    'mciRecord.Command = "save"
    lngRecordEnd = mciRecord.position
    'mciRecord.Command = "close"
    mciRecord.Command = "stop"
    blRecording = False
End Sub

Private Sub cmdReRead_Click()
    Call ReadPlay
End Sub

Private Sub cmdTract_Click()
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    Dim myFSO As New FileSystemObject
    If myFSO.FileExists(lvwDividePoint.SelectedItem.SubItems(2)) = False Then
        MsgBox "所在文件不存在……", , "不存在"
        Exit Sub
    End If
    
    Dim intCurrentPosition As Long
    intCurrentPosition = mciRead.position
    strFileName = lvwDividePoint.SelectedItem.SubItems(2)
    mciRead.FileName = strFileName
    mciRead.Notify = False
    mciRead.Wait = True
    mciRead.Command = "close"
    mciRead.Command = "open"
    mciRead.From = GetCurrentDividePoint
    mciRead.To = intCurrentPosition
    mciRead.Notify = True
    mciRead.Wait = False
    mciRead.Command = "play"
End Sub



Private Sub cmdSelectText_Click()
    lvwLesson.Visible = True
    txtNewWord.Visible = False
End Sub

Private Sub cmdSavePointOne_Click()
    Call SavePoint(lvwLesson.SelectedItem.Text + ".one")
End Sub

Private Sub cmdSavePointThree_Click()
    Call SavePoint(lvwLesson.SelectedItem.Text + ".three")
End Sub

Private Sub cmdSavePointTwo_Click()
    Call SavePoint(lvwLesson.SelectedItem.Text + ".two")
End Sub

Private Sub cmdSaveText_Click()
    If MsgBox("在这会覆盖原文本的,是否确定要保存？", vbYesNo) = vbNo Then Exit Sub

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(App.Path + "\第四册（自编）\" + Trim(lvwLesson.SelectedItem.Text) + ".txt", ForWriting, True)
    ts.Write txtTRY.Text
    
    ts.Close
End Sub

Private Sub cmdTiming_Click()
    frmTime.Show 0
End Sub

Private Sub cmdTiming_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtTRY.SetFocus
End Sub

Private Sub Form_Load()
    '皮肤
    Call SkinH_Attach
    
    lvwLesson.ColumnHeaders.Add , , "课文", 3000
    lvwLesson.ColumnHeaders.Add , , "地址", 3000
    lvwLesson2.ColumnHeaders.Add , , "课文", 3000
    lvwLesson2.ColumnHeaders.Add , , "地址", 3000
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim myListItem As ListItem
    If fso.FileExists(App.Path + "\列表.llk") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\列表.llk")
        Do While Not ts.AtEndOfStream
            Set myListItem = lvwLesson.ListItems.Add
            myListItem.Text = ts.ReadLine
            myListItem.SubItems(1) = ts.ReadLine
        Loop
        ts.Close
    End If
        
    If fso.FileExists(App.Path + "\列表2.llk") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\列表2.llk")
        Do While Not ts.AtEndOfStream
            Set myListItem = lvwLesson2.ListItems.Add
            myListItem.Text = ts.ReadLine
            myListItem.SubItems(1) = ts.ReadLine
        Loop
        ts.Close
    End If
    
    If lvwLesson.ListItems.Count >= 1 Then
        'txtTRY.FileName = App.Path + "\文本库\" + lvwLesson.SelectedItem.Text + ".rtf"
    End If
    
    lvwDividePoint.ColumnHeaders.Add , , "断点", 1 '2000
    lvwDividePoint.ColumnHeaders.Add , , "断点时间", 5000
    lvwDividePoint.ColumnHeaders.Add , , "所在文件", 100000
    
    
    blPB = True
    blSave = True
    
    blTimeOverPause = False
    
    txtNewWord.Visible = False
    'Me.WindowState = 2
    
    blRecording = False
    blRecordPlaying = False
    
    intScreenType = STANDARD_SCREEN
    IsQuickCompare = False
    
    
    blShowCarefullListenForm = False
    
    showEnglishLearnWindows = True
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    Set hk = New clsRegHotKeys
    'hk.RegHotKeys Me.hwnd, CtrlKey, vbKeyC, "C"
    hk.RegHotKeys Me.hWnd, 0, vbKeyPageDown, "PageDown"
    hk.RegHotKeys Me.hWnd, 0, vbKeyPageUp, "PageUp"
    
    hk.RegHotKeys Me.hWnd, 0, vbKeyInsert, "Insert"
    'hk.RegHotKeys Me.hwnd, 0, vbKeyDelete, "Delete"
    hk.RegHotKeys Me.hWnd, ShiftKeys.ctrlKey, vbKeyDelete, "Delete"
    
    hk.RegHotKeys Me.hWnd, 0, vbKeyAdd, "Add"
    hk.RegHotKeys Me.hWnd, 0, vbKeySubtract, "Subtract"
    
    hk.RegHotKeys Me.hWnd, 0, vbKeyDivide, "Divide"
    hk.RegHotKeys Me.hWnd, 0, vbKeyMultiply, "Multiply"
    
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyAdd, "Alt_Add"
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeySubtract, "Alt_Subtract"
    
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyDivide, "Alt_Divide"
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyMultiply, "Alt_Multiply"
    
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyReturn, "Alt_Enter"
    
    
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyPageDown, "Alt_PageDown"
    hk.RegHotKeys Me.hWnd, ShiftKeys.altKey, vbKeyPageUp, "Alt_PageUp"
    
    hk.RegHotKeys Me.hWnd, 0, 192, "key~" '192是~键
    
    hk.RegHotKeys Me.hWnd, 0, 12, "key5" '12是小键盘5键
    hk.RegHotKeys Me.hWnd, 0, 9, "keyTab" '9是小键盘Tab键
    
    Me.Show   '这个不能省略，否则窗体无法显示出来！
    
    hk.WaitMsg
    ''''''''''''''''''''''''''''''''''''''''''''''''']
End Sub

''''''''''''''''''注册全局键''''''''''''''''''''''''''''''']

''''''''''''''''''注册全局键''''''''''''''''''''''''''''''']

''''''''''''''''''注册全局键''''''''''''''''''''''''''''''']

''''''''''''''''''注册全局键''''''''''''''''''''''''''''''']

''''''''''''''''''注册全局键''''''''''''''''''''''''''''''']
  Private Sub hk_HotKeysDown(Key As String)
    Dim pos As Single
    Dim rate As Double
    Select Case Key
        Case "Alt_PageDown"
            If blShowCarefullListenForm = True Then
                
                rate = Val(txtRate.Text)
                txtRate.Text = str(rate - 0.1)
                pos = WindowsMediaPlayer1.Controls.currentPosition
                Call EnglishPlay(pos, Val(txtRate.Text))
            End If
        Case "key5"
            Call cmdLoadLrc_Click
        Case "keyTab"
            Call mnuEditDelDividePoint_Click
            lvwDividePoint.Refresh
            
            '后面加的
            SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
            Call cmdReRead_Click
        Case "key~"
            Dim hwndDict, hwndPDF As Long
            Dim hwndWord As Long
            Dim rtTitle As String * 256
            Dim rtClassName As String * 256
                        
            hwndDict = FindWindow("YodaoMainWndClass", vbNullString)
            hwndPDF = FindWindow("classFoxitReader", vbNullString)
            hwndWord = FindWindow("OpusApp", vbNullString)

            If showEnglishLearnWindows = True Then
                ShowWindow hwndDict, SW_HIDE
                ShowWindow hwndPDF, SW_HIDE
                Me.Visible = False
                
                Do While hwndWord <> 0
                    GetClassName hwndWord, rtClassName, 255
                    If InStr(1, Trim(rtClassName), "OpusApp") > 0 Then
                        GetWindowText hwndWord, rtTitle, 255
                        If InStr(1, rtTitle, "WPS") > 0 Then
                            ShowWindow hwndWord, SW_HIDE
                        End If
                    End If
                    hwndWord = GetNextWindow(hwndWord, GW_HWNDNEXT)
                Loop
                
                showEnglishLearnWindows = False
            Else
                ShowWindow hwndDict, SW_SHOW
                ShowWindow hwndPDF, SW_SHOW
                Me.Visible = True
                
                Do While hwndWord <> 0
                    GetClassName hwndWord, rtClassName, 255
                    If InStr(1, Trim(rtClassName), "OpusApp") > 0 Then
                        GetWindowText hwndWord, rtTitle, 255
                        If InStr(1, rtTitle, "WPS") > 0 Then
                            ShowWindow hwndWord, SW_SHOW
                        End If
                    End If
                    hwndWord = GetNextWindow(hwndWord, GW_HWNDNEXT)
                Loop
                SetActiveWindow hwndWord
                SetForegroundWindow hwndWord
                showEnglishLearnWindows = True
            End If
            
        Case "Alt_PageUp"
            If blShowCarefullListenForm = True Then
                rate = Val(txtRate.Text)
                txtRate.Text = str(rate + 0.1)
                pos = WindowsMediaPlayer1.Controls.currentPosition
                Call EnglishPlay(pos, Val(txtRate.Text))
            End If
        Case "Alt_Enter"
            On Error Resume Next
            WindowsMediaPlayer1.Controls.stop
            blShowCarefullListenForm = False
            Unload frmCarefullListen
        Case "Alt_Add"
            If blShowCarefullListenForm = True Then
                If WindowsMediaPlayer1.playState = wmppsPlaying Then
                    WindowsMediaPlayer1.Controls.pause
                Else
                    WindowsMediaPlayer1.Controls.play
                End If
            End If
        Case "Alt_Subtract"
            On Error Resume Next
            WindowsMediaPlayer1.Controls.stop
            blShowCarefullListenForm = False
            Unload frmCarefullListen
            Call cmdCarefulListen_Click
        Case "Alt_Divide"
            If blShowCarefullListenForm = True Then
                pos = WindowsMediaPlayer1.Controls.currentPosition
                If pos - 2 > 0 Then
                    Call EnglishPlay(pos - 2, Val(txtRate.Text))
                Else
                    Call EnglishPlay(0, Val(txtRate.Text))
                End If
            End If
        Case "Alt_Multiply"
            If blShowCarefullListenForm = True Then
                pos = WindowsMediaPlayer1.Controls.currentPosition
                If pos + 2 < WindowsMediaPlayer1.currentMedia.duration Then
                    Call EnglishPlay(pos + 2, Val(txtRate.Text))
                Else
                    Call EnglishPlay(WindowsMediaPlayer1.currentMedia.duration - 1, Val(txtRate.Text))
                End If
            End If
        Case "Divide"
            Call JumpPlay(mciRead.position - 5000)
        Case "Multiply"
            Call JumpPlay(mciRead.position + 5000)
        Case "PageDown"
            SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_DOWN, 0
            Call cmdReRead_Click
            lngPauseTime = 0

Case "PageUp"

SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
Call cmdReRead_Click
lngPauseTime = 0

Case "Insert"

cmdPlay2_Click
SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
Call cmdReRead_Click
lngPauseTime = 0
Case "Delete"
    If chkQuickPress.Value = Checked Then
    
        Call mnuEditDelDividePoint_Click
        lvwDividePoint.Refresh
        
        '后面加的
        SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
        Call cmdReRead_Click
    
    End If
Case "Add"

cmdPause_Click
Case "Subtract"

Call cmdReRead_Click
lngPauseTime = 0
End Select
  End Sub
  
''''''''''''''''''''''''''''''''''''''''''''''''']


Private Sub Form_Resize()
    On Error Resume Next
    If intScreenType = STANDARD_SCREEN Then
        
        If chkNote.Value = Checked Then
            lvwLesson.Height = (Me.ScaleHeight - lvwLesson.top) / 2
            txtNote.top = lvwLesson.Height + lvwLesson.top + 100
            txtNote.Height = lvwLesson.Height - 100
            txtNote.Visible = True
        Else
            txtNote.Visible = False
            lvwLesson.Height = (Me.ScaleHeight - lvwLesson.top)
        End If
        
        lvwDividePoint.Height = (Me.ScaleHeight - lvwDividePoint.top)
        
        'txtTRY.top = 2160
        'txtTRY.left = 2040
        txtTRY.Width = Me.ScaleWidth - txtTRY.left - 100
        txtTRY.Height = Me.ScaleHeight - txtTRY.top - 100
        
        txtCharOffSet.top = 1800
        myFont.top = 1800
    End If
    
    
    If intScreenType = FULL_SCREEN Then
        txtTRY.top = 10
        txtTRY.left = 10
        txtTRY.Width = Me.ScaleWidth - txtTRY.left - 10
        txtTRY.Height = Me.ScaleHeight - txtTRY.top - 10
        txtTRY.ZOrder (0)
        
        txtCharOffSet.top = 0
        txtCharOffSet.ZOrder (0)
        myFont.top = 0
        myFont.ZOrder (0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
'    frmQuit.Show 1
'    If frmQuit.blQuitCancel = True Then
'        Cancel = True
'        Exit Sub
'    End If
    Call SavePoint(lvwLesson.SelectedItem.Text + ".pit")
     
    If lvwLesson.ListItems.Count >= 1 Then
        Call SaveText(lvwLesson.SelectedItem.Text + ".llk")
        Call SaveNote(lvwLesson.SelectedItem.Text + ".llk")
        txtTRY.SaveFile App.Path + "\文本库\" + lvwLesson.SelectedItem.Text + ".rtf"
    End If
       
    On Error Resume Next
    Unload frmMain3
    Unload frmDic
    Unload frmTime
    Unload frmCarefullListen
    
    hk.UnWaitMsg
    hk.UnRegAllHotKeys
    
        
End Sub



Private Sub lvwDividePoint_DblClick()
    If lvwDividePoint.ListItems.Count < 1 Then Exit Sub
    
    Call ReadPlay
End Sub

Private Sub lvwDividePoint_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case KEY_ENTER
'            Call cmdReRead
'        Case KEY_SPACE
'            Call cmdPlay
'        Case KEY_END
'            Call cmdPause
'    End Select
End Sub

Private Sub lvwDividePoint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case MOUSE_LEFT
            'Call mnuSoundPlay_Click
        Case MOUSE_RIGHT
            PopupMenu mnuDel
    End Select
End Sub

Private Sub dblClickPlay()
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    '''''''''''''''''
    If lvwLesson2.ListItems.Count > 5 Then
        lvwLesson2.ListItems.Remove (1)
    End If
    Dim lstItem As ListItem
    Set lstItem = lvwLesson2.ListItems.Add(, , lvwLesson.SelectedItem.Text)
    lstItem.SubItems(1) = lvwLesson.SelectedItem.SubItems(1)
    ''''''''''''''''''
    
    Call cmdPlay_Click
       
End Sub

Private Sub lvwLesson_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
    Shift = 0
End Sub

Private Sub lvwLesson_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub lvwLesson_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
    Shift = 0
End Sub

Private Sub lvwLesson_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        Call SavePoint(lvwLesson.SelectedItem.Text + ".pit")
        Call SaveOnlyText(lvwLesson.SelectedItem.Text + ".llk")
        Call SaveNote(lvwLesson.SelectedItem.Text + ".llk")
        txtTRY.SaveFile App.Path + "\文本库\" + lvwLesson.SelectedItem.Text + ".rtf"
        
    End If
End Sub

Private Sub SaveNote(FileName As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    ''''''''''''''Note_Begin'''''''''''''''''
    If fso.FileExists(App.Path + "\文本库\" + FileName + ".not") = False Then
        Set ts = fso.CreateTextFile(App.Path + "\文本库\" + FileName + ".not")
    Else
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName + ".not", ForWriting)
    End If
    ts.Write txtNote.Text
    ts.Close
    ''''''''''''''Note_End'''''''''''''''''
End Sub

Private Sub SavePoint(FileName As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    ''''''''''''''Point_Begin'''''''''''''''''
    If fso.FileExists(App.Path + "\文本库\" + FileName + ".point") = False Then
        Set ts = fso.CreateTextFile(App.Path + "\文本库\" + FileName + ".point")
    Else
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName + ".point", ForWriting)
    End If
    Dim myItem As ListItem
    Dim i As Integer
    i = 1
    For i = 1 To lvwDividePoint.ListItems.Count
        ts.WriteLine lvwDividePoint.ListItems(i).Text
        ts.WriteLine lvwDividePoint.ListItems(i).SubItems(1)
        ts.WriteLine lvwDividePoint.ListItems(i).SubItems(2)
    Next
    ts.Close
    ''''''''''''''Point_End'''''''''''''''''
End Sub
Private Sub OpenPoint(FileName As String)
    'On Error Resume Next
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    ''''''''''''''Point_Begin'''''''''''''''''
    If fso.FileExists(App.Path + "\文本库\" + FileName + ".point") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName + ".point")
        If ts.AtEndOfStream = False Then
            lvwDividePoint.ListItems.Clear
            Dim myItem As ListItem
            Do While ts.AtEndOfStream = False
                Set myItem = lvwDividePoint.ListItems.Add
                myItem.Text = ts.ReadLine
                myItem.SubItems(1) = ts.ReadLine
                myItem.SubItems(2) = ts.ReadLine
            Loop
        Else
            ''
        End If
        ts.Close
    Else
        ''
       
    End If
    ''''''''''''''Point_End''''''''''''''''
End Sub


Private Sub OpenNote(FileName As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    ''''''''''''''Note_Begin'''''''''''''''''
    If fso.FileExists(App.Path + "\文本库\" + FileName + ".not") = True Then
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName + ".not")
        If ts.AtEndOfStream = False Then
            txtNote.Text = ts.ReadAll
        Else
            txtNote.Text = ""
            
        End If
        ts.Close
    Else
        txtNote.Text = ""
       
    End If
    ''''''''''''''Note_End''''''''''''''''
End Sub

Private Sub lvwLesson_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call dblClickPlay
    Call ReadPlay
    
    Call OpenPoint(lvwLesson.SelectedItem.Text + ".pit")

    Call OpenNote(lvwLesson.SelectedItem.Text + ".llk")
    
    If Button = 1 Then
        'Call OpenText(lvwLesson.SelectedItem.Text + ".llk")
        On Error GoTo ErrHdl
        txtTRY.FileName = App.Path + "\文本库\" + lvwLesson.SelectedItem.Text + ".rtf"
        'txtTRY.SelStart = 0
        'txtTRY.SelLength = Len(txtTRY.Text)
        'txtTRY.SelFontName = "Times New Roman"
        'txtTRY.SelCharOffset = 110
        'txtTRY.SelBold = True
        'txtTRY.SelFontSize = 14
        'txtTRY.SelStart = 0
        'txtTRY.SelLength = 0
        Exit Sub
ErrHdl:
        If OpenText(lvwLesson.SelectedItem.Text + ".llk") = False Then '返回false表示没有这个文件
            txtTRY.SelStart = 0
            txtTRY.SelLength = Len(txtTRY.Text)
            txtTRY.SelFontName = "Times New Roman"
            txtTRY.SelBold = True
            txtTRY.SelCharOffset = 110
            txtTRY.SelFontSize = 14
            txtTRY.SelStart = 0
            txtTRY.SelLength = 0
        End If
        Exit Sub
    End If
End Sub

Private Sub lvwLesson2_DblClick()
    If lvwLesson2.ListItems.Count < 1 Then Exit Sub
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    Dim i As Integer
    For i = 1 To lvwLesson.ListItems.Count
        If lvwLesson.ListItems(i).Text = lvwLesson2.SelectedItem.Text And _
            lvwLesson.ListItems(i).SubItems(1) = lvwLesson2.SelectedItem.SubItems(1) Then
            lvwLesson.ListItems(i).Selected = True
            Exit For
        End If
    Next
    
    Call cmdPlay_Click
End Sub

Private Sub lvwLesson2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwLesson2.ListItems.Count < 1 Then Exit Sub
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    Call lvwLesson_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lvwLesson2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvwLesson2.ListItems.Count < 1 Then Exit Sub
    If lvwLesson.ListItems.Count < 1 Then Exit Sub
    
    Dim i As Integer
    For i = 1 To lvwLesson.ListItems.Count
        If lvwLesson.ListItems(i).Text = lvwLesson2.SelectedItem.Text And _
            lvwLesson.ListItems(i).SubItems(1) = lvwLesson2.SelectedItem.SubItems(1) Then
            lvwLesson.ListItems(i).Selected = True
            Exit For
        End If
    Next
    
    Call lvwLesson_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub mnuFullScreen_Click()
    intScreenType = FULL_SCREEN
    Call Form_Resize
End Sub

Private Sub mnuOpen_Click()
    
    
    On Error GoTo ErrHandler
    
    dlgFile.CancelError = True
    dlgFile.FileName = ""
    dlgFile.MaxFileSize = 10000
    dlgFile.Flags = cdlOFNNoChangeDir Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    dlgFile.Filter = "英语录音 (*.mp3;*.wav)|*.mp3;*.wav"
    dlgFile.ShowOpen
    
    Dim lstItem As ListItem
    Dim strDialog_File_Name As String
    Dim strFiles_Name() As String
    Dim strFiles_ShortName() As String
    Dim blnIsExist_File As Boolean
    Dim i, j As Long
    
    strDialog_File_Name = Trim(dlgFile.FileName)
    If Len(strDialog_File_Name) = 0 Then
        Exit Sub
    End If
          
    strFiles_Name = Split(strDialog_File_Name, Chr(0))
    strFiles_ShortName = Split(strDialog_File_Name, Chr(0))
    If UBound(strFiles_Name) = 0 Then
        ReDim strFiles_Name(1)
        strFiles_Name(1) = strDialog_File_Name
        ReDim strFiles_ShortName(1)
        strFiles_ShortName(1) = right(strDialog_File_Name, Len(strDialog_File_Name) - InStrRev(strDialog_File_Name, "\"))
    Else
        For i = 1 To UBound(strFiles_Name)
            strFiles_Name(i) = strFiles_Name(0) + "\" + strFiles_Name(i) '连接路径和文件名，组成文件数组
        Next i
    End If
    
    For j = 1 To UBound(strFiles_Name)
        blnIsExist_File = False
        For i = 1 To lvwLesson.ListItems.Count
            If lvwLesson.ListItems(i).SubItems(1) = Trim(strFiles_Name(j)) Then
                MsgBox " " + vbCrLf + "这个文件" + strFiles_Name(j) + "已经有了!" + vbCrLf + " "
                blnIsExist_File = True
            End If
        Next
        If blnIsExist_File = False Then
            Set lstItem = lvwLesson.ListItems.Add(, , Trim(strFiles_ShortName(j)))
            lstItem.SubItems(1) = strFiles_Name(j)
        End If
    Next j
    
    Exit Sub
ErrHandler:
    
End Sub

Private Sub mnuStandardScreen_Click()
    intScreenType = STANDARD_SCREEN
    Call Form_Resize
End Sub

Private Sub myFont_Click()
    On Error Resume Next
    
    'dlgFile.CancelError = True
    dlgFile.Flags = cdlCFBoth '+ cdlCFPrinterFonts + cdlCFScreenFonts ' + cdlCFForceFontExist
    
    dlgFile.Color = txtTRY.SelColor
    dlgFile.FontName = txtTRY.SelFontName
    dlgFile.FontBold = txtTRY.SelBold
    dlgFile.FontItalic = txtTRY.SelItalic
    dlgFile.FontSize = txtTRY.SelFontSize
    dlgFile.FontStrikethru = txtTRY.SelStrikeThru
    dlgFile.FontUnderline = txtTRY.SelUnderline
    
    dlgFile.ShowFont
    
    txtTRY.SelColor = dlgFile.Color
    txtTRY.SelFontName = dlgFile.FontName
    txtTRY.SelBold = dlgFile.FontBold
    txtTRY.SelItalic = dlgFile.FontItalic
    txtTRY.SelFontSize = dlgFile.FontSize
    txtTRY.SelStrikeThru = dlgFile.FontStrikethru
    txtTRY.SelUnderline = dlgFile.FontUnderline
    
    txtTRY.SetFocus
End Sub

Private Sub picColor_Click()
    txtTRY.SelColor = picColor.BackColor
    txtTRY.SetFocus
End Sub

Private Sub picSetColor_DblClick()
    On Error GoTo ErrHandler
    
    dlgFile.ShowColor
    txtTRY.SelColor = dlgFile.Color
    picColor.BackColor = dlgFile.Color
    txtTRY.SetFocus
ErrHandler:
    
End Sub

Private Sub Timer2_Timer()
    '''''''''''新加2018-2-8'''''''''''
    On Error Resume Next
    'lblRecord.Caption = GetPlayTime()
    ''''''''''''''''''''''''''''''''''
End Sub

Private Sub tmrPlay_Timer()

    If blShowCarefullListenForm = True Then Exit Sub

    If blMoveSlider = True Then Exit Sub
        
    If chkAddTime.Value = Checked Then
        txtAddTime.Visible = True
        lblAddTime.Visible = True
    Else
        txtAddTime.Visible = False
        lblAddTime.Visible = False
    End If
    
    If chkPauseTime.Value = Checked Then
        lblPauseTime.Visible = True
        txtPauseTime.Visible = True
    Else
        lblPauseTime.Visible = False
        txtPauseTime.Visible = False
    End If
        
    sldSound.Value = mciRead.position
    
    On Error GoTo ErrHandler
    If mciRead.position >= mciRead.To Then
        If chkPauseTime.Value = Checked Then

            If lngPauseTime <= Val(txtPauseTime.Text) * (1000 / tmrPlay.Interval) Then
                lngPauseTime = lngPauseTime + 1
                lblPauseTime.Caption = str(lngPauseTime \ (1000 / tmrPlay.Interval)) + "/"
                
                mciRead.Command = "stop"
                'blTimeOverPause = True
                
                Exit Sub
            Else
                lngPauseTime = 0
            End If
        Else

        End If
        Call ReadPlay
    Else
        'blTimeOverPause = False
    End If
    Exit Sub
ErrHandler:
End Sub

Private Function GetPlayTime() As String
    Dim Msecond As Long
    Dim Hours As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
        
    Msecond = mciRead.position
    Seconds = (Msecond \ 1000) Mod 60
    Minutes = (Msecond \ 60000) Mod 60
    Hours = (Msecond \ 3600000)
    
    GetPlayTime = CStr(Hours) + "小时" + CStr(Minutes) + "分" + CStr(Seconds) + "秒"
End Function

Private Function GetPlayTime_Ex(position As Long) As String
    Dim Msecond As Long
    Dim Hours As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
        
    Msecond = position
    Seconds = (Msecond \ 1000) Mod 60
    Minutes = (Msecond \ 60000) Mod 60
    Hours = (Msecond \ 3600000)
    
    GetPlayTime_Ex = CStr(Hours) + "小时" + CStr(Minutes) + "分" + CStr(Seconds) + "秒"
End Function

Private Function PositionToString(Msecond As Long) As String
    Dim Hours As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
        
    'Msecond = mciRead.Position
    Seconds = (Msecond \ 1000) Mod 60
    Minutes = (Msecond \ 60000) Mod 60
    Hours = (Msecond \ 3600000)
    
    PositionToString = str(Hours) + "小时" + str(Minutes) + "分" + str(Seconds) + "秒"
End Function

Public Function GetCurrentDividePoint() As Long
    
    If lvwDividePoint.ListItems.Count < 1 Then Exit Function
    
    GetCurrentDividePoint = Val(lvwDividePoint.SelectedItem.Text)
    
End Function

Public Function GetNextDividePoint() As Long
    
    If lvwDividePoint.ListItems.Count < 1 Then Exit Function
    
    If lvwDividePoint.SelectedItem.Index < lvwDividePoint.ListItems.Count Then
        GetNextDividePoint = Val(lvwDividePoint.ListItems(lvwDividePoint.SelectedItem.Index + 1).Text)
    Else
        GetNextDividePoint = mciRead.Length
    End If
End Function

Private Sub sldSound_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blMoveSlider = True
End Sub

Private Sub sldSound_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mciRead.From = sldSound.Value
    mciRead.Command = "play"
    blMoveSlider = False
End Sub

Private Sub tmrRecord_Timer()

    
    If blRecordPlaying = True Then
        'sldSound.Width = 4455
        lblRecord.Visible = True
        lblRecord.Caption = "播放"
        If mciRecord.position >= lngRecordEnd Then blRecordPlaying = False
    ElseIf blRecording = True Then
        'sldSound.Width = 4455
        lblRecord.Visible = True
        lblRecord.Caption = str((mciRecord.position - lngRecordStart) / 1000) + "秒"
    Else
        'sldSound.Width = 4455 '5415
        lblRecord.Visible = True 'False
        'lblRecord.Caption = "- -"
        
        Dim Msecond As Long
        Dim Hours As Integer
        Dim Minutes As Integer
        Dim Seconds As Integer
            
        Msecond = mciRead.position
        Seconds = (Msecond \ 1000) Mod 60
        Minutes = (Msecond \ 60000) Mod 60
        Hours = (Msecond \ 3600000)
        
        lblRecord.Caption = str(Minutes) + "分" + str(Seconds) + "秒"
        
    End If
End Sub

Private Sub tmrStatus_Timer()
    If blMoveSlider = True Then Exit Sub
    
    sldSound.Value = mciRead.position
    
End Sub

Private Sub mnuEditDelAllDividePoint_Click()
    lvwDividePoint.ListItems.Clear
    Dim myListItem As ListItem
    Set myListItem = lvwDividePoint.ListItems.Add
    myListItem.Text = Format(0, "000000000")
    myListItem.SubItems(1) = PositionToString(0)
    myListItem.SubItems(2) = strFileName
End Sub

Private Sub mnuEditDelDividePoint_Click()
    Dim myListItem As ListItem
    On Error GoTo ErrHandler
    lvwDividePoint.ListItems.Remove lvwDividePoint.SelectedItem.Index
    If lvwDividePoint.ListItems.Count < 1 Then
        lvwDividePoint.ListItems.Clear
        Set myListItem = lvwDividePoint.ListItems.Add
        myListItem.Text = Format(0, "000000000")
        myListItem.SubItems(1) = PositionToString(0)
        myListItem.SubItems(2) = strFileName
    End If
    Exit Sub
ErrHandler:
    'MsgBox "错误操作：" + vbCrLf + vbCrLf + "没有选中断点，或没有断点。"
    lvwDividePoint.ListItems.Clear
    Set myListItem = lvwDividePoint.ListItems.Add
    myListItem.Text = Format(0, "000000000")
    myListItem.SubItems(1) = PositionToString(0)
    myListItem.SubItems(2) = strFileName
End Sub


Private Sub txtCharOffSet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_ENTER Then
        If Val(txtCharOffSet.Text) > 0 Then
            txtTRY.SelCharOffset = Val(txtCharOffSet.Text)
        End If
    End If
    If KeyCode = KEY_HOME Then
        txtTRY.SetFocus
    End If
End Sub

Private Sub txtTRY_KeyDown(KeyCode As Integer, Shift As Integer)
'    Label1.Caption = str(KeyCode)
'    Static blStartRecord As Boolean
'    Select Case KeyCode
'        Case KEY_UPPAGE
'            KeyCode = 0
'            If Shift = SHIFT_CTRL Then
'                Call JumpPlay(mciRead.Position - 5000)
'            Else
'                SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_UP, 0
'                Call cmdReRead_Click
'                lngPauseTime = 0
'            End If
'        Case KEY_DOWNPAGE
'            KeyCode = 0
'            If Shift = SHIFT_CTRL Then
'                Call JumpPlay(mciRead.Position + 5000)
'            Else
'                SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_DOWN, 0
'                Call cmdReRead_Click
'                lngPauseTime = 0
'            End If
'        Case KEY_LEFT
'            If Shift = SHIFT_CTRL Or chkQuickPress.Value = Checked Then
'                KeyCode = 0
'                Call JumpPlay(mciRead.Position - 5000)
'            End If
'        Case KEY_RIGHT
'            If Shift = SHIFT_CTRL Or chkQuickPress.Value = Checked Then
'                KeyCode = 0
'                Call JumpPlay(mciRead.Position + 5000)
'            End If
'        Case KEY_UP
'            If Shift = 2 Or Me.chkQuickPress.Value = Checked Then
'                SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_UP, 0
'                Call cmdReRead_Click
'            End If
'        Case KEY_DOWN
'            If Shift = 2 Or Me.chkQuickPress.Value = Checked Then
'                SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_DOWN, 0
'                Call cmdReRead_Click
'            End If
''        Case KEY_UP
'''            If chkQuickPress.Value = Checked Then
'''                KeyCode = 0
'''                cmdPlay2_Click
'''                SendMessage lvwDividePoint.hWnd, WM_KEYDOWN, VK_UP, 0
'''                Call cmdReRead_Click
'''            End If
''        Case KEY_DOWN
'''            If chkQuickPress.Value = Checked Then
'''                KeyCode = 0
'''                Shift = 0
'''                Call cmdPause_Click
'''            End If
'        Case KEY_PAUSE
'            KeyCode = 0
'            cmdPause_Click
'        Case KEY_INSERT
'            KeyCode = 0
'            cmdPlay2_Click
'            SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_UP, 0
'            Call cmdReRead_Click
'            lngPauseTime = 0
'        Case KEY_ENTER
'            If Shift = SHIFT_CTRL Then
'                KeyCode = 0
'                Shift = 0
'                Call cmdQuickCompare_Click
'            ElseIf Shift = 0 Then
'                KeyCode = 0
'                cmdPause_Click
'            ElseIf Shift = SHIFT_SHIFT Then
'                blPB = False
'            Else
'            End If
'        Case KEY_DEL
'            If chkQuickPress.Value = Checked Or Shift = 2 Or Shift = SHIFT_SHIFT Then
'                KeyCode = 0
'                Shift = 0
'                Call mnuEditDelDividePoint_Click
'                lvwDividePoint.Refresh
'            End If
'        Case 107, 229, 219, 221  '大键盘]键 '小键盘+键,大键盘[键 '小键盘+键
'            KeyCode = 0
'            Shift = 0
'            Call cmdReRead_Click
'            lngPauseTime = 0
''        Case 109 '小建盘-键
''            KeyCode = 0
''            cmdPlay2_Click
''            SendMessage lvwDividePoint.hwnd, WM_KEYDOWN, VK_UP, 0
''            Call cmdReRead_Click
'        Case 110  '小键盘del键
'            If chkCompare.Value = Checked Then
'                If blStartRecord = False Then
'                    Call cmdRecord_Click
'                Else
'                    Call cmdRecordStop_Click
'                End If
'                blStartRecord = Not blStartRecord
'            End If
'
'        Case 96 '小键盘0键
'            If chkCompare.Value = Checked Then
'                Call OnlyPlayRecord
'            End If
'        Case 97 '小建盘1键
'            If chkCompare.Value = Checked Then
'                Call cmdRecordPlay_Click
'            End If
'        Case Else
'    End Select
End Sub
Private Sub JumpPlay(ByVal varFrom As Long)
    If varFrom < 1 Then
        varFrom = 0
    Else
        mciRead.From = varFrom
    End If
    mciRead.Command = "play"
End Sub

Private Sub txtTRY_KeyPress(KeyAscii As Integer)
    If blPB = True Then
        If KeyAscii = ASC_KEY_ENTER Then KeyAscii = 0       '屏蔽Enter键
        If KeyAscii = ASC_KEY_CTRL_ENTER Then KeyAscii = 0  '屏蔽Ctrl+Enter键
    End If
    If KeyAscii = 43 Then KeyAscii = 0 '小键盘+键
'    If KeyAscii = 45 Then KeyAscii = 0 '小键盘-键
    
    If KeyAscii = 93 Or KeyAscii = 91 Then '大键盘[,]键
        DoEvents
        KeyAscii = 0
    End If
    
    If chkQuickPress.Value = Checked Then KeyAscii = 0
    If chkCompare.Value = Checked Then KeyAscii = 0
    blPB = True
End Sub

Public Sub SaveText(FileName As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    If fso.FileExists(App.Path + "\列表.llk") = False Then
        Set ts = fso.CreateTextFile(App.Path + "\列表.llk")
    Else
        Set ts = fso.OpenTextFile(App.Path + "\列表.llk", ForWriting)
    End If
    
    Dim i As Integer
    For i = 1 To lvwLesson.ListItems.Count
        ts.WriteLine lvwLesson.ListItems(i).Text
        ts.WriteLine lvwLesson.ListItems(i).SubItems(1)
    Next
    ts.Close
        
    ''''''''''''''''
    If fso.FileExists(App.Path + "\列表2.llk") = False Then
        Set ts = fso.CreateTextFile(App.Path + "\列表2.llk")
    Else
        Set ts = fso.OpenTextFile(App.Path + "\列表2.llk", ForWriting)
    End If
        
    For i = 1 To lvwLesson2.ListItems.Count
        ts.WriteLine lvwLesson2.ListItems(i).Text
        ts.WriteLine lvwLesson2.ListItems(i).SubItems(1)
    Next
    ts.Close
    '''''''''''''''''
        
    If fso.FileExists(App.Path + "\文本库\" + FileName) = False Then
        Set ts = fso.CreateTextFile(App.Path + "\文本库\" + FileName)
    Else
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName, ForWriting)
    End If
    ts.Write txtTRY.Text
    ts.Close
End Sub

Public Sub SaveOnlyText(FileName As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    If fso.FileExists(App.Path + "\文本库\" + FileName) = False Then
        Set ts = fso.CreateTextFile(App.Path + "\文本库\" + FileName)
    Else
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName, ForWriting)
    End If
    ts.Write txtTRY.Text
    ts.Close
    

End Sub

Function OpenText(FileName As String) As Boolean
    OpenText = False
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    If fso.FileExists(App.Path + "\文本库\" + FileName) = True Then
        Set ts = fso.OpenTextFile(App.Path + "\文本库\" + FileName)
        If ts.AtEndOfStream = False Then
            txtTRY.Text = ts.ReadAll
        Else
            txtTRY.Text = ""
        End If
        ts.Close
        OpenText = True
    Else
        txtTRY.Text = ""
    End If
    


End Function

Public Function GetNum(strItem As String) As String
    Dim i, intStart, intEnd As Integer
    Dim blNum As Boolean
    blNum = False
    intStart = 0
    intEnd = 0
    
    For i = 1 To Len(strItem)
        If blNum = False Then
            If Asc(Mid(strItem, i, 1)) >= 48 And Asc(Mid(strItem, i, 1)) <= 57 Then
                intStart = i
                blNum = True
            End If
        Else
            If Asc(Mid(strItem, i, 1)) < 48 Or Asc(Mid(strItem, i, 1)) > 57 Then
                intEnd = i
                blNum = False
                Exit For
            End If
        End If
    Next
    If blNum = True Then intEnd = Len(strItem)
    
    If intStart = 0 Then
        GetNum = ""
    Else
        GetNum = Replace(str(Val(Mid(strItem, intStart, intEnd - intStart))), " ", "")
    End If
End Function

