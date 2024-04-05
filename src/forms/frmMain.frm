VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "QZhi TxtToBmp"
   ClientHeight    =   6615
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9360
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picProgress 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   609
      TabIndex        =   5
      Top             =   6120
      Width           =   9135
   End
   Begin VB.PictureBox picFrame 
      Height          =   5895
      Left            =   120
      ScaleHeight     =   389
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   605
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.PictureBox picMask 
         Height          =   240
         Left            =   8760
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   4
         Top             =   5520
         Width           =   240
      End
      Begin VB.HScrollBar hsbHScroll 
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   5520
         Width           =   8775
      End
      Begin VB.VScrollBar vsbVScroll 
         Height          =   5535
         Left            =   8760
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picImg 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   3
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileImport 
         Caption         =   "导入(&I)..."
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "导出(&E)..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "选项(&O)"
      Begin VB.Menu mnuOptionsBackgroundColor 
         Caption         =   "背景色(&B)"
      End
      Begin VB.Menu mnuOptionsForegroundColor 
         Caption         =   "前景色(&F)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpLicense 
         Caption         =   "Apache License, 2.0(&L)"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于 TxtToBmp(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 2024 QZhi Studio
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.


Option Explicit

Private fntFont(255, 15, 7) As Byte

Private lngImgWidth As Long
Private lngImgHeight As Long

Private bytFileData() As Byte

Private clrColor(1) As OLE_COLOR

Private xCursor As Long
Private yCursor As Long

Private hProgress As Long

Private Sub Form_Initialize()
    hProgress = CreateWindowExA(0, "msctls_progress32", "", WS_CHILD Or WS_VISIBLE, 0, 0, 16, 16, picProgress.hwnd, 0, App.hInstance, 0)
End Sub

Private Sub Form_Load()
    LoadFont
    clrColor(1) = vbWhite
    picMask.BorderStyle = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picFrame.Move 120, 120, Me.ScaleWidth - 240, Me.ScaleHeight - 720
    picProgress.Move 120, Me.ScaleHeight - 480, Me.ScaleWidth - 240, 360
    
    vsbVScroll.Move picFrame.ScaleWidth - 16, 0, 16, picFrame.ScaleHeight - 16
    hsbHScroll.Move 0, picFrame.ScaleHeight - 16, picFrame.ScaleWidth - 16, 16
    picMask.Move picFrame.ScaleWidth - 16, picFrame.ScaleHeight - 16, 16, 16
    
    If picImg.Width > picFrame.ScaleWidth Then
        hsbHScroll.Visible = True
        hsbHScroll.Enabled = True
        hsbHScroll.Max = picImg.Width - picFrame.ScaleWidth + 16
    Else
        hsbHScroll.Enabled = False
        hsbHScroll.Visible = False
    End If
    
    If picImg.Height > picFrame.ScaleHeight Then
        vsbVScroll.Visible = True
        vsbVScroll.Enabled = True
        vsbVScroll.Max = picImg.Height - picFrame.ScaleHeight + 16
    Else
        vsbVScroll.Enabled = False
        vsbVScroll.Visible = False
    End If
    
    If (hsbHScroll.Visible = True) Or (vsbVScroll.Visible = True) Then
        picMask.Visible = True
    Else
        picMask.Visible = False
    End If
    
    DoEvents
End Sub

Private Sub LoadFont()

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim bytTemp() As Byte
    Dim lngTemp As Long

    ReDim bytTemp(4095)
    bytTemp = LoadResData("ASC16", 10)
    
    For i = 0 To 255
        For j = 0 To 15
            lngTemp = bytTemp(i * 16 + j)
            For k = 0 To 7
                fntFont(i, j, 7 - k) = lngTemp And 1&
                lngTemp = lngTemp \ 2
            Next k
        Next j
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub hsbHScroll_Change()
    picImg.Left = -hsbHScroll.Value
End Sub

Private Sub hsbHScroll_Scroll()
    hsbHScroll_Change
End Sub

Private Sub mnuFileExport_Click()
    Dim drRet As DLGRET
    
    drRet = GetSaveFile(Me.hwnd, "位图文件 (*.bmp)" & Chr(0) & "*.bmp" & Chr(0) & Chr(0), App.Path, "")
    
    If drRet.blnError = True Then
        Exit Sub
    End If
    
    SavePicture picImg.Image, drRet.strFileName
End Sub

Private Sub mnuFileImport_Click()
    Dim drRet As DLGRET
    
    drRet = GetOpenFile(Me.hwnd, "所有文件 (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0), App.Path, "")
    
    If drRet.blnError = True Then
        Exit Sub
    End If
    
    Dim intFileNum As Integer

    intFileNum = FreeFile
    
    Open drRet.strFileName For Binary As #intFileNum
    
        If LOF(intFileNum) = 0 Then
            Close #intFileNum
            Exit Sub
        End If
        
        ReDim bytFileData(LOF(intFileNum) - 1)

        Get #intFileNum, , bytFileData
    
    Close #intFileNum
    
    TxtToBmp
End Sub

Private Sub TxtToBmp()
    Dim i As Long
    
    xCursor = -1
    yCursor = 0
    
    picImg.Move 0, 0, 8, 16
    picImg.Cls
    
    SendMessageA hProgress, PBM_SETRANGE32, 0, UBound(bytFileData)
    SendMessageA hProgress, PBM_SETPOS, 0, 0
    
    For i = 0 To UBound(bytFileData)
        
        If bytFileData(i) = Asc(vbLf) Then
            yCursor = yCursor + 1
            picImg.Height = picImg.Height + 16
        ElseIf bytFileData(i) = Asc(vbCr) Then
            xCursor = -1
        Else
        
            xCursor = xCursor + 1
            If (xCursor + 1) * 8 > picImg.ScaleWidth Then
                picImg.Width = picImg.Width + 8
            End If
        
            putchar bytFileData(i), xCursor * 8, yCursor * 16
        End If
        
        'If i Mod 512 = 0 Then
            picImg.Refresh
            DoEvents
        'End If
        
        SendMessageA hProgress, PBM_SETPOS, i, 0
        
    Next i
    
    SendMessageA hProgress, PBM_SETPOS, 0, 0
    
    picImg.Refresh
    DoEvents

End Sub


Private Function putchar(ByVal char As Byte, x As Long, y As Long) As Boolean

    On Error Resume Next

    Dim i As Long
    Dim j As Long

    For i = 0 To 7
        For j = 0 To 15
            picImg.PSet (x + i, y + j), clrColor(fntFont(char, j, i))
        Next j
    Next i
    
    putchar = True
    
    Exit Function
    
FuncError:
    
End Function

Private Sub mnuHelpAbout_Click()
    ShellAboutA Me.hwnd, App.ProductName, "一个将文本转换为 DOS 风格图片的小工具（仅支持 437 代码页）。" & vbCrLf & "By QZhi Studio", Me.Icon
End Sub

Private Sub mnuHelpLicense_Click()
    Dim fLicense As New frmLicense
    fLicense.Show vbModal
End Sub

Private Sub mnuOptionsBackgroundColor_Click()
    clrColor(0) = ChooseColorDlg(Me.hwnd, clrColor(0))
    picImg.BackColor = clrColor(0)
End Sub

Private Sub mnuOptionsForegroundColor_Click()
    clrColor(1) = ChooseColorDlg(Me.hwnd, clrColor(1))
End Sub

Private Sub picFrame_Resize()
    Form_Resize
End Sub

Private Sub picImg_Resize()
    Form_Resize
End Sub

Private Sub picProgress_Resize()
    MoveWindow hProgress, 0, 0, picProgress.ScaleWidth, picProgress.ScaleHeight, True
End Sub

Private Sub vsbVScroll_Change()
    picImg.Top = -vsbVScroll.Value
End Sub

Private Sub vsbVScroll_Scroll()
    vsbVScroll_Change
End Sub
