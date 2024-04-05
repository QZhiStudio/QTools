Attribute VB_Name = "modWINAPI"
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

Public Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long

Public Declare Function GetOpenFileNameA Lib "comdlg32.dll" (ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileNameA Lib "comdlg32.dll" (ByRef pOpenfilename As OPENFILENAME) As Long

Public Declare Function ChooseColorA Lib "comdlg32.dll" (ByRef pChoosecolor As CHOOSECOLOR) As Long

Public Declare Function ShellAboutA Lib "shell32.dll" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Public Type CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type DLGRET ' 自定义返回类型
    lngStat As Long
    strFileName As String
    blnError As Boolean
End Type

Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000

Public Const WM_USER = &H400&
Public Const PBM_SETPOS = (WM_USER + 2&)
Public Const PBM_SETRANGE32 = (WM_USER + 6&)

Private crCustColors(15) As OLE_COLOR

' lpstrInitialDir 初始地址
Public Function GetOpenFile(ByVal hwndOwner As Long, ByVal lpstrFilter As String, ByVal lpstrInitialDir As String, ByVal lpstrTitle As String) As DLGRET

    Dim ofnOpenFileName As OPENFILENAME
    Dim dlrReturn As DLGRET
    
    Dim i As Long
    
    With ofnOpenFileName
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = lpstrFilter
        .lpstrFile = Space(&HFE)
        .nMaxFile = &HFF
        .lpstrFileTitle = Space(&HFE)
        .nMaxFileTitle = &HFF
        .lpstrInitialDir = lpstrInitialDir
        .lpstrTitle = lpstrTitle
        .flags = &H1804
        .lStructSize = Len(ofnOpenFileName)
    End With
    
    dlrReturn.lngStat = GetOpenFileNameA(ofnOpenFileName)
    If dlrReturn.lngStat >= 1 Then
        dlrReturn.strFileName = ofnOpenFileName.lpstrFile
        dlrReturn.blnError = False
        
        dlrReturn.strFileName = Replace(dlrReturn.strFileName, Chr(0), "")
        For i = Len(dlrReturn.strFileName) To 1 Step -1
            If Mid(dlrReturn.strFileName, i, 1) = " " Then
                dlrReturn.strFileName = Mid(dlrReturn.strFileName, 1, Len(dlrReturn.strFileName) - 1)
            Else
                Exit For
            End If
        Next i
    Else
        dlrReturn.strFileName = vbNullString
        dlrReturn.blnError = True
    End If
    
    GetOpenFile = dlrReturn
    
End Function

' lpstrInitialDir 初始地址
Public Function GetSaveFile(ByVal hwndOwner As Long, ByVal lpstrFilter As String, ByVal lpstrInitialDir As String, ByVal lpstrTitle As String) As DLGRET

    Dim ofnSaveFileName As OPENFILENAME
    Dim dlrReturn As DLGRET
    
    With ofnSaveFileName
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = lpstrFilter
        .lpstrFile = Space(&HFE)
        .nMaxFile = &HFF
        .lpstrFileTitle = Space(&HFE)
        .nMaxFileTitle = &HFF
        .lpstrInitialDir = lpstrInitialDir
        .lpstrTitle = lpstrTitle
        .flags = &H1804
        .lStructSize = Len(ofnSaveFileName)
    End With
    
    dlrReturn.lngStat = GetSaveFileNameA(ofnSaveFileName)
    If dlrReturn.lngStat >= 1 Then
        dlrReturn.strFileName = ofnSaveFileName.lpstrFile
        dlrReturn.blnError = False
    Else
        dlrReturn.strFileName = vbNullString
        dlrReturn.blnError = True
    End If
    
    GetSaveFile = dlrReturn
    
End Function

Public Function ChooseColorDlg(ByVal hwndOwner As Long, ByVal rgbInit As OLE_COLOR) As OLE_COLOR
    Dim ccChooseColor As CHOOSECOLOR
    
    With ccChooseColor
        .lStructSize = LenB(ccChooseColor)
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .rgbResult = rgbInit
        .lpCustColors = VarPtr(crCustColors(0))
        .flags = 1
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = ""
    End With
    
    ChooseColorA ccChooseColor
    
    ChooseColorDlg = ccChooseColor.rgbResult
End Function
