VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmLicense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apache License, 2.0"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6960
   Icon            =   "frmLicense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin SHDocVwCtl.WebBrowser brwLicense 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   7646
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblText 
      Caption         =   "本程序受 Apache 2.0 协议保护。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmLicense"
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

Dim WithEvents hdocLicense As HTMLDocument
Attribute hdocLicense.VB_VarHelpID = -1

Private Sub brwLicense_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Set hdocLicense = brwLicense.Document
End Sub

Private Sub brwLicense_DownloadBegin()
    brwLicense.Silent = True
End Sub

Private Sub brwLicense_DownloadComplete()
    brwLicense.Silent = True
End Sub

Private Sub Form_Load()
    brwLicense.Navigate "res://" & App.EXEName & ".exe/License.html"
End Sub

Private Function hdocLicense_oncontextmenu() As Boolean
    hdocLicense_oncontextmenu = False
End Function

