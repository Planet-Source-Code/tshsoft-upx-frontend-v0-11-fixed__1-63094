VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin UPXFE.ReadOutput ReadOutput1 
      Left            =   3900
      Top             =   4020
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   4815
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   90
      Width           =   5205
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sVer As String 'UPX Version

Private Sub Form_Load()
Dim S As String

If FileExists(cPath & "upx.exe") Then
   ReadOutput1.SetCommand = "upx -V"
   ReadOutput1.ProcessCommand
   DoEvents
Else
   sVer = ""
End If

S = "Written by TSHsoft" & vbCrLf
S = S & vbCrLf
S = S & "The Ultimate Packer for eXecutables " & sVer & vbCrLf
S = S & "Copyright (c) 1996-2004 Markus Oberhumer & Laszlo Molnar" & vbCrLf
S = S & "http://upx.sourceforge.net" & vbCrLf
S = S & vbCrLf
S = S & "UPX is a portable, extendable, high-performance executable" & vbCrLf
S = S & "packer for several different executable formats. It achieves" & vbCrLf
S = S & "an excellent compression ratio and offers **very** fast" & vbCrLf
S = S & "decompression. Your executables suffer no memory overhead" & vbCrLf
S = S & "or other drawbacks for most of the formats supported." & vbCrLf
S = S & vbCrLf
S = S & vbCrLf
S = S & "UPX FrontEnd v0.11  01/11/2005" & vbCrLf
S = S & "- Bug Fixed" & vbCrLf
S = S & vbCrLf
S = S & "UPX FrontEnd v0.10  31/10/2005" & vbCrLf

Text1.Text = S
End Sub

Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)

 If sChunk <> "" Then
    sVer = UCase(Left(sChunk, 8))
 End If
 
End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

