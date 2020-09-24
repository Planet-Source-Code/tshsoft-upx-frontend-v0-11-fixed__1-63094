VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPX (Ultimate Packer for eXecutables) FrontEnd v"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
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
   ScaleHeight     =   5775
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   3060
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   24
      Top             =   5970
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame4 
      Caption         =   "Command Function:"
      Height          =   2685
      Left            =   90
      TabIndex        =   9
      Top             =   2970
      Width           =   7665
      Begin MSComctlLib.ListView ListView1 
         Height          =   2235
         Left            =   270
         TabIndex        =   21
         Top             =   270
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   3942
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Filename"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Packed"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ratio"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2822
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Command Function:"
      Height          =   1965
      Left            =   4020
      TabIndex        =   8
      Top             =   990
      Width           =   3735
      Begin VB.CheckBox ckBackup 
         Caption         =   "Keep backup files"
         Height          =   195
         Left            =   300
         TabIndex        =   22
         ToolTipText     =   "Backup as Filename.ex~"
         Top             =   900
         Width           =   1635
      End
      Begin VB.CommandButton cmdDecompress 
         Caption         =   "Decompress"
         Height          =   345
         Left            =   2370
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompress 
         Caption         =   "Compress"
         Height          =   345
         Left            =   2370
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1770
         Max             =   9
         Min             =   1
         TabIndex        =   12
         Top             =   390
         Value           =   1
         Width           =   1155
      End
      Begin VB.CheckBox ckOverwrite 
         Caption         =   "Overwrite exist files"
         Height          =   225
         Left            =   300
         TabIndex        =   11
         Top             =   1380
         Width           =   1965
      End
      Begin VB.TextBox txtQuality 
         Height          =   285
         Left            =   750
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "1"
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Better"
         Height          =   225
         Left            =   3030
         TabIndex        =   17
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lblQuality 
         Caption         =   "Quality:"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Faster"
         Height          =   255
         Left            =   1260
         TabIndex        =   15
         Top             =   420
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
      Height          =   345
      Left            =   2730
      TabIndex        =   3
      Top             =   1350
      Width           =   1095
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   345
      Left            =   2730
      TabIndex        =   7
      Top             =   2310
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Execute Files"
      Height          =   1965
      Left            =   90
      TabIndex        =   5
      Top             =   990
      Width           =   3885
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Left            =   2640
         TabIndex        =   18
         Top             =   810
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   330
         Width           =   2385
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   645
      Left            =   4380
      TabIndex        =   4
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   7665
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   345
         Left            =   4950
         TabIndex        =   20
         Top             =   330
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   345
         Left            =   5340
         TabIndex        =   19
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtOutputDir 
         Height          =   345
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   330
         Width           =   3465
      End
      Begin VB.Label Label1 
         Caption         =   "TSHsoft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6510
         TabIndex        =   23
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Output Folder:"
         Height          =   225
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   1155
      End
   End
   Begin UPXFE.ReadOutput ReadOutput1 
      Left            =   6240
      Top             =   5910
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Ultimate Packer for eXecutables FrontEnd
'UPX FrontEnd v0.10 31/10/2005
'Copyright (C) 2005 by TSHsoft
'Internal Dos program UPX v1.25

Option Explicit
Private c As cFileDialog
Dim strOutput As String
Dim Temp(6) As String 'Output data(status)


Private Sub cmdBrowse_Click()
Dim strResFolder As String

strResFolder = BrowseForFolder(hwnd, "Please select a folder.")

If strResFolder <> "" Then
   txtOutputDir.Text = strResFolder
End If
End Sub

Private Sub cmdHelp_Click()
  frmHelp.Show
End Sub

Private Sub cmdAdd_Click()
On Error GoTo cmdClassError
Dim sFiles() As String
Dim filecount As Long
Dim sDir As String
Dim i As Long
    
    With c
        .DialogTitle = "Choose Executable Files"
        .CancelError = False
        .Filename = "" 'clear
        .hwnd = Me.hwnd
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_ALLOWMULTISELECT
        .InitDir = App.Path
        .Filter = "Executable Files (*.exe)|*.exe"
        .FilterIndex = 1
        .ShowOpen

        If .Filename = "" Then Exit Sub
        .ParseMultiFileName sDir, sFiles(), filecount
        If UBound(sFiles) = 0 Then
           List1.AddItem sFiles(0)
           lstFiles.AddItem .Filename
        Else
           For i = 0 To filecount - 1
               If Mid(sDir, Len(sDir), 1) <> "\" Then
                  lstFiles.AddItem sDir & "\" & sFiles(i)
               Else
                  lstFiles.AddItem sDir & sFiles(i)
               End If
               List1.AddItem sFiles(i)
           Next i
        End If
    End With
    
Exit Sub

cmdClassError:
    If (Err.Number <> 20001) Then
        MsgBox "Error: " & Err.Description, vbCritical, "Add"
    End If
    
End Sub

Private Sub cmdClearAll_Click()
  lstFiles.Clear
  List1.Clear
  Text1.Text = ""
End Sub

Private Sub cmdCompress_Click()
On Error GoTo errHandler
Dim strCommand, strOption As String
Dim strRun, sOutput, TMP As String
Dim i, j, n As Integer
  
  strCommand = ""
  strCommand = "-" & txtQuality.Text & " "
  strOption = ""
  
  If ckOverwrite.Value = 1 Then
     strOption = "-f "
  End If
  
  If Right(txtOutputDir.Text, 1) <> "\" Then
       sOutput = txtOutputDir.Text & "\"
    Else
       sOutput = txtOutputDir.Text
    End If
    
  'if upx.exe not found
  If FileExists(cPath & "upx.exe") = False Then
     MsgBox "upx.exe not found!", vbExclamation, "Compress"
  Else
  'if upx.exe found
     ListView1.ListItems.Clear
     ListView1.ColumnHeaders(3).Text = "Packed"
     Me.MousePointer = 11 'busy
     For i = 0 To List1.ListCount - 1
         strRun = "upx " & strCommand & strOption & "-o " & Chr(34) & _
                  sOutput & List1.List(i) & Chr(34) & Chr(32) _
                  & Chr(34) & lstFiles.List(i) & Chr(34)
         
         If ckBackup.Value = 1 Then
            If FileExists(sOutput & List1.List(i)) Then
               FileCopy sOutput & List1.List(i), Mid(sOutput & List1.List(i), 1, Len(sOutput & List1.List(i)) - 1) & "~"
            End If
         End If
         
         'compressing...
         Text1.Text = "" 'clear
         ReadOutput1.SetCommand = strRun
         ReadOutput1.ProcessCommand
         DoEvents
         j = InStrRev(Text1, "upx:")
         If j = 0 Then 'if j=0 mean not found error.
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            n = StringTokenizer(Trim(TMP) & " ")
            For n = 1 To (n / 6)
                With ListView1.ListItems
                .Add(n).Text = Mid(Temp(6 * n), 1, Len(Temp(6 * n)) - 6) 'Filename
                .item(n).SubItems(1) = FileByteFormat(CLng(Temp(6 * (n - 1) + 1))) 'Size
                .item(n).SubItems(2) = FileByteFormat(CLng(Temp(3 * (n + n - 1)))) 'Packed
                .item(n).SubItems(3) = Temp(3 * (n + n - 1) + 1) 'Ratio
                .item(n).SubItems(4) = "Done!" 'Status
                End With
            Next n
         Else
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            j = InStrRev(TMP, "AlreadyPacked")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Already Packed" 'Status
                End With
            End If
            j = InStrRev(TMP, "File exists")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "File exists" 'Status
                End With
            End If
            j = InStrRev(TMP, "Permission denied")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Packed
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Permission denied" 'Status
                End With
            End If
            
         End If
         
     Next i
     'Complete!
     Me.MousePointer = 0 'default
  End If
  
Exit Sub
errHandler:
MsgBox (Err.Description & " - " & Err.Source & " - " & CStr(Err.Number)), vbCritical, "Compress"
End Sub

Private Sub cmdDecompress_Click()
On Error GoTo errHandler
Dim strCommand, strOption As String
Dim strRun, sOutput, TMP As String
Dim i, j, n As Integer
  
  strCommand = ""
  strCommand = "-d "
  strOption = ""
  
  If ckOverwrite.Value = 1 Then
     strOption = "-f "
  End If
  
  If Right(txtOutputDir.Text, 1) <> "\" Then
     sOutput = txtOutputDir.Text & "\"
  Else
     sOutput = txtOutputDir.Text
  End If
    
  'if upx.exe not found
  If FileExists(cPath & "upx.exe") = False Then
     MsgBox "upx.exe not found!", vbExclamation, "Decompress"
  Else
  'if upx.exe found
     ListView1.ListItems.Clear
     ListView1.ColumnHeaders(3).Text = "Unpacked"
     Me.MousePointer = 11 'busy
     For i = 0 To List1.ListCount - 1
         strRun = "upx " & strCommand & strOption & "-o " & Chr(34) & _
                  sOutput & List1.List(i) & Chr(34) & Chr(32) _
                  & Chr(34) & lstFiles.List(i) & Chr(34)
         
         If ckBackup.Value = 1 Then
            If FileExists(sOutput & List1.List(i)) Then
               FileCopy sOutput & List1.List(i), Mid(sOutput & List1.List(i), 1, Len(sOutput & List1.List(i)) - 1) & "~"
            End If
         End If
         
         'decompressing...
         Text1.Text = "" 'clear
         ReadOutput1.SetCommand = strRun
         ReadOutput1.ProcessCommand
         DoEvents
         
         j = InStrRev(Text1, "upx:")
         If j = 0 Then 'if j=0 mean not found error.
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            n = StringTokenizer(Trim(TMP) & " ")
            For n = 1 To (n / 6)
                With ListView1.ListItems
                .Add(n).Text = Mid(Temp(6 * n), 1, Len(Temp(6 * n)) - 8) 'Filename
                .item(n).SubItems(1) = FileByteFormat(CLng(Temp(3 * (n + n - 1)))) 'Size
                .item(n).SubItems(2) = FileByteFormat(CLng(Temp(6 * (n - 1) + 1))) 'Unpacked
                .item(n).SubItems(3) = Temp(3 * (n + n - 1) + 1) 'Ratio
                .item(n).SubItems(4) = "Done!" 'Status
                End With
            Next n
         Else
            TMP = Mid(Text1, 327, Len(Text1) - 327)
            j = InStrRev(TMP, "NotPacked")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Unpacked
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Not Packed" 'Status
                End With
            End If
            j = InStrRev(TMP, "File exists")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Unpacked
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "File exists" 'Status
                End With
            End If
            j = InStrRev(TMP, "Permission denied")
            If j <> 0 Then
               With ListView1.ListItems
                .Add(1).Text = List1.List(i) 'Filename
                .item(1).SubItems(1) = FileByteFormat(FileLen(lstFiles.List(i))) 'Size
                .item(1).SubItems(2) = "0" 'Unpacked
                .item(1).SubItems(3) = "0" 'Ratio
                .item(1).SubItems(4) = "Permission denied" 'Status
                End With
            End If
            
         End If
         
     Next i
     'Complete!
     Me.MousePointer = 0 'default
  End If

Exit Sub
errHandler:
MsgBox (Err.Description & " - " & Err.Source & " - " & CStr(Err.Number)), vbCritical, "Decompress"
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

  If List1.ListCount = 0 Then
     MsgBox "Please select executable files to delete.", vbExclamation, "Delete"
  End If
  
  Do While i < List1.ListCount
      If List1.Selected(i) = True Then
         List1.RemoveItem i
         lstFiles.RemoveItem i
      Else
         i = i + 1
      End If
      DoEvents
  Loop
  
End Sub

Private Sub txtQuality_Change()
  HScroll1.Value = txtQuality.Text
End Sub

Private Sub txtQuality_KeyPress(KeyAscii As Integer)
  'if input 1 to 9 or Backspace
If KeyAscii >= 49 And KeyAscii <= 57 Or KeyAscii = 8 Then
   Else
   KeyAscii = 0
End If
End Sub

Private Sub HScroll1_Change()
  txtQuality.Text = HScroll1.Value
End Sub

Private Sub ReadOutput1_GotChunk(ByVal sChunk As String, ByVal LastChunk As Boolean)
 
 Text1 = Text1 & sChunk
 
End Sub

Private Sub ReadOutput1_Error(ByVal Error As String, LastDLLError As Long)
    MsgBox "Error!" & vbNewLine & _
            "Description: " & Error & vbNewLine & _
            "LastDLLError: " & LastDLLError, vbCritical, "Error"
End Sub

Private Sub Form_Load()

    Me.Caption = Me.Caption & App.Major & "." & App.Minor & App.Revision
    Set c = New cFileDialog
    
    txtOutputDir.Text = App.Path
    
    'copy it to WINDOWS Directory, so can run it anyway!
    If FileExists(Get_WinPath & "upx.exe") = False Then
       If FileExists(cPath & "upx.exe") = True Then
          FileCopy cPath & "upx.exe", Get_WinPath & "upx.exe"
       Else
          MsgBox "upx.exe not found! Please download at http://upx.sourceforge.net", vbInformation, "UPX FrontEnd"
          End
       End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set c = Nothing
End Sub

Function StringTokenizer(DATA As String) As Integer
On Error Resume Next
Dim i As Integer, j As Integer, t As Integer
Dim S As String

j = 1 'set data count to zero
t = 1 'set start to zero

For i = 1 To Len(DATA)
    S = Mid$(DATA, i, 1)
    If S = " " Then
       If Trim(Mid$(DATA, t, i - t)) <> "" Then
          Temp(j) = Trim(Mid$(DATA, t, i - t))
          t = i + 1
          j = j + 1
       End If
    End If
Next i

StringTokenizer = j - 1
End Function

Public Function FileByteFormat(FileBytes As Long) As String
On Error Resume Next
Dim nFileNum As Integer
Dim TempNum As Single

If FileBytes > 0 Then
    ' Get file's length
    FileByteFormat = FileBytes / 1024
    
    ' Round number
    TempNum = FileByteFormat - Int(FileByteFormat)
    
    ' Use different scale according to the size of the file
    Select Case Val(FileByteFormat)
        Case Is > 1024 ' Use Mega Byte
            FileByteFormat = Format(FileByteFormat / 1000, "#.##MB")
        Case Else  ' Use Kilo Byte
            ' All values are to round up
            FileByteFormat = Format(FileByteFormat + (1 - TempNum), "###KB")
    End Select
Else
    FileByteFormat = "0KB"
End If

End Function
