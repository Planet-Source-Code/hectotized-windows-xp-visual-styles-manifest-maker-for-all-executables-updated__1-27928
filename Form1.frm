VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XP Manifest Maker"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tbxPath 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   5760
      ScaleHeight     =   4095
      ScaleWidth      =   5175
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   5175
      Begin ComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Application's Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Application Path"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Restore"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   3600
         Width           =   1695
      End
   End
   Begin VB.TextBox tbx2 
      Height          =   2895
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frx":1CFA
      Top             =   8040
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   240
      ScaleHeight     =   4095
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   600
      Width           =   5175
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add To Right-Click Context Menu (Windows Explorer)"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit XP Manifest Maker After Creating Manifest."
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3000
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run Application After Creating Manifest."
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox tbxApp 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Make Manifest"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   3600
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application's Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1350
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tbx1 
      Height          =   1605
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":1EA6
      Top             =   7560
      Width           =   6015
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
      TabWidthStyle   =   2
      TabFixedWidth   =   2999
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Make Manifest"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Restore"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FSO As New fileSystemObject

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub Check3_Click()
SaveRegLong HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker", "ContxMenu", Check3.Value


If Check3.Value = 1 Then
SaveRegString HKEY_CLASSES_ROOT, "exefile\shell\Disable Visual Styles\Command", "", App.Path & App.EXEName & ".exe /D%L\"""
SaveRegString HKEY_CLASSES_ROOT, "exefile\shell\Enable Visual Styles\Command", "", App.Path & App.EXEName & ".exe /E%L\"""

Else
DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Disable Visual Styles\Command"
DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Disable Visual Styles"
DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Enable Visual Styles\Command"
DeleteKey HKEY_CLASSES_ROOT, "exefile\shell\Enable Visual Styles"
End If
End Sub

Private Sub Command1_Click()






    With CommonDialog1
    .DialogTitle = "Choose Any .exe"
    .CancelError = False
    .FileName = ""
    .InitDir = "C:\"
    .Filter = "VB Executible|*.exe"
    .MaxFileSize = 32000
    .ShowOpen
    
    
    End With
    
    tbxApp.Text = CommonDialog1.FileTitle
    tbxPath.Text = CommonDialog1.FileName
    
    If tbxApp.Text = "" Then
    Command2.Enabled = False
    Else
    Command2.Enabled = True
    End If
    
End Sub

Private Sub Command2_Click()
If tbxApp.Text = "" Or tbxPath.Text = "" Then

MsgBox "Not all Fields are populated"
Exit Sub
End If


If FSO.FileExists(tbxPath.Text & ".manifest") = True Then
FSO.DeleteFile (tbxPath.Text & ".manifest")
End If


Open tbxPath.Text & ".manifest" For Append As 1

Print #1, tbx1.Text & tbxApp.Text & """"
Print #1, tbx2.Text




Close 1
'save to manifest settings
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & tbxApp.Text, "FileName", tbxApp.Text
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & tbxApp.Text, "FilePath", tbxPath.Text

'save to win reg
SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", tbxPath.Text, "WIN2000"

'load info to listview
With ListView1.ListItems.Add(, , Left(tbxApp.Text, Len(tbxApp.Text) - 4))
    .SubItems(1) = tbxPath.Text
End With

 
 If Check1.Value = 1 Then Shell tbxPath.Text, vbNormalFocus

tbxApp.Text = ""
tbxPath.Text = ""
Command2.Enabled = False

Command3.Enabled = True
If Check2.Value = 1 Then End
End Sub
 

Private Sub tbxDescription_Change()
 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next

If ListView1.SelectedItem.Selected = False Then
MsgBox "Please Select Application To Restore"
Exit Sub
End If


'Delete  from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)

'delete manifest
FSO.DeleteFile (ListView1.SelectedItem.SubItems(1) & ".manifest")

'Delete from listview
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

If ListView1.ListItems.Count = 0 Then Command3.Enabled = False

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()


    
    Dim cmd, cmFileName As String, cmCommand

On Error Resume Next
'''this code is here for the context menus
cmd = Trim$(Command$)
cmd = Right(cmd, Len(cmd) - 1)
cmd = Left(cmd, Len(cmd) - 1)

cmCommand = Left(cmd, 1)
cmFileName = Right(cmd, Len(cmd) - 1)
cmFileName = Left(cmFileName, Len(cmFileName) - 1)

If cmd > "" Then
Select Case cmCommand

Case "E"

If FSO.FileExists(cmFileName & ".manifest") = True Then
FSO.DeleteFile (cmFileName & ".manifest")
End If


Open cmFileName & ".manifest" For Append As 1
'fix this
Print #1, tbx1.Text & ShortFileName(cmFileName) & """"
Print #1, tbx2.Text
Close 1
'save to manifest settings
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FileName", ShortFileName(cmFileName)
SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FilePath", cmFileName



'save to win reg
SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName, "WIN2000"

Case "D"
'Delete  from registry
DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName)
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName

'delete manifest
FSO.DeleteFile (cmFileName & ".manifest")

End Select

End

End If

















Dim fApp
Dim rApp, fName, fPath

Picture1.Left = 240
Picture2.Left = 240

Check3.Value = GetRegLong(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker", "ContxMenu")

'load saved alerts to listview

countit = CountRegKeys(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests")


For i = 0 To countit - 1
fApp = GetRegKey(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests", i)



fName = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & fApp, "FileName")
fPath = GetRegString(HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & fApp, "FilePath")




    With ListView1.ListItems.Add(, , Left(fName, Len(fName) - 4))
        .SubItems(1) = fPath
    End With


Next i

If ListView1.ListItems.Count = 0 Then Command3.Enabled = False

End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
Case 1
Picture1.Visible = True
Picture2.Visible = False
Case 2
Picture2.Visible = True
Picture1.Visible = False


End Select
End Sub

Public Function ShortFileName(ByVal sFileName As String) As String
    For i = 0 To Len(sFileName)
        If Left(Right(sFileName, i), 1) = "\" Then
        ShortFileName = Right(sFileName, i - 1)
        i = Len(sFileName)
        End If
    Next i
    
End Function

