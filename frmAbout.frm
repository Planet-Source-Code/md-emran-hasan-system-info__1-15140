VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B6C1EA38-375B-11D4-93AB-E7C32384627A}#3.0#0"; "FREELIB.OCX"
Object = "{BE516F6C-863D-11D2-BF9B-8AF16ECF9476}#1.4#0"; "ECLIPSEUI99.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4320
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "frmAbout.frx":0CCA
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox Picture5 
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   3795
      TabIndex        =   41
      Top             =   600
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3615
      ScaleWidth      =   3855
      TabIndex        =   11
      Top             =   60000
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2355
         ScaleWidth      =   3795
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
         Begin VB.Label lblDiskSpace 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   25
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblFreeVirtual 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   24
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblTotalVirtual 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   23
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblFreeMemory 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   22
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblTotalMemory 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   21
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblOperating 
            BackStyle       =   0  'Transparent
            Caption         =   "Windows 98 4.10.2222"
            Height          =   255
            Left            =   2100
            TabIndex        =   20
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Available Disk Space (C:)"
            Height          =   195
            Left            =   60
            TabIndex        =   19
            Top             =   1920
            Width           =   1905
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Available virtual memory"
            Height          =   195
            Left            =   60
            TabIndex        =   18
            Top             =   1560
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Total virtual memory"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   1200
            Width           =   1590
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Available physical memory"
            Height          =   195
            Left            =   60
            TabIndex        =   16
            Top             =   840
            Width           =   1995
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Total physical memory"
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   480
            Width           =   1710
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Operating System"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   120
            Width           =   1410
         End
      End
      Begin MSComctlLib.ListView lv 
         Height          =   975
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1720
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList"
         ForeColor       =   16777215
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2355
         ScaleWidth      =   3795
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Windows version"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Registered Owner"
            Height          =   195
            Left            =   60
            TabIndex        =   37
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Registered Organisation"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   840
            Width           =   1860
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Current Resolution"
            Height          =   195
            Left            =   60
            TabIndex        =   35
            Top             =   1200
            Width           =   1470
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• Windows directory"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "• System directory"
            Height          =   195
            Left            =   60
            TabIndex        =   33
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblWinVer 
            BackStyle       =   0  'Transparent
            Caption         =   "Windows 98 4.10.2222"
            Height          =   255
            Left            =   2100
            TabIndex        =   32
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblOwner 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   31
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblOrg 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   30
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblRes 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   29
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lblWinDir 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   28
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblSysDir 
            BackStyle       =   0  'Transparent
            Caption         =   "64928 KB"
            Height          =   255
            Left            =   2100
            TabIndex        =   27
            Top             =   1920
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3615
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   48000
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "&Register"
         Height          =   350
         Left            =   2640
         TabIndex        =   40
         Top             =   1500
         Width           =   1095
      End
      Begin EclipseUI99.EclipseLink EclipseLink1 
         Height          =   255
         Left            =   555
         Top             =   3120
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Click here to visit Tools & Utilities Web Site"
         URL             =   "www.emranhasan.com"
      End
      Begin VB.Label Label6 
         Caption         =   $"frmAbout.frx":0D3D
         Height          =   975
         Left            =   600
         TabIndex        =   10
         Top             =   2040
         Width           =   3255
      End
      Begin EclipseUI99.EclipseDivider EclipseDivider2 
         Height          =   30
         Left            =   600
         TabIndex        =   9
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSerial 
         Caption         =   "www.emranhasan.com"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblUser 
         Caption         =   "Emran Hasan"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "This product is licensed to:"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   1905
      End
      Begin EclipseUI99.EclipseDivider EclipseDivider1 
         Height          =   30
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Copyright © 2000-2001, Md Emran Hasan."
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   3090
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmAbout.frx":0E00
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tools && Utilities 2001"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   2085
      End
   End
   Begin FreeLibSrc.FreeLib F2 
      Height          =   480
      Left            =   1560
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IniFile         =   ""
      IniSection      =   ""
      IniSize         =   ""
      IniDefault      =   ""
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":1ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":23A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "System Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Credits"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   4440
      Width           =   1335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As New AllInOne
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Function to retrive the Windows Directory
Function GetWindowsDir() As String
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If
End Function
Private Sub SetSysInformation()
Dim WinDir As String
Dim temp2 As Long

temp2 = f.GetPhysMemTotal / 1024
lblTotalMemory.Caption = Format(temp2, "##,###") & " KB"
temp2 = f.GetPhysMemAvailable / 1024
lblFreeMemory.Caption = Format(temp2, "##,###") & " KB"

temp2 = f.GetVirtMemTotal / 1024
lblTotalVirtual.Caption = temp2 & " KB"
temp2 = f.GetVirtMemAvailable / 1024
lblFreeVirtual.Caption = temp2 & " KB"

Dim temp3 As String

Dim temp4 As String
temp4 = f.GetWinResXY
lblRes.Caption = temp4

GetWinVer
lblWinVer.Caption = wVer
lblOperating.Caption = wVer

lblOwner.Caption = modRegistry.RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
lblOrg.Caption = modRegistry.RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")

WinDir = GetWindowsDir

lblWinDir.Caption = WinDir
lblSysDir.Caption = GetSystemDirectory

Dim temp5
temp5 = f.GetHdiskSpace("C:\", True)
lblDiskSpace.Caption = temp5
End Sub
'Function 2 add backslash
Public Function AddBackslash(S As String) As String
If Len(S) > 0 Then
If Right$(S, 1) <> "\" Then
AddBackslash = S + "\"
Else
AddBackslash = S
End If
Else
AddBackslash = "\"
End If
End Function
'Function to retrive the System directory
Public Function GetSystemDirectory() As String
   Dim S As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   S = Space(i)
   Call GetSystemDirectoryA(S, i)
   GetSystemDirectory = AddBackslash(Left$(S, i - 1))
End Function

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmRegister.Show 1
Dim Name As String
Dim code As String

Name = GetSetting(App.Title, "Settings", "Name")
code = GetSetting(App.Title, "Settings", "Code")

End Sub

Private Sub Form_Load()
Dim Name As String
Dim code As String

Name = GetSetting(App.Title, "Settings", "Name")
code = GetSetting(App.Title, "Settings", "Code")

lblUser.Caption = Cipher.decrypt(Name, "1234567890")
lblSerial.Caption = Cipher.decrypt(code, "0987654321")

lv.ListItems.Add , "Comp", "System Info", 1
lv.ListItems.Add , "Win", "Windows Info", 2

SetSysInformation

End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lv.SelectedItem.Key = "Comp" Then
    Picture3.Visible = True
    Picture4.Visible = False
ElseIf lv.SelectedItem.Key = "Win" Then
    Picture3.Visible = False
    Picture4.Visible = True
End If
End Sub

Private Sub Tab1_Click()
If Tab1.SelectedItem.Caption = "About" Then
    Picture1.Visible = True
    Picture2.Visible = False
Else
    Picture1.Visible = False
    Picture2.Visible = True
End If
End Sub
