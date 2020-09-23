VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplorer 
   Caption         =   "Explorer"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   9060
   Icon            =   "Explorer.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   6540
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   794
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   41
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BackFolder"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewB"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewS"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewL"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button41 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ViewD"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   4800
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2895
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":058A
            Key             =   "DOS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":08DE
            Key             =   "Drive"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0C32
            Key             =   "avi"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":0F86
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":1412
            Key             =   "bmp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":1766
            Key             =   "cpp"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":1ABA
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":1E0E
            Key             =   "dat"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":2162
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":24B6
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":280A
            Key             =   "frm"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":2B5E
            Key             =   "gif"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":2EB2
            Key             =   "h"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":3206
            Key             =   "hlp"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":355A
            Key             =   "inf"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":38AE
            Key             =   "Others"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":3C02
            Key             =   "sys"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":3F56
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":42AA
            Key             =   "vbg"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":45FE
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":4952
            Key             =   "vbp"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":4CA6
            Key             =   "bas"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":4FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":534E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":56A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":59F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":5D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":609E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":63F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":6746
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":6A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":6DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":7142
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Explorer.frx":7496
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar3 
      Height          =   375
      Left            =   3405
      TabIndex        =   4
      Top             =   435
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6165
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11165
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "All Folders"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   3400
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7858
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7858
      _Version        =   393217
      Indentation     =   471
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuTool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect To Network Drive"
      End
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect  Network Drive"
      End
      Begin VB.Menu mnuline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnugo 
         Caption         =   "Go ->"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "?"
   End
End
Attribute VB_Name = "frmExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'Explorer Program Made By Carl Harvey.
'Date  : 199-09-24
'Planet Nick : Carlos
'Thanks for keeping these line here
'***************************************

Dim SizeOn As Boolean
Dim OldX, InitialFormWith, InitialFormHeigth
Dim MoveOn As Boolean
Public Sub GetTreeStructure(ByVal path As String, ByVal ftype As String, ByVal NodeTo As String)
       Dim hFile As Long, ts As String, WFD As WIN32_FIND_DATA
       Dim result As Long, sAttempt As String, szPath As String
       Dim strtemp
       Dim nod1 As Node
       szPath = path & "*.*" & Chr$(0)
       hFile = FindFirstFile(szPath, WFD)
       Do
         If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
          'Hey look, we've got a directory!
             ts = StripNull(WFD.cFileName)
             If Not (ts = "." Or ts = "..") Then
                 'Don't look for hidden or system directories
                 If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
                    rep = InStr(1, WFD.cFileName, Chr(0), vbBinaryCompare)
                    strtemp = Mid(WFD.cFileName, 1, rep - 1)
                     Dim str2 As String
                    str2 = "R" & strtemp & TreeView1.Nodes.Count
                    With TreeView1
                      Set nod1 = .Nodes.Add(NodeTo, tvwChild, str2, strtemp)
                    End With
                    nod1.Image = "DOS"
                 End If
             End If
           End If
           WFD.cFileName = ""
           result = FindNextFile(hFile, WFD)
        Loop Until result = 0
       FindClose hFile
End Sub
'**********************************************
'* Function ChowFromFolder is From Planet-Source-Code
'* Modified by Carlos 09-10-99
'* Modified by Carlos 09-24-99
'***********************************************
Private Sub ChowFromFolder(ByVal zpath As String, ByVal FileType As String)
       Dim hFile As Long, result As Long, szPath As String
       Dim WFD As WIN32_FIND_DATA
       Dim TMP As ListItem
       Dim pos1
       ListView1.SortOrder = 0 ' Set to Icon view
       szPath = zpath & FileType & Chr$(0)
       'Start asking windows for files.
       hFile = FindFirstFile(szPath, WFD)
       Do
           ts = StripNull(WFD.cFileName)
           If Not (ts = "." Or ts = "..") Then
             
             If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                'if its a folder
                rep = "DOS"
             Else
                pos1 = InStr(1, WFD.cFileName, ".", vbBinaryCompare)
                rep = GetImage(Mid(WFD.cFileName, pos1 + 1, 3))
             End If
             pos1 = InStr(1, WFD.cFileName, Chr$(0), vbBinaryCompare)
             If Trim(Mid(WFD.cFileName, 1, pos1 - 1)) <> "" Then
                Set TMP = ListView1.ListItems.Add(, , Trim(WFD.cFileName), rep, rep)
                If rep <> "DOS" Then TMP.SubItems(1) = WFD.nFileSizeLow / 1000 & " Kb   "
                Dim strtemp As Variant
                On Error Resume Next
                strtemp = WFD.ftCreationTime.dwHighDateTime
                TMP.SubItems(3) = strtemp
                strtemp = "File"
                If WFD.dwFileAttributes = 16 Then strtemp = "Folder"
                TMP.SubItems(2) = strtemp
             End If
           End If
             WFD.cFileName = ""
             result = FindNextFile(hFile, WFD)
       Loop Until result = 0
       FindClose hFile
End Sub
Private Function GetImage(ByVal imgstr As String) As String
Select Case LCase(imgstr)
Case "zip", "dll", "inf", "exe", "bas", "gif", "txt", "dat", "bat", "hlp", "bmp", "led", "frm", "vbg", "h", "cpp", "vbp", "ctl", "avi": GetImage = LCase(imgstr) 'Normal case
Case "jpg": GetImage = "gif"           'if jpg same icon as gif
Case "com": GetImage = "exe"           'if com same icon as exe
Case "mpg", "mov": GetImage = "avi"    'if mpg or mov same icon as avi
Case "ini": GetImage = "inf"           'if ini same icon as inf
Case Else: GetImage = "Others"         'All other files
End Select
End Function

Private Sub PutInTree(ByVal Node As String, ByVal nb)
With TreeView1
  myfullpath = Mid(.Nodes.Item(Node).FullPath, 9, Len(.Nodes.Item(Node).FullPath) - 8) & "\"
  If myfullpath <> "a:\" Then 'dont automaticly check a:
   If Not .Nodes.Item(Node).Children > 0 Then 'if not already explored
     GetTreeStructure myfullpath, "*.*", .Nodes.Item(Node).Key
   End If
  End If
  If nb = 1 Then Exit Sub 'End of recursivity
  '***********************************
  'Recursive call
   PutInTree .Nodes.Item(Node).Next.Key, nb - 1
End With
End Sub

Private Sub GetFreeSpace(ByVal pathn As String)
       'Calls other functions to provide the info.
       'Data is stored in my own user-defined type.
       Dim RDI As RANDYS_OWN_DRIVE_INFO
       Dim r As Long, nbtemp As Integer
       r& = rgbGetDiskFreeSpaceRDI(pathn$, RDI)
       'show the results
       With ListView1
        For i = 1 To ListView1.ListItems.Count
         nbt = nbt + Val(.ListItems(i).ListSubItems(1))
        Next
      End With
       nbtv = 1
       If nbt > 1000 Then nbtv = 1000
       If nbt = 0 Then
         StatusBar2.Panels(2).Text = Format$(nbt, "###,##0") & IIf(nbtv = 0, " Kb", " Mb") & "  ( Drive Space Free : " & Format$(RDI.DrvSpaceFree / 1000000, "###0") & " Mb )"
       Else
         StatusBar2.Panels(2).Text = Format$(nbt / nbtv, "###,##0") & IIf(nbtv = 0, " Kb", " Mb") & "  ( Drive Space Free : " & Format$(RDI.DrvSpaceFree / 1000000, "###0") & " Mb )"
       End If
 End Sub

Private Sub Form_Load()
Dim nod1 As Node
Set nod1 = TreeView1.Nodes.Add(, , "RDesktop", "Desktop")
 nod1.Image = "DOS"
For i = 0 To Drive1.ListCount - 1
 Set nod1 = TreeView1.Nodes.Add("RDesktop", tvwChild, "R" & Mid(Drive1.List(i), 1, 2) & TreeView1.Nodes.Count, Mid(Drive1.List(i), 1, 2))
 nod1.Image = "Drive"
Next
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If SizeOn Then
 OldX = x
 MoveOn = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MoveOn Then
 'moveon = True
 If x <> OldX Then
  TreeView1.Width = TreeView1.Width - (OldX - x)
  StatusBar1.Width = TreeView1.Width
  
  ListView1.Left = TreeView1.Width + 40
  ListView1.Width = Me.Width
 StatusBar3.Left = ListView1.Left
 StatusBar3.Width = ListView1.Width
 End If
 OldX = x
End If
If x = TreeView1.Width Or x = TreeView1.Width + 15 And Not SizeOn Then
 Me.MousePointer = 9
 SizeOn = True
ElseIf SizeOn Then
 Me.MousePointer = 0
 SizeOn = False
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MoveOn = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = ListView1.Width + (Width - (ListView1.Left + ListView1.Width)) - 80
StatusBar3.Width = ListView1.Width
ListView1.Height = (Me.Height - 2030)
TreeView1.Height = (Me.Height - 2030)
End Sub

Private Sub Form_Terminate()
End
End Sub


Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If SizeOn Then
 Me.MousePointer = 0
 SizeOn = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "ViewD": ListView1.View = lvwReport
    Case "ViewL": ListView1.View = lvwList
    Case "ViewS": ListView1.View = lvwSmallIcon
    Case "ViewB": ListView1.View = lvwIcon
    Case "BackFolder"
   If TreeView1.SelectedItem.Parent.Text <> "Desktop" Then
     TreeView1.SelectedItem.Parent.Selected = True
     ListView1.ListItems.Clear
     StatusBar3.SimpleText = "Content Of : " & Mid(TreeView1.SelectedItem.FullPath, 9, Len(TreeView1.SelectedItem.FullPath) - 8)
     myfullpath = Mid(TreeView1.SelectedItem.FullPath, 9, Len(TreeView1.SelectedItem.FullPath) - 8) & "\"
     ' Expand the Treeview object with sub folders
     ChowFromFolder myfullpath, "*.*"
     StatusBar2.Panels(1).Text = ListView1.ListItems.Count & "  Objects found"
     GetFreeSpace Mid(myfullpath, 1, 3)
End If
 
End Select

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
With Node
 .Sorted = True
 PutInTree .Child.Key, .Children
End With
End Sub
'**************************************************************************************
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
' When a ColumnHeader object is clicked, the ListView control is
' sorted by the subitems of that column.
' Set the SortKey to the Index of the ColumnHeader - 1
If ListView1.SortOrder = 0 Then
 ListView1.SortOrder = 1
 Else   ' Set Sorted to True to sort the list.
 ListView1.SortOrder = 0
End If
 ListView1.Sorted = True
End Sub
Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If SizeOn Then
 Me.MousePointer = 0
 SizeOn = False
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Text <> "Desktop" Then
   ListView1.ListItems.Clear
   StatusBar3.SimpleText = "Content Of : " & Mid(Node.FullPath, 9, Len(Node.FullPath) - 8)
   myfullpath = Mid(Node.FullPath, 9, Len(Node.FullPath) - 8) & "\"
   'Expand the Treeview object with sub folders
   ChowFromFolder myfullpath, "*.*"
   StatusBar2.Panels(1).Text = Format(ListView1.ListItems.Count, "###,##0") & "  Objects found"
   GetFreeSpace Mid(myfullpath, 1, 3)
End If
End Sub

'****************************************************




'*********************************************



Private Function rgbGetDiskFreeSpaceRDI(RootPathName$, RDI As RANDYS_OWN_DRIVE_INFO) As Long
       'returns data about the selected drive.
       'Passed is the RootPathName$; the other
       'variables are filled in here.
       Dim r As Long
       
       r& = GetDiskFreeSpace(RootPathName$, RDI.DrvSectors, RDI.DrvBytesPerSector, RDI.DrvFreeClusters, RDI.DrvTotalClusters)
       
       RDI.DrvSpaceTotal = (RDI.DrvSectors * RDI.DrvBytesPerSector * RDI.DrvTotalClusters)
       RDI.DrvSpaceFree = (RDI.DrvSectors * RDI.DrvBytesPerSector * RDI.DrvFreeClusters)
       RDI.DrvSpaceUsed = RDI.DrvSpaceTotal - RDI.DrvSpaceFree
       
       rgbGetDiskFreeSpaceRDI& = r&
   End Function
Private Function rgbGetLogicalDriveStrings() As String
       'returns a single string of available drive
       'letters, each separated by a space
       '(i.e. a:\ c:\ d:\), suitable for display
       Dim r As Long
       Dim i As Integer
       Dim lpBuffer As String
       
       lpBuffer$ = Space$(64)
       
       r& = GetLogicalDriveStrings(Len(lpBuffer$), lpBuffer$)
       
       lpBuffer$ = Trim$(lpBuffer$)
       rgbGetLogicalDriveStrings = lpBuffer$
   End Function
Private Function StripNulls(startStrg$) As String
       'Take a string separated by a Chr$(0), and split off 1 item, and
       'shorten the string so that the next item is ready for removal.
       Dim c As Integer
       Dim Item As String
       c% = 1
       Do
           If Mid$(startStrg$, c%, 1) = Chr$(0) Then
               
               Item$ = Mid$(startStrg$, 1, c% - 1)
               startStrg$ = Mid$(startStrg$, c% + 1, Len(startStrg$))
               StripNulls$ = Item$
               Exit Function
           End If
           c% = c% + 1
       Loop
End Function


