VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "网讯浏览器"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16305
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   16305
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame settings 
      BackColor       =   &H00404040&
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   14400
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label info_box 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "关于"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   50
         Top             =   960
         Width           =   735
      End
      Begin VB.Label homeset_box 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "主页"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   49
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   1560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Image homeset 
         Height          =   450
         Left            =   120
         Picture         =   "Form1.frx":10CA
         Top             =   360
         Width           =   450
      End
      Begin VB.Image info 
         Height          =   450
         Left            =   120
         Picture         =   "Form1.frx":1BD4
         Top             =   960
         Width           =   450
      End
   End
   Begin VB.PictureBox plus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15240
      Picture         =   "Form1.frx":26DE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   52
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox close 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15720
      Picture         =   "Form1.frx":2E8C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   51
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox mem4 
      Height          =   270
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox mem3 
      Height          =   270
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox mem2 
      Height          =   270
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox mem1 
      Height          =   270
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox search_box 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Text            =   "请输入网址或搜索内容"
      Top             =   120
      Width           =   10215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      Tab             =   8
      TabsPerRow      =   10
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":363A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Web1"
      Tab(0).Control(2)=   "url_tab0_box"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":3656
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":3672
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "Form1.frx":368E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "Form1.frx":36AA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "Form1.frx":36C6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "Form1.frx":36E2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "Form1.frx":36FE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "Form1.frx":371A
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "Frame9"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "Form1.frx":3736
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame10"
      Tab(9).ControlCount=   1
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   -75120
         TabIndex        =   43
         Top             =   240
         Width           =   16215
         Begin VB.TextBox url_tab9_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin VB.TextBox box9 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   16095
         End
         Begin SHDocVwCtl.WebBrowser Web10 
            Height          =   7335
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   0
         TabIndex        =   39
         Top             =   240
         Width           =   16095
         Begin VB.TextBox box8 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   120
            Width           =   15975
         End
         Begin VB.TextBox url_tab8_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin SHDocVwCtl.WebBrowser Web9 
            Height          =   7335
            Left            =   0
            TabIndex        =   40
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   -75000
         TabIndex        =   35
         Top             =   240
         Width           =   16095
         Begin VB.TextBox box7 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   15975
         End
         Begin VB.TextBox url_tab7_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin SHDocVwCtl.WebBrowser Web8 
            Height          =   7335
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   8055
         Left            =   -75000
         TabIndex        =   31
         Top             =   240
         Width           =   16095
         Begin VB.TextBox box6 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   15975
         End
         Begin VB.TextBox url_tab6_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            TabIndex        =   33
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   7335
         End
         Begin SHDocVwCtl.WebBrowser Web7 
            Height          =   7335
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         ForeColor       =   &H00000000&
         Height          =   8055
         Left            =   -75000
         TabIndex        =   27
         Top             =   240
         Width           =   16095
         Begin VB.TextBox url_tab5_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin VB.TextBox box5 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   120
            Width           =   15975
         End
         Begin SHDocVwCtl.WebBrowser Web6 
            Height          =   7335
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   -75000
         TabIndex        =   21
         Top             =   240
         Width           =   16095
         Begin VB.TextBox url_tab4_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin VB.TextBox box4 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   120
            Width           =   15975
         End
         Begin SHDocVwCtl.WebBrowser Web5 
            Height          =   7335
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   -75000
         TabIndex        =   16
         Top             =   240
         Width           =   16095
         Begin VB.TextBox box3 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   120
            Width           =   15975
         End
         Begin VB.TextBox url_tab3_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin SHDocVwCtl.WebBrowser web4 
            Height          =   7335
            Left            =   0
            TabIndex        =   17
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8055
         Left            =   -75000
         TabIndex        =   12
         Top             =   240
         Width           =   16095
         Begin VB.TextBox url_tab2_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin VB.TextBox box2 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   120
            Width           =   15975
         End
         Begin SHDocVwCtl.WebBrowser web3 
            Height          =   7335
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   -75000
         TabIndex        =   8
         Top             =   240
         Width           =   16095
         Begin VB.TextBox url_tab1_box 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "。。。。。。"
            Top             =   7800
            Width           =   9255
         End
         Begin SHDocVwCtl.WebBrowser Web2 
            Height          =   7335
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   15975
            ExtentX         =   28178
            ExtentY         =   12938
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
            Location        =   ""
         End
         Begin VB.TextBox box1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   120
            Width           =   15975
         End
      End
      Begin VB.TextBox url_tab0_box 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   -75000
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "。。。。。。"
         Top             =   8040
         Width           =   9255
      End
      Begin SHDocVwCtl.WebBrowser Web1 
         Height          =   7335
         Left            =   -75000
         TabIndex        =   3
         Top             =   600
         Width           =   15975
         ExtentX         =   28178
         ExtentY         =   12938
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
         Location        =   ""
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   -75000
         TabIndex        =   5
         Top             =   240
         Width           =   16095
         Begin VB.TextBox box0 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   15975
         End
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   15480
      X2              =   15480
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6720
      TabIndex        =   46
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image forward 
      Height          =   450
      Left            =   1920
      Picture         =   "Form1.frx":3752
      Top             =   120
      Width           =   450
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00E0E0E0&
      X1              =   1800
      X2              =   1800
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Image back 
      Height          =   450
      Left            =   1320
      Picture         =   "Form1.frx":425C
      Top             =   120
      Width           =   450
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Image refresh 
      Height          =   450
      Left            =   720
      Picture         =   "Form1.frx":4D66
      Top             =   120
      Width           =   450
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   600
      X2              =   600
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   13680
      X2              =   13680
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Image set 
      Height          =   450
      Left            =   15600
      Picture         =   "Form1.frx":5870
      Top             =   120
      Width           =   450
   End
   Begin VB.Image Home 
      Height          =   450
      Left            =   120
      Picture         =   "Form1.frx":637A
      Top             =   120
      Width           =   450
   End
   Begin VB.Label GO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12840
      TabIndex        =   48
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************** ______  ____                __  __
'*                                    * |    | \ | |  \  \ | |||--  || |||
'*       该程序由清遥Singal制作       * |\/\/| _ | |   \--|| ||| \  _____\
'*                                    * |/\/\| |-+-|  --\/|| /---\   /  \
'*https://space.bilibili.com/314017356* |    | | | |   //\|| || ||  __ __
'*                                    * |   \| |/| |/ /   \|  / |_/ |_||_|
'**************************************
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Sub close_Click()
Dim b As Integer
b = SSTab1.Tab
c = "web" & Str(b + 1)
d = Replace(c, " ", "")
Controls(Trim(d)).Navigate "http:///"
SSTab1.TabVisible(b) = False
End Sub

Private Sub Form_Load()
mem2.Text = "0"
Label1.Caption = "你找到了彩蛋"
Unload Form2
Unload frmAbout
Set Form2 = Nothing
Set frmAbout = Nothing
SSTab1.TabCaption(0) = "导航已取消"
SSTab1.TabCaption(1) = "导航已取消"
SSTab1.TabCaption(2) = "导航已取消"
SSTab1.TabCaption(3) = "导航已取消"
SSTab1.TabCaption(4) = "导航已取消"
SSTab1.TabCaption(5) = "导航已取消"
SSTab1.TabCaption(6) = "导航已取消"
SSTab1.TabCaption(7) = "导航已取消"
SSTab1.TabCaption(8) = "导航已取消"
SSTab1.TabCaption(9) = "导航已取消"
SSTab1.Tab = 0
SSTab1.TabVisible(0) = True
SSTab1.TabVisible(1) = False
SSTab1.TabVisible(2) = False
SSTab1.TabVisible(3) = False
SSTab1.TabVisible(4) = False
SSTab1.TabVisible(5) = False
SSTab1.TabVisible(6) = False
SSTab1.TabVisible(7) = False
SSTab1.TabVisible(8) = False
SSTab1.TabVisible(9) = False
settings.Visible = False
Form1.box0.Locked = True
Form1.box1.Locked = True
Form1.box2.Locked = True
Form1.box3.Locked = True
Form1.box4.Locked = True
Form1.box5.Locked = True
Form1.box6.Locked = True
Form1.box7.Locked = True
Form1.box8.Locked = True
Form1.box9.Locked = True
Form1.url_tab0_box.Locked = True
Form1.url_tab1_box.Locked = True
Form1.url_tab2_box.Locked = True
Form1.url_tab3_box.Locked = True
Form1.url_tab4_box.Locked = True
Form1.url_tab5_box.Locked = True
Form1.url_tab6_box.Locked = True
Form1.url_tab7_box.Locked = True
Form1.url_tab8_box.Locked = True
Form1.url_tab9_box.Locked = True
Form1.search_box.Locked = False
Web1.Silent = True
Web2.Silent = True
web3.Silent = True
web4.Silent = True
Web5.Silent = True
Web6.Silent = True
Web7.Silent = True
Web8.Silent = True
Web9.Silent = True
Web10.Silent = True
End Sub

Private Sub GO_Click()
    If search_box.Text = "" Then
        search_box.Text = "请输入网址或搜索内容"
    Else
        Dim ADDR As String
            ADDR = search_box.Text
        Dim SEAR As String
            SEAR = "https://www.baidu.com/s?wd=" & search_box.Text & "&rsv_spt=1&issp=1&rsv_bp=0&tn=baiduhome_pg&rsv_sug3=3&rsv_sug4=85&rsv_sug1=3&rsv_sug2=0&inputT=1861"
        Dim tabs As Integer
            tabs = SSTab1.Tab
        Dim tabelse As Integer
            tabelse = 9 - tabs
        For i = 0 To 9
            If i = 9 Then
                MsgBox "标签页过多,关掉一些吧"
                Exit For
            ElseIf SSTab1.TabVisible(i) = False Or SSTab1.TabCaption(i) = "导航已取消" Or SSTab1.TabCaption(i) = "地址无效" Then
                webref = "web" + Str(i + 1)
                webrep = Replace(webref, " ", "")
                SSTab1.TabVisible(i) = True
                SSTab1.Tab = i
                If InStr(LCase(search_box.Text), ".") > 0 Then
                    Controls(Trim(webrep)).Navigate ADDR
                    Exit For
                Else
                    Controls(Trim(webrep)).Navigate SEAR
                    Exit For
                End If
            Else
            End If
        Next i
    End If
End Sub

Private Sub Home_Click()
If mem1.Text = "" Or mem1.Text = "0" Then
    Dim tabs As Integer
    tabs = SSTab1.Tab
    webref = "web" + Str(tabs + 1)
    webrep = Replace(webref, " ", "")
    Controls(Trim(webrep)).GoHome
End If
End Sub

Private Sub homeset_Click()
If mem2.Text = "0" Then
    search_box.Text = "请输入您想设置的主页，结束时再次点击设置-主页"
    mem2.Text = "1"
ElseIf InStr(LCase(search_box.Text), ".") > 0 Then
    Dim hKey As Long, S As String, M As Integer
    S = search_box.Text
    M = Len(S) + 1
    RegCreateKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main\Start", hKey
    RegSetValueEx hKey, "Start Page", 0, REG_SZ, ByVal S, M
    mem2.Text = "0"
    search_box.Text = "请输入网址或搜索内容"
    MsgBox "设置成功"
    Dim e As Integer
    e = SSTab1.Tab
    e = e + 1
    Controls(Trim("web" & e)).GoHome
Else
    search_box.Text = "请输入有效网址，结束时再次点击设置-主页"
    mem2.Text = "1"
End If
End Sub

Private Sub info_Click()
settings.Visible = False
Dim times As Integer
times = Val(mem4.Text)
times = times + 1
If times >= 5 Then
    back_mem_4 = MsgBox("是否确定开启调试模式？" & Chr(13) & Chr(10) & "注意：更改调试模式中任何数值皆可能令程序崩溃，点击“是”表示您愿意承担调试模式带来的不稳定，如拒绝，请点击“否”", 4 + 48, "warning!!!")
        If back_mem_4 = vbYes Then
            Load Form2
            SSTab1.TabVisible(0) = True
            SSTab1.TabVisible(1) = True
            SSTab1.TabVisible(2) = True
            SSTab1.TabVisible(3) = True
            SSTab1.TabVisible(4) = True
            SSTab1.TabVisible(5) = True
            SSTab1.TabVisible(6) = True
            SSTab1.TabVisible(7) = True
            SSTab1.TabVisible(8) = True
            SSTab1.TabVisible(9) = True
            mem1.Visible = True
            mem2.Visible = True
            mem3.Visible = True
            mem4.Visible = True
            mem1.Locked = False
            settings.Visible = True
            Form2.Visible = True
            times = 0
            Else
                times = 0
                Load frmAbout
                frmAbout.Visible = True
                settings.Visible = False
        End If
    Else
        mem4.Text = times
        Load frmAbout
        frmAbout.Visible = True
        settings.Visible = False
End If
End Sub

Private Sub Label1_Click()
Shell "explorer.exe ""https://space.bilibili.com/314017356"
End Sub

Private Sub plus_Click()
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Controls(Trim(webrep)).GoHome
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub refresh_Click()
Dim tabs As Integer
tabs = SSTab1.Tab
webref = "web" & Str(tabs + 1)
webrep = Replace(webref, " ", "")
Controls(Trim(webrep)).refresh
End Sub

Private Sub search_box_Click()
search_box.Text = ""
End Sub

Private Sub set_Click()
If settings.Visible = True Then
    settings.Visible = False
    Else
        settings.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
myexit = MsgBox("您是否确定退出？", vbYesNo + vbDefaultButton2 + vbQuestion, "退出确认...")
If myexit = vbNo Then
    Cancel = True
    Else
        Unload Form2
        Set Form2 = Nothing
        Unload frmAbout
        Set Form2 = Nothing
        Unload Me
        Set Form1 = Nothing
End If
End Sub

Private Sub Web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab0_box.Text = Web1.LocationURL
End Sub

Private Sub Web2_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab1_box.Text = Web2.LocationURL
End Sub

Private Sub Web3_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab2_box.Text = web3.LocationURL
End Sub

Private Sub Web4_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab3_box.Text = web4.LocationURL
End Sub

Private Sub Web5_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab4_box.Text = Web5.LocationURL
End Sub

Private Sub Web6_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab5_box.Text = Web6.LocationURL
End Sub

Private Sub Web7_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab6_box.Text = Web7.LocationURL
End Sub

Private Sub Web8_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab7_box.Text = Web8.LocationURL
End Sub

Private Sub Web9_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab8_box.Text = Web9.LocationURL
End Sub

Private Sub Web10_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    url_tab9_box.Text = Web10.LocationURL
End Sub

Private Sub Web1_DownloadComplete()
    SSTab1.TabCaption(0) = Left(Web1.Document.Title, 6)
    box0.Text = Web1.Document.Title
End Sub

Private Sub Web2_DownloadComplete()
    SSTab1.TabCaption(1) = Left(Web2.Document.Title, 6)
    box1.Text = Web2.Document.Title
End Sub

Private Sub Web3_DownloadComplete()
    SSTab1.TabCaption(2) = Left(web3.Document.Title, 6)
    box2.Text = web3.Document.Title
End Sub

Private Sub Web4_DownloadComplete()
    SSTab1.TabCaption(3) = Left(web4.Document.Title, 6)
    box3.Text = web4.Document.Title
End Sub

Private Sub Web5_DownloadComplete()
    SSTab1.TabCaption(4) = Left(Web5.Document.Title, 6)
    box4.Text = Web5.Document.Title
End Sub

Private Sub Web6_DownloadComplete()
    SSTab1.TabCaption(5) = Left(Web6.Document.Title, 6)
    box5.Text = Web6.Document.Title
End Sub

Private Sub Web7_DownloadComplete()
    SSTab1.TabCaption(6) = Left(Web7.Document.Title, 6)
    box6.Text = Web7.Document.Title
End Sub

Private Sub Web8_DownloadComplete()
    SSTab1.TabCaption(7) = Left(Web8.Document.Title, 6)
    box7.Text = Web8.Document.Title
End Sub

Private Sub Web9_DownloadComplete()
    SSTab1.TabCaption(8) = Left(Web9.Document.Title, 6)
    box8.Text = Web9.Document.Title
End Sub

Private Sub Web10_DownloadComplete()
    SSTab1.TabCaption(9) = Left(Web10.Document.Title, 6)
    box9.Text = Web10.Document.Title
End Sub

Private Sub Web1_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web2_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web3_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web4_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web5_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web6_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web7_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web8_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web9_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub

Private Sub Web10_NewWindow2(ppDisp As Object, Cancel As Boolean)
For a = 0 To 9
    If a = 9 Then
        MsgBox "标签页过多，关掉一些吧。"
        Exit For
    ElseIf SSTab1.TabVisible(a) = False Then
        SSTab1.TabVisible(a) = True
        webref = "web" + Str(a + 1)
        webrep = Replace(webref, " ", "")
        Set ppDisp = Controls(Trim(webrep)).Application
        SSTab1.Tab = a
        Exit For
        Else
    End If
Next a
End Sub
