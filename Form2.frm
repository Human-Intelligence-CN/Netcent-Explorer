VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "调试模式"
   ClientHeight    =   2835
   ClientLeft      =   23235
   ClientTop       =   12000
   ClientWidth     =   3765
   LinkTopic       =   "Form2"
   ScaleHeight     =   2835
   ScaleWidth      =   3765
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Text            =   "更多功能尚未完善..."
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "状态栏"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton Command7 
         Caption         =   "状态栏数值归零"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "解锁所有状态栏"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "文本框"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
      Begin VB.CommandButton Command5 
         Caption         =   "恢复默认"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "解锁所有文本框"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "标签页"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton Command2 
         Caption         =   "仅显示tab0"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "隐藏整个sstab"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Command1_Click()
If Command1.Caption = "隐藏整个sstab" Then
Form1.SSTab1.Visible = False
Command1.Caption = "显示整个sstab"
Else
Form1.SSTab1.Visible = True
Command1.Caption = "隐藏整个sstab"
End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "仅显示tab0" Then
Form1.SSTab1.TabVisible(0) = True
Form1.SSTab1.TabVisible(1) = False
Form1.SSTab1.TabVisible(2) = False
Form1.SSTab1.TabVisible(3) = False
Form1.SSTab1.TabVisible(4) = False
Form1.SSTab1.TabVisible(5) = False
Form1.SSTab1.TabVisible(6) = False
Form1.SSTab1.TabVisible(7) = False
Form1.SSTab1.TabVisible(8) = False
Form1.SSTab1.TabVisible(9) = False
Form1.SSTab1.Visible = True
Command1.Caption = "隐藏整个sstab"
Command2.Caption = "显示所有标签页"
Else
Form1.SSTab1.TabVisible(0) = True
Form1.SSTab1.TabVisible(1) = True
Form1.SSTab1.TabVisible(2) = True
Form1.SSTab1.TabVisible(3) = True
Form1.SSTab1.TabVisible(4) = True
Form1.SSTab1.TabVisible(5) = True
Form1.SSTab1.TabVisible(6) = True
Form1.SSTab1.TabVisible(7) = True
Form1.SSTab1.TabVisible(8) = True
Form1.SSTab1.TabVisible(9) = True
Form1.SSTab1.Visible = True
Command1.Caption = "隐藏整个sstab"
Command2.Caption = "仅显示tab0"
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "解锁所有文本框" Then
Form1.box0.Locked = False
Form1.box1.Locked = False
Form1.box2.Locked = False
Form1.box3.Locked = False
Form1.box4.Locked = False
Form1.box5.Locked = False
Form1.box6.Locked = False
Form1.box7.Locked = False
Form1.box8.Locked = False
Form1.box9.Locked = False
Form1.url_tab0_box.Locked = False
Form1.url_tab1_box.Locked = False
Form1.url_tab2_box.Locked = False
Form1.url_tab3_box.Locked = False
Form1.url_tab4_box.Locked = False
Form1.url_tab5_box.Locked = False
Form1.url_tab6_box.Locked = False
Form1.url_tab7_box.Locked = False
Form1.url_tab8_box.Locked = False
Form1.url_tab9_box.Locked = False
Form1.search_box.Locked = False
Text1.Locked = False
Command3.Caption = "锁定所有文本框"
Command5.Visible = True
Else
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
Form1.search_box.Locked = True
Text1.Locked = True
Command3.Caption = "解锁所有文本框"
Command5.Visible = True
End If
End Sub

Private Sub Command5_Click()
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
Text1.Locked = True
Command5.Visible = False
End Sub

Private Sub Command6_Click()
If Command6.Caption = "解锁所有状态栏" Then
Form1.mem1.Locked = False
Form1.mem2.Locked = False
Form1.mem3.Locked = False
Form1.mem4.Locked = False
Command6.Caption = "锁定所有状态栏"
Else
Form1.mem1.Locked = True
Form1.mem2.Locked = True
Form1.mem3.Locked = True
Form1.mem4.Locked = True
Command6.Caption = "解锁所有状态栏"
End If
End Sub

Private Sub Command7_Click()
Form1.mem1.Text = ""
Form1.mem2.Text = "0"
Form1.mem3.Text = "0"
Form1.mem4.Text = "0"
End Sub

Private Sub Form_Load()
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command5.Visible = False
Command6.Visible = True
Command7.Visible = True
Text1.Visible = True
Command1.Caption = "隐藏整个sstab"
Command2.Caption = "仅显示tab0"
Command3.Caption = "解锁所有文本框"
Command6.Caption = "解锁所有状态栏"
Text1.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
Form1.mem1.Visible = False
Form1.mem2.Visible = False
Form1.mem3.Visible = False
Form1.mem4.Visible = False
Form1.SSTab1.Visible = True
Form1.mem4.Text = "0"
Unload Me
Set Form2 = Nothing
End Sub
