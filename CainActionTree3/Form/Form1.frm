VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{50ACB2FB-FCAC-4FAC-8AB2-DB8564563320}#31.0#0"; "ActionTreeView.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command9 
      Caption         =   "Family Tree"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Tree and Info"
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Tree Only with Header"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tree Only without Header"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin ActionTreeView.ActionTree ActionTree1 
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Keine Icons"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Icons Small"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Icons Large"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   3360
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2368
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borderstyle 3D"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borderstyle Flat"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   4080
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8434
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E500
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EDDA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub ActionTree1_ItemClick(Button As Integer, Shift As Integer, NodeItem As ActionTreeView.Node)
    
    Me.Caption = NodeItem.Caption
    
End Sub

Private Sub Command1_Click()
    
    ActionTree1.BorderStyle = 0
    
End Sub

Private Sub Command2_Click()
    ActionTree1.BorderStyle = 1
End Sub

Private Sub Command3_Click()
    ActionTree1.Icons Nothing

End Sub

Private Sub Command4_Click()
    ActionTree1.Icons ImageList1(0)

End Sub

Private Sub Command5_Click()
    ActionTree1.Icons ImageList1(1)

End Sub

Private Sub Command6_Click()

    ActionTree1.Clear_Items
    ActionTree1.Clear_Columns

    ActionTree1.AT_Nodes.Add "zz", "", "Level 1", 1
    ActionTree1.AT_Nodes.Add "pu", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f34", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f33", "zz", "Level 2", 5
    
    ActionTree1.AT_Nodes.Add "sg", "f34", "Level 3", 3
    ActionTree1.AT_Nodes.Add "sfg", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "vbxc", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "sdfg", "vbxc", "Level 4", 1
    
    ActionTree1.AT_Nodes.Add "oo", "", "Alexander", 4
    
    ActionTree1.AT_Nodes.Add "pp", "oo", "The", 5
    
    ActionTree1.AT_Nodes.Add "ee", "", "Great", 3
    ActionTree1.AT_Nodes.Add "d345", "ee", "of", 2
    ActionTree1.AT_Nodes.Add "tt", "ee", "Greece", 1
    
    ActionTree1.AT_Nodes.Add "qq", "", "Spanish", 4
    
    ActionTree1.Refresh

End Sub

Private Sub Command7_Click()

    ActionTree1.Clear_Items
    ActionTree1.Clear_Columns

    ActionTree1.AT_Nodes.Add "zz", "", "Level 1", 1
    ActionTree1.AT_Nodes.Add "pu", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f34", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f33", "zz", "Level 2", 5
    
    ActionTree1.AT_Nodes.Add "sg", "f34", "Level 3", 3
    ActionTree1.AT_Nodes.Add "sfg", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "vbxc", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "sdfg", "vbxc", "Level 4", 1
    
    ActionTree1.AT_Nodes.Add "oo", "", "Alexander", 4
    
    ActionTree1.AT_Nodes.Add "pp", "oo", "The", 5
    
    ActionTree1.AT_Nodes.Add "ee", "", "Great", 3
    ActionTree1.AT_Nodes.Add "d345", "ee", "of", 2
    ActionTree1.AT_Nodes.Add "tt", "ee", "Greece", 1
    
    ActionTree1.AT_Nodes.Add "qq", "", "Spanish", 4

    ActionTree1.AT_Columns.Add 80, "Actions"
    
    ActionTree1.Refresh
    
End Sub

Private Sub Command8_Click()

    ActionTree1.Clear_Items
    ActionTree1.Clear_Columns
    
    ActionTree1.AT_Nodes.Add "zz", "", "Level 1", 1
    ActionTree1.AT_Nodes.Add "pu", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f34", "zz", "Level 2", 3
    ActionTree1.AT_Nodes.Add "f33", "zz", "Level 2", 5
    
    ActionTree1.AT_Nodes.Add "sg", "f34", "Level 3", 3
    ActionTree1.AT_Nodes.Add "sfg", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "vbxc", "f34", "Level 3", 2
    ActionTree1.AT_Nodes.Add "sdfg", "vbxc", "Level 4", 1
    
    ActionTree1.AT_Nodes.Add "oo", "", "Alexander", 4
    
    ActionTree1.AT_Nodes.Add "pp", "oo", "The", 5
    
    ActionTree1.AT_Nodes.Add "ee", "", "Great", 3
    ActionTree1.AT_Nodes.Add "d345", "ee", "of", 2
    ActionTree1.AT_Nodes.Add "tt", "ee", "Greece", 1
    
    ActionTree1.AT_Nodes.Add "qq", "", "Spanish", 4
    
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter" & vbCrLf & "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Nodes("qq").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("qq").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("tt").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Columns.Add 80, "Actions"
    ActionTree1.AT_Columns.Add 80, "Info"
    ActionTree1.AT_Columns.Add 80, "Status"
    ActionTree1.AT_Columns.Add 80, "Kommentar"
    
    
    ActionTree1.Refresh
    

End Sub

Private Sub Command9_Click()

    ActionTree1.Clear_Items
    ActionTree1.Clear_Columns
    
    ActionTree1.AT_Nodes.Add "zz", "", "Level 1", 1, "tt3"
    ActionTree1.AT_Nodes.Add "pu", "zz", "Level 2", 3, "tt2"
    ActionTree1.AT_Nodes.Add "f34", "zz", "Level 2", 3, "tt3"
    ActionTree1.AT_Nodes.Add "f33", "zz", "Level 2", 5, "tt4"
    
    ActionTree1.AT_Nodes.Add "sg", "f34", "Level 3", 3, "tt5"
    ActionTree1.AT_Nodes.Add "sfg", "f34", "Level 3", 2, "tt3"
    ActionTree1.AT_Nodes.Add "vbxc", "f34", "Level 3", 2, "tt5"
    ActionTree1.AT_Nodes.Add "sdfg", "vbxc", "Level 4", 1, "tt5"
    
    ActionTree1.AT_Nodes.Add "oo", "", "Alexander", 4, "tt4"
    
    ActionTree1.AT_Nodes.Add "pp", "oo", "The", 5, "tt6"
    
    ActionTree1.AT_Nodes.Add "tt4", "tt4", "Great", 3
    ActionTree1.AT_Nodes.Add "ttewr", "tt4", "of", 2
    ActionTree1.AT_Nodes.Add "tt3", "tt3", "Greece", 1
    
    ActionTree1.AT_Nodes.Add "tt6", "tt6", "Spanish", 4
    
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("oo").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("vbxc").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t3"
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t4"
    ActionTree1.AT_Nodes("sdfg").Child.Add "The sharpshooter", "t5"
    
    ActionTree1.AT_Columns.Add 80, "Actions"
    ActionTree1.AT_Columns.Add 80, "Info"
    ActionTree1.AT_Columns.Add 80, "Status"
    ActionTree1.AT_Columns.Add 80, "Kommentar"
    
    
    ActionTree1.Refresh
End Sub

Private Sub Form_Load()

    ActionTree1.BackColor = vbWhite
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    ActionTree1.Height = Me.ScaleHeight - 500
    ActionTree1.Width = Me.ScaleWidth / 2

End Sub
