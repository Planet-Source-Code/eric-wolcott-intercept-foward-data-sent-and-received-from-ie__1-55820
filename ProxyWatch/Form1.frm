VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin Project1.Proxy Proxy1 
      Left            =   4875
      Top             =   330
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4950
      Left            =   30
      TabIndex        =   12
      Top             =   2460
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   8731
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Proxy"
      Height          =   225
      Left            =   4095
      TabIndex        =   11
      Top             =   2130
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Height          =   1920
      Left            =   60
      TabIndex        =   2
      Top             =   135
      Width           =   3930
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   1500
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   2550
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   975
         TabIndex        =   6
         Text            =   "0"
         Top             =   195
         Width           =   1620
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   975
         TabIndex        =   5
         Text            =   "0"
         Top             =   525
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Current Host:"
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   1275
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Status: "
         Height          =   270
         Left            =   405
         TabIndex        =   8
         Top             =   855
         Width           =   2460
      End
      Begin VB.Label Label2 
         Caption         =   "Outgoing: "
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label Label1 
         Caption         =   "Incomming: "
         Height          =   270
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4335
      Top             =   300
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   120
      Left            =   6390
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Proxy"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   2115
      Width           =   3930
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5385
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0584
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E52
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
Dim LastKey As String

Private Sub Command1_Click()
    '---------------------------START PROXY-----------------------'
    Proxy1.Start
    TreeView1.Nodes.Add , , "Start" & Time(), "Proxy Started (" & Time() & ")", 3, 3
End Sub

Private Sub Command2_Click()
    '---------------------------STOP PROXY-----------------------'
    Proxy1.StopProxy
    TreeView1.Nodes.Add , , "Stop" & Time(), "Proxy Stopped (" & Time() & ")", 3, 3
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    '---------------------------DISPLAY DATA-----------------------'
    If Node.SelectedImage = 2 Then
    MsgBox Node.Key
    End If
End Sub

Private Sub Proxy1_FowaredRequest(RequestID As Integer, SockType As FindIndex, Data As String)
    '---------------------------INBOUND/OUTBOUND DATA-----------------------'
    Dim K As Boolean, L As Integer: K = True
    If SockType = i_Outgoing Then 'Send To IE
        Text2.Text = Val(Text3.Text) + Len(Data)
        Proxy1.AcceptRequest RequestID, SockType, Replace(Data, "Yahoo", "Planet-Source-Code") 'Foward Data On
        Do Until K = False
            L = L + 1
            K = CheckKeyExist(Page & "Incomming" & L)
        Loop
        TreeView1.Nodes.Add LastKey, tvwChild, Page & "Incomming" & L, "Incomming", 4, 4
        TreeView1.Nodes.Add Page & "Incomming" & L, tvwChild, Page & "Incomming" & L & Data, Data, 2, 2
    ElseIf SockType = i_Incomming Then 'Send To Host
        Text3.Text = Val(Text3.Text) + Len(Data)
        Proxy1.AcceptRequest RequestID, SockType, Data 'Foward Data On
        Do Until K = False
            L = L + 1
            K = CheckKeyExist(Page & "Outgoing" & L)
        Loop
        TreeView1.Nodes.Add LastKey, tvwChild, Page & "Outgoing" & L, "Outgoing", 4, 4
        TreeView1.Nodes.Add Page & "Outgoing" & L, tvwChild, Page & "Outgoing" & L & Data, Data, 2, 2
    End If
End Sub

Private Sub Proxy1_PageChange(Page As String)
    '---------------------------IE REQUEST DIFFERENT SERVER-----------------------'
    Text4.Text = Page
    Dim K As Boolean, L As Integer: K = True
    If CheckKeyExist(Page) = True Then
        Do Until K = False
            L = L + 1
            K = CheckKeyExist(Page & L)
        Loop
        TreeView1.Nodes.Add Page, tvwChild, Page & L, Time(), 5, 6
        LastKey = Page & L
    Else
        TreeView1.Nodes.Add , , Page, Page, 1, 1
        K = True: L = 0
        Do Until K = False
            L = L + 1
            K = CheckKeyExist(Page & L)
        Loop
        TreeView1.Nodes.Add Page, tvwChild, Page & L, Time(), 5, 6
        LastKey = Page & L
    End If
End Sub

Private Sub Proxy1_StatusChangeOutgoing(Current_State As SockStatus)
    '---------------------------STATUS CHANGED-----------------------'
    If Current_State = i_Closed Then
    Text1.Text = "Closed"
    ElseIf Current_State = i_Opened Then: Text1.Text = "Opened"
    ElseIf Current_State = i_Listening Then: Text1.Text = "Listening"
    ElseIf Current_State = i_Pending Then: Text1.Text = "Pending"
    ElseIf Current_State = i_Resolving Then: Text1.Text = "Resolving"
    ElseIf Current_State = i_Resolved Then: Text1.Text = "Resolved"
    ElseIf Current_State = i_Connecting Then: Text1.Text = "Connecting"
    ElseIf Current_State = i_Connected Then: Text1.Text = "Connected"
    ElseIf Current_State = i_Closing Then: Text1.Text = "Closed"
    ElseIf Current_State = i_Error Then: Text1.Text = "Error"
    End If
End Sub
Function CheckKeyExist(Key As String) As Boolean
    CheckKeyExist = False
    Dim i: For i = 1 To TreeView1.Nodes.Count
    If UCase(TreeView1.Nodes(i).Key) = UCase(Key) Then CheckKeyExist = True: Exit Function
    Next
End Function
