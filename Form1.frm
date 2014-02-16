VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.PictureBox NutTron 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   5640
      MousePointer    =   15  'Size All
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1965
      TabIndex        =   3
      Top             =   3960
      Width           =   1965
   End
   Begin VB.PictureBox NutTronMove 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   4200
      Picture         =   "Form1.frx":1CDB
      ScaleHeight     =   1995
      ScaleWidth      =   1965
      TabIndex        =   5
      Top             =   3480
      Width           =   1965
   End
   Begin VB.PictureBox NutTronDown 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   4680
      Picture         =   "Form1.frx":392E
      ScaleHeight     =   1995
      ScaleWidth      =   1965
      TabIndex        =   4
      Top             =   5040
      Width           =   1965
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   1560
   End
   Begin BirthDay.UniTextBox TextBox1 
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4895
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   ""
      Locked          =   -1  'True
      Enabled         =   0   'False
      BorderStyle     =   0
      Scrollbar       =   2
      Alignment       =   2
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   3600
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   21
      Left            =   13080
      Top             =   10080
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   20
      Left            =   13560
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   10200
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   18
      Left            =   10560
      Top             =   9000
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   17
      Left            =   7560
      Top             =   9960
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   16
      Left            =   8040
      Top             =   7560
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   15
      Left            =   5760
      Top             =   9840
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   14
      Left            =   6240
      Top             =   6000
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   13
      Left            =   2880
      Top             =   5280
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   12
      Left            =   3240
      Top             =   8760
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   11
      Left            =   240
      Top             =   9720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   10
      Left            =   720
      Top             =   7320
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   9
      Left            =   11880
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   8
      Left            =   12360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   7
      Left            =   2520
      Top             =   5520
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   960
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   9000
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   9360
      MousePointer    =   15  'Size All
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   6360
      Top             =   4440
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   6840
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   1320
      MousePointer    =   15  'Size All
      Top             =   3240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Const IDC_HAND As Long = &H7F89
Dim hCursor As Long

Dim fNum As Long, B() As Byte, fp
Dim arr, i As Integer, s As String
    


Private Sub Command1_Click()

If Timer1.Enabled = False Then
Timer1.Enabled = True
Command1.Caption = "Stop"
Else
Timer1.Enabled = False
Command1.Caption = "Start"
End If
End Sub



Private Sub Command3_Click()
End
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
Source.Left = x
Source.Top = y
End Sub

Private Sub Form_Load()
App.TaskVisible = False
hCursor = LoadCursor(ByVal 0&, IDC_HAND)

With TextBox1
.Left = 0
.Top = Me.Height / 2
.Width = Me.Width
.Height = 1500
End With

With NutTron
.Top = Me.Height / 2 - .Height - 1000
.Left = Me.Width / 2 - .Width / 2
End With
With NutTronMove
.Top = Me.Height / 2 - .Height - 1000
.Left = Me.Width / 2 - .Width / 2
End With

With NutTronDown
.Top = Me.Height / 2 - .Height - 1000
.Left = Me.Width / 2 - .Width / 2
End With

SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
Randomize
For a = 0 To Image1.Count - 1
Image1(a).MousePointer = 15
Image1(a).Picture = LoadResPicture(Int((15 * Rnd) + 1), 0)
Next a
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
NutTron.Visible = True
NutTronMove.Visible = True
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Image1(Index).Drag
End Sub

Private Sub NutTron_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
NutTron.Visible = False
End Sub



Private Sub NutTronMove_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SetCursor hCursor
NutTronMove.Visible = False
End Sub

Private Sub NutTronMove_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
SetCursor hCursor
End Sub

Private Sub NutTronMove_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    fp = App.Path & "\Text.txt"
    fNum = FreeFile()
    Open fp For Binary Access Read As #fNum
        ReDim B(LOF(fNum))
    Get #fNum, , B
    Close #fNum

    
    arr = Split(B, vbCrLf)
    s = arr(i)
    
TextBox1.Text = s
i = i + 1
If i > UBound(arr) Then i = 0

NutTron.Visible = True
NutTronMove.Visible = True
End Sub

Private Sub Timer1_Timer()


For x = 1 To Image1.Count - 1
Image1(x).Top = Image1(x).Top - x * 5
    If Image1(x).Top < 0 - Image1(x).Height Then Image1(x).Top = Me.Height
Next x
Image1(0).Top = Image1(0).Top - 50
    If Image1(0).Top < 0 - Image1(0).Height Then Image1(0).Top = Me.Height

End Sub
