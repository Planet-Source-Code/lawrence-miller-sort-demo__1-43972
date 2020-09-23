VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   1620
   ClientTop       =   1935
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9180
   Begin VB.CommandButton cmdShell 
      Caption         =   "Shell"
      Height          =   525
      Left            =   2730
      TabIndex        =   7
      Top             =   4170
      Width           =   1005
   End
   Begin VB.CommandButton cmdQuick 
      Caption         =   "Quick"
      Height          =   525
      Left            =   2730
      TabIndex        =   6
      Top             =   3570
      Width           =   1005
   End
   Begin VB.CommandButton cmdInsertion 
      Caption         =   "Insertion"
      Height          =   525
      Left            =   2730
      TabIndex        =   5
      Top             =   2970
      Width           =   1005
   End
   Begin VB.CommandButton cmdHeap 
      Caption         =   "Heap"
      Height          =   525
      Left            =   2730
      TabIndex        =   4
      Top             =   2370
      Width           =   1005
   End
   Begin VB.CommandButton cmdExchange 
      Caption         =   "Exchange"
      Height          =   525
      Left            =   2730
      TabIndex        =   3
      Top             =   1770
      Width           =   1005
   End
   Begin VB.CommandButton cmdBubble 
      Caption         =   "Bubble"
      Height          =   525
      Left            =   2730
      TabIndex        =   2
      Top             =   1170
      Width           =   1005
   End
   Begin VB.ListBox List2 
      Height          =   5325
      Left            =   3870
      TabIndex        =   1
      Top             =   510
      Width           =   2475
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Elapsed Time"
      Height          =   285
      Left            =   6420
      TabIndex        =   14
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblShell 
      Height          =   315
      Left            =   6390
      TabIndex        =   13
      Top             =   4110
      Width           =   2685
   End
   Begin VB.Label lblQuick 
      Height          =   315
      Left            =   6390
      TabIndex        =   12
      Top             =   3570
      Width           =   2685
   End
   Begin VB.Label lblInsertion 
      Height          =   315
      Left            =   6390
      TabIndex        =   11
      Top             =   3030
      Width           =   2685
   End
   Begin VB.Label lblHeap 
      Height          =   315
      Left            =   6390
      TabIndex        =   10
      Top             =   2430
      Width           =   2685
   End
   Begin VB.Label lblExchange 
      Height          =   315
      Left            =   6390
      TabIndex        =   9
      Top             =   1770
      Width           =   2685
   End
   Begin VB.Label lblBubble 
      Height          =   315
      Left            =   6390
      TabIndex        =   8
      Top             =   1200
      Width           =   2685
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SArray(10000) As Integer

Private Sub cmdBubble_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Bubble SArray()
    lblBubble.Caption = Timer - Start
    lblBubble.Refresh
    LoadList
    Set s = Nothing
End Sub

Private Sub cmdExchange_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Exchange SArray()
    lblExchange.Caption = Timer - Start 'DateDiff("s", Start, Now)
    lblExchange.Refresh
    LoadList
    Set s = Nothing

End Sub

Private Sub cmdHeap_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Heap SArray()
    lblHeap.Caption = Timer - Start 'DateDiff("s", Start, Now)
    lblHeap.Refresh
    LoadList
    Set s = Nothing

End Sub

Private Sub cmdInsertion_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Insertion SArray()
    lblInsertion.Caption = Timer - Start 'DateDiff("s", Start, Now)
    lblInsertion.Refresh
    LoadList
    Set s = Nothing

End Sub

Private Sub cmdQuick_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  Dim Max  As Long
  
    Max = UBound(SArray)
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Quick SArray(), 1, Max
    lblQuick.Caption = Timer - Start 'DateDiff("s", Start, Now)
    lblQuick.Refresh
    LoadList
    Set s = Nothing

End Sub

Private Sub cmdShell_Click()
  Dim i As Integer
  Dim s As Sort
  Dim Start As Double
  
    List2.Clear
    Set s = New Sort
    Start = Timer
    s.Shell SArray()
    lblShell.Caption = Timer - Start 'DateDiff("s", Start, Now)
    lblShell.Refresh
    LoadList
    Set s = Nothing

End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim j As Integer
  Dim Max As Integer
  
    Max = UBound(SArray)
    List1.Clear
    For i = 1 To Max
      j = Int(Rnd * Max) + 1
      SArray(i) = j
      List1.AddItem j
      If SArray(i) = 0 Then Debug.Print i
    Next

End Sub

Private Sub LoadList()
  Dim i As Integer
  Dim Max As Integer
  
    Max = UBound(SArray)
    For i = 1 To Max
      List2.AddItem SArray(i)
    Next
    ReloadArray
End Sub

Private Sub ReloadArray()
  Dim i As Integer
  Dim j As Integer
  
    j = List1.ListCount - 1
    For i = 0 To j
      SArray(i + 1) = List1.List(i)
    Next
End Sub
