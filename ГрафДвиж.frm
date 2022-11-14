VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17550
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Печать"
      Height          =   615
      Left            =   6600
      TabIndex        =   17
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
      Caption         =   "down"
      Height          =   615
      Left            =   11760
      TabIndex        =   16
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "left"
      Height          =   615
      Left            =   11160
      TabIndex        =   15
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "right"
      Height          =   615
      Left            =   12360
      TabIndex        =   14
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "up"
      Height          =   615
      Left            =   11760
      TabIndex        =   13
      Top             =   5640
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   9480
      Max             =   100
      TabIndex        =   11
      Top             =   8760
      Value           =   70
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Form 2"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   13440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-->"
      Height          =   975
      Left            =   11160
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Height          =   975
      Left            =   11160
      TabIndex        =   7
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   13560
      TabIndex        =   5
      Text            =   "S:\ABC\VBA2.txt"
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "парсим файл с матрицами"
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   1575
      Left            =   11160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   13320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Заполнить массив"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   7800
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   240
      ScaleHeight     =   5715
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Подюор угла оси z"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Путь к файлу с матрицами"
      Height          =   255
      Left            =   11160
      TabIndex        =   6
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teststring(2) As String
Dim angle As Double


Dim path As String
Dim filenumber As Integer

Dim matr(5, 3) As Double        'исходная матрица
Dim matrwork(5, 3) As Double    'матрица для изменений
Dim batr
Dim i, j As Integer
Dim bb As Integer









Private Sub Command1_Click()






Me.Picture1.Cls
Me.Picture1.Scale (-5, -5)-(5, 5)

 

batr = Array(0, 1, 1, 2, 2, 3, 3, 0, 1, 4, 2, 4, 0, 4, 3, 4, 1, 5, 2, 5, 0, 5, 3, 5)

For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
    
    Me.Picture1.Line (matrwork(mn, 0), matrwork(mn, 1))-(matrwork(bn, 0), matrwork(bn, 1))
Next





End Sub



Private Sub Command10_Click()
Me.PrintForm

End Sub



Private Sub Command2_Click()
For i = 0 To UBound(matrwork, 1)
    matrwork(i, 1) = matrwork(i, 1) - 0.1
Next
Call Command1_Click
End Sub
Private Sub Command9_Click()
For i = 0 To UBound(matrwork, 1)
    matrwork(i, 1) = matrwork(i, 1) + 0.1
Next
Call Command1_Click
End Sub
Private Sub Command8_Click()
For i = 0 To UBound(matrwork, 1)
    matrwork(i, 0) = matrwork(i, 0) - 0.1
Next
Call Command1_Click
End Sub
Private Sub Command7_Click()
For i = 0 To UBound(matrwork, 1)
    matrwork(i, 0) = matrwork(i, 0) + 0.1
Next
Call Command1_Click
End Sub


Private Sub Command3_Click() ' чтение файла с матрицам Работате хорошо не трогать
path = Me.Text5.Text
ff = FreeFile
Open path For Input As #ff

Do While Not EOF(ff)

    Line Input #ff, s

    Me.Text4.Text = Me.Text4.Text & s & vbCrLf
Loop
Close filenumber
End Sub

Private Sub Command4_Click() ' показываем матрицы
For i = 0 To UBound(matr, 1)
    For j = 0 To UBound(matr, 2)
        Me.Text1.Text = Me.Text1.Text & matr(i, j) & "@"
    Next
    Me.Text1.Text = Me.Text1.Text & vbCrLf
Next
End Sub

Private Sub Command5_Click() ' отображаем отспличиную херню

c = Split(Me.Text4.Text, "/")



Me.Text6.Text = c(0) & c(1) & c(2) & c(3) & c(4) & c(5) & c(6) ''ИЗМЕНИТЬ



For i = 1 To UBound(matr, 1) + 1    'заполняем матрицу
    ngh = Split(c(i), "!")
    For j = 0 To UBound(matr, 2)
        nn = CDbl(ngh(j))
        matr(i - 1, j) = nn
        
    Next
Next


End Sub

Public Sub anglell()                                         'функция меняющая угол и вызывающая отображение
For i = 0 To UBound(matr, 1)
    For j = 0 To UBound(matr, 2)
        matrwork(i, j) = matr(i, j)
    Next
Next


For k = 0 To UBound(matr, 1)                                 'ось z
    If matr(k, 2) > 0 Then
        matrwork(k, 0) = matrwork(k, 0) + matrwork(k, 2) * angle
        matrwork(k, 1) = matrwork(k, 1) + matrwork(k, 2) * (1 - angle)
    End If
Next
Call Command1_Click
End Sub

Private Sub Command6_Click() 'кнопка перехода на следующую стр. (NOT USED)
Me.Hide
Form2.Show

End Sub


Private Sub HScroll1_Change()
angle = Me.HScroll1.Value / 100
Call anglell

End Sub




