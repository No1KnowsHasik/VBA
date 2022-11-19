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
   Begin VB.PictureBox Picture1 
      Height          =   6405
      Left            =   480
      ScaleHeight     =   6345
      ScaleWidth      =   8475
      TabIndex        =   28
      Top             =   120
      Width           =   8540
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   615
      Left            =   8160
      Max             =   600
      Min             =   40
      TabIndex        =   26
      Top             =   6840
      Value           =   350
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Парсим карту"
      Height          =   405
      Left            =   12120
      TabIndex        =   25
      Top             =   8760
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   10200
      TabIndex        =   24
      Text            =   "S:\ABC\MAP1.txt"
      Top             =   8760
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   735
      Left            =   1080
      Max             =   180
      TabIndex        =   19
      Top             =   6960
      Value           =   32
      Width           =   3135
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   735
      Left            =   1080
      Max             =   180
      TabIndex        =   18
      Top             =   7800
      Value           =   60
      Width           =   3135
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   735
      Left            =   1080
      Max             =   180
      TabIndex        =   17
      Top             =   8640
      Width           =   3135
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Печать"
      Height          =   315
      Left            =   15120
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "down"
      Height          =   615
      Left            =   5040
      TabIndex        =   15
      Top             =   8520
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "left"
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "right"
      Height          =   615
      Left            =   5640
      TabIndex        =   13
      Top             =   7920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "up"
      Height          =   615
      Left            =   5040
      TabIndex        =   12
      Top             =   7320
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   12000
      Max             =   100
      TabIndex        =   10
      Top             =   5160
      Value           =   70
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Form 2"
      Height          =   195
      Left            =   12600
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   13440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-->"
      Height          =   975
      Left            =   11160
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Height          =   1455
      Left            =   6840
      TabIndex        =   6
      Top             =   7680
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   10200
      TabIndex        =   4
      Text            =   "S:\ABC\VBA2.txt"
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Парсим матрицу"
      Height          =   405
      Left            =   12120
      TabIndex        =   3
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   1575
      Left            =   11160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   13320
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Заполнить массив"
      Height          =   615
      Left            =   11640
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Scale"
      Height          =   375
      Left            =   6480
      TabIndex        =   27
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Путь к файлу с картой"
      Height          =   255
      Left            =   10200
      TabIndex        =   23
      Top             =   8400
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "Z"
      Height          =   615
      Left            =   480
      TabIndex        =   22
      Top             =   8640
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Y"
      Height          =   615
      Left            =   480
      TabIndex        =   21
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Подбор угла изометрии"
      Height          =   375
      Left            =   13080
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Путь к файлу с матрицами"
      Height          =   255
      Left            =   10200
      TabIndex        =   5
      Top             =   7560
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teststring(2) As String
Dim angle As Double
Dim X, Y As Double
Dim scalee As Integer
Dim path As String
Dim filenumber As Integer
Dim matr(5, 3) As Double        'исходная матрица
Dim matrwork(5, 3) As Double    'матрица для изменений
Dim batr
Dim i, j As Integer
Dim bb As Integer

Private Sub Command1_Click()
Me.Picture1.Cls
Me.Picture1.Scale (-20, -15)-(60, 45)
For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
    
    Me.Picture1.Line (matrwork(mn, 0) * Me.HScroll5.Value / 100, matrwork(mn, 1) * Me.HScroll5.Value / 100)-(matrwork(bn, 0) * Me.HScroll5.Value / 100, matrwork(bn, 1) * Me.HScroll5.Value / 100)
Next

vinos_x = matrwork(2, 0) * Me.HScroll5.Value / 100 + ((matrwork(1, 0) - matrwork(2, 0)) * Me.HScroll5.Value / 100) / 2
vinos_y = matrwork(2, 1) * Me.HScroll5.Value / 100 + ((matrwork(1, 1) - matrwork(2, 1)) * Me.HScroll5.Value / 100) / 2

Me.Picture1.Line (vinos_x, vinos_y)-(vinos_x - 5, vinos_y - 5)
Me.Picture1.Line (vinos_x - 5, vinos_y - 5)-(vinos_x - 20, vinos_y - 5)
Me.Picture1.CurrentX = vinos_x - 20
Me.Picture1.CurrentY = vinos_y - 7
Me.Picture1.Print "Сторона октаэдра ="

'Printer.Print Me.Picture1.Image
'Printer.EndDoc

End Sub

Private Sub Command11_Click() 'парсим мапу
path = Me.Text2.Text
ff = FreeFile
Open path For Input As #ff

Line Input #ff, s
batr = Split(s, "#")
Close filenumber
End Sub

Private Sub Command2_Click()  'перемещение объекта вверх
Y = Y - 0.3
Call rotate
End Sub

Private Sub Command9_Click()  'перемещение объекта вниз
Y = Y + 0.3
Call rotate
End Sub

Private Sub Command8_Click()  'перемещение объекта влево
X = X - 0.3
Call rotate
End Sub

Private Sub Command7_Click()  'перемещение объекта вправо
X = X + 0.3
Call rotate
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

Private Sub Command4_Click() ' показываем матрицы и рисунок
Call Command3_Click
Call Command5_Click
Call Command11_Click
Call rotate

For i = 0 To UBound(matr, 1)
    For j = 0 To UBound(matr, 2)
        Me.Text1.Text = Me.Text1.Text & matr(i, j) & "@"
    Next
    Me.Text1.Text = Me.Text1.Text & vbCrLf
Next
End Sub

Private Sub Command5_Click() ' отображаем результат парса

c = Split(Me.Text4.Text, "/")

For l = 0 To UBound(c)
    Me.Text6.Text = Me.Text6.Text & c(l)
Next
For i = 1 To UBound(matr, 1) + 1    'заполняем матрицу
    ngh = Split(c(i), "!")
    For j = 0 To UBound(matr, 2)
        nn = CDbl(ngh(j))
        matr(i - 1, j) = nn
        
    Next
Next
End Sub

Private Sub HScroll4_Change()
Call rotate
End Sub

Private Sub HScroll3_Change()
Call rotate
End Sub

Private Sub HScroll2_Change()
Call rotate
End Sub

Private Sub rotate()
For i = 0 To UBound(matr, 1)
    matrwork(i, 0) = matr(i, 0) + matr(i, 0) * Cos(Me.HScroll2.Value * 3.14 / 180) - matr(i, 1) * Sin(Me.HScroll2.Value * 3.14 / 180) + matr(i, 0) * Cos(Me.HScroll3.Value * 3.14 / 180) + matr(i, 2) * Sin(Me.HScroll3.Value * 3.14 / 180) + X
    matrwork(i, 1) = matr(i, 1) - matr(i, 0) * Sin(Me.HScroll2.Value * 3.14 / 180) + matr(i, 1) * Cos(Me.HScroll2.Value * 3.14 / 180) + matr(i, 1) * Cos(Me.HScroll4.Value * 3.14 / 180) + matr(i, 2) * Sin(Me.HScroll4.Value * 3.14 / 180) + Y
    matrwork(i, 2) = matr(i, 2) - matr(i, 0) * Sin(Me.HScroll3.Value * 3.14 / 180) + matr(i, 2) * Cos(Me.HScroll3.Value * 3.14 / 180) - matr(i, 1) * Sin(Me.HScroll4.Value * 3.14 / 180) + matr(i, 2) * Cos(Me.HScroll4.Value * 3.14 / 180)
Next
Call Command1_Click
End Sub

Private Sub Command6_Click() 'кнопка перехода на следующую стр. (NOT USED)
Me.Hide
Form2.Show
End Sub

Private Sub HScroll5_Change()
Call Command3_Click
Call Command5_Click
Call Command11_Click
Call rotate
End Sub
