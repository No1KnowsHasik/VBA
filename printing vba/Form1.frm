VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   475
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   709
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "печать"
      Height          =   495
      Left            =   6960
      TabIndex        =   29
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Следующая"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Предыдущая"
      Height          =   495
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "ООО ""Последний звонок"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Число строк"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   11
      Left            =   4560
      TabIndex        =   30
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   3975
      Left            =   240
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800080&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   408
      X2              =   408
      Y1              =   120
      Y2              =   408
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   368
      X2              =   624
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Дисплей"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   10
      Left            =   4560
      TabIndex        =   28
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Тип экрана"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   27
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Особенности"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Размер"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   25
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Вес"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Антенна"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   23
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Конструкция"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   22
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Тип корпуса"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Диапазоны частот"
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Index           =   12
      Left            =   1680
      TabIndex        =   19
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   11
      Left            =   6240
      TabIndex        =   18
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   10
      Left            =   6240
      TabIndex        =   17
      Top             =   5400
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   9
      Left            =   6240
      TabIndex        =   16
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   855
      Index           =   8
      Left            =   6240
      TabIndex        =   15
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   14
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   13
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   12
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   11
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   10
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Центровка
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Рабочее место продавца мобил"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      BorderWidth     =   3
      Height          =   5655
      Left            =   4320
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Модель"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Центровка
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Прозрачно
      Caption         =   "руб."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Центровка
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Мобильный телефон"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Технические характеристики"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Цена"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   360
      Top             =   1920
      Width           =   3750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim MassiveData(12) As String
Dim FileName As String
Dim F As Long
Dim X As Long
Dim QuantityFonts As Long

Private Sub Form_Load()
FileName = "1"
DataLoading
End Sub

Private Sub DataLoading()
Image1.Picture = LoadPicture(App.Path & "\" & FileName & ".jpg")
Image1.Left = 24
Image1.Top = 128
If Image1.Width > Image1.Height Then Image1.Top = Image1.Top + (250 - Image1.Height) / 2
If Image1.Height > Image1.Width Then Image1.Left = Image1.Left + (250 - Image1.Width) / 2
F = FreeFile
Open App.Path & "\" & FileName & ".txt" For Input As #F
For X = 1 To 12
Line Input #F, MassiveData(X)
Label9(X).Caption = MassiveData(X)
Next X
Close #F

End Sub

Private Sub Command1_Click()
If Val(FileName) > 1 Then
FileName = Str(Val(FileName) - 1)
FileName = Right(FileName, Len(FileName) - 1)
DataLoading
End If
End Sub

Private Sub Command2_Click()
FileName = Str(Val(FileName) + 1)
FileName = Right(FileName, Len(FileName) - 1)
If Dir(App.Path & "\" & FileName & ".txt") <> "" Or Dir(App.Path & "\" & FileName & ".jpg") <> "" Then
DataLoading
Else
FileName = Str(Val(FileName) - 1)
FileName = Right(FileName, Len(FileName) - 1)
DataLoading
End If
End Sub

Private Sub Command3_Click()
Dim OrientTelefona As Single
Dim Fonts() As String
Dim VertCoord As Single
Dim NumberFont As Long
QuantityFonts = 0


'находим сколько есть общеупотребимых шрифтов из 4
   For NumberFont = 0 To Printer.FontCount - 1
       If Printer.Fonts(NumberFont) = "Arial Cyr" Or Printer.Fonts(NumberFont) = "Times New Roman" _
       Or Printer.Fonts(NumberFont) = "Courier New" Or Printer.Fonts(NumberFont) = "MS Sans Serif" Then
            QuantityFonts = QuantityFonts + 1
            ReDim Preserve Fonts(QuantityFonts)
            Fonts(QuantityFonts) = Printer.Fonts(NumberFont)
            
       End If
    Next NumberFont


'устанавливаем шрифт
If QuantityFonts > 0 Then Printer.FontName = Fonts(1)
'устанавливаем размерность в сантиметрах
Printer.ScaleMode = vbCentimeters

'размер шрифта
Printer.FontSize = 12
'устанавливаем подчеркивание
Printer.Font.Underline = True
'устананавливаем качество печати - среднее
Printer.PrintQuality = 3
'прижимаем вправо первую строчку с названием фирмы
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Label1.Caption)
'запоминаем, сколько высоты заняла первая строка
VertCoord = Printer.TextHeight(Label1.Caption)
            'встраиваем в страницу первую строчку
            Printer.Print Label1.Caption
'устанавливаем большой размер шрифта для названия телефона
Printer.FontSize = 20
'отменяем подчеркивание
Printer.Font.Underline = False
'зато ставим жирный
Printer.Font.Bold = True
'делаем отступ вниз на 1,5 см от первой строчки, меняя вертикальную координату
Printer.CurrentY = VertCoord + 1.5
'выравниваем вторую строчку по-середине
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Label9(1).Caption & "Лидер продаж - телефон ")) / 2
'запоминаем, сколько высоты мы использовали на  обе строчки вместе с отступами
VertCoord = VertCoord + Printer.TextHeight(Label9(1).Caption) + 1.5
            'встраиваем в страницу вторую строчку
            Printer.Print "Лидер продаж - телефон " & Label9(1).Caption
            

 
'устанавливаем размер шрифта поменьше для заглавия технических характеристик телефона
Printer.FontSize = 14
'добавляем курсив
Printer.Font.Italic = True
'отступаем вниз еще на 2 см
Printer.CurrentY = VertCoord + 2
'запоминаем, сколько высоты мы использовали на  три строчки вместе с отступамидобавляем высоту второй строки
VertCoord = VertCoord + Printer.TextHeight(Label4.Caption) + 2
'оставляем 8 см слева под картинку, а остальное справа - под текст
'строчку центрируем на ее Printer.ScaleWidth - 8 см
Printer.CurrentX = (Printer.ScaleWidth - 8 - Printer.TextWidth(Label4.Caption)) / 2
            'встраиваем в страницу третью строчку
            Printer.Print Label4.Caption
            

'Printer.ScaleWidth-8 и берем текущую координату по вертикали VertCoord
'печатаем картинку, центрируя ее на своей площади
If Image1.Height > Image1.Width Then
OrientTelefona = 2.25
Else
OrientTelefona = 0.5
End If

Printer.PaintPicture Image1.Picture, Printer.ScaleWidth - ScaleX(Image1.Width, vbPixels, vbCentimeters) - OrientTelefona, VertCoord
            
'отменяем курсив и жирность
Printer.Font.Italic = False
Printer.Font.Bold = False
'устанавливаем размер шрифта еще меньше
Printer.FontSize = 12
'отступаем вниз от 3-ей строки на 1 см
Printer.CurrentY = VertCoord + 1
VertCoord = VertCoord + 1
'теперь в цикле распечатываем технические данные (всего 10 строк)

For X = 2 To 11
'отступаем слева по сантиметру
Printer.CurrentX = 1
Printer.CurrentY = VertCoord
Printer.Print Label8(X).Caption

Printer.CurrentX = 5
Printer.CurrentY = VertCoord
Printer.Print Label9(X).Caption
VertCoord = VertCoord + 0.6

Next X
'отступаем еше 1 см и проводим линию
VertCoord = VertCoord + 1
Printer.Line (1, VertCoord)-(17.5, VertCoord + 0.7), vbBlack, BF

Printer.CurrentY = VertCoord + 0.1
Printer.CurrentX = 6
'теперь пишем цену
Printer.Font.Bold = True
Printer.Print "Самая лучшая цена: " & Label9(12).Caption & " руб."
Printer.Font.Bold = False
Printer.EndDoc

End Sub



