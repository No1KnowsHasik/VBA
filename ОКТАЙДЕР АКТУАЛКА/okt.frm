VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OCTAHEDRON"
   ClientHeight    =   10050
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "okt.frx":0000
   ScaleHeight     =   10050
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FF80FF&
      Caption         =   "SVG"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FF80FF&
      Caption         =   "BMP"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "DRAW"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Picture         =   "okt.frx":41EB42
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8760
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5321
      Left            =   120
      ScaleHeight     =   7601.006
      ScaleMode       =   0  'User
      ScaleWidth      =   7808.988
      TabIndex        =   17
      Top             =   70
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "C++"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Picture         =   "okt.frx":421B84
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "ARCHIVE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      MaskColor       =   &H0000C000&
      Picture         =   "okt.frx":424BC6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   5760
      TabIndex        =   14
      Top             =   5400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   16711935
      ForeColor       =   16777215
      BackColorFixed  =   16711935
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483638
      ForeColorSel    =   16777215
      BackColorBkg    =   16744703
      GridColorFixed  =   16711935
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "Сторона | Объем"
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Add new value"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      MaskColor       =   &H00C000C0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "PRINTER"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "okt.frx":427C08
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "COPY"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      Picture         =   "okt.frx":42AC4A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "VOICE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      OLEDropMode     =   1  'Manual
      Picture         =   "okt.frx":42DC8C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8760
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8520
      MaskColor       =   &H00FF00FF&
      Picture         =   "okt.frx":430CCE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "OPEN"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7320
      Picture         =   "okt.frx":433D10
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "SOLIDWORKS"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6120
      Picture         =   "okt.frx":436D52
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "EXPLOYER"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      Picture         =   "okt.frx":439D94
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "PP"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      Picture         =   "okt.frx":43CDD6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "EXCEL"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2520
      Picture         =   "okt.frx":43FE18
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "WORD"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      Picture         =   "okt.frx":442E5A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5655
      Begin VB.TextBox Text1 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   1935
         Left            =   2160
         TabIndex        =   2
         Text            =   "6"
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Side of the Octahedron"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Mode"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   7800
      TabIndex        =   22
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Menu mmSort 
      Caption         =   "Sort"
      Begin VB.Menu mSortA 
         Caption         =   "SortA"
      End
      Begin VB.Menu mSortV 
         Caption         =   "SortV"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim print_mode As Integer
Dim draw_mode As Boolean

'Dim vlist()

Sub draw() 'отрисовка


If draw_mode = True Then
    Me.Picture1.BackColor = &HFFFFFF
    Me.Picture1.ForeColor = &H0&
Else
    Me.Picture1.BackColor = &HFF80FF
    Me.Picture1.ForeColor = &HFFFFFF
End If

    
    
    
Me.Picture1.Visible = True
Me.Command15.Visible = True
Dim matr(5, 2) As Double
Dim batr

matr(0, 0) = 5.001
matr(0, 1) = 4.028
matr(1, 0) = 0
matr(1, 1) = 4.028
matr(2, 0) = 1.732
matr(2, 1) = 5.087
matr(3, 0) = 6.732
matr(3, 1) = 5.087
matr(4, 0) = 3.366
matr(4, 1) = 8.586
matr(5, 0) = 3.366
matr(5, 1) = 0.53
batr = Array(0, 1, 1, 2, 2, 3, 3, 0, 1, 4, 2, 4, 0, 4, 3, 4, 1, 5, 2, 5, 0, 5, 3, 5)

Me.Picture1.Cls
Me.Picture1.Scale (-20, -15)-(60, 45)
'Me.Picture1.draw


For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
'    Me.Picture1
    Me.Picture1.Line (matr(mn, 0) * 5, matr(mn, 1) * 5)-(matr(bn, 0) * 5, matr(bn, 1) * 5)
Next

vinos_x = matr(2, 0) * 5 + ((matr(1, 0) - matr(2, 0)) * 5) / 2
vinos_y = matr(2, 1) * 5 + ((matr(1, 1) - matr(2, 1)) * 5) / 2

Me.Picture1.Line (vinos_x, vinos_y)-(vinos_x - 7, vinos_y - 7)
Me.Picture1.Line (vinos_x - 7, vinos_y - 7)-(vinos_x - 16, vinos_y - 7)
Me.Picture1.CurrentX = vinos_x - 15
Me.Picture1.CurrentY = vinos_y - 10
Me.Picture1.FontSize = 13
Me.Picture1.Print Format(CDbl(Me.Text1.Text), "0.000")

Me.Picture1.CurrentX = -10
Me.Picture1.CurrentY = -10
a = (Me.Text1 ^ 3 * Sqr(2)) / 3
Me.Picture1.Print "Объем октаэдера со стороной " & Format(CDbl(Me.Text1.Text), "0.000") & vbCrLf & "   равен " & Format(a, "0.000")


End Sub
Private Sub Command1_Click() 'word
Dim w As Object

On Error GoTo noword
Set w = CreateObject("word.application")
 On Error GoTo 0
 w.Visible = True
 w.documents.Add
 w.selection.typetext "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
 w.Activate
 Set w = Nothing
 Exit Sub
noword:
    MsgBox "noword"
End Sub

Private Sub Command11_Click() 'заполнение таблица объемами
Me.MSFlexGrid1.Rows = Me.MSFlexGrid1.Rows + 1
Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 0) = Me.Text1.Text
Me.MSFlexGrid1.TextMatrix(Me.MSFlexGrid1.Rows - 1, 1) = Format((Me.Text1 ^ 3 * Sqr(2)) / 3, "0.000")
End Sub

Private Sub Command12_Click()
Call arhiv
End Sub

Sub prr() 'печать вектор

Dim matr(5, 2) As Double
Dim batr

matr(0, 0) = 5.001
matr(0, 1) = 4.028
matr(1, 0) = 0
matr(1, 1) = 4.028
matr(2, 0) = 1.732
matr(2, 1) = 5.087
matr(3, 0) = 6.732
matr(3, 1) = 5.087
matr(4, 0) = 3.366
matr(4, 1) = 8.586
matr(5, 0) = 3.366
matr(5, 1) = 0.53
batr = Array(0, 1, 1, 2, 2, 3, 3, 0, 1, 4, 2, 4, 0, 4, 3, 4, 1, 5, 2, 5, 0, 5, 3, 5)



Printer.ScaleMode = 0
Printer.DrawWidth = 10



Printer.Scale (-15, -20)-(45, 60)
For k = 0 To UBound(batr) Step 2
    mn = Int(batr(k))
    bn = Int(batr(k + 1))
    ff = matr()
    Printer.DrawStyle = vbDash
    Printer.Line (matr(mn, 0) * 5, matr(mn, 1) * 5)-(matr(bn, 0) * 5, matr(bn, 1) * 5)
Next

vinos_x = matr(2, 0) * 5 + ((matr(1, 0) - matr(2, 0)) * 5) / 2
vinos_y = matr(2, 1) * 5 + ((matr(1, 1) - matr(2, 1)) * 5) / 2

Printer.Line (vinos_x, vinos_y)-(vinos_x - 7, vinos_y - 7)
Printer.Line (vinos_x - 7, vinos_y - 7)-(vinos_x - 16, vinos_y - 7)
Printer.CurrentX = vinos_x - 15
Printer.CurrentY = vinos_y - 10
Printer.FontSize = 13
Printer.Print Format(CDbl(Me.Text1.Text), "0.000")

Printer.CurrentX = -10
Printer.CurrentY = -10
a = (Me.Text1 ^ 3 * Sqr(2)) / 3
Printer.Print "Объем октаэдера со стороной " & Format(CDbl(Me.Text1.Text), "0.000") & vbCrLf & "   равен " & Format(a, "0.000")
Printer.EndDoc








End Sub

Sub arhiv() 'архивация
Me.CommonDialog1.FileName = ""
Me.CommonDialog1.ShowSave

If Me.CommonDialog1.FileName <> "" Then
    Dim ShellApp As Object
    Open Me.CommonDialog1.FileName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
    
    Set ShellApp = CreateObject("shell.application")
    filetozip = "C:\1\фыр.txt" 'путь к вашему файлу
    ShellApp.namespace(Me.CommonDialog1.FileName).copyhere filetozip
End If

End Sub

Private Sub Command13_Click() 'Chrome c++
Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://onlinegdb.com/OeTv1VrMWG/"
End Sub

Private Sub Command14_Click()
Call draw
End Sub

Private Sub Command15_Click()
Me.Picture1.Visible = False
Me.Command15.Visible = False
End Sub

Private Sub Command16_Click()
print_mode = 1
Me.Command16.BackColor = &HFF00&
Me.Command16.Enabled = False
Me.Command17.BackColor = &HC0FFC0
Me.Command17.Enabled = True




End Sub

Private Sub Command17_Click()
print_mode = 2
Me.Command17.BackColor = &HFF00&
Me.Command17.Enabled = False
Me.Command16.Enabled = True
Me.Command16.BackColor = &HC0FFC0

End Sub

Private Sub Command2_Click() 'excel
Dim e As Object

On Error GoTo noexcel
Set e = CreateObject("excel.application")
 On Error GoTo 0
 e.Visible = True
 e.workbooks.Add
 e.ActiveSheet.Range("A1").Value = "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
 Set e = Nothing
 Exit Sub
noexcel:
    MsgBox "noexcel"

End Sub

Private Sub Command3_Click() 'PP
Dim p As Object

On Error GoTo nopowerpoint
Set p = CreateObject("powerpoint.application")
 On Error GoTo 0
 p.Visible = True
 p.Presentations.Add
 Set newslide = p.activepresentation.slides.Add(1, 11)
 Set bb = p.activepresentation.slides(1)
 bb.shapes(1).textframe.textrange.Text = "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
 p.Activate
 Set p = Nothing
 Exit Sub
nopowerpoint:
    MsgBox "nopowerpoint"

End Sub

Private Sub Command4_Click() 'Explorer
Dim IExplorer As Object
Set IExplorer = CreateObject("InternetExplorer.Application")
IExplorer.Visible = True
IExplorer.Navigate "объём октаэдра = " & (Me.Text1 ^ 3 * Sqr(2)) / 3
Set IExplorer = Nothing
End Sub

Private Sub Command5_Click() 'Solid
Dim part As Object
Dim longstatus As Long
On Error Resume Next
Set swapp = CreateObject("sldworks.application")
swapp.Visible = True
Set part = swapp.NewDocument("C:\Program Files\SolidWorks Corp\SolidWorks\lang\russian\Tutorial\part.prtdot", 0, 0, 0)
swapp.ActivateDoc2 "Деталь2", False, longstatus
Set part = swapp.ActiveDoc
Dim myModelView As Object
Set myModelView = part.ActiveView
myModelView.FrameState = swWindowState_e.swWindowMaximized
boolstatus = part.Extension.SelectByID2("Сверху", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
part.SketchManager.InsertSketch True
part.ClearSelection2 True
Dim vSkLines As Variant
vSkLines = part.SketchManager.CreateCenterRectangle(0, 0, 0, Me.Text1.Text / 1000, Me.Text1.Text / 1000, 0)
part.ClearSelection2 True
part.SketchManager.InsertSketch True
part.ShowNamedView2 "*Триметрия", 8
part.ClearSelection2 True
boolstatus = part.Extension.SelectByID2("Эскиз1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
Dim myFeature As Object
Set myFeature = part.FeatureManager.FeatureExtrusion2(False, False, False, 0, 0, Me.Text1.Text / 1000, Me.Text1.Text / 1000, True, True, False, False, 0.78539816339745, 0.78539816339745, False, False, False, False, True, True, True, 0, 0, False)
part.SelectionManager.EnableContourSelection = False
Set swmass = part.Extension.createmassproperty
dvolume = swmass.Volume
swapp.sendmsgtouser ("объём октаэдра = " & ((Me.Text1.Text ^ 3) * Sqr(2))) / 3
Me.Label2.Caption = (((Me.Text1.Text ^ 3) * Sqr(2)) / 3)
End Sub

Private Sub Command6_Click() 'open

Dim inData

Me.CommonDialog1.Filter = "Tекстовый файл (.txt)|*.txt"
Me.CommonDialog1.ShowOpen
  Open Me.CommonDialog1.FileName For Input As #1
  Input #1, inData
  bb = inData
  Close #1

For i = 1 To Len(bb)
    yy = Mid(bb, i, 1) Like "#"
    If yy = True Then kk = i
Next

Me.Text1.Text = kk

End Sub

Private Sub Command7_Click() 'save
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Me.CommonDialog1.Filter = "TXT|*.txt"
Me.CommonDialog1.ShowSave
Open Me.CommonDialog1.FileName For Output As #2
  Print #2, "объем октаэдра с длиной ребра " & " " & Me.Text1.Text & " " & " мм равен   " & s
  Close #2
End Sub

Private Sub Command8_Click() 'Voice
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
Dim sss As Object
Set sss = CreateObject("SAPI.SpVoice")
sss.Speak "volume of octahedron is " & (Replace(Format(s, "0.0"), ",", "."))

Const SAFT48kHz16BitStereo = 39
Const SSFMCreateForWrite = 3 ' Creates file even if file exists and so destroys or overwrites the existing file

Dim oFileStream, oVoice

Set oFileStream = CreateObject("SAPI.SpFileStream")
oFileStream.Format.Type = SAFT48kHz16BitStereo
oFileStream.Open "C:\Test\Sample.wav", SSFMCreateForWrite

Set oVoice = CreateObject("SAPI.SpVoice")
Set oVoice.AudioOutputStream = oFileStream
oVoice.Speak "volume of octahedron is " & (Replace(Format(s, "0.0"), ",", ".")) & "                                                                                         "

oFileStream.Close

End Sub


Private Sub Command9_Click()
s = ((Me.Text1.Text ^ 3) * Sqr(2)) / 3
a = "volume of octahedron so storonoy " & " " & Me.Text1.Text & " " & " mm is " & s

Clipboard.SetText (a)
tet = Clipboard.GetText()
MsgBox (tet)
End Sub

Private Sub Command10_Click()


If print_mode <> 0 Then
    If print_mode = 1 Then
        Call printr
    End If
    If print_mode = 2 Then
        Call prr
    End If
Else
    MsgBox "Выберите Printer Mode", vbCritical, "ERROR"
End If





    
End Sub


Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu mmSort
End If
End Sub
Sub printr() 'printer
draw_mode = True
Call draw
Printer.ScaleMode = vbCentimeters
Printer.PrintQuality = 10
Printer.PaperSize = 5
Printer.Orientation = 1
Printer.PaintPicture Me.Picture1.Image, 0, 0, (Printer.ScaleWidth * 4) / 4, (Printer.ScaleWidth * 3) / 4
draw_mode = False
Call draw
Printer.EndDoc


End Sub
Private Sub mSortV_Click()
Me.MSFlexGrid1.Sort = 1
End Sub

