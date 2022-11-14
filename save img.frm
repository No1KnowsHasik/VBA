VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   14115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   14115
   ScaleWidth      =   14310
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   3120
      Width           =   4935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   1815
      Left            =   840
      TabIndex        =   3
      Top             =   9960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   7680
      Width           =   7335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1575
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type DEVMODE
dmDeviceName As String * CCHDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCHFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type

Public Sub Capture(control_hWnd As Long, fNAME As String, Optional OnlyToClipBoard As Boolean = False)
On Error GoTo ErrorCapture
Dim sp As RECT, x As Long
If fNAME <> "" Then
x = GetWindowRect(control_hWnd, sp)
ScrnCap sp.Left, sp.Top, sp.Right, sp.Bottom
If OnlyToClipBoard = False Then
SavePicture Clipboard.GetData, fNAME


'Printer.Print Clipboard.GetData
'Printer.EndDoc

End If
End If
Exit Sub
ErrorCapture:
MsgBox Err & ":Error in Caputre(). Error Message:" & Err.Description, vbCritical, "Warning"
Exit Sub
End Sub

Private Sub ScrnCap(Lt, Top, Rt, Bot)
On Error GoTo ErrorScrnCap
Dim rWIDTH As Long, rHEIGHT As Long
Dim SourceDC As Long, DestDC As Long, bHANDLE As Long, Wnd As Long
Dim dHANDLE As Long, dm As DEVMODE
rWIDTH = Rt - Lt
rHEIGHT = Bot - Top
SourceDC = CreateDC("DISPLAY", 0&, 0&, dm)
DestDC = CreateCompatibleDC(SourceDC)
bHANDLE = CreateCompatibleBitmap(SourceDC, rWIDTH, rHEIGHT)
SelectObject DestDC, bHANDLE
BitBlt DestDC, 0, 0, rWIDTH, rHEIGHT, SourceDC, Lt, Top, &HCC0020
Wnd = 0
OpenClipboard Wnd
EmptyClipboard
SetClipboardData 2, bHANDLE
CloseClipboard
DeleteDC DestDC
ReleaseDC dHANDLE, SourceDC

'Printer.Print GetDesktopWindow
'Printer.EndDoc
'
Exit Sub
ErrorScrnCap:
MsgBox Err & ":Error in ScrnCap(). Error Message:" & Err.Description, vbCritical, "Warning"
Exit Sub
End Sub

Public Sub CaptureDesktop()
On Error GoTo ErrorCaptureDesktop
Dim dhWND As Long, sp As RECT, x As Long
dhWND = GetDesktopWindow
If dhWND <> 0 Then
x = GetWindowRect(dhWND, sp)
ScrnCap sp.Left, sp.Top, sp.Right, sp.Bottom
End If
Exit Sub
ErrorCaptureDesktop:
MsgBox Err & ":Error in CaptureDesktop. Error Message: " & Err.Description, vbCritical, "Warning"
Exit Sub
End Sub

Private Sub Form_Load()
Command1.Caption = "Экран"
Command2.Caption = "Форма"
Command3.Caption = "Кнопка"
Command4.Caption = "Текстовое окно"
End Sub
Private Sub Command1_Click()
On Error Resume Next
Call CaptureDesktop
SavePicture Clipboard.GetData, "C:\1\desktop.bmp"
MsgBox "Картинка экрана сохранена в C:\1\desktop.bmp"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Call Capture(Me.hwnd, "C:\1\form.bmp")
MsgBox "Картинка формы сохранена в C:\1\form.bmp"
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call Capture(Me.Command1.hwnd, "C:\1\button.bmp")
MsgBox "Картинка кнопки сохранена в C:\1\button.bmp"
End Sub

Private Sub Command4_Click()
On Error Resume Next
Call Capture(Me.Dir1.hwnd, "C:\1\drv.bmp")
MsgBox "Картинка DriveListBox сохранена в C:\1\drv.bmp"
End Sub
