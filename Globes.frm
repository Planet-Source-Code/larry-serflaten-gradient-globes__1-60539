VERSION 5.00
Begin VB.Form Globes 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2175
   ClientLeft      =   2100
   ClientTop       =   3345
   ClientWidth     =   2040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   -15
      Picture         =   "Globes.frx":0000
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   28800
   End
End
Attribute VB_Name = "Globes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Quit As Boolean
Private Colors() As Long
Private Delay As Long

Const ChangeOver As Long = 20000


Private Sub Form_Activate()
Dim msg As String
    msg = "Left click on screen to set delay time." & vbCrLf
    msg = msg & "(Click higher for faster speeds, lower for slower speeds)"
    msg = msg & vbCrLf & vbCrLf
    msg = msg & "Right click on screen to end program."
    MsgBox msg, vbInformation, "Gradient Globes"
End Sub

Private Sub Form_Load()
Dim i As Long
    Randomize
    WindowState = vbMaximized
    DrawWidth = 4
    Delay = 55000
    Timer1.Interval = 10
    
    ReDim Colors(0 To Picture1.ScaleWidth - 1)
    For i = 0 To UBound(Colors)
      Colors(i) = Picture1.Point(i, 0)
    Next
End Sub

' User input to set speed and exit program

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' Exit on any key press
    Quit = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
    Case vbLeftButton
      ' Set delay time
      Timer1.Enabled = False
      Delay = y * 500
      Timer1.Interval = Delay \ 200
      Timer1.Enabled = (Delay > ChangeOver)
      Timer1_Timer
    Case vbRightButton
      ' Exit program
      Quit = True
    End Select
End Sub

' Timer handles slower speeds, inner loop handles fast speeds

Private Sub Timer1_Timer()

    If Delay <= ChangeOver Then
        ' Tight loop for fastest speeds
        Timer1.Enabled = False
        Do
          DrawCircle
          Wait
          If Delay > ChangeOver Then Exit Do
        Loop Until Quit
        Timer1.Enabled = True
    Else
        ' Else let Timer handle it
        DrawCircle
    End If
    
    Timer1.Interval = Delay \ 200
    
    If Quit Then
      Timer1.Enabled = False
      Unload Me
    End If

End Sub

Private Sub DrawCircle()
Dim klr&, wid&, rad&, hgt&, x&, y As Long
    ' Cache properties
    wid = ScaleWidth
    hgt = ScaleHeight
    ' Pick position
    x = Int(Rnd * wid * 0.8) + (wid * 0.1)
    y = Int(Rnd * hgt * 0.8) + (hgt * 0.1)
    ' Pick radius, color
    rad = Int(Rnd * hgt * 0.4) + (hgt * 0.1)
    klr = Int(Rnd * 12) * 160 + 6
    If klr < 5 Then klr = 4
    ' Draw globe
    Do While rad > 0
      Circle (x, y), rad, Colors(klr)
      klr = klr + 1
      rad = rad - 3
    Loop
    
' Set to True to see values
#If False Then
  ShowDelay
#End If
    
End Sub

Private Sub Wait()
Dim cnt As Long
    Do While cnt <= Delay
      DoEvents
      cnt = cnt + 1
      If Quit Then Exit Sub
    Loop
End Sub

Private Sub ShowDelay()
    ' For debugging change over from timer to fast loop
    If Timer1.Enabled Then
      Line (0, 0)-(128, 12), vbYellow, BF
    Else
      Line (0, 0)-(128, 12), vbWhite, BF
    End If
    PSet (4, 0), vbWhite
    Print "T="; Timer1.Interval, "D="; Delay
End Sub
