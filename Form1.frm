VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualization Demo using MS Agent's Direct Speech Recognition (DSR)  -  by CovertLoop"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   960
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DirectSR1 
      Height          =   375
      Left            =   6120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   6360
      X2              =   10080
      Y1              =   3840
      Y2              =   8400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   2760
      X2              =   5640
      Y1              =   360
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   6000
      X2              =   480
      Y1              =   4320
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   11520
      X2              =   6000
      Y1              =   720
      Y2              =   3600
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      Height          =   855
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Dash-Dot-Dot
      Height          =   855
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The "Commands.txt" file is used
'for recognizing the sound the
'computer is playing.
'The entire alphabet was used to
'cover most of the music/sound
'playing.  Generally, if you want
'the program to recognize a single
'word, you would put that word in
'the "Commands.txt" file.  Since
'the program never knows what sound
'or word will come out of a song,
'I have chosen every letter in the
'alphabet... which seems to work fine.

Private Sub DirectSR1_VUMeter(ByVal beginhi As Long, ByVal beginlo As Long, ByVal level As Long)
'This gets the value of the
'sound being played and sets
'both of the shapes' size and
'postion accordingly.

'Hide the shapes
Shape1.Visible = False
Shape2.Visible = False

'Set the width to the sound
    Shape1.Width = level / 20
    Shape2.Width = level / 10
    
'Set the height to the sound
    Shape1.Height = level / 20
    Shape2.Height = level / 10
    
'Center the shapes vertically
    Shape1.Top = Form1.Height / 2 - Shape1.Height / 2
    Shape2.Top = Form1.Height / 2 - Shape2.Height / 2
    
'Center the shapes horizontally
    Shape1.Left = Form1.Width / 2 - Shape1.Width / 2
    Shape2.Left = Form1.Width / 2 - Shape2.Width / 2

'Show the shapes
Shape1.Visible = True
Shape2.Visible = True
End Sub


Private Sub Form_Load()
'Set up the DSR
    Call DirectSR1.Deactivate
    Call DirectSR1.GrammarFromFile(App.Path & "\Commands.Txt")
    Call DirectSR1.Activate

'Reveal the form
Show
End Sub


Private Sub Level1_Change()
Change1.Caption = Val(Change1.Caption) + 1
End Sub

Private Sub Level2_Change()
Change2.Caption = Val(Change2.Caption) + 1
End Sub

Private Sub Level3_Change()
Change3.Caption = Val(Change3.Caption) + 1
End Sub

Private Sub Level4_Change()
Change4.Caption = Val(Change4.Caption) + 1
End Sub

Private Sub Level5_Change()
Change5.Caption = Val(Change5.Caption) + 1
End Sub

Private Sub Level6_Change()
Change6.Caption = Val(Change6.Caption) + 1
End Sub

Private Sub Timer1_Timer()
'This keeps the ends of the lines
'on the top, bottom, left, and right
'of Shape2
    Line1.Y2 = Shape2.Top
    Line1.X2 = Shape2.Left + (Shape2.Width / 2)
    Line2.Y1 = Shape2.Top + Shape2.Height
    Line2.X1 = Shape2.Left + (Shape2.Width / 2)

    Line3.X2 = Shape2.Left
    Line3.Y2 = Shape2.Top + (Shape2.Height / 2)
    Line4.X1 = Shape2.Left + Shape2.Width
    Line4.Y1 = Shape2.Top + (Shape2.Height / 2)

End Sub


Private Sub Timer2_Timer()

End Sub


