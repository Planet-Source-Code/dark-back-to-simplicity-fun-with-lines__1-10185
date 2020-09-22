VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Back to simplicity: Fun with lines"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   1440
      X2              =   2640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   3480
      X2              =   2280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   2640
      X2              =   3480
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   1440
      X2              =   2280
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TButton2"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TButton1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   3840
      X2              =   4680
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3000
      Y1              =   1080
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   3000
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   240
      X2              =   1920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   1080
      X2              =   240
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   1080
      X2              =   1920
      Y1              =   240
      Y2              =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cHighlight = &H80000014 'nice little constant for the highlight color
Const cShadow = &H80000010 ' nice little constant for the shadow color
Dim Lb1Posx As Integer, Lb1Posy As Integer ' these will be the positions (x,y) of label1
Dim Lb2Posx As Integer, Lb2Posy As Integer ' these will be the positions (x,y) of label2
Dim Lb3Posx As Integer, Lb3Posy As Integer ' these will be the positions (x,y) of label3
'
' This can be easily edited and made into a user control for even easier usage
' The mouse up/down events are obviously triggered by the labels, so that sort of limits
'   the mouse event "real estate", but this could also be fixed with a few invisible labels
'   who in turn triger the respective events.. there is also a small glitch when a user double
'   clicks an area, (even when i *did* have a doubleclick event trigger.. dont know why, but
'   its not that big of a deal..
'
'You can change the code to even make these into check-buttons that stay down/up on a click event,
'or make a "floating button" (MS-style) by making the lines invisible until the mouse is over
'the label..
'
'Remeber, Skill will get you in the door... imagination will make you stand out from the rest.


Private Sub Form_Load()
Lb1Posx = Label1.Left ' save ther positions in a variable
Lb1Posy = Label1.Top ' dito
Lb2Posx = Label2.Left ' yadayada
Lb2Posy = Label2.Top ' i'm sure its obvious by now..
Lb3Posx = Label3.Left
Lb3Posy = Label3.Top

Label3.Caption = vbCrLf & "Tbutton3" ' this is just to add a cr/lf
        'to center the text (on the y axis) otherwise it would be too high
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'now lets swap the colors so the "3D" effect looks decent
Line1.BorderColor = cHighlight
Line2.BorderColor = cShadow
Line3.BorderColor = cHighlight
'now let's move the label down a tad to
'   compliment the "down" effect
Label1.Top = Lb1Posy + 10
Label1.Left = Lb1Posx + 10

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'swap the colors back
Line1.BorderColor = cShadow
Line2.BorderColor = cHighlight
Line3.BorderColor = cShadow
'move the label back
Label1.Top = Lb1Posy
Label1.Left = Lb1Posx

End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'the rest is the same, 'cept its for the 2nd label
'   and of course, considering the angle, 2 lines
'   are shadowed instead of one..
Line4.BorderColor = cShadow
Line6.BorderColor = cHighlight
Line5.BorderColor = cShadow
Label2.Top = Lb2Posy + 10
Label2.Left = Lb2Posx + 10


End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line4.BorderColor = cHighlight
Line6.BorderColor = cShadow
Line5.BorderColor = cHighlight
Label2.Top = Lb2Posy
Label2.Left = Lb2Posx

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line7.BorderColor = cShadow
Line10.BorderColor = cShadow
Line8.BorderColor = cHighlight
Line9.BorderColor = cHighlight
Label3.Left = Lb3Posx + 10
Label3.Top = Lb3Posy + 10
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line7.BorderColor = cHighlight
Line10.BorderColor = cHighlight
Line8.BorderColor = cShadow
Line9.BorderColor = cShadow
Label3.Left = Lb3Posx
Label3.Top = Lb3Posy
End Sub
