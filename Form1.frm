VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Colour Change"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBGR 
      Caption         =   "Sea Green"
      Height          =   375
      Index           =   4
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optBGR 
      Caption         =   "Yellow"
      Height          =   375
      Index           =   3
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optBGR 
      Caption         =   "Blue"
      Height          =   375
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optBGR 
      Caption         =   "Green"
      Height          =   375
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optBGR 
      Caption         =   "Red"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton optRGB 
      Caption         =   "Sea Green"
      Height          =   375
      Index           =   4
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optRGB 
      Caption         =   "Yellow"
      Height          =   375
      Index           =   3
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optRGB 
      Caption         =   "Blue"
      Height          =   375
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optRGB 
      Caption         =   "Green"
      Height          =   375
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optRGB 
      Caption         =   "Red"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   92
      Left            =   0
      TabIndex        =   92
      Top             =   11040
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   91
      Left            =   0
      TabIndex        =   91
      Top             =   10920
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   90
      Left            =   0
      TabIndex        =   90
      Top             =   10800
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   89
      Left            =   0
      TabIndex        =   89
      Top             =   10680
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   88
      Left            =   0
      TabIndex        =   88
      Top             =   10560
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   87
      Left            =   0
      TabIndex        =   87
      Top             =   10440
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   86
      Left            =   0
      TabIndex        =   86
      Top             =   10320
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   85
      Left            =   0
      TabIndex        =   85
      Top             =   10200
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   84
      Left            =   0
      TabIndex        =   84
      Top             =   10080
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   83
      Left            =   0
      TabIndex        =   83
      Top             =   9960
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   82
      Left            =   0
      TabIndex        =   82
      Top             =   9840
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   81
      Left            =   0
      TabIndex        =   81
      Top             =   9720
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   80
      Left            =   0
      TabIndex        =   80
      Top             =   9600
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   79
      Left            =   0
      TabIndex        =   79
      Top             =   9480
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   78
      Left            =   0
      TabIndex        =   78
      Top             =   9360
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   77
      Left            =   0
      TabIndex        =   77
      Top             =   9240
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   76
      Left            =   0
      TabIndex        =   76
      Top             =   9120
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   75
      Left            =   0
      TabIndex        =   75
      Top             =   9000
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   74
      Left            =   0
      TabIndex        =   74
      Top             =   8880
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   73
      Left            =   0
      TabIndex        =   73
      Top             =   8760
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   72
      Left            =   0
      TabIndex        =   72
      Top             =   8640
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   71
      Left            =   0
      TabIndex        =   71
      Top             =   8520
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   70
      Left            =   0
      TabIndex        =   70
      Top             =   8400
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   69
      Left            =   0
      TabIndex        =   69
      Top             =   8280
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   68
      Left            =   0
      TabIndex        =   68
      Top             =   8160
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   67
      Left            =   0
      TabIndex        =   67
      Top             =   8040
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   66
      Left            =   0
      TabIndex        =   66
      Top             =   7920
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   65
      Left            =   0
      TabIndex        =   65
      Top             =   7800
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   64
      Left            =   0
      TabIndex        =   64
      Top             =   7680
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   63
      Left            =   0
      TabIndex        =   63
      Top             =   7560
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   62
      Left            =   0
      TabIndex        =   62
      Top             =   7440
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   61
      Left            =   0
      TabIndex        =   61
      Top             =   7320
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   60
      Left            =   0
      TabIndex        =   60
      Top             =   7200
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   59
      Left            =   0
      TabIndex        =   59
      Top             =   7080
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   58
      Left            =   0
      TabIndex        =   58
      Top             =   6960
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   57
      Left            =   0
      TabIndex        =   57
      Top             =   6840
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   56
      Left            =   0
      TabIndex        =   56
      Top             =   6720
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   55
      Left            =   0
      TabIndex        =   55
      Top             =   6600
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   54
      Left            =   0
      TabIndex        =   54
      Top             =   6480
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   53
      Left            =   0
      TabIndex        =   53
      Top             =   6360
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   52
      Left            =   0
      TabIndex        =   52
      Top             =   6240
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   51
      Left            =   0
      TabIndex        =   51
      Top             =   6120
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   50
      Left            =   0
      TabIndex        =   50
      Top             =   6000
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   49
      Left            =   0
      TabIndex        =   49
      Top             =   5880
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   48
      Left            =   0
      TabIndex        =   48
      Top             =   5760
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   47
      Left            =   0
      TabIndex        =   47
      Top             =   5640
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   46
      Left            =   0
      TabIndex        =   46
      Top             =   5520
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   45
      Left            =   0
      TabIndex        =   45
      Top             =   5400
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   44
      Left            =   0
      TabIndex        =   44
      Top             =   5280
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   43
      Left            =   0
      TabIndex        =   43
      Top             =   5160
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   42
      Left            =   0
      TabIndex        =   42
      Top             =   5040
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   41
      Left            =   0
      TabIndex        =   41
      Top             =   4920
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   40
      Left            =   0
      TabIndex        =   40
      Top             =   4800
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   39
      Left            =   0
      TabIndex        =   39
      Top             =   4680
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   38
      Left            =   0
      TabIndex        =   38
      Top             =   4560
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   37
      Left            =   0
      TabIndex        =   37
      Top             =   4440
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   36
      Left            =   0
      TabIndex        =   36
      Top             =   4320
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   35
      Left            =   0
      TabIndex        =   35
      Top             =   4200
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   34
      Left            =   0
      TabIndex        =   34
      Top             =   4080
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   33
      Left            =   0
      TabIndex        =   33
      Top             =   3960
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   32
      Left            =   0
      TabIndex        =   32
      Top             =   3840
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   31
      Left            =   0
      TabIndex        =   31
      Top             =   3720
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   30
      Left            =   0
      TabIndex        =   30
      Top             =   3600
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   29
      Left            =   0
      TabIndex        =   29
      Top             =   3480
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   28
      Left            =   0
      TabIndex        =   28
      Top             =   3360
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   27
      Left            =   0
      TabIndex        =   27
      Top             =   3240
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   26
      Left            =   0
      TabIndex        =   26
      Top             =   3120
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   25
      Left            =   0
      TabIndex        =   25
      Top             =   3000
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   24
      Left            =   0
      TabIndex        =   24
      Top             =   2880
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   23
      Left            =   0
      TabIndex        =   23
      Top             =   2760
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   22
      Left            =   0
      TabIndex        =   22
      Top             =   2640
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   21
      Left            =   0
      TabIndex        =   21
      Top             =   2520
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   20
      Left            =   0
      TabIndex        =   20
      Top             =   2400
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   19
      Left            =   0
      TabIndex        =   19
      Top             =   2280
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   18
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   17
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   16
      Left            =   0
      TabIndex        =   16
      Top             =   1920
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   15
      Left            =   0
      TabIndex        =   15
      Top             =   1800
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   14
      Left            =   0
      TabIndex        =   14
      Top             =   1680
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   13
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   15
   End
   Begin VB.Label lblRGB 
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColour As Integer             'stores the currently selected option button
Dim iCount As Integer              'temp variable used for the for/next loops
Dim iWidth As Integer              'to store the width of the form
Dim iHeight As Integer             'to store the height of the form
Dim iTop As Integer                'to store the distance from the top of the form for each label
Dim bColour As Boolean             'to store the current direction of the colours

Private Sub Form_Load()
iColour = 2                        'sets the starting colour as blue
Reset_Controls                     'set the colours when the form is first loaded
End Sub

Private Sub Reset_Controls()
'
Form1.Hide                         'hide the form while it's being updated
'
iWidth = Form1.ScaleWidth          'Pickup the size of the form
iHeight = Form1.ScaleHeight / 92   'pickup the height of the form and divide it by the number of labels
iTop = 0                           'set the distance from the top to zero
'
For iCount = 0 To 92               'start the for/next loop to cycle through all the labels
    lblRGB(iCount).Width = iWidth  'set the width of the labels to match the width of the form
    lblRGB(iCount).Top = iTop      'set the distance from the top for this label
    iTop = iTop + iHeight          'increase the distance from the top by adding the current distance to the value stored in iHeight
Next iCount                        'loop back to the start until all the labels have been processed
'
If bColour = False Then            'check current direction of the colours
    Colour_Change                  'if downwards run the Colour_Change sub procedure
    Else                           'if not
    Colour_Reverse                 'then run the Colour_Reverse sub procedure
End If                             'end the if statement
'
Form1.Show                         'now that the updating is done, show the form again
'
End Sub

Private Sub Form_Resize()
Reset_Controls
End Sub

Private Sub optRGB_Click(Index As Integer)
iColour = Index                   'store the option button selected in iColour
bColour = False                   'stores the direction of the colour change
Colour_Change                     'proceed to the Colour_Change sub procedure
End Sub

Private Sub optBGR_Click(Index As Integer)
iColour = Index                   'store the option button selected in iColour
bColour = True                    'stores the direction of the colour change
Colour_Reverse                    'proceed to the Colour_Reverse sub procedure
End Sub

Private Sub Colour_Change()
'
'Cycles through each of the labels, having checked which option button has been selected (iColour)
'and then sets the backcolor accordingly. Colour starts off black and works it's way DOWN the screen
'to the desired colour.
'
For iCount = 0 To 92
    If iColour = 0 Then lblRGB(iCount).BackColor = RGB(iCount * 2.741935, 0, 0)
    If iColour = 1 Then lblRGB(iCount).BackColor = RGB(0, iCount * 2.741935, 0)
    If iColour = 2 Then lblRGB(iCount).BackColor = RGB(0, 0, iCount * 2.741935)
    If iColour = 3 Then lblRGB(iCount).BackColor = RGB(iCount * 2.741935, iCount * 2.741935, 0)
    If iColour = 4 Then lblRGB(iCount).BackColor = RGB(0, iCount * 2.741935, iCount * 2.741935)
Next iCount
End Sub

Private Sub Colour_Reverse()
'
'Cycles through each of the labels, having checked which option button has been selected (iColour)
'and then sets the backcolor accordingly.  Colour starts off black and works its way UP the screen
'to the desired colour.
'
For iCount = 92 To 0 Step -1
    If iColour = 0 Then lblRGB(iCount).BackColor = RGB((92 - iCount) * 2.741935, 0, 0)
    If iColour = 1 Then lblRGB(iCount).BackColor = RGB(0, (92 - iCount) * 2.741935, 0)
    If iColour = 2 Then lblRGB(iCount).BackColor = RGB(0, 0, (92 - iCount) * 2.741935)
    If iColour = 3 Then lblRGB(iCount).BackColor = RGB((92 - iCount) * 2.741935, (92 - iCount) * 2.741935, 0)
    If iColour = 4 Then lblRGB(iCount).BackColor = RGB(0, (92 - iCount) * 2.741935, (92 - iCount) * 2.741935)
Next iCount
End Sub


'
' Footnote
'

' This project works by resizing a set of labels which were created as a control array.  The labels
' were then resized to have a width of 1, thus appearing invisible on the form. Now you can happily
' add your own controls and resize the form and the project will do the colour at the form load
' stage, ensuring that the form is filled.
'
' I know this can be done other ways, the most common example being line fills, however, this is
' simple and effective and doesn't have refresh problems unlike some methods.
'
' I've commented the code VERY heavily but if theres anything that you don't understand then let me
' know.
'
' Please remember, this is aimed at beginners, just to show some simple effects...
'
' ----------------------------------------------------------------------------------------------------


