VERSION 5.00
Begin VB.Form frmFlames 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flames by SebaMix"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   162
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   60
      Top             =   60
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "www.infotrade.it/sebamix"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1260
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "sebamix@hotmail.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SebaMix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   900
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   660
      TabIndex        =   0
      Top             =   840
      Width           =   915
   End
End
Attribute VB_Name = "frmFlames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This small intro was coded by SebaMix at the begin of 2002;
' You can use this code as you wish, but remember that I don't take
' any kind of responsability about it. All the code is give "As is".
' Just remember to put my name on the credits of your proggie... :)
' If you wanna write me this is the address : sebamix@hotmail.com
' Remember that I can answer you question only on my free-time, so don't
' wait for a fast reply.

' I don't know why I write this code... I think that I'll put it into some
' eggs of my programs... ^_^

' NOTES
' This program explain how to put a graphical flame effect to a form.
' The flame is small only for a speed problem (bigger flames makes slower
' the proggie).
' The form's backcolor depend by the MyPalette(1) element.
' I write it with VB6, but I think that it work fine with other versions of VB.

' You can download this code from http://www.planetsourcecode.com and from
' my site (the current address is change from http://sebamix.supereva.it to
' http://www.infotrade.it/sebamix)

' Enjoy!

' P.S. Sorry my bad english... ^_^

Option Explicit

' Use this vector to store the colors of the palette;
Private MyPalette(1 To 255) As Long
' Use this bi-dimensional vector to store the flame pixels;
Private MyField(1 To 100, 1 To 25) As Long

' Here is the begin of all;
Private Sub Form_Load()
    ' Var. that I use for storing the green gradient of the palette;
    Dim g As Long
    ' Var. that I use for storing the red gradient of the palette;
    Dim r As Long
    ' A simple counter;
    Dim i As Integer
    
    ' The palette structure:
    '  0______________________255
    ' |________________________|
    '  ^          ^           ^
    ' Black       Red         Yellow
    ' ^ (Becouse black is the background of the form)
    
    ' Now I must build the palette; you can change this algorithm to obtain
    ' different colors combinations;
    g = 0 ' Green gradient start from 0;
    r = 50 ' Red gradient start form 50;
    For i = 1 To 255
        ' Increase the red gradient;
        r = r + 1
        ' If red gradient major thank 200 lets add some green,
        ' so we obtain some yellow at the end of the palette;
        If r > 200 Then
            g = g + 4 ' Green gradiend increase by 4;
        End If
        
        ' Now build the color and store it inside the palette's vector.
        ' I use the standard RGB function for this;
        MyPalette(i) = RGB(r, g, 0)
        ' By using this line you will see a blue flame...
        'MyPalette(i) = RGB(0, g, r)
        ' Or use this one... ^_^
        'MyPalette(i) = RGB(g, 0, r)
    Next i
    
    ' By uncommenting these lines you will see the palette print on the form;
    'For i = 255 To 1 Step -1
    '    Me.Line (i, 10)-(i + 1, 15), MyPalette(i)
    'Next i
    ' Let's set the back color of the form;
    Me.BackColor = MyPalette(1)
    ' Let's the timer starts his work;
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    
    Dim y As Long
    Dim x As Long
    
    Dim cx As Long
    Dim cy As Long
    
    Dim v1 As Long
    Dim v2 As Long
    Dim v3 As Long
    Dim MidVal As Long
    
    ' Stop the timer;
    Timer1.Enabled = False
    
    ' We must calculate a random line of pixels on the bottom
    ' of the flame.
    For i = 1 To 25
        ' The random number will be between 101 and 255;
        ' Put the random color on the bottom of our flame-vector;
        MyField(100, i) = 100 + (Rnd * 154) + 1
    Next i
    
    ' TIP : try to uncomment this block of code for see a cool effect;
    'MyField(90, 10) = MyPalette(1)
    'MyField(90, 11) = MyPalette(1)
    'MyField(90, 12) = MyPalette(1)
    'MyField(90, 13) = MyPalette(1)
    'MyField(90, 14) = MyPalette(1)
    ' TIP2 : or this block of code;
    'MyField(90, 10) = MyPalette(255)
    'MyField(90, 11) = MyPalette(255)
    'MyField(90, 12) = MyPalette(255)
    'MyField(90, 13) = MyPalette(255)
    'MyField(90, 14) = MyPalette(255)
    
    cy = 99
    
    For y = 99 To 1 Step -1
        cx = 10
        For x = 1 To 25
            
            ' This is how the flame works:
            ' We have a matrix of pixels; we only must analize each pixel
            ' and recalculate it by taking the nearest values and dividing
            ' them by 3.
            '  ___ ___ ___ ___ ___ _...
            ' |   |   |   |   |   |
            ' |___|___|___|___|___|_...
            ' |   | V2| X | V3|   |
            ' |___|___|___|___|___|_... ^
            ' |   |   | V1|   |   |     |
            ' |___|___|___|___|___|_... |
            '
            ' X is the pixel that we are analyzing;
            ' X = (V1 + V2 + V3) / 3
            ' We must start from the bottom of the matrix (y=99);
            ' When we calculate the new value of a pixel, we print it
            ' to the form. This give the movement effect;
            
            ' Reset all the values for calculate the new pixel value;
            v1 = 0
            v2 = 0
            v3 = 0
            
            ' Now start read each value to calculate the new pixel's color;
            
            v1 = MyField(y + 1, x)
        
            If x > 1 Then
                v2 = MyField(y, x - 1)
            End If
            
            If x < 25 Then
                v3 = MyField(y, x + 1)
            End If
            
            ' Calculate the value;
            MidVal = (v1 + v2 + v3) / 3
            
            ' Make sure that the new value is a good value (>=1 and <= 255);
            If MidVal <= 0 Then
                MidVal = 1
            ElseIf MidVal > 255 Then
                MidVal = 255
            End If
            
            ' Store the new value to the matrix;
            ' Remember: on the matrix I store the ordinal value of the MyPalette
            ' vector. For example we can find a value >=1 and <= 255;
            ' If we want the color, we must take the MyPalette(MyField(y,x));
            MyField(y, x) = MidVal
            ' Write the pixel. I don't know why, but I must draw a 2 pixel
            ' line, or I don't see nothing... :(
            Me.Line (cx, cy)-(cx + 1, cy), MyPalette(MyField(y, x))
            ' Increase the x value;
            cx = cx + 1
        Next x
        ' Increase the y value;
        cy = cy - 1
    Next y
    
    ' Restart the timer;
    Timer1.Enabled = True
End Sub

' That's all, folks! ^_^
' Enjoy it!!!
