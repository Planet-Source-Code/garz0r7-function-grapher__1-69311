VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Grapher v1.0"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   15240
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "SAVE GRAPH "
      Height          =   495
      Left            =   6240
      TabIndex        =   28
      Top             =   9480
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PLOT !"
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   8880
      Width           =   1575
   End
   Begin VB.ListBox roots 
      Height          =   1815
      Left            =   12840
      TabIndex        =   20
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox steps 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Text            =   "400"
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CheckBox autoy 
      Caption         =   "AUTO Y ?"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   8400
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin MSScriptControlCtl.ScriptControl tmpScript 
      Left            =   8400
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSScriptControlCtl.ScriptControl vbscript 
      Left            =   9000
      Top             =   9480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ComboBox txtfunc 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Text            =   "3*SIN(X)*COS(2*X)"
      Top             =   7560
      Width           =   5655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "DISPLAY DERIVATIVE?"
      Height          =   615
      Left            =   6360
      TabIndex        =   11
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ZOOM -"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ZOOM +"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   8040
      Width           =   855
   End
   Begin VB.TextBox ymx 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Text            =   "20"
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox ymn 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "-20"
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox xmx 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Text            =   "10"
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox xmn 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Text            =   "-10"
      Top             =   8040
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8325
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox display 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H00E0E0E0&
      Height          =   6975
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   6915
      ScaleWidth      =   15075
      TabIndex        =   0
      Top             =   360
      Width           =   15135
   End
   Begin VB.Label f_inf 
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   27
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label mouse_inf 
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   26
      Top             =   7560
      Width           =   4095
   End
   Begin VB.Label ye 
      Caption         =   "Label8"
      Height          =   255
      Left            =   9360
      TabIndex        =   25
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label xe 
      Caption         =   "Label7"
      Height          =   255
      Left            =   8040
      TabIndex        =   24
      Top             =   9600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label RO 
      Alignment       =   2  'Center
      Caption         =   "ROOTS:"
      Height          =   255
      Left            =   12840
      TabIndex        =   21
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label ymaxval 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   9240
      Width           =   3855
   End
   Begin VB.Label yminval 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   9600
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Function's Max Value :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   9240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Function's Min Value :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "STEPS :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Ymin , Ymax VALUES :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Xmin , Xmax VALUES :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "DISPLAY FUNCTION :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Menu ext 
      Caption         =   "Exit.."
   End
   Begin VB.Menu abt 
      Caption         =   "About.."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mousex As Double
Dim mousey As Double
Dim yymax As Double
Dim yymin As Double
Option Explicit

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
Const factor = 0.1
Dim min As Double
Dim max As Double
Dim dx As Double
min = xmn
max = xmx
dx = max - min
xmn = Round(min + factor * dx, 4)
xmx = Round(max - factor * dx, 4)

vbscript.Reset
vbscript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
plot xmn, xmx, ymn, ymx, steps
End Sub

Private Sub Command2_Click()
Const factor = 0.1
Dim min As Double
Dim max As Double
Dim dx As Double
min = xmn
max = xmx
dx = max - min
xmn = Round(min - factor * dx, 4)
xmx = Round(max + factor * dx, 4)

vbscript.Reset
vbscript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
plot xmn, xmx, ymn, ymx, steps
End Sub

Private Sub Command3_Click()
vbscript.Reset
vbscript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
txtfunc.AddItem (txtfunc)
plot xmn, xmx, ymn, ymx, steps
End Sub

Private Sub Command4_Click()
Dim path As String
path = App.path + "\LastGraph" + Str(Timer) + ".bmp"
SavePicture display.Image, path
End Sub

Private Sub display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousex = XDISPLAY(X, 0, display.Width, xmn, xmx)
xe = Round(mousex, 4)
mousey = YDISPLAY(Y, 0, display.Height, yymax, yymin)
'mousey = vbscript.Eval("f(" + Str(xe) + ")")
ye = Round(mousey, 4)
mouse_inf = "Mouse Pos (x,y) : ( " + xe + " , " + ye + " )"

mousey = vbscript.Eval("f(" + Str(xe) + ")")
ye = Round(mousey, 4)
f_inf = "F( " + xe + " ) = " + ye

End Sub

Private Sub Form_Load()
Form1.Show
display.ScaleWidth = display.Width
display.ScaleHeight = display.Height

vbscript.Reset
vbscript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
txtfunc.AddItem (txtfunc)
plot xmn, xmx, ymn, ymx, steps
'a = YDISPLAY(-8, 1, -5, -2, 5)

End Sub
Function XDISPLAY(ByVal x1 As Double, xs1 As Double, xe1 As Double, xs2 As Double, xe2 As Double) As Double
'normal display, sign(xe1-xs1)=sign(xe2-xs2)
'xs1----x1------------xe1
'TO
'xs2--------XDISPLAY-------------xe2

XDISPLAY = (x1 - xs1) * ((xe2 - xs2) / (xe1 - xs1)) + xs2
End Function
Function YDISPLAY(ByVal y1 As Double, ys1 As Double, ye1 As Double, ys2 As Double, ye2 As Double) As Double
On Error GoTo er
'inverse display, sign(ye1-ys1)=-sign(ye2-ys2)
YDISPLAY = (y1 - ys1) * ((ye2 - ys2) / (ye1 - ys1)) + ys2
er:
End Function
Sub BULD_DISPLAY(xmin As Double, xmax As Double, ymin As Double, ymax As Double)
'On Error Resume Next
Const dxsize = 800
Const dysize = 800
Const digits = 2
Dim xaxis As Long
Dim yaxis As Long
Dim col As Long
Dim k As Integer
display.Cls
col = display.ForeColor

'get y-axis position
yaxis = XDISPLAY(0, xmin, xmax, 0, display.Width)
'get x-axis position
xaxis = YDISPLAY(0, ymax, ymin, 0, display.Height)

'BUILD X-AXES!
If xaxis >= 0 And xaxis <= display.Height Then
'up
    For k = 0 To Int(xaxis / dxsize)
        display.Line (0, xaxis - k * dxsize)-(display.Width, xaxis - k * dxsize)
    Next
'down
    For k = 1 To Int((display.Height - xaxis) / dxsize)
        display.Line (0, xaxis + k * dxsize)-(display.Width, xaxis + k * dxsize)
    Next
Else
    For k = 0 To Int(display.Height / dxsize)
        display.Line (0, k * dxsize)-(display.Width, k * dxsize)
    Next
End If

'BUILD Y-AXES!
If yaxis >= 0 And yaxis <= display.Width Then
    'left
    For k = 0 To Int(yaxis / dysize)
        display.Line (yaxis - k * dysize, 0)-(yaxis - k * dysize, display.Height)
    Next
    'right
    For k = 1 To Int((display.Width - yaxis) / dysize)
        display.Line (yaxis + k * dysize, 0)-(yaxis + k * dysize, display.Height)
    Next
Else
    For k = 0 To Int(display.Width / dysize)
        display.Line (k * dysize, 0)-(k * dysize, display.Height)
    Next
End If

'BUILD MAIN X AND Y AXES!
display.DrawWidth = 2
display.Line (XDISPLAY(xmin, xmin, xmax, 0, display.Width), YDISPLAY(0, ymax, ymin, 0, display.Height))-(XDISPLAY(xmax, xmin, xmax, 0, display.Width), YDISPLAY(0, ymax, ymin, 0, display.Height)), vbBlack
display.Line (XDISPLAY(0, xmin, xmax, 0, display.Width), YDISPLAY(ymax, ymax, ymin, 0, display.Height))-(XDISPLAY(0, xmin, xmax, 0, display.Width), YDISPLAY(ymin, ymax, ymin, 0, display.Height)), vbBlack
display.DrawWidth = 1

'PRINT X AND Y AXES VALUES
If xaxis >= 0 And xaxis <= display.Height Then
'up
    For k = 0 To Int(xaxis / dxsize)
        display.PSet (0, xaxis - k * dxsize)
        display.ForeColor = vbBlack
        If k <> Int(xaxis / dxsize) Then
            display.Print Round(YDISPLAY(xaxis - k * dxsize, 0, display.Height, ymax, ymin), digits)
        End If
        display.ForeColor = col
    Next
'down
    For k = 1 To Int((display.Height - xaxis) / dxsize)
        display.PSet (0, xaxis + k * dxsize)
        display.ForeColor = vbBlack
        display.Print Round(YDISPLAY(xaxis + k * dxsize, 0, display.Height, ymax, ymin), digits)
        display.ForeColor = col
    Next
Else
    For k = 0 To Int(display.Height / dxsize)
        display.PSet (0, k * dxsize)
        display.ForeColor = vbBlack
        If k > 0 Then
            display.Print Round(YDISPLAY(k * dxsize, 0, display.Height, ymax, ymin), digits)
        End If
        display.ForeColor = col
    Next
End If

If yaxis >= 0 And yaxis <= display.Width Then
'left
    For k = 0 To Int(yaxis / dysize)
        display.PSet (yaxis - k * dysize, 0)
        display.ForeColor = vbBlack
        If k <> Int(yaxis / dysize) Then
            display.Print Round(XDISPLAY(yaxis - k * dysize, 0, display.Width, xmin, xmax), digits)
        End If
        display.ForeColor = col
Next
'right
    For k = 1 To Int((display.Width - yaxis) / dysize)
        display.PSet (yaxis + k * dysize, 0)
        display.ForeColor = vbBlack
        display.Print Round(XDISPLAY(yaxis + k * dysize, 0, display.Width, xmin, xmax), digits)
        display.ForeColor = col
    Next
Else
    For k = 0 To Int(display.Width / dysize)
        display.PSet (k * dysize, 0)
        display.ForeColor = vbBlack
        If k > 0 Then
            display.Print Round(XDISPLAY(k * dysize, 0, display.Width, xmin, xmax), digits)
        End If
        display.ForeColor = col
    Next
End If


End Sub

Private Sub txtfunc_Change()
On Error GoTo er
Dim tmp As Integer
txtfunc.BackColor = vbWhite
tmpScript.Reset
tmpScript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
tmp = tmpScript.Eval("f(" + Str(Rnd() * 10) + ")")
er:
'z = Err.Description + Str(Err.Number)
If Err.Number <> 6 And Err.Number <> 0 Then
txtfunc.BackColor = vbRed
End If
End Sub

Private Sub txtfunc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtfunc.BackColor = vbWhite Then
vbscript.Reset
vbscript.AddCode "Function f(x)" & vbCrLf & "f=" & txtfunc.Text & vbCrLf & "End Function"
txtfunc.AddItem (txtfunc)
plot xmn, xmx, ymn, ymx, steps
End If
End Sub
Sub plot(xmin As Double, xmax As Double, ymin As Double, ymax As Double, steps As Long)
'On Error Resume Next
Const factor = 0.1 'for room between ymax and realmax
Dim X As Double
Dim xstep As Double
Dim Y As Double
Dim ymin1 As Double
Dim ymax1 As Double
Dim prevy As Double
Dim prevx As Double
Dim table() As Double
Dim col As Long
Dim k As Long
Dim kmax As Long
'step1:function analisis
roots.Clear
xstep = (xmax - xmin) / steps
X = xmin - xstep
Y = vbscript.Eval("f(" + Str(X) + ")") 'for prevy

ymin1 = vbscript.Eval("f(" + Str(X) + ")")
ymax1 = vbscript.Eval("f(" + Str(X) + ")")
yminval = ymin1
ymaxval = ymax1
ReDim table(1 To steps + 2, 2) As Double
k = 0
Do
DoEvents

prevx = X
X = X + xstep
k = k + 1
table(k, 1) = XDISPLAY(X, xmin, xmax, 0, display.Width)
If X > xmax Then
X = xmax
End If
If X < xmin Then
X = xmin
End If
prevy = Y

Y = vbscript.Eval("f(" + Str(X) + ")")

If prevy * Y <= 0 Then
roots.AddItem ("INTO [ " + Str(Round(prevx, 4)) + "  , " + Str(Round(X, 4)) + " ]")
End If
If Y < ymin1 Then

ymin1 = Y
yminval = Y
End If
If Y > ymax1 Then
ymax1 = Y
ymaxval = Y
End If
Loop Until X >= xmax
yminval = Round(ymin1, 5)
ymaxval = Round(ymax1, 5)
If autoy.Value = vbChecked Then
ymin = ymin1 - factor * (ymax1 - ymin1)
ymax = ymax1 + factor * (ymax1 - ymin1)
If ymin = ymax Then
'ymin = ymin - factor * Abs(ymin)
'ymax = ymax + factor * Abs(ymax)
End If
End If

kmax = k

'step2:find y's!!
X = xmin
For k = 1 To kmax
DoEvents
X = X + xstep
Y = vbscript.Eval("f(" + Str(X) + ")")
table(k, 2) = YDISPLAY(Y, ymax, ymin, 0, display.Height)
Next

If xmin = xmax Then
xmin = xmax - 0.01
xmax = xmax + 0.01
End If
If ymin = ymax Then
ymin = ymax - 0.01
ymax = ymax + 0.01
End If
BULD_DISPLAY xmin, xmax, ymin, ymax
yymax = ymax
yymin = ymin
col = display.ForeColor
display.ForeColor = vbBlue
'step3:plot it!
For k = 2 To kmax
DoEvents
PB = (k * 100 / kmax)
If table(k, 2) < display.Height And table(k - 1, 2) < display.Height Then
display.Line (table(k, 1), table(k, 2))-(table(k - 1, 1), table(k - 1, 2))
End If
Next
display.ForeColor = col


End Sub
