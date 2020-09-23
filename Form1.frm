VERSION 5.00
Begin VB.Form LSystems 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "LSystems"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   Begin VB.Frame Run 
      Caption         =   "Run"
      Height          =   975
      Left            =   2400
      TabIndex        =   41
      Top             =   4440
      Width           =   3855
      Begin VB.CommandButton Generate 
         Caption         =   "Generate -->"
         Height          =   375
         Left            =   2040
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Clear 
         Caption         =   "Clear all"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame FKleiner 
      Caption         =   "10, 5, 2.5,..."
      Height          =   735
      Left            =   2400
      TabIndex        =   35
      Top             =   3600
      Width           =   3855
      Begin VB.TextBox Kleiner 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Text            =   ".75"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox kleinertmp 
         Alignment       =   2  'Zentriert
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   36
         Text            =   "not used"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Options 
      Caption         =   "Options"
      Height          =   855
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   2055
      Begin VB.CheckBox notused3 
         Caption         =   "not used"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1045
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox notused2 
         Caption         =   "not used"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1045
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox notused1 
         Caption         =   "not used"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox notext 
         Caption         =   "No TXT"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame StartPoint 
      Caption         =   "Start Point"
      Height          =   1695
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   2055
      Begin VB.TextBox xbox 
         Height          =   285
         Left            =   480
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox ybox 
         Height          =   285
         Left            =   480
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Refresh 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   480
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton DefXY 
         Caption         =   "Default"
         Height          =   315
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X: "
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   195
      End
   End
   Begin VB.Frame SL 
      Caption         =   "Save/Load"
      Height          =   735
      Left            =   2400
      TabIndex        =   23
      Top             =   2760
      Width           =   3855
      Begin VB.CommandButton Speicher 
         Caption         =   "["
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Lade 
         Caption         =   "]"
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Turn 
      Caption         =   "+/-"
      Height          =   1335
      Left            =   2400
      TabIndex        =   17
      Top             =   1320
      Width           =   3855
      Begin VB.CommandButton Plus 
         Caption         =   "+"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Minus 
         Caption         =   "-"
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox WinkelP 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Text            =   "45"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox WinkelN 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Text            =   "-30"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Default"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FFrame 
      Caption         =   "F"
      Height          =   1215
      Left            =   2400
      TabIndex        =   13
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton F 
         Caption         =   "F"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox init 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Text            =   "-100"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Default1 
         Caption         =   "Default"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Steps 
      Caption         =   "Steps"
      Height          =   5655
      Left            =   6360
      TabIndex        =   10
      Top             =   0
      Width           =   1575
      Begin VB.ListBox Liste2 
         Height          =   5325
         ItemData        =   "Form1.frx":0442
         Left            =   720
         List            =   "Form1.frx":0449
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox Liste1 
         Height          =   5325
         ItemData        =   "Form1.frx":0451
         Left            =   120
         List            =   "Form1.frx":0458
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton tree5 
      Caption         =   "Tree III"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton tree4 
      Caption         =   "Koch-Kurve"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton tree3 
      Caption         =   "Tree II"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton tree2 
      Caption         =   "Tree I"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Examples 
      Caption         =   "Examples"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton tree1 
         Caption         =   "Krautartige Pflanze I"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox Axiom2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6240
      Width           =   7815
   End
   Begin VB.CommandButton Hinzu 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Cnt 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Text            =   "6"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Axiom 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "F[+F][-F]"
      Top             =   5760
      Width           =   6375
   End
End
Attribute VB_Name = "LSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pi = 3.14159265358979
Dim x As Integer
Dim y As Integer
Dim a As Double
Dim xnew As Double
Dim ynew As Double
Dim SpeicherX(1 To 255)
Dim SpeicherY(1 To 255)
Dim SpeicherA(1 To 255)
Dim SpeicherXnew(1 To 255)
Dim SpeicherYnew(1 To 255)
Dim n As Integer
Dim memory As String

Private Sub Clear_Click()
Reset
End Sub
Private Function Reset()
Ausgabe.Baum.Cls
Liste1.Clear
Liste2.Clear
n = 0
x = Ausgabe.Baum.Width / 2
y = Ausgabe.Baum.Height
xnew = 0
a = 0
Axiom2 = ""
End Function

Private Sub Default1_Click()
init = -100
End Sub

Private Sub DefXY_Click()
x = Ausgabe.Baum.Width / 2
y = Ausgabe.Baum.Height
End Sub

Private Sub F_Click()
If notext.Value = False Then
    Liste1.AddItem "F"
    Liste2.AddItem Sqr(xnew ^ 2 + ynew ^ 2)
End If
Ausgabe.Baum.Line (x, y)-Step(xnew, ynew) ', RGB(0, 255 - Sqr(xnew ^ 2 + ynew ^ 2) * 2, 0)
x = x + xnew: y = y + ynew
xnew = xnew * Kleiner
ynew = ynew * Kleiner
End Sub

Private Sub Form_Load()
Ausgabe.Show
Reset
Refresh_Click
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Ausgabe
End Sub

Private Sub Generate_Click()
DefXY_Click
n = 0
x = Ausgabe.Baum.Width / 2
y = Ausgabe.Baum.Height
xnew = 0
a = 0
Axiom2 = ""
x = xbox
y = ybox
ynew = init
Ausgabe.Baum.Cls
Dim i As Integer
For i = 0 To Liste1.ListCount
    Select Case Liste1.List(i)
        Case "F"
            F_Click
        Case "+"
            Plus_Click
        Case "-"
            Minus_Click
        Case "["
            Speicher_Click
        Case "]"
            Lade_Click
    End Select
Next i
End Sub

Private Sub Hinzu_Click()
On Error GoTo fehler
Clear_Click
x = xbox
y = ybox
ynew = init
Dim tmp As String
Dim result As String
result = "F"
tmp = Trim$(Axiom.Text)

Dim i As Integer
For i = 1 To Int(Cnt)
    result = Replace(result, "F", tmp)
Next i
If notext.Value = False Then
    Axiom2 = Trim$(result)
Else
    memory = Trim$(result)
End If

If notext.Value = False Then
For i = 1 To Len(Axiom2)
    Select Case Mid$(Axiom2, i, 1)
        Case "F"
            F_Click
        Case "["
            Speicher_Click
        Case "]"
            Lade_Click
        Case "+"
            Plus_Click
        Case "-"
            Minus_Click
    End Select
Next i
Else
For i = 1 To Len(memory)
    Select Case Mid$(memory, i, 1)
        Case "F"
            F_Click
        Case "["
            Speicher_Click
        Case "]"
            Lade_Click
        Case "+"
            Plus_Click
        Case "-"
            Minus_Click
    End Select
Next i
End If

Exit Sub
fehler:
MsgBox "überlauf. kleinerer wert eingeben"
Clear_Click
End Sub

Private Sub Lade_Click()
If notext.Value = False Then
    Liste1.AddItem "]"
    Liste2.AddItem n
End If
x = SpeicherX(n)
y = SpeicherY(n)
a = SpeicherA(n)
xnew = SpeicherXnew(n)
ynew = SpeicherYnew(n)
n = n - 1
End Sub

Private Sub Liste1_Click()
Liste2.ListIndex = Liste1.ListIndex
End Sub

Private Sub Liste2_Click()
Liste1.ListIndex = Liste2.ListIndex
End Sub
Private Sub Minus_Click()
If notext.Value = False Then
    Liste1.AddItem "-"
    Liste2.AddItem WinkelN
End If
Dim tmp As Double
tmp = xnew
xnew = Cos(2 * WinkelN / 360 * pi) * tmp - Sin(2 * WinkelN / 360 * pi) * ynew
ynew = Sin(2 * WinkelN / 360 * pi) * tmp + Cos(2 * WinkelN / 360 * pi) * ynew
End Sub

Private Sub notext_Click()
If notext.Value <> 1 Then
    Generate.Enabled = False
Else
    Generate.Enabled = True
End If
End Sub

Private Sub Plus_Click()
If notext.Value = False Then
    Liste1.AddItem "+"
    Liste2.AddItem WinkelP
End If
Dim tmp As Double
tmp = xnew
xnew = Cos(2 * WinkelP / 360 * pi) * tmp - Sin(2 * WinkelP / 360 * pi) * ynew
ynew = Sin(2 * WinkelP / 360 * pi) * tmp + Cos(2 * WinkelP / 360 * pi) * ynew
End Sub

Private Sub Refresh_Click()
xbox = x
ybox = y
End Sub

Private Sub Speicher_Click()
n = n + 1
SpeicherX(n) = x
SpeicherY(n) = y
SpeicherA(n) = a
SpeicherXnew(n) = xnew
SpeicherYnew(n) = ynew
If notext.Value = False Then
    Liste1.AddItem "["
    Liste2.AddItem n
End If
End Sub

Private Sub tree1_Click()
Axiom = "F[+F]F[-F]F"
WinkelP = -25.7
WinkelN = 25.7
Cnt = 4
Kleiner = 1
init = -5
End Sub
Private Sub tree2_Click()
Axiom = "FF+[+F-F-F]-[-F+F+F]"
Cnt = 4
Kleiner = 1
init = -10
WinkelP = -25
WinkelN = 25
notext.Value = 1
Generate.Enabled = False
End Sub
Private Sub tree3_Click()
Cnt = 6
WinkelP = 45
WinkelN = -30
Axiom = "F[+F][-F]"
Kleiner = 0.618 + 0.1
init = -100
End Sub
Private Sub tree4_Click()
Cnt = 4
Axiom = "F+F--F+F"
Kleiner = 1
init = -1
WinkelP = -45
WinkelN = 45
End Sub
Private Sub tree5_Click()
Cnt = 6
WinkelP = 30
WinkelN = -45
Axiom = "F[+FF][-F]"
Kleiner = 0.618 + 0.1
init = -100
End Sub
