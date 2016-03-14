VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Zufällige Würfel"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "L o h n t  e s s i c h ?"
      Height          =   2535
      Left            =   6720
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Input1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Input2 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton copyright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "M"
      Height          =   315
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   14  'Pfeil und Fragezeichen
      TabIndex        =   5
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zwei Würfel werfen"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Würfeln bis eine 6 kommt"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Würfel 1x werfen"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label LblAugensumme 
      Caption         =   "Augensumme?"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label LblWieOft 
      Caption         =   "Wie oft?"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LabelWuerfel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":7A08
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":EC88
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":157B3
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":1BB67
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   4800
      Picture         =   "Wuerfel.frx":217A4
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3015
      Left            =   4680
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    wuerfel = randomNr(1, 6)
    hideEveryImage
    ' Da nicht jedes Bild automatisch auf der obersten Ebene angezeigt wird,
    ' werden sicherheitshalber alle Bilder aus dem Sichtfeld entfernt
    Select Case wuerfel
        Case 1
            Image1.Visible = True ' Image1 = eine 1, Image2 = eine 2, ...
            LabelWuerfel.Caption = "Schade, nur eine 1!"
        Case 2
            Image2.Visible = True
            LabelWuerfel.Caption = "Schade, wieder so wenig!"
        Case 3
            Image3.Visible = True
            LabelWuerfel.Caption = "3 mal 2 ist 6! Nächstes mal klappts bestimmt!"
        Case 4
            Image4.Visible = True
            LabelWuerfel.Caption = "4 plus 2 ist 6!"
        Case 5
            Image5.Visible = True
            LabelWuerfel.Caption = "Das ist doch schonmal ein Anfang!"
        Case 6
            Image6.Visible = True
            LabelWuerfel.Caption = "Wer eine 6 würfelt, darf nochmal!"
    End Select
End Sub

Private Sub Command2_Click()
    hideEveryImage
    Image6.Visible = True
    i = 0
    Do
        dieSechs = randomNr(1, 6)
        i = i + 1
    Loop While (dieSechs <> 6) ' Würfele so lange, bis wir eine 6 haben. i zählt mit, wie oft wir schon gewürfelt haben.
    LabelWuerfel.Caption = "Du brauchst " & i & " mal um eine " & dieSechs & " zu würfeln!"
    
End Sub

Private Sub Command3_Click()
    If Input1.Text = "" Or Input2.Text = "" Then
        MsgBox "Bitte geben Sie jeweils eine Zahl ein!"
    Else
        n = Int(Input1.Text)
        Sum = Int(Input2.Text)
        WieOft = 0
        For i = 1 To n
            a = randomNr(1, 6)
            b = randomNr(1, 6)
            If (a + b) = Sum Then ' Sobald a + b = Summe: Erhöhe WieOft um 1
                WieOft = WieOft + 1
            End If
        Next i
        hideEveryImage
        LabelWuerfel.Caption = "Du bekommst die Augensumme " & WieOft & " mal!"
    End If
    
End Sub

Private Sub Command4_Click()
    n = 1000000 ' Mach das 1 Mio. mal
    For i = i To n
        a = randomNr(1, 6) ' Würfel 1 bis 6
        b = randomNr(1, 6)
        c = randomNr(1, 6)
        d = randomNr(1, 6)
        
        If (a = 6) Or (b = 6) Or (c = 6) Or (d = 6) Then ' Wenn auch nur einer der Würfel eine 6 ist
            didItHappen = didItHappen + 1
        End If
    Next i
    If (didItHappen / n > 0.5) Then ' Wenn wir ca. eine Chance von 1/2 haben
        MsgBox "Es lohnt sich. Chance: " & didItHappen / n * 100 & "%"
    Else
        MsgBox "Es lohnt sich nicht. Chance: " & didItHappen / n * 100 & "%"
    End If
End Sub

Private Sub copyright_Click()
    MsgBox "Copyright by Marc Nitzsche, 2016"
End Sub

Function randomNr(Min As Integer, Max As Integer)
    ' Die Gestaltung als Funktion mit Start- und Endwertist zukunftssicherer als ein einfacher Würfel
    ' da wir diese Funktion dann auch für andere Zufallszahlen (nicht 1 bis 6) benutzen können
    randomNr = Int(Rnd() * Max) + Min
End Function
Function hideEveryImage() ' Alle Bilder verschwinden lassen
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
End Function
Private Sub Input1_Validate(Cancel As Boolean) ' Wie oft?
If (Not IsNumeric(Input1.Text)) And (Not Input1.Text = "") Then
    MsgBox "Bitte geben Sie jeweils eine Zahl ein!"
    Cancel = True
End If
End Sub

Private Sub Input2_Validate(Cancel As Boolean) ' Augensumme
If (Not IsNumeric(Input2.Text)) And (Not Input2.Text = "") Then
    MsgBox "Bitte geben Sie eine Zahl zwischen 1 und 12 ein!"
    Cancel = True
End If
End Sub
