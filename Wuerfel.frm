VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Zufällige Würfel"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton copyright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "M"
      Height          =   315
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   14  'Pfeil und Fragezeichen
      TabIndex        =   3
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zwei Würfel werfen"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Würfeln bis eine 6 kommt"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Würfel 1x werfen"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
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
      Left            =   6840
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   3015
      Left            =   4560
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image6 
      Height          =   1695
      Left            =   6840
      Picture         =   "Wuerfel.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image5 
      Height          =   1695
      Left            =   6840
      Picture         =   "Wuerfel.frx":7A08
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1695
      Left            =   6840
      Picture         =   "Wuerfel.frx":EC88
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1695
      Left            =   6840
      Picture         =   "Wuerfel.frx":157B3
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   6840
      Picture         =   "Wuerfel.frx":1BB67
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   6840
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
      Left            =   6720
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
    Select Case wuerfel
        Case 1
            Image1.Visible = True
            LabelWuerfel.Caption = "Schade, nur eine 1!"
        Case 2
            Image2.Visible = True
            LabelWuerfel.Caption = "Schade, wieder so wenig!"
        Case 3
            Image3.Visible = True
            LabelWuerfel.Caption = "3 mal 2 ist 6! Nächstes mal klappts bestimmt!"
        Case 4
            Image4.Visible = True
            LabelWuerfel.Caption = "4 mal 1,5 ist 6!"
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
    Loop While (dieSechs <> 6)
    LabelWuerfel.Caption = "Du brauchst " & i & " mal um eine " & dieSechs & " zu würfeln!"
    
End Sub

Private Sub copyright_Click()
    MsgBox "Copyright by Marc Nitzsche, 2016"
End Sub

Function randomNr(Min As Integer, Max As Integer)
    randomNr = Int(Rnd() * Max) + Min
End Function
Function hideEveryImage()
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
End Function


