VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1875
   LinkTopic       =   "Form1"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   125
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picStrip 
      BackColor       =   &H00808080&
      Height          =   6180
      Left            =   150
      ScaleHeight     =   6120
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.Frame Frame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'Kein
         Height          =   4620
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   4935
         Width           =   1545
         Begin VB.Label lbl 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00808080&
            Caption         =   "Label4"
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   15
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'Kein
         Height          =   4620
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   3915
         Width           =   1545
         Begin VB.Label lbl 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00808080&
            Caption         =   "Label3"
            Height          =   255
            Index           =   3
            Left            =   135
            TabIndex        =   11
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'Kein
         Height          =   4620
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   3180
         Width           =   1545
         Begin VB.Label lbl 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00808080&
            Caption         =   "Label2"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   14
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'Kein
         Height          =   4620
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   2310
         Width           =   1545
         Begin VB.Label lbl 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00808080&
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   13
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'Kein
         Height          =   4620
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   1515
         Width           =   1545
         Begin VB.Label lbl 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00808080&
            Caption         =   "Label0"
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.CommandButton strip 
         Caption         =   "Button 4"
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CommandButton strip 
         Caption         =   "Button 3"
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   900
         Width           =   1515
      End
      Begin VB.CommandButton strip 
         Caption         =   "Button 2"
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   1515
      End
      Begin VB.CommandButton strip 
         Caption         =   "Button 1"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   300
         Width           =   1515
      End
      Begin VB.CommandButton strip 
         Caption         =   "Button 0"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim move_strip  As Integer
Dim x, i        As Integer
Dim LastIndex   As Integer
Dim pos()       As String



Private Sub Form_Load()
Dim x           As Integer

    ReDim pos(strip.Count - 1, 0)
    For x = 0 To strip.Count - 1
        pos(x, 0) = "top"
        Frame(x).Width = picStrip.ScaleWidth
        Frame(x).Height = picStrip.ScaleHeight - (strip.Count * strip(0).Height)
        Frame(x).Left = strip(0).Left
        Frame(x).Top = strip(0).Top + (strip(0).Height * (x + 1))
        Frame(x).Visible = False
    Next
    Frame(strip.Count - 1).Visible = True
    lbl(strip.Count - 1).Caption = strip(strip.Count - 1).Caption
    LastIndex = strip.Count - 1

End Sub



Private Sub strip_Click(Index As Integer)

    If pos(Index, 0) = "top" Then
        Call move_down(Index)
    Else
        Call move_up(Index)
    End If
    
End Sub

Private Sub move_down(Index As Integer)

    If Not LastIndex = Index Then
        move_strip = strip.Count - 1
        x = 1
        Do Until move_strip = Index
            strip(move_strip).Top = picStrip.ScaleHeight - (strip(move_strip).Height * x)
            pos(move_strip, 0) = "bottom"
            x = x + 1
            move_strip = move_strip - 1
        Loop
        
        Frame(Index).Visible = True
        lbl(Index).Caption = strip(Index).Caption
        Frame(LastIndex).Visible = False
        LastIndex = Index
    End If
    
    picStrip.SetFocus
    
End Sub

Private Sub move_up(Index As Integer)

    If Not LastIndex = Index Then
        move_strip = 0
        x = 0
        Do While move_strip <= Index
            strip(move_strip).Top = 0 + (strip(move_strip).Height * x)
            pos(move_strip, 0) = "top"
            x = x + 1
            move_strip = move_strip + 1
        Loop
        
        Frame(Index).Visible = True
        lbl(Index).Caption = strip(Index).Caption
        Frame(LastIndex).Visible = False
        LastIndex = Index
    End If
    
    picStrip.SetFocus
    
End Sub

