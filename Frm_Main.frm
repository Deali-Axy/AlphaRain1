VERSION 5.00
Begin VB.Form Frm_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14835
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFC0&
   Icon            =   "Frm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   12  'No Drop
   ScaleHeight     =   9645
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   960
      Top             =   3000
   End
   Begin VB.Label LBL_Version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ver:1.0"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   17.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   450
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label LBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   17.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   450
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   9120
      Width           =   240
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: AlphaRain1
'Author: Deali-Axy
'Email: deali@live.com
'MicroBlog: http://weibo.com/dealiaxy

Option Explicit
Dim i As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27
            End
        Case Else
            Dim R, G, B As Integer
            Randomize
            i = i + 1
            LBL(0).Caption = Str(i)
            LBL(0).Left = 0
            Load LBL(i)

            R = Int(255 * Rnd) + 1
            G = Int(Rnd * 255) + 1
            B = Int(Rnd * 255) + 1

            If i Mod 10 = 0 Then
                Me.Cls
            End If

            Me.ForeColor = RGB(R, G, B)

            Print "[Form_KeyPress]Get KeyBoard Char+" & Chr(KeyAscii)
            Print "[Form_KeyPress]Load LBL+" & i
            Print "[Form_KeyPress]Load Tmr+" & i
            Print "[Form_KeyPress]Get Random Color+" & RGB(R, G, B)

            With LBL(i)
                .Left = Int(Screen.Width * Rnd) + 1
                '.Top = Int(Screen.Height * Rnd) + 1
                .Top = 0
                .AutoSize = True
                .BackStyle = 0
                .ForeColor = RGB(R, G, B)
                .Caption = Chr(KeyAscii)
                .Font.Size = Int(Rnd * 180) + 1
                .Visible = True
            End With
            Load Tmr(i)
            Tmr(i).Enabled = True
    End Select
End Sub

Private Sub Form_Load()
    With LBL(0)
        .Left = 0
        .Top = Screen.Height - .Height
    End With

    With LBL_Version
        .AutoSize = True
        .Caption = "Ver: " & App.Major & "." & App.Minor & "." & App.Revision
        .Left = Screen.Width - .Width - 100
        .Top = 100
    End With
End Sub

Private Sub Tmr_Timer(Index As Integer)
    Dim R, G, B As Integer
    Randomize

    R = Int(255 * Rnd) + 1
    G = Int(Rnd * 255) + 1
    B = Int(Rnd * 255) + 1

    Me.ForeColor = RGB(R, G, B)

    If LBL(Index).Top > Screen.Height Then
        Tmr(Index).Enabled = False
        LBL(Index).Visible = False
        Unload LBL(Index)
        Print "[Tmr_Timer]Unload LBL+" & Index
        Print "[Tmr_Timer]Unload Tmr+" & Index
        Unload Tmr(Index)
        Exit Sub
    Else
        LBL(Index).Top = LBL(Index).Top + 100
    End If
End Sub
