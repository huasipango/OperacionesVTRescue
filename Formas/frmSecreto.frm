VERSION 5.00
Begin VB.Form frmSecreto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave para guardar ajuste"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Left            =   773
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frmSecreto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tecleado As String, palabra_secreta As String
Private Sub Form_Load()
  Dim x As Integer, s As String, y As Integer
  palabra_ok = False
  palabra_secreta = ""
  Text1.Text = ""
  For y = 1 To 8
      s = ""
      Randomize
      If y Mod 2 = 0 Then
         s = Chr(Int(9 * Rnd + 48))
      Else
         s = Chr(Int(26 * Rnd + 65))
      End If
      palabra_secreta = palabra_secreta & s
  Next
  Label1.Caption = UCase(palabra_secreta)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Len(Text1.Text) = 8 Then
     tecleado = UCase(Text1.Text)
     If UCase(tecleado) = UCase(palabra_secreta) Then
        palabra_ok = True
        Unload Me
     Else
        palabra_ok = False
        MsgBox "Clave Incorrecta", vbCritical, "Error"
        Exit Sub
     End If
  End If
End Sub
