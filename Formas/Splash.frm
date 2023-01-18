VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vale Total, S.A. de C.V."
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image3 
      Height          =   1740
      Left            =   240
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1980
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Versión "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Operación Vale Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lblEspere 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Conectando a Servidor, espere un momento..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   3975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    Call ChecaRegWin
End Sub

