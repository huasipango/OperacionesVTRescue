VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmActivBanco 
   Caption         =   "Actividades del Banco"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   10200
         Picture         =   "frmActivBanco.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9720
         Picture         =   "frmActivBanco.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtHist 
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Dias Historial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10815
      Begin FPSpread.vaSpread spdArchivos 
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "frmActivBanco.frx":0204
         TabIndex        =   1
         Top             =   240
         Width           =   10455
      End
   End
End
Attribute VB_Name = "frmActivBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated

Private Sub cmdBuscarC_Click()
    If Trim(txtHist.Text) <> "" Then
        CargaDatos
    Else
        txtHist.Text = 1
        CargaDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    txtHist.Text = 1
    CargaDatos
End Sub

Sub CargaDatos()
Dim i As Integer
sqls = "exec sp_LoNuevo_Sel 'IMP', " & Val(txtHist.Text) & ",1"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
With spdArchivos
 .Col = -1
 .Row = -1
 .Action = 12
 .MaxRows = 0
 
 i = 0
 Do While Not rsBD.EOF
    i = i + 1
    .MaxRows = i
    .Row = i
    .Col = 1
    .Text = Format(rsBD!Fecha, "mm/dd/yy")
    .Col = 2 'descrip
    .Text = rsBD!Texto
    .Col = 3 'archivo
    .Text = rsBD!Version
    rsBD.MoveNext
 Loop

End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtHist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtHist.Text) <> "" Then
            CargaDatos
        End If
    End If
End Sub
