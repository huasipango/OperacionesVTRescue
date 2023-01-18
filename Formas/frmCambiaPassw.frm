VERSION 5.00
Begin VB.Form frmCambiaPass 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cambio de Password"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Picture         =   "frmCambiaPassw.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Aceptar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frmCambiaPassw.frx":0236
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox p3 
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox p2 
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox p1 
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4440
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Estrangelo Edessa"
         Size            =   15.75
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaxLength       =   15
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme Password Nuevo:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   3360
      Width           =   3405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password Nuevo:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2280
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password actual:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   1110
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "frmCambiaPassw.frx":0465
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4080
   End
End
Attribute VB_Name = "frmCambiaPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated

Private Sub Aceptar_Click()
Dim PasswordAnt
  If p1 = "" Or p2 = "" Or p3 = "" Then
     MsgBox "Se le sugiere que capture bien sus datos", vbExclamation, "Informacion incompleta"
     p1.SetFocus
     Exit Sub
  End If
  sql = "SELECT dbo.DesEncriptar(Password) AS Password FROM USUARIOS WHERE USUARIO='" & Text1.Text & "'"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
  If Not rsBD.EOF Then
    PasswordAnt = rsBD!Password
     If Trim(PasswordAnt) = Trim(p1.Text) Then
        If Trim(p2) <> Trim(p3) Then
           MsgBox "Confirme correctamente su nuevo password", vbExclamation, "Confirme su password correctamente"
           p3.SetFocus
           rsBD.Close
           Set rsBD = Nothing
           Exit Sub
        Else
            If Trim(PasswordAnt) = Trim(p2.Text) Then
                MsgBox "El password no puede ser igual al anterior", vbExclamation, "Confirme su password correctamente"
                p2.SetFocus
                rsBD.Close
                Set rsBD = Nothing
                Exit Sub
            End If
            If Not Comprobar_Contraseña(p2.Text) Then
                MsgBox "El password debe ser de 8 a 10 digitos y debe contener al menos una mayuscula una minuscula y un caracter especial.", vbExclamation, "Confirme su password correctamente"
                p2.SetFocus
                rsBD.Close
                Set rsBD = Nothing
                Exit Sub
            End If
           sql = "EXEC sp_Cambia_Pwd " & Usuario & ",'" & p2 & "'"
           cnxBD.Execute sql
           MsgBox "Su password ha sido cambiado correctamente...En la proxima sesion recuerde que entrara con su nuevo password", vbInformation, "Cambio de Password correcto"
           Exit Sub
        End If
     Else
        MsgBox "El Password actual no es valido, verifiquelo...", vbExclamation, "Error en password actual"
        p1.SetFocus
        rsBD.Close
        Set rsBD = Nothing
        Exit Sub
     End If
  Else
     MsgBox "No se encontro al empleado", vbCritical, "No existe empleado"
     Exit Sub
  End If
End Sub

Private Sub Cancelar_Click()
  Unload Me
End Sub


Private Sub Form_Load()
  Set mclsAniform = New clsAnimated
'  Set cnxBD = New ADODB.Connection
'  cnxBD.CommandTimeout = 6000
'  cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
  Text1 = gstrUsuario
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  cnxBD.Close
'  Set cnxBD = Nothing
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub p1_KeyPress(KeyAscii As Integer)
 If p1.Text <> "" And KeyAscii = 13 Then
    p2.SetFocus
 End If
End Sub
