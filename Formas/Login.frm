VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.ComboBox cboBodegas 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2325
   End
   Begin VB.Timer tmrKey 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Cancel          =   -1  'True
      Height          =   600
      Left            =   2400
      Picture         =   "Login.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "Cancelar"
      ToolTipText     =   "Salir del Sistema"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000016&
      Default         =   -1  'True
      DownPicture     =   "Login.frx":030A
      Height          =   600
      Left            =   1080
      Picture         =   "Login.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Aceptar"
      ToolTipText     =   "Entrar al Sistema"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Contraseña"
      Top             =   480
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      ToolTipText     =   "Nombre de Usuario"
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Sucursal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Tag             =   "&Contraseña:"
      Top             =   890
      Width           =   1080
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   15
      Left            =   150
      Picture         =   "Login.frx":0B8E
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   14
      Left            =   150
      Picture         =   "Login.frx":10D7
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   13
      Left            =   150
      Picture         =   "Login.frx":15ED
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   12
      Left            =   150
      Picture         =   "Login.frx":1AA4
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   11
      Left            =   150
      Picture         =   "Login.frx":1EE8
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   10
      Left            =   150
      Picture         =   "Login.frx":213E
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   9
      Left            =   150
      Picture         =   "Login.frx":23DF
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   8
      Left            =   150
      Picture         =   "Login.frx":2874
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   7
      Left            =   150
      Picture         =   "Login.frx":2D41
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   6
      Left            =   150
      Picture         =   "Login.frx":327F
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   5
      Left            =   150
      Picture         =   "Login.frx":373D
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   4
      Left            =   150
      Picture         =   "Login.frx":3A05
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   3
      Left            =   150
      Picture         =   "Login.frx":3C76
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   2
      Left            =   150
      Picture         =   "Login.frx":4113
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgKey 
      Height          =   300
      Index           =   1
      Left            =   150
      Picture         =   "Login.frx":4654
      Top             =   200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Tag             =   "&Contraseña:"
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuari&o:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Tag             =   "Usuari&o:"
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Public OK As Boolean
Dim actual As Integer

Private Sub cboBodegas_Click()

    sqls = "SELECT bodega, SERVIDOR, formatofactura FROM BODEGAS"
    sqls = sqls & " WHERE BODEGA = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
           gstrServidor = gstrServidor
        gnBodega = rsBD!Bodega
        FormatoFactura = IIf(IsNull(rsBD!FormatoFactura), 0, rsBD!FormatoFactura)
    Else
        MsgBox "Problemas al buscar el servidor, contacte a Sistemas"
        End
    End If
        
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    actual = 1
    txtUserName = Usuario
    
    CargaBodegas cboBodegas
    If TipoEntrada = "SC" Then
        cboBodegas.Visible = False
        lblLabels(2).Visible = False
    End If
End Sub
Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim Perfil As Integer
Dim PwdNuevo As String

    Screen.MousePointer = vbHourglass
    ' vsp Cambiar rsBD por rsBD
    On Error GoTo ErrBuscar
    
    If cboBodegas.Text = "" Then
        MsgBox "Por favor seleccione la sucursal...!!", vbInformation, "Sucursal"
        cboBodegas.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
'    Set cnxBD = New ADODB.Connection
'    cnxBD.CommandTimeout = 6000
'    cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
   
'    Set rsBD = New ADODB.Recordset
'    SQLS = "Select a.Usuario, a.Nombre, a.Password, a.Bodega, a.status,  b.UENNumero, a.Administrador, A.MULTIBODEGA, A.MULTIUEN, A.MULTIVENDEDOR"
'    SQLS = SQLS & " FROM Usuarios a, Bodegas b"
'    SQLS = SQLS & " WHERE a.Usuario = '" & txtUserName & "'"
'    SQLS = SQLS & " and b.Bodega = a.Bodega"
'    rsBD.Open SQLS, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    sqls = "Exec Sp_Login '" & txtUserName.Text & "', 'FBE'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    If rsBD.EOF Then
        MsgBox "No existe el usuario.", vbInformation, "Inicio de sesión"
        txtUserName.SetFocus
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
        Exit Sub
    End If
    
    On Error GoTo 0
    If Trim(rsBD!Password) <> Trim(txtPassword) Then
        MsgBox "La contraseña no es válida; vuelva a intentarlo", vbCritical, "Inicio de sesión"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.Text)
    Else
        OK = True
        Usuario = txtUserName
        sqls = "SELECT * FROM Derechos Where Modulo='MST' AND Usuario=" & Val(Usuario)
        Set rsBD3 = New ADODB.Recordset
        rsBD3.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        If rsBD3.EOF Then
           user_master = False
        Else
           user_master = True
        End If
        rsBD3.Close
        Set rsBD3 = Nothing
                
        nomUsuario = rsBD![Nombre]
        gnBodega = rsBD![Bodega]
        
        If user_master = False Then
           If gnBodega <> cboBodegas.ItemData(cboBodegas.ListIndex) Then
              MsgBox "Su usuario no tiene permisos para conectarse a otra sucursal", vbCritical, "Usuario restringido"
              Exit Sub
           End If
        End If
        
        gnBodega = cboBodegas.ItemData(cboBodegas.ListIndex)
        gstrUsuario = txtUserName
        gnUEN = rsBD![UENNumero]
        gnMultiBodega = rsBD![MB]
        gnMultiUEN = rsBD![MultiUEN]
        gnMultiVend = rsBD!multivendedor
        gnAdministrador = rsBD![Administrador]
        mdiMain.sbStatusBar.Panels(5).Text = gstrUsuario
        mdiMain.sbStatusBar.Panels(6).Text = gstrPC 'gstrUsuario
        Perfil = IIf(Not IsNull(rsBD!cveperfil), rsBD!cveperfil, 0)
        If DateDiff("d", rsBD!FechaUltPwd, Now) > 90 Then
            PwdNuevo = InputBox("Es necesario cambiar su contraseña", "Expiracion de contraseña")
            If PwdNuevo = txtPassword Then
                PwdNuevo = InputBox("El password no puede ser igual al anterior", "Expiracion de contraseña")
            End If
            Do While Not Comprobar_Contraseña(PwdNuevo)
                PwdNuevo = InputBox("El password debe ser de 8 a 10 digitos y debe contener al menos una mayuscula una minuscula y un caracter especial.", "Expiracion de contraseña")
            Loop
            sqls = "EXEC sp_Cambia_Pwd " & Usuario & ",'" & PwdNuevo & "'"
            Set rsBD3 = New ADODB.Recordset
            rsBD3.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
            Set rsBD3 = Nothing
            MsgBox "Password vigente por los siguientes 3 meses!"
        End If
        
        If gnAdministrador = "S" Then
          
          Call ActivaMenus
        Else
          Call UserMenu("FBE", Perfil)
        End If
        
        sqls = "select tipobodega from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
                   
        sqls = " sp_Bodegas_Sel " & cboBodegas.ItemData(cboBodegas.ListIndex)
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
        If Not rsBD.EOF Then
               elserver = ""
               gsPathFE = "C:\Facturacion"   'rsBD!PathFE"
            gsTipoBodega = rsBD!TipoBodega
        Else
               gsPathFE = "C:\Facturacion"   'rsBD!PathFE"
            gsTipoBodega = "P"
        End If
            
        Call SaveSetting(gstrSistema, gstrConfigSist, gstrKeyLastUser, Usuario)
        Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
        Me.Hide
    End If
    rsBD.Close
    Set rsBD = Nothing
    Exit Sub
ErrBuscar:
    MsgBox "Error al buscar en el servidor. Favor de avisar a sistemas!" & vbCr & "Error: " & Str(ERR.Number) & ", " & ERR.Description, vbCritical, "Error..."
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub tmrKey_Timer()
    imgKey(actual).Visible = False
    If actual = 15 Then
        actual = 1
        imgKey(1).Visible = True
    Else
        actual = actual + 1
        imgKey(actual).Visible = True
    End If
End Sub

