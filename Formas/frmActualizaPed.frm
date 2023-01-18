VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmActualizaPed 
   Caption         =   "Actualizacion de Pedidos"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread sprDatos 
      Height          =   3375
      Left            =   360
      OleObjectBlob   =   "frmActualizaPed.frx":0000
      TabIndex        =   8
      Top             =   2760
      Width           =   8415
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   7680
      Picture         =   "frmActualizaPed.frx":119F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actualizar pedidos seleccionados con la nueva fecha de dispersion"
      Height          =   1215
      Left            =   480
      TabIndex        =   9
      Top             =   6240
      Width           =   6015
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Grabar"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   4560
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmActualizaPed.frx":12A1
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   2640
         TabIndex        =   10
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Fecha Dispersion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   11
         Top             =   540
         Width           =   2010
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mostrando informacion del producto y fecha de dispersion seleccionada"
      Height          =   2235
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin VB.OptionButton opt3 
         Caption         =   "Por Liberar"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Liberados"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Todos"
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   4680
         Picture         =   "frmActualizaPed.frx":13A3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cboBodegas 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmActualizaPed.frx":14A5
         Left            =   1800
         List            =   "frmActualizaPed.frx":14AC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmActualizaPed.frx":14CE
         Left            =   1800
         List            =   "frmActualizaPed.frx":14D8
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1800
         TabIndex        =   5
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mostrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Dispersion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   1380
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   4
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   435
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmActualizaPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim inicia As Boolean
Dim prod As Byte, SI As Boolean
Dim sEmail As String
Dim fecha_ant As Date, nomandes As Boolean, gran_mail As String
Dim elcliente As Integer
Dim Sucursal As Integer, pedido As Long

Private Sub cboProducto_Click()
Dim aqui As Byte, yano As Boolean
  yano = False
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
      'CargaDatos
  End If
End Sub

Private Sub cmdBuscarC_Click()
On Error GoTo ERR:
    CargaDatos
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdGrabar_Click()
Dim v As Integer, conta As Integer
On Error GoTo ERR:
With sprDatos
     If Format(mskFechaFin.Text, "MM/DD/YYYY") < Format(Date, "MM/DD/YYYY") Then
        MsgBox "Error en la fecha de dispersion del pedido, No puede ser anterior al dia de hoy", vbCritical, "Fecha inválida"
        Exit Sub
     End If
     .Col = 7
     conta = 0
     For v = 1 To .MaxRows
         .Row = v
         If .value = 1 Then
            conta = conta + 1
         End If
     Next
     If conta > 0 Then
        If MsgBox("¿Esta completamente seguro de cambiarle la fecha de dispersion a estos pedidos seleccionados?", vbQuestion + vbDefaultButton2 + vbYesNo, "Modificando fechas de dispersion") = vbYes Then
           For v = 1 To .MaxRows
               .Row = v
               .Col = 7
               If .value = True Then
                  .Col = 1
                  Sucursal = .Text
                  .Col = 2
                  elcliente = Val(.Text)
                  .Col = 4
                  pedido = .Text
                  If CboProducto.Text = "PAGO-UNIFORME BE" Then
                     Product = 8
                  End If
                  '--Evalua que no este dispesado antes
                  sqls = "sp_PedidosVarios @Sucursal=" & Val(Sucursal)
                  sqls = sqls & " ,@Pedido=" & pedido
                  sqls = sqls & " ,@Accion='PrevActualizaPedido'"
                  Set rsBD = New ADODB.Recordset
                  rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly, adCmdText
                  If Not rsBD.EOF Then
                     If rsBD!bon_ped_status = 12 Then
                        'MsgBox "El pedido ya fue dispersado el dia " & Format(rsBD!bon_ped_fechaent, "dd/mmm/yyyy"), vbExclamation, "Imposible moverlo de fecha"
                     Else
'                         If Trim(elserver) = "SBMONTERREY" Then
                            sqls = "sp_PedidosVarios Null,Null,Null," & Product & ",'ActualizaPedido','" & Format(mskFechaFin.Text, "MM/DD/YYYY") & "'" & _
                            " ," & Val(Sucursal) & "," & Val(pedido)
                            cnxbdMty.Execute sqls, intRegistros
                            Call Manda_aviso_de_cambio
'                         Else
'                            sqls = "sp_PedidosVarios Null,Null,Null," & Product & ",'ActualizaPedido','" & Format(mskFechaFin.Text, "MM/DD/YYYY") & "'" & _
'                            " ," & Val(sucursal) & "," & Val(pedido)
'                            cnxBD.Execute sqls, intRegistros
'
'                            sqls = "sp_PedidosVarios Null,Null,Null," & Product & ",'ActualizaPedido','" & Format(mskFechaFin.Text, "MM/DD/YYYY") & "'" & _
'                            " ," & Val(sucursal) & "," & Val(pedido)
'                            cnxbdMty.Execute sqls, intRegistros
'                            Call Manda_aviso_de_cambio
'                         End If
                     End If
                  End If
               End If
           Next
           MsgBox "Se han actualizado de fecha los pedidos selecccionados y enviado notificacion a los clientes", vbInformation, "Pedidos actualizados"
           CargaDatos
        End If
     Else
        MsgBox "No ha seleccionado ningun pedido para modificar fecha de dispersion", vbExclamation, "Pedidos no seleccionados"
     End If
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Form_Load()
    Set mclsAniform = New clsAnimated
    Call CboPosiciona(cboBodegas, 0)
    mskFechaIni.Text = Date + 1
    mskFechaFin.Text = Date + 1
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart") 'aqui lo omiti
    If user_master = False Then
       CargaBodegasS2 cboBodegas 'CargaBodegasS2
    Else
       CargaBodegas cboBodegas
    End If
    opt1.value = True
    'CargaDatos
End Sub

Sub CargaDatos()
Dim Status As Integer

Status = 0
If opt1.value = True Then
    Status = 0
ElseIf opt2.value = True Then
    Status = 10
ElseIf opt3.value = True Then
    Status = 1
End If
 
If cboBodegas.ItemData(cboBodegas.ListIndex) = 0 And user_master = False Then
   cboBodegas.ItemData(0) = gnBodega
End If
sqls = "exec spr_FactAutorizaDisp " & cboBodegas.ItemData(cboBodegas.ListIndex) & ", '" & Format(mskFechaIni.Text, "MM/DD/YYYY") & "', '" & Format(mskFechaIni.Text, "MM/DD/YYYY") & "'," & Status & "," & Product & ",'F'"

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
With sprDatos
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
    .Text = Val(rsBD!Bodega)
    .Col = 2
    .Text = Val(rsBD!cliente)
    .Col = 3
    .Text = Trim(rsBD!Nombre)
    .Col = 4
    .Text = Val(rsBD!pedido)
    .Col = 5
    .Text = CDbl(rsBD!valor)
    .Col = 6
    .Text = Format(rsBD!fechadisp, "dd/mmm/yyyy")
    rsBD.MoveNext
Loop
End With
End Sub

Sub Manda_aviso_de_cambio()
If gblnSendMail Then
      Call Checa_tumail
      If nomandes = True Then
         Set posendmail = New clsSendMail
         If posendmail.IsValidEmailAddress(gran_mail) Then
            
            sAsunto = "Vale Total Modificacion de aplicacion de saldos"
            
            sTexto = "Estimado cliente le informamos que la fecha de aplicacion de los saldos de su Pedido " & pedido & " ha sido modificada." & vbCrLf & vbCrLf
            sTexto = sTexto & "Fecha de Aplicacion: " & Format(mskFechaIni, "DD/MM/YYYY") & vbCrLf
            sTexto = sTexto & "Nueva Fecha de Aplicacion: " & Format(mskFechaFin, "DD/MM/YYYY") & vbCrLf
            sTexto = sTexto & "Gracias por su preferencia" & vbCrLf
            sTexto = sTexto & "Vale Total" & vbCrLf & vbCrLf
            sTexto = sTexto & "Para cualquier duda o aclaración favor de llamar al 8000 0000 Ext. 5 para el resto de la República" & vbCrLf & vbCrLf
            sTexto = sTexto & "Los saldos se verán reflejados en las tarjetas en el horario de 8:A.M A 2:00 P.M, horario del centro de la republica" & vbCrLf & vbCrLf
            sTexto = sTexto & "Este es un correo electrónico de confirmación, favor de no responder." & vbCrLf

            With posendmail
               .SMTPHost = gstrMailSMTPHost
               .SMTPPort = gstrMailSMTPPort
               .From = gstrMailFrom
               .Recipient = Trim(gran_mail)
               .Subject = sAsunto
               .Message = sTexto
               
               .UseAuthentication = True
               .Username = gstrMailUser
               .Password = gstrMailPassword
               .PersistentSettings = False
               
               .Send
            End With
           
         End If
         Set posendmail = Nothing
         'MsgBox "Se envió notificación del cambio al cliente.", vbInformation, "Envio de notificación."
      Else
        'MsgBox "No se enviará notificación del cambio al cliente porque no tiene Email capturado", vbExclamation, "Mail no se enviará"
        Exit Sub
      End If
   End If
End Sub

Sub Checa_tumail()
  gran_mail = ""
  sqls = "SELECT TOP 1 Isnull(MailFETo,'')MailFETo FROM Clientesconfig Where cliente=" & Val(elcliente)
  sqls = sqls & " and MailFE=1 ORDER BY Bodega"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
  If rsBD.EOF Then
     nomandes = False
  Else
     nomandes = True
     gran_mail = rsBD!MailFETo
  End If
  Set rsBD = Nothing
  Exit Sub
End Sub

