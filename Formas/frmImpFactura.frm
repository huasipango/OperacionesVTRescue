VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmImpFactura 
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   7200
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   6600
      Top             =   960
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   3975
      Begin VB.OptionButton Option2 
         Caption         =   "Notas de Consumo"
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Facturas"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1560
      Top             =   120
   End
   Begin VB.CommandButton Alerta 
      BackColor       =   &H0080C0FF&
      Caption         =   "¡ ATENCION !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   550
      Left            =   120
      TabIndex        =   31
      Top             =   1560
      Width           =   6135
      Begin VB.ComboBox cboProducto 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmImpFactura.frx":0000
         Left            =   1800
         List            =   "frmImpFactura.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   8400
      Width           =   8295
      Begin VB.CommandButton Modificar 
         BackColor       =   &H00C0C0C0&
         Height          =   405
         Left            =   5280
         Picture         =   "frmImpFactura.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Modifica el No.Cliente antes de Facturar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   7680
         Picture         =   "frmImpFactura.frx":045E
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   7080
         Picture         =   "frmImpFactura.frx":0560
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   6480
         Picture         =   "frmImpFactura.frx":0662
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancelar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdReimp 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   4680
         Picture         =   "frmImpFactura.frx":0764
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Reimprimir"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   405
         Left            =   5880
         Picture         =   "frmImpFactura.frx":0866
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Actualizar Pedidos"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame frmReimp 
      Caption         =   "Reimpresión de Facturas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   1455
      TabIndex        =   10
      Top             =   3720
      Width           =   5655
      Begin VB.TextBox txtserie 
         Height          =   285
         Left            =   5400
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdReimp2 
         Caption         =   "Reimprimir"
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtFactura 
         Height          =   315
         Left            =   4080
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboBodegasReimp 
         Height          =   315
         ItemData        =   "frmImpFactura.frx":0E80
         Left            =   1200
         List            =   "frmImpFactura.frx":0E87
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblPedido 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblImporte 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedidos No Facturados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   8295
      Begin FPSpread.vaSpread spdPedidos 
         Height          =   4935
         Left            =   120
         OleObjectBlob   =   "frmImpFactura.frx":0EA9
         TabIndex        =   5
         Top             =   360
         Width           =   8055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   8295
      Begin VB.TextBox txtFechaFac 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkSelTodas 
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmImpFactura.frx":1F1C
         Left            =   1200
         List            =   "frmImpFactura.frx":1F23
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Facturacion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Bodega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      X1              =   1560
      X2              =   8400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   240
      Picture         =   "frmImpFactura.frx":1F45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Facturación Bono Electrónico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmImpFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim sEmail As String
Dim prod As Byte, sal As Byte
Dim cte_interno As String, nop As Boolean, pednum As String, fec_ent As Date
Dim cierrabe As Boolean, continua As Boolean
Dim ahora As Date
Dim ahora2 As Date, dif As Date
Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess&, ByVal bInheritHandle&, ByVal dwProcessId&) _
As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) _
As Long
Private Sub Alerta_Click()
  MsgBox "Existen Notas de Consumo pendientes por autorizar en Pago GAS", vbExclamation, "Notas de consumo pendientes"
End Sub

Private Sub Alerta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Timer1.Enabled = False
  Alerta.Visible = True
End Sub

Private Sub cboBodegas_Click()
  CargaPedidos
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Sub InicializaForma()
    frmReimp.Visible = False
    If user_master = False Then
       CargaBodegasS2 cboBodegas
       CargaBodegasS2 cboBodegasReimp
    Else
       CargaBodegasServ cboBodegas
       CargaBodegas cboBodegasReimp
    End If

    'CargaBodegasServ cboBodegas
    'CargaBodegas cboBodegasReimp
    chkSelTodas.value = 0
    CargaPedidos
    frmReimp.Visible = False
    ImpUnaVez = False
End Sub

Private Sub cmdCancelar_Click()
Dim Pedido As Long
Dim cliente As Integer
Dim valor As Double
 With spdPedidos
      .Col = 1
      .Row = .ActiveRow
      Pedido = Val(.Text)
      
      If MsgBox("Desea cancelar el pedido " & Pedido, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         .Row = .ActiveRow
         .Col = 2
         cliente = Val(.Text)
         .Col = 4
         valor = Val(.Text)
         
         sqls = "sp_FactBEVarios '" & gstrUsuario & "'," & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex)) & _
         "," & cliente & "," & Pedido & "," & Product & ",'Cancela'"
         
         
'         sqls = "update bon_pedido"
'         sqls = sqls & " set bon_ped_status = 2, bon_ped_usuario = '" & gstrUsuario & "'"
'         sqls = sqls & " Where BON_PED_SUCURSAL = " & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex))
'         sqls = sqls & "  and BON_PED_CLIENTE = " & Cliente
'         sqls = sqls & "  and BON_PED_NUMERO = " & pedido
'         sqls = sqls & " AND BON_PED_PRODUCTO =" & Product '6"
         
         
         cnxBD.Execute sqls, intRegistros
         
         MsgBox "Pedido Cancelado!!", vbInformation, "Pedido Cancelado"
         
         sqls = "select identificacion from clientes where cliente = " & Val(cliente)
         
         Set rsBD = New ADODB.Recordset
         rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
         
         '---
         .Action = 5
         CargaPedidos
         '---
         If Not rsBD.EOF Then
            sEmail = Trim(rsBD!identificacion)
         Else
            sEmail = ""
         End If
         If gblnSendMail Then
            If sEmail <> "" Then
            Set posendmail = New clsSendMail
            If posendmail.IsValidEmailAddress(sEmail) Then
              
                sAsunto = "VALE TOTAL: Cancelación de Pedido"
                  
               sTexto = "Vale Total agradece su preferencia" & vbCrLf & vbCrLf
               sTexto = sTexto & "Su pedido de Bono Electrónico ha sido cancelado" & vbCrLf
               sTexto = sTexto & "No. Pedido:  " & Pedido & vbCrLf
               sTexto = sTexto & "Valor del Pedido: " & Format(valor, "########0.00")

               With posendmail
                  .SMTPHost = gstrMailSMTPHost
                  .SMTPPort = gstrMailSMTPPort
                  .From = gstrMailFrom
                  .Recipient = Trim(sEmail)
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
         End If
        End If
       '.Action = 5
      End If
      
   'CargaPedidos
      
 End With
End Sub

Sub busca_fecha_pedido()
sqls = "SELECT bon_ped_fechaent FROM BON_PEDIDO "
sqls = sqls & " WHERE bon_ped_tipocom=3 AND bon_ped_cliente=" & cte_interno
sqls = sqls & " AND bon_ped_numero=" & pednum
sqls = sqls & " AND bon_ped_status=0"
sqls = sqls & " AND bon_ped_producto=7"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   fec_ent = rsBD!bon_ped_fechaent
End If
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ERR:
    cmdGrabar.Enabled = False
    If Option1.value = True Then
       GrabaFactura
    End If
    If Option2.value = True And Product = 2 Then
       If continua = True Then
          Graba_Nota_consumo
          investiga_sihayNC
       Else
          MsgBox "Existe un usuario en el modulo de ajustes...espere un momento e intente de nuevo", vbExclamation, "Excepcion controlada"
          Exit Sub
       End If
    End If
    ImpUnaVez = False
    cmdGrabar.Enabled = True
Exit Sub
ERR:
  MsgBox "Errores presentados-> " & ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub Graba_Nota_consumo()
Dim Folio As Integer, valor As Double
Dim Fecha As String
On Error GoTo err_gral
With spdPedidos
 
   For i = 1 To .MaxRows
      .Row = i
      .Col = 5
      If .value = 1 Then
         .Col = 1
         Pedido = .Text
         pednum = Pedido
         .Col = 2
         cliente = Val(.Text)
         .Col = 4
         valor = Val(.Text)
         Folio = BuscaFolio
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         Call busca_fecha_pedido
         sqls = " exec sp_AjustesBe @Folio = " & Val(Folio) & _
                           " , @Cliente = " & Val(cliente) & _
                           " , @concepto = 0 " & _
                           " , @Cargo = 0 " & _
                           " , @Abono = " & Val(valor) & _
                           " , @Usuario   = '" & gstrUsuario & "'" & _
                           " , @Fecha   =  '" & Format(fec_ent, "MM/DD/YYYY") & "'" & _
                           " , @Producto=" & Product
         cnxBD.Execute sqls, intRegistros
                 
         Fecha = Format(fec_ent, "MM/DD/YYYY")   'Format(fec_ent, "MM/DD/YYYY")
         sqls = " exec sp_AjustesDetBE_Consumo @Folio = " & Val(Folio) & _
                           " , @Cliente = " & cliente & _
                           " , @FechaDisp  = '" & Fecha & "'" & _
                           " , @Pedido=" & Pedido & _
                           " , @Producto=" & Product
         cnxBD.Execute sqls, intRegistros
                 
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         sqls = "exec spb_GrabaNotaConsumoBE @Sucursal =" & gnBodega & _
           ",@Producto=" & Product & _
           ",@TipoPedido = 11" & _
           ",@Pasan = 0" & _
           ",@Motivo = ''" & _
           ",@TipoFactura = 3" & _
           ",@Cliente = " & Val(cliente) & _
           ",@nValorPedido = " & Val(valor) & _
           ",@Usuario   = '" & gstrUsuario & "'" & _
           ",@PedidoN   = " & Pedido
         Set rsBD = New ADODB.Recordset
         rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
         sqls = " exec Sp_Folio_Sel_Upd 'UPD',0, 'AJU',  " & Val(Folio) & ""
         cnxBD.Execute sqls, intRegistros
                 
         blnreimp = False

      End If
   Next i
End With
ImpUnaVez = False

rsBD.Close
Set rsBD = Nothing

CargaPedidos

Exit Sub
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmImpFactura.Graba_Nota_consumo")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Impresion de Nota de Consumo"
   MsgBar "", False
End Sub


Sub GrabaFactura()
Dim PRIMERA As Boolean
Dim resp, dob_factura As Byte
Dim impuesto As Double
Dim rsimp As ADODB.Recordset
Dim nfol_fac As Long, Pedido As Long, sserie As String, serie2 As String

On Error GoTo err_gral

PRIMERA = True
NTEMPSUC = 0
NTEMPSUC = cboBodegas.ItemData(cboBodegas.ListIndex)

If NTEMPSUC = 0 Then
   MsgBox "Problemas al encontrar la serie de la Factura ", vbOKOnly, "Avisar a Sistemas"
   Exit Sub
End If

With spdPedidos

 
   For i = 1 To .MaxRows
      .Row = i
      .Col = 5
      If .value = 1 Then
            
        .Col = 1
        Pedido = .Text
        .Col = 2
        cliente = Val(.Text)
        
                
        sqls = "SELECT * FROM Bon_cte_Prod Where bon_cpd_cte=" & cliente & " and bon_cpd_prod=" & Product
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        If Not rsBD.EOF Then
            sqls = " EXEC SP_FACTURABE @BODEGA = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
               " ,@CLIENTE = " & cliente & _
               " ,@PEDIDO = " & Pedido & _
               " ,@FECHAFAC = '" & Format(txtFechaFac, "mm/dd/yyyy") & "'" & _
               " ,@Usuario = '" & Usuario & "'" & _
               " ,@Producto =" & Product
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
                 
            
             If Not rsBD.EOF Then
                dob_factura = Val(rsBD!Dob)
                Folio = rsBD!folioFac
                Serie = Trim(rsBD!Serie)
                serie2 = Trim(rsBD!serie2)
                Call Producto_actual
                If dob_factura = 0 Then
                   MsgBox "Factura " & rsBD!folioFac & " generada.", vbInformation, "Facturacion"
                   Call doGenArchFE(cboBodegas.ItemData(cboBodegas.ListIndex), CStr(Serie), Val(Folio), Val(Folio), 7)
                Else
                   MsgBox "Facturas " & rsBD!folioFac - 1 & " y " & rsBD!folioFac & "  generadas.", vbInformation, "Despensa"
                   Call doGenArchFE(cboBodegas.ItemData(cboBodegas.ListIndex), CStr(Serie), Val(Folio) - 1, Val(Folio), 7)
                End If
             Else
               MsgBox "Error al generar la factura", vbCritical, "Error..."
               Exit Sub
             End If
             blnreimp = False
        Else
          MsgBox "El cliente No." & cliente & " tiene comision cero y asi no se puede facturar", vbCritical, "Verifique la comision de este cliente antes de facturar"
        End If
      End If
   Next i

End With
ImpUnaVez = False

rsBD.Close
Set rsBD = Nothing

CargaPedidos

Exit Sub

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmImpFactura.Grabafactura")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Impresion de Facturas a Clientes"
   MsgBar "", False
End Sub

Function BuscaFolio()
sqls = " exec Sp_Folio_Sel_Upd 'SEL', 0, 'AJU'"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
    BuscaFolio = rsBD!Folio
Else
    BuscaFolio = 1
End If
End Function

Private Sub Encabezado()
    ' Imprime Encabezado
    importe = rsBD!importe
    iva_total = rsBD!iva_total
    Subtotal = rsBD!Subtotal
    Comision = rsBD!Comision
    val_comision = rsBD!val_comision
    iva_comision = rsBD!iva_comision
    If IsNull(iva_comision) = True Then iva_comision = 0
    If IsNull(val_comision) = True Then val_comision = 0
    total_comision = val_comision + iva_comision
    
    'xFuncion.letras (rsbd!total3)
    Leyenda (rsBD!total3), ""
    svalor = letrero
    nTotal_Partida = Right(Space(6) & Trim(Format(rsBD!total_cantidad_partida, "#####")), 5)
    nTotal1 = Right(Space(12) & Trim(Format(rsBD!total1, "#,###,##0.00")), 12)
    nTotal2 = Right(Space(10) & Trim(Format(rsBD!total2, "###,##0.00")), 10)
    nTotal3 = Right(Space(12) & Trim(Format(rsBD!total3, "#,###,##0.00")), 12)
    
    If nPrimerFactura = 1 Then
        Print #2,
    End If
    Print #2,
    Print #2,
    If cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 1 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 10 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 4 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 5 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 6 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 8 Then Print #2,
    If cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 1 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 10 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 4 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 5 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 6 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 12 Or cmbSucursal.ItemData(cmbSucursal.ListIndex) <> 8 Then Print #2,
    
    Print #2,
    Print #2,
    Print #2,
    Print #2,
    Print #2,
    Print #2,
    Print #2, Tab(3); Left(Trim(Mid(rsBD!Nombre, 1, 30)) & Space(30), 30); _
    Tab(37); Left(Trim(rsBD!Rfc) & Space(15), 15); _
    Tab(56); Left(Trim(Format(rsBD!cliente, "#####")) & Space(5), 5); _
    Tab(63); "Pedido " & Right(Space(7) & Trim(Format(rsBD!Pedido, "########")), 7); _
    Tab(88); Left(Trim(Format(rsBD!BON_FAC_FECHAENT, "YYYY-MM-DD")) & Space(10), 10); _
    Tab(100); Right(Space(2) & Trim(rsBD!bon_fac_serie), 2) & Left(Trim(Format(rsBD!BON_FAC_NUMERO, "#######")) & Space(7), 7)
    Print #2, Tab(3); Left(Trim(Mid(rsBD!Nombre, 31, 30)) & Space(30), 30)
    Print #2, Tab(3); rsBD!Domicilio
    Print #2, Tab(3); rsBD!Colonia
    Print #2, Tab(3); Left(Trim(rsBD!poblacion) & Space(30), 30); Tab(34); Left(Trim(rsBD!CodigoPostal) & Space(6), 6); Tab(41); Left(Trim(rsBD!Telefono) & Space(14), 14); Tab(55); nVendedor
    Print #2,
    Print #2,
    Print #2,
    Print #2, Tab(8); "------------------- DESGLOSE DEL PEDIDO ------------------- "
    Print #2, Tab(20); "DEL :                    AL:"
End Sub

Private Sub Detalle()

    Print #2, Tab(8); rsBD!Producto & " " & rsBD!iva & "%  " & Right(Space(13) & Trim(rsBD!fol_inicial), 13) & "          " & Right(Space(13) & Trim(rsBD!fol_final), 13); _
    Tab(61); Right(Space(5) & Trim(Format(rsBD!cantidad_partida, "#,###")), 5); _
    Tab(69); Right(Space(6) & Trim(Format(rsBD!valor_partida, "##0.00")), 6); _
    Tab(76); Right(Space(12) & Trim(Format(rsBD!valor_menos_iva, "#,###,##0.00")), 12); _
    Tab(88); Right(Space(10) & Trim(Format(rsBD!valor_iva, "###,##0.00")), 10); _
    Tab(99); Right(Space(12) & Trim(Format(rsBD!valor_mas_iva, "#,###,##0.00")), 12)
    
End Sub
Sub PieDePagina()
    Print #2, Tab(20); "SUBTOTAL : "; _
    Tab(76); Right(Space(12) & Trim(Format(importe, "#,###,##0.00")), 12); _
    Tab(88); Right(Space(10) & Trim(Format(iva_total, "###,##0.00")), 10); _
    Tab(99); Right(Space(12) & Trim(Format(Subtotal, "#,###,##0.00")), 12)
    
    Print #2, Tab(20); "SERVICIO:   " & Format(Comision, "#0.00") & " %"; _
    Tab(76); Right(Space(12) & Trim(Format(val_comision, "#,###,##0.00")), 12); _
    Tab(88); Right(Space(10) & Trim(Format(iva_comision, "###,##0.00")), 10); _
    Tab(99); Right(Space(12) & Trim(Format(total_comision, "#,###,##0.00")), 12)
    Print #2,
    Print #2, Tab(61); Right(Space(6) & Trim(Format(nTotal_Partida, "##,###")), 6); _
    Tab(76); Right(Space(12) & Trim(Format(nTotal1, "#,###,##0.00")), 12); _
    Tab(88); Right(Space(10) & Trim(Format(nTotal2, "###,##0.00")), 10); _
    Tab(99); Right(Space(12) & Trim(Format(nTotal3, "#,###,##0.00")), 12)
    
    Print #2,
    Print #2,
    Print #2,
    Print #2, Tab(17); sgMotivo
    Print #2,
    Print #2, Tab(70); Right(Space(12) & Trim(Format(nTotal3, "#,###,##0.00")), 12); Tab(65); Right(Space(40) & Mid(Trim(svalor), 1, 40), 40)
    Print #2, Tab(17); "www.valetotal.com"; Tab(65); Right(Space(40) & Mid(Trim(svalor), 41, 40), 40)
    Print #2, Tab(17); "PAGO EN UNA SOLA EXHIBICION"
    Print #2,
    Print #2,
    Print #2, ""
End Sub

Private Sub cmdRefresh_Click()
 CargaPedidos
End Sub

Private Sub cmdReimp_Click()
    frmReimp.Visible = True
    Call CboPosiciona(cboBodegasReimp, gnBodega)
    cmdReimp2.Enabled = False
End Sub

Private Sub cmdReimp2_Click()
    If Val(txtFactura) <> 0 And Trim(lblCliente) <> "" Then
     blnreimp = True
     Call Imprimefactura(cboBodegas.ItemData(cboBodegas.ListIndex), Val(txtFactura), txtSerie, Val(lblPedido))
     frmReimp.Visible = False
     blnreimp = False
     sTipoArch = "FACTURAS"
     strPuerto = doFindPrinter(gstrPC, sTipoArch)
     sFileFact = "C:\facturacion\f" & Val(txtFactura) & ".lis  "
     nFileBat = FreeFile()
     sFileBat = "c:\facturacion\bat\factura.bat"
     Close #nFileBat
     Open sFileBat For Output As #nFileBat
     If gsOS = "XP" Then
       Print #nFileBat, "PRINT /D:" & strPuerto & " " & sFileFact
     Else
       Print #nFileBat, "COPY " & sFileFact & " " & strPuerto
     End If
     Close #nFileBat
    
                  
    If MsgBox("Desea mandar la factura a la impresora", vbQuestion + vbYesNo, "Impresion") = vbYes Then
   '   If ImpUnaVez = False Then
        resp = InputBox("Por favor teclee el siguiente folio fisico de la factura", "Folio de factura")
        sqls = "exec  sp_AsigPapFact_Sel " & Val(sUbicacion) & ", '" & Impresora & "', '" & txtSerie & "', 1,1"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        If Not rsBD.EOF Then
            If Val(rsBD!FolioAct) <> Val(resp) Then
                MsgBox "El folio que capturó, no coincide con el folio fisico de la factura, NO SE IMPRIMIRÁ LA FACTURA " & nfol_fac & "", vbCritical, "Error en folio"
                ImpUnaVez = False
            Else
                ImpUnaVez = True
            End If
        Else
            MsgBox "Error en folios fisicos", vbCritical, "Error en folios"
        End If
    '    End If
    
        If doValImpFact(Val(sUbicacion), Impresora, 1, "") <> 0 Then
               nResp = vbNo
               ImpUnaVez = False
        Else
            If ImpUnaVez = True Then
               ' If MsgBox("Desea mandar la factura a la impresora", vbYesNo) = vbYes Then
                    EsperarShell (sFileBat) ', vbMinimizedFocus
                    sqls = "select bodegaimprime BI from bodegas where bodega = " & Val(gnBodega)
                    Set rsbod = New ADODB.Recordset
                    rsbod.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
                    
                          If Not rsbod.EOF Then
                              BodegaImp = rsbod!BI
                          Else
                              BodegaImp = Val(gnBodega)
                          End If
                    
                    strsql = "sp_GrabaProcImpFact " & BodegaImp & ", " & _
                        "'" & Impresora & "', " & _
                        "2," & _
                        "'" & Trim(txtSerie) & "', " & _
                        Val(txtFactura) & ", " & _
                        "'" & gstrUsuario & "', " & _
                         "'" & Trim(txtSerie) & "'"
                    Set rstPrc = New ADODB.Recordset
                    rstPrc.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
                    DoEvents
                    Set rstPrc = Nothing
                'End If
            End If
            Kill sFileFact
        End If
        
    End If

    Else
        MsgBox "No ha capturado el numero de factura", vbExclamation, "Falta No. de Factura"
        txtFactura.SetFocus
    End If
End Sub

Private Sub cmdReimprime_Click()
    frmReimp.Visible = True
    Call CboPosiciona(cboBodegasReimp, gnBodega)
    cmdReimp2.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub chkSelTodas_Click()
    If chkSelTodas.value = 1 Then
        For i = 1 To spdPedidos.MaxRows
            spdPedidos.Row = i
            spdPedidos.Col = 5
            spdPedidos.value = 1
        Next i
    Else
        For i = 1 To spdPedidos.MaxRows
            spdPedidos.Row = i
            spdPedidos.Col = 5
            spdPedidos.value = 0
        Next i
    End If
End Sub

Private Sub Command1_Click()
 'Producto_factura = 1
 'Call doGenArchFE(1, "AA", 85777, 85777, 7)
  Reimp_FCNC = 7
  frmFact.Show
End Sub

Private Sub Command2_Click()
    frmReimp.Visible = False
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub

Sub ocupa_usuario()
Dim Nombre As String
    continua = True
    sql = "SELECT Nombre from Usuarios where Usuario='" & Trim(gstrUsuario) & "'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       Nombre = Mid(Trim(rsBD!Nombre), 1, 30)
       sql = "SELECT COUNT(Usuario) Total FROM usuario_ajustes"
       Set rsBD = New ADODB.Recordset
       rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
       If rsBD!total <= 0 Then
          sql = "INSERT INTO usuario_ajustes VALUES('" & Trim(gstrUsuario) & "','" & Trim(Nombre) & "')"
          cnxbdMty.Execute sql
          continua = True
       Else
          sql = "SELECT TOP 1 Isnull(Nombre,'USUARIO DESCONOCIDO') Usuario FROM usuario_ajustes"
          Set rsBD = New ADODB.Recordset
          rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
          MsgBox "¿El usuario " & Trim(rsBD!Usuario) & " esta ocupando el modulo de ajustes, y estara restringido mientras" & vbCrLf & _
          " termina sus ajustes para evitar mezclar informacion", vbInformation, "Opcion restringida Temporalmente"
          Frame6.Enabled = False
          Frame5.Enabled = False
          Frame2.Enabled = False
          Frame1.Enabled = False
          Frame3.Enabled = False
          Frame4.Enabled = False
          continua = False
          'Unload Me
       End If
    Else
      continua = False
    End If
End Sub

Private Sub Form_Load()
    ahora = Now
    Set mclsAniform = New clsAnimated
    nop = False
    cierrabe = False
    cte_interno = "19620" 'pruebas 10622
    CboProducto.Clear
    Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
    CboProducto.Text = UCase("Winko Mart")
    
    If user_master = False Then
       CargaBodegasS2 cboBodegas
       CargaBodegasS2 cboBodegasReimp
    Else
       CargaBodegasServ cboBodegas
       CargaBodegas cboBodegasReimp
    End If
    
    'CargaBodegasServ cboBodegas
    'CargaBodegas cboBodegasReimp
    CargaPedidos
    frmReimp.Visible = False
    ImpUnaVez = False
    sal = 0
    Call ocupa_usuario
    Call STAT
    Call investiga_sihayNC
    
    If Dir("C:\Facturacion\Paso", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema de facturaciòn en este equipo, favor de comunicarse a sistemas"
        Command1.Enabled = False
        cmdReimp.Enabled = False
        cmdRefresh.Enabled = False
        Modificar.Enabled = False
        cmdCancelar.Enabled = False
        cmdGrabar.Enabled = False
    End If
    If Dir("C:\Facturacion\Pruebas", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema de facturaciòn en este equipo, favor de comunicarse a sistemas"
        Command1.Enabled = False
        cmdReimp.Enabled = False
        cmdRefresh.Enabled = False
        Modificar.Enabled = False
        cmdCancelar.Enabled = False
        cmdGrabar.Enabled = False
    End If
    If Dir("C:\GoDir", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema GOSOFT, favor de comunicarse a sistemas"
        Command1.Enabled = False
        cmdReimp.Enabled = False
        cmdRefresh.Enabled = False
        Modificar.Enabled = False
        cmdCancelar.Enabled = False
        cmdGrabar.Enabled = False
    End If
End Sub

Sub Imprimefactura(Bodega As Integer, Factura As Long, Serie As String, Pedido As Long)
Dim RengImp As Integer
Dim ArchImpre
Close #2

ArchImpre = "c:\Facturacion\f" & Factura & ".lis"
Open ArchImpre For Output As #2
                
sql = "select  b.BON_FAC_NUMERO factura , b.BON_FAC_CLIENTE  cliente  , C.NOMBRE, ISNULL(f.Calle, C.Domicilio) Domicilio, C.COLONIA, C.POBLACION,  C.CODIGOPOSTAL, C.TELEFONO, C.VENDEDOR,  b.BON_FAC_VALOR valor ,C.RFC, " & _
" b.BON_FAC_FECHAEMI BON_FAC_FECHAEMI , b.BON_FAC_FECHAENT BON_FAC_FECHAENT , b.BON_FAC_FECHAVEN BON_FAC_FECHAVEN ,  " & _
" isnull(b.BON_FAC_BONEXE,0) BON_FAC_BONEXE ,      isnull(b.BON_FAC_BONGRA,0) BON_FAC_BONGRA     , isnull(b.BON_FAC_IVAGRA,0)  BON_FAC_IVAGRA , " & _
" isnull(b.BON_FAC_COMISION,0)  BON_FAC_COMISION ,     isnull(b.BON_FAC_IVACOMIS,0) BON_FAC_IVACOMIS,       isnull(b.BON_FAC_VALOR,0) BON_FAC_VALOR, " & _
" isnull(b.BON_FAC_PEDIDO,0) BON_FAC_PEDIDO , isnull(b.BON_FAC_USUARIO,0) BON_FAC_USUARIO , B.BON_FAC_SERIE, e.DESCCORTA estado, d.descripcion poblacion, isnull(c.rutanumero,0) rutanumero" & _
"   from bon_pedido a , bon_factura b, CLIENTES C , POBLACIONES D, estados e, clientesdg f" & _
    " Where a.bon_ped_sucursal  = " & Bodega & _
    " and   a.bon_ped_numero    = " & Pedido & _
    " and   a.bon_ped_producto      =" & Product & _
    " and   b.bon_fac_sucursal      = a.bon_ped_sucursal " & _
    " and   b.bon_fac_numero        = " & Factura & _
    " and   b.bon_fac_pedido    = a.bon_ped_numero     " & _
    " and   b.BON_FAC_CLIENTE   = a.bon_ped_cliente     " & _
    " and   b.BON_FAC_TPOBON    = a.bon_ped_producto  " & _
    " AND   B.BON_FAC_CLIENTE = C.CLIENTE " & _
    " AND   c.CLIENTE *= f.CLIENTE " & _
    " AND   C.POBLACION = D.POBLACION" & _
    " and   C.ESTADO = E.ESTADO" & _
    " order by b.BON_FAC_NUMERO "
                        
    Set rsbd1 = New ADODB.Recordset
    rsbd1.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    
           
    fecha_emi$ = Format(rsbd1!bon_fac_fechaemi, "yyyy-mm-dd")
    fecha_ent$ = Format(rsbd1!BON_FAC_FECHAENT, "yyyy-mm-dd")
    fecha_ven$ = Format(rsbd1!BON_FAC_FECHAVEN, "yyyy-mm-dd")
    nBON_FAC_BONEXE = Format(rsbd1!bon_fac_bonexe, "########0.00")
    Subtotal = Format(rsbd1!bon_fac_bonexe, "########0.00")
    ivaexe = Format(0, "########0.00")
    total1 = Format(rsbd1!bon_fac_bonexe, "########0.00")
    val_comision = Format(rsbd1!bon_fac_comision, "########0.00")
    ivacomis = Format(rsbd1!BON_FAC_IVACOMIS, "########0.00")
    total2 = Format(rsbd1!bon_fac_comision + rsbd1!BON_FAC_IVACOMIS, "########0.00")
    total = Format(rsbd1!bon_fac_bonexe + rsbd1!bon_fac_comision, "########0.00")
    TotalIva = Format(rsbd1!BON_FAC_IVACOMIS, "########0.00")
    total3 = Format(rsbd1!bon_fac_bonexe + rsbd1!bon_fac_comision + rsbd1!BON_FAC_IVACOMIS, "########0.00")
    
    
    svalor = Leyenda(rsbd1!bon_fac_valor)
     
    GoSub Encabezado
    GoSub Detalle
    GoSub PieDePagina
    
    Close #14
    
    '----------------------------
       
      DoEvents
      
    Exit Sub

Encabezado:
   RengImp = 5
   sEdo = Trim(rsbd1!estado & "")
   spobl = Trim(rsbd1!poblacion & "")
   scp = Trim(rsbd1!CodigoPostal & "")
   scp = IIf(scp = "0", "", scp)
   spobl = Left(spobl, 40 - (Len(sEdo) + 2) - (Len(scp) + 2)) & ", " & sEdo & "  " & scp
   sserie = rsbd1!bon_fac_serie

   For lngI = 1 To 11
      Print #2,
   Next
   Print #2, Tab(5); Pad(Trim(Mid(rsbd1!Nombre, 1, 44)), 45, " ", "R"); _
               Tab(51); Pad(Trim(IIf(IsNull(rsbd1!Rfc), " ", rsbd1!Rfc)), 15, " ", "R"); _
               Tab(75); Pad(Trim(Format(rsbd1!cliente, "#####")), 5, " ", "R"); _
               Tab(86); Pad(Trim(Format(rsbd1!bon_fac_fechaemi, "YYYY-MM-DD")), 10, " ", "R"); _
               Tab(100); Pad(Trim(rsbd1!bon_fac_serie), 2, " ", "R") & Pad(Trim(Format(rsbd1!Factura, "#######")), 7, " ", "R")
   Print #2, Tab(5); Pad(Trim(Mid(rsbd1!Nombre, 45, 90)), 45, " ", "R")
   Print #2, Tab(5); rsbd1!Domicilio
   Print #2, Tab(5); rsbd1!Colonia
   
   If rsbd1!cliente = 17412 Then
        Print #2, Tab(5); Pad(Trim(spobl), 40, " ", "R"); _
                  Tab(86); "PAGO: DE CONTADO"
   Else
        Print #2, Tab(5); Pad(Trim(spobl), 40, " ", "R")
   End If
   
   Print #2, Tab(51); Pad(Trim(IIf(IsNull(rsbd1!Telefono), "", rsbd1!Telefono)), 14, " ", "R"); _
             Tab(75); Pad(Trim(Format(rsbd1!bon_fac_pedido, "########")), 8, " ", "R"); _
               Tab(107); "R" & rsbd1!RutaNumero
   
   If rsbd1!cliente = 17412 Then
        Print #2, Tab(86); "CREDITO: 0 DIAS"
        RengImp = 4
   End If
               
   For lngI = 1 To RengImp
      Print #2,
   Next
   Print #2, Tab(5); "----------------------- DESGLOSE DEL PEDIDO ----------------------- "
   Print #2,

   Return
   
Detalle:

    sqls = " select cliente as BON_CLIEMP_DEPTO, count(numempl) as totemp, sum(valor) AS TOTALVALOR"
    sqls = sqls & " From recibos"
    sqls = sqls & " Where Sucursal = " & Bodega
    sqls = sqls & " and pedido = " & rsbd1!bon_fac_pedido
    sqls = sqls & " group by cliente"

    Set rsdet = New ADODB.Recordset
    rsdet.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Do While Not rsdet.EOF
        cont = cont + 1
        
        DescProd = ""
        If Product = 1 Then
              DescProd = "Winko Mart"
        ElseIf Product = 2 Then
              DescProd = "VALE TOTAL COMB NACIONAL"
        ElseIf Product = 3 Then
              DescProd = "VALE TOTAL COMB CONTROL"
        End If
        
        Print #2, Tab(5); DescProd; Tab(34); rsdet!bon_cliemp_depto; _
        Tab(50); Pad(Trim(Format(rsdet!TotEmp, "##,##0")), 6, " ", "L"); _
        Tab(56); " Empleados"; Tab(72); Pad(Trim(Format(rsdet!TotalValor, "###,##0.00")), 12, " ", "L"); _
        Tab(85); Pad(Trim(Format(0, "##,##0.00")), 9, " ", "L"); Tab(96); Pad(Trim(Format(rsdet!TotalValor, "##,###,##0.00")), 13, " ", "L") _
        
        rsdet.MoveNext
    Loop

    For X = 1 To 20 - cont
        Print #2,
    Next X

Return

PieDePagina:
        
    sqls = "select bon_cpd_comision from bon_cte_prod"
    sqls = sqls & " Where bon_cpd_prod =" & Product '6"
    sqls = sqls & " and bon_cpd_cte = " & rsbd1!cliente

    Set rscom = New ADODB.Recordset
    rscom.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

    If Not rscom.EOF Then
        Comision = rscom!bon_cpd_comision
    Else
        Comision = 0
    End If
    
    sqls = "select impuestointerior iva from bodegas"
    sqls = sqls & " Where bodega = " & Bodega & ""
    
    Set rscom = New ADODB.Recordset
    rscom.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

    If Not rscom.EOF Then
        iva = rscom!iva
    Else
        iva = 0
    End If
    
   sServicio = "SERVICIO:  " & Format(Comision, "##0.00") & " %   IVA:  " & iva & "%"
  
   Print #2, Tab(20); "SUBTOTAL : "; _
                     Tab(72); Pad(Trim(Format(Subtotal, "#,###,##0.00")), 12, " ", "L"); _
                     Tab(85); Pad(Trim(Format(ivaexe, "##,##0.00")), 9, " ", "L"); _
                     Tab(96); Pad(Trim(Format(total1, "##,###,##0.00")), 13, " ", "L")
   Print #2, Tab(20); sServicio; _
                     Tab(72); Pad(Trim(Format(val_comision, "###,##0.00")), 12, " ", "L"); _
                     Tab(85); Pad(Trim(Format(ivacomis, "##,##0.00")), 9, " ", "L"); _
                     Tab(96); Pad(Trim(Format(total2, "##,###,##0.00")), 13, " ", "L")
   Print #2,
   Print #2,
   Print #2,
   
   Print #2, Tab(17); "PAGO EN UNA SOLA EXHIBICION"; Tab(72); Pad(Trim(Format(total, "#,###,##0.00")), 12, " ", "L"); _
                     Tab(85); Pad(Trim(Format(TotalIva, "##,##0.00")), 9, " ", "L"); _
                     Tab(96); Pad(Trim(Format(total3, "##,###,##0.00")), 13, " ", "L")
   For lngI = 1 To 5
      Print #2,
   Next
'   Print #2, Tab(17); sgMotivo
   Print #2,
  
   Print #2, Tab(71); Format(rsbd1!BON_FAC_FECHAENT, "D"); Tab(77); Format(rsbd1!BON_FAC_FECHAENT, "MMMM"); Tab(91); Format(rsbd1!BON_FAC_FECHAENT, "YYYY")
   Print #2,
   svalor = Pad(Trim(Format(total3, "#,###,##0.00")), 12, " ", "L") & " " & svalor
   If Len(svalor) < 40 Then
        Print #2, Tab(17); Mid(svalor, 1, 40);  'Right(Space(40) & Mid(Trim(svalor), 1, 40), 40)
   Else
        Print #2, Tab(17); Mid(svalor, 1, 40);
   End If
   
   If Len(svalor) >= 40 Then
       Print #2, Tab(17); Mid(svalor, 41, Len(svalor) - 40); ""
   End If
'   Print #2,

   'Print #2, ""
   
   Close #2
  
   Return

End Sub

Sub CargaPedidos()

sql = "sp_Bodegas_sel " & cboBodegas.ItemData(cboBodegas.ListIndex)

Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
   txtFechaFac = rsBD!FechaOp
Else
   MsgBox "Error al cargar la fecha de facturación, no se pueden cargar los pedidos de esta sucursal", vbCritical, "Errores presentados"
   Exit Sub
End If
'Pedidos autorizados
If Option1.value = True Then
   sql = "sp_FactBEVarios Null," & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex)) & _
   ",Null,Null," & Product & ",'CargaPedidos'"
End If

If Option2.value = True Then
   sql = "sp_FactBEVarios Null," & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex)) & _
   "," & cte_interno & ",Null," & Product & ",'CargaPedidos2'"
End If
 
If Option1.value = False And Option2.value = False Then
   sql = "sp_FactBEVarios Null," & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex)) & _
   ",Null,Null," & Product & ",'CargaPedidos'"
   Option1.value = True
End If

Call STAT

'sql = " SELECT BON_PED_NUMERO as pedido, BON_PED_CLIENTE as cliente ,NOMBRE as nombre , BON_PED_VALOR as valor"
'sql = sql & " From BON_PEDIDO, CLIENTES"
'sql = sql & " Where BON_PED_CLIENTE = Cliente"
'sql = sql & " AND BON_PED_SUCURSAL = " & IIf(cboBodegas.ItemData(cboBodegas.ListIndex) = -1, 0, cboBodegas.ItemData(cboBodegas.ListIndex))
'sql = sql & " AND BON_PED_PRODUCTO =" & Product '6"
'sql = sql & " AND BON_PED_STATUS=0"


Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    
With spdPedidos
If rsBD.EOF Then
  '  MsgBox "No hay pedidos pendientes por facturar", vbInformation
    .Col = -1
    .Row = -1
    .Action = 12
    .MaxRows = 0
Else
    i = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = Val(rsBD!Pedido)
        .Col = 2
        .Text = Val(rsBD!cliente)
        .Col = 3
        .Text = rsBD!Nombre
        .Col = 4
        .Text = CDbl(rsBD!valor)
        rsBD.MoveNext
    Loop
    
End If
End With
End Sub
Sub GeneraReporte()
 If lstTipo.ItemData(lstTipo.ListIndex) = 12 Then
        
        sql = "  exec spr_imprime_factura_all " & _
        NTEMPSUC & " , " & _
        nfol_fac & " ," & _
        cmbProducto.ItemData(cmbProducto.ListIndex) & "," & npedido & "," & txtnNum_Cte & "," & nTipoCom
        
         
        Set rsBD = New ADODB.Recordset
        
        rsBD.Open sql, objSqlConexion, adOpenForwardOnly, adLockOptimistic
        If rsBD.EOF = True Then
            objSqlConexion.RollbackTrans
            MsgBar "", False
            MsgBox "Error al Generar archivo de Factura, Avisar a Sistemas !", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        

        
                Print #50, Tab(0); Right("00" + Trim(cmbProducto.ItemData(cmbProducto.ListIndex)), 2) & _
                Right("00000" + Trim(CStr(rsBD!cliente)), 5) & _
                Right("000000" + Trim(CStr(rsBD!BON_FAC_NUMERO)), 6) & _
                Right("000000000000" + Trim(CStr(Round(rsBD!total3, 0))), 12) & _
                Right("00" + Trim(CStr(rsBD!total3 - Round(rsBD!total3, 0))), 2)
                        
                        
                        
        Nombre = "": Domicilio = "": Colonia = "": poblacion = "": estado = "": CodigoPostal = "": Rfc = ""
        bon_fac_serie = "": Producto = "": fol_inicial = "": fol_final = 0
        cantidad_partida = 0: cliente = 0: vendedor = 0
        BON_FAC_NUMERO = 0
        
        iva = 0: valor_partida = 0: valor_menos_iva = 0: valor_iva = 0: valor_mas_iva = 0: Comision = 0: importe = 0
        
        
        iva_total = 0: Subtotal = 0: val_comision = 0: iva_comision = 0: total_comision = 0

        GdaFicha = 0
        Contt = 1
        'ArchImpre = "c:\bonos\resp\" & Trim(CStr(nFol_fac)) & "_facturas.lis"
        'Open ArchImpre For Output As #2
        Do While rsBD.EOF() = False
            'Salto De Paginas
            If Nombre = "" Then
                Contt = 1
                Encabezado
                Nombre = rsBD!Nombre
            End If

            If Contt > 16 Then
                'pie de pagina sin letrero
                For i = 1 To 8
                    Print #2,
                Next
                'Termina Avance
                Contt = 1
                'Encabezado
            End If
            Contt = Contt + 1
            Detalle
            rsBD.MoveNext
        Loop
        'For i = 1 To (19 - Contt)
        For i = 1 To (28 - Contt)
            Print #2,
        Next
        PieDePagina
        End If
        Close #2
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If nop = True Then
     Timer1.Enabled = True
  End If
  'Alerta.Visible = True
End Sub

Private Sub Form_Terminate()
If continua = True Then
  sql = "SELECT TOP 1 * FROM usuario_ajustes"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rsBD.EOF Then
      If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
         sql = "DELETE usuario_ajustes"
         cnxbdMty.Execute sql
     End If
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If continua = True Then
   sql = "SELECT TOP 1 * FROM usuario_ajustes"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
   If Not rsBD.EOF Then
      If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
         sql = "DELETE usuario_ajustes"
         cnxbdMty.Execute sql
     End If
   End If
End If
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If nop = True Then
     Timer1.Enabled = True
  End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If nop = True Then
     Timer1.Enabled = True
  End If
End Sub


Private Sub Modificar_Click()
Dim i As Integer
Dim Pedido As Integer, cliente As Integer
On Error GoTo ERR:
   
    If spdPedidos.MaxRows = 0 Then
       MsgBox "No hay pedidos por facturar", vbExclamation, "No hay pedidos"
       Exit Sub
    End If
    
    With spdPedidos
    For i = 1 To .MaxRows
      .Row = i
      .Col = 5
      If .value = 1 Then
         Bodegp = cboBodegas.ItemData(cboBodegas.ListIndex)
         .Col = 1
         pedidop = .Text
         .Col = 2
         clientepp = Val(.Text)
         frmCambioCliente.Show 1
         Exit Sub
      End If
    Next
    End With
    
    MsgBox "No ha seleccionado ningun pedido para modificarle el No. de Cliente", vbInformation, "No ha seleccionado nada"
  
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub Option1_Click()
   CargaPedidos
End Sub

Private Sub Option2_Click()
  CargaPedidos
End Sub

Private Sub spdPedidos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If nop = True Then
     Timer1.Enabled = True
  End If
End Sub


Private Sub Timer1_Timer()
If sal = 0 Then
   Alerta.Visible = True
   sal = 1
Else
   Alerta.Visible = False
   sal = 0
End If
End Sub

Private Sub Timer2_Timer()
ahora2 = Now
dif = ahora2 - ahora
If Format(dif, "HH:MM:SS") >= "00:05:00" Then
   '---
   If continua = True Then
      sql = "SELECT TOP 1 * FROM usuario_ajustes"
      Set rsBD = New ADODB.Recordset
      rsBD.Open sql, cnxbdMty, adOpenForwardOnly, adLockReadOnly
      If Not rsBD.EOF Then
         If Trim(rsBD!Usuario) = Trim(gstrUsuario) Then
            sql = "DELETE usuario_ajustes"
            cnxbdMty.Execute sql
         End If
      End If
   End If
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
   '--
   MsgBox "Excedio el limite de tiempo para permanecer en este modulo", vbCritical, "Expulsion temporal"
   Unload Me
End If
End Sub

Private Sub Timer3_Timer()
Dim min5 As Date
ahora2 = Now
min5 = "00:05:00"
dif = min5 - (ahora2 - ahora)
Me.Caption = "Facturacion Electonica BE restan: " & Format(dif, "HH:MM:SS")
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtFactura) <> "" Then
        
            sqls = "sp_FactBEVarios Null," & cboBodegas.ItemData(cboBodegas.ListIndex) & _
            ",Null," & Val(txtFactura) & ",Null,'BuscaFactura'"
            
            
'            sqls = "select bon_fac_sucursal, bon_fac_cliente, bon_fac_numero, bon_fac_status" & _
'                   " ,bon_fac_valor, bon_fac_pedido ,bon_fac_tpobon,bon_Fac_serie, nombre  from bon_factura , clientes" & _
'                   " Where bon_fac_sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'                   " and bon_fac_numero = " & Val(txtFactura) & _
'                   " and bodega = bon_Fac_sucursal and cliente = bon_fac_cliente"

                   
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                If rsBD!bon_fac_tpobon <> 6 And rsBD!bon_fac_tpobon <> 7 And rsBD!bon_fac_tpobon <> 8 Then
                    MsgBox "Esta factura no es de Bono electronico, no es posible reimprimirla", vbCritical, "Error en factura"
                    txtFactura = ""
                    Exit Sub
                End If
                
                If rsBD!bon_Fac_Status = 2 Then
                    MsgBox "Esta factura no se puede reimprimir ya que esta cancelada", vbCritical, "Factura cancelada"
                    txtFactura = ""
                    Exit Sub
                End If
                
                lblCliente = Trim(rsBD!Nombre)
                lblImporte = rsBD!bon_fac_valor
                lblPedido = rsBD!bon_fac_pedido
                txtSerie = rsBD!bon_fac_serie
                cmdReimp2.Enabled = True
                
            Else
                MsgBox "Factura Invalida !!", vbCritical, "Factura invalida!"
                cmdReimp2.Enabled = False
            End If
                
                
        Else
            MsgBox "Primero capture el numero de factura a reimprimir", vbInformation, "Capture No. de factura"
            cmdReimp2.Enabled = False
            
        End If
    End If
End Sub

Private Sub txtFactura_LostFocus()
    txtFactura_KeyPress (13)
End Sub

Sub investiga_sihayNC()
    sqls = "UPDATE BON_PEDIDO"
    sqls = sqls & " SET bon_ped_tipocom=3"
    sqls = sqls & " WHERE Bon_ped_producto=7 "
    sqls = sqls & " AND bon_ped_status=0 "
    sqls = sqls & " AND bon_ped_tipocom=1 "
    sqls = sqls & " AND bon_ped_cliente IN (" & cte_interno & ")"
   ' sqls = sqls & " AND bon_ped_fechaped>='03/30/2010'" 'PROVISIONAL PARA HACER PRUEBAS
    cnxBD.Execute sqls, intRegistros
'--------------------
    rsBD.Close
    sqls = "SELECT * FROM BON_PEDIDO WHERE Bon_ped_producto=7 "
    sqls = sqls & " AND bon_ped_status=0 "
    sqls = sqls & " AND bon_ped_tipocom=3 "
    sqls = sqls & " AND Bon_ped_producto=7 "
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    If rsBD.EOF Then
       nop = False  'no hay notas de consumo
       Timer1.Enabled = False
       Alerta.Visible = False
    Else
       nop = True
       Timer1.Enabled = True
    End If
End Sub

Sub STAT()
'    sql = "SELECT * FROM CONFIGBODEGAS WHERE Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
'    Set rsBD = New ADODB.Recordset
'    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
'    If Not rsBD.EOF Then
'       If Month(Date) <> Month(rsBD!FechaCierre) Then
'          cmdGrabar.Enabled = False
'          MsgBox "Se tiene que cerrar el mes anterior antes de poder facturar", vbCritical
'       Else
'          cmdGrabar.Enabled = True
'       End If
'    Else
'       MsgBox "Error no hay fecha de cierre actual para esa Sucursal", vbCritical, "Sin fecha"
'    End If
End Sub
Function doValImpFact(nBodega As Integer, sMaquina As String, nCantidad As Integer, Serie As String) As Integer
Dim rstTmp2 As ADODB.Recordset, strsql As String
'Set cnxBD = New ADODB.Connection
'    cnxBD.CommandTimeout = 6000
'    cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
   strsql = "sp_ValImpFact " & nBodega & ", "
   strsql = strsql & "'" & sMaquina & "', " & nCantidad
   If Serie <> "" Then
    strsql = strsql & ",'" & Serie & "'"
   End If
   Set rstTmp2 = New ADODB.Recordset
   rstTmp2.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   nResp = rstTmp2!resp
   DoEvents
   Set rstTmp2 = Nothing
   Select Case nResp
   Case 1
      MsgBox "No tiene suficiente papeleria de Facturas asignadas a la impresora " & gstrImpFact, vbInformation, "Validacion"
   Case 2
      MsgBox "La impresora " & gstrImpFact & " actualmente esta imprimiendo.", vbInformation, "Validacion"
   End Select
   doValImpFact = nResp
End Function
Sub EsperarShell(sCmd As String)
Dim hShell As Long
Dim hProc As Long
Dim codExit As Long

' ejecutar comando
hShell = Shell(Environ$("Comspec") & " /c " & sCmd, 2)

' esperar a que se complete el proceso
hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)

Do
    GetExitCodeProcess hProc, codExit
    DoEvents
Loop While codExit = STILL_ACTIVE
End Sub
