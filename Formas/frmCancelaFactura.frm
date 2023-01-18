VERSION 5.00
Begin VB.Form frmCancelaFactura 
   Caption         =   "Cancelación de FACTURA"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   120
      TabIndex        =   16
      Top             =   240
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
         ItemData        =   "frmCancelaFactura.frx":0000
         Left            =   1800
         List            =   "frmCancelaFactura.frx":000A
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
         TabIndex        =   17
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   6495
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Salir"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5640
         Picture         =   "frmCancelaFactura.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancel"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4800
         Picture         =   "frmCancelaFactura.frx":011E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancelar alta"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmCancelaFactura.frx":0220
         Left            =   1200
         List            =   "frmCancelaFactura.frx":0227
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtFactura 
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtserie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   360
         Width           =   375
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
         TabIndex        =   12
         Top             =   360
         Width           =   975
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
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Width           =   4575
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
         TabIndex        =   9
         Top             =   840
         Width           =   855
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
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblImporte 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblPedido 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
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
         Left            =   3480
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCancelaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim Cte As Integer

Private Sub cboBodegas_Click()

sqls = "select servidor from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    Server = rsBD!Servidor
End If
Call STAT
End Sub

Private Sub cmdCancelar_Click()
'en duda para ver si en el store se añade un parametro
'para producto de entrada parece no ser necesario
resp = MsgBox("Esta seguro de que desea cancelar la factura " & txtFactura.Text & " ?", vbYesNo)

If resp = vbYes Then

'    If Product <> 2 Then
       sqls = "exec sp_cancelafactura @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
           ", @Pedido = " & lblPedido.Caption & _
           ", @TipoCompr = 1" & _
           ", @Serie = " & Trim(txtSerie.Text) & _
           ", @Factura = " & txtFactura & _
           ", @Usuario = '" & gstrUsuario & "'"
           
        cnxBD.Execute sqls, intRegistros
        Call doGenArchCanc(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(txtSerie.Text), txtFactura.Text, 1)
'    ElseIf Product = 2 Then
'        sqls = "exec sp_cancelafactura @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'           ", @Pedido = " & lblPedido.Caption & _
'           ", @TipoCompr = 1" & _
'           ", @Serie = " & Trim(txtserie.Text) & _
'           ", @Factura = " & txtFactura & _
'           ", @Usuario = '" & gstrUsuario & "'"
'        cnxBD.Execute sqls, intRegistros
'        Call doGenArchCanc(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(txtserie.Text), txtFactura.Text, 1)
      
'        sqls = "exec sp_cancelafactura @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'           ", @Pedido = " & lblPedido.Caption & _
'           ", @TipoCompr = 1" & _
'           ", @Serie = " & "Y" & Mid(Trim(txtserie.Text), 2, 1) & _
'           ", @Factura = " & txtFactura & _
'           ", @Usuario = '" & gstrUsuario & "'"
'        cnxBD.Execute sqls, intRegistros
'        Call doGenArchCanc(cboBodegas.ItemData(cboBodegas.ListIndex), "Y" & Mid(Trim(txtserie.Text), 2, 1), txtFactura.Text, 1)
'    End If
      'inicia seccion de mandar mail
      sqls = "sp_PedidosVarios Null," & Val(Cte) & ",Null,Null,'BuscaCliente',Null," & cboBodegas.ItemData(cboBodegas.ListIndex) & ""
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
        
      If rsBD.EOF Then
         lblImporte = ""
         lblPedido = ""
         lblCliente = ""
         MsgBox " Cliente no existe o no esta dado de alta para esta sucursal", vbExclamation, "No se pudo enviar notificacion de cancelacion x mail"
         MsgBox "Factura " & txtFactura & " ha sido cancelada", vbInformation, "Factura cancelada"
         lblCliente.Caption = ""
         lblImporte.Caption = ""
         lblPedido.Caption = ""
         txtSerie.Text = ""
         txtFactura.Text = ""
         Exit Sub
      Else
         sEmail = IIf(IsNull(rsBD!identificacion), "", rsBD!identificacion)
      End If
            
      If gblnSendMail Then
      If sEmail <> "" Then
         Set posendmail = New clsSendMail
         If posendmail.IsValidEmailAddress(sEmail) Then
            
            sAsunto = "Vale Total: Cancelacion de Factura"
            sTexto = "Vale Total agradece su preferencia" & vbCrLf & vbCrLf
            sTexto = sTexto & "Y asi mismo le informa que la factura: " & Trim(txtSerie.Text) & "-" & txtFactura
            sTexto = sTexto & " con No. de Pedido:  " & Trim(lblPedido.Caption) & vbCrLf
            sTexto = sTexto & " y Valor del Pedido: " & Format(lblImporte.Caption, "########0.00") & vbCrLf & _
                              "HA SIDO CANCELADA."

            With posendmail
               .SMTPHost = gstrMailSMTPHost
               .SMTPPort = gstrMailSMTPPort
               .From = gstrMailFrom
               .Recipient = Trim(sEmail) & ";vsoto@valetotal.com"
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
   
   '--------fin seccion de mandar mail
      
        MsgBox "Factura " & txtFactura & " ha sido cancelada", vbInformation, "Factura cancelada"
          
        lblCliente.Caption = ""
        lblImporte.Caption = ""
        lblPedido.Caption = ""
        txtSerie.Text = ""
        txtFactura.Text = ""
        
Else
    MsgBox "No se canceló la factura " & txtFactura.Text & "", vbInformation, "Factura no pudo ser cancelada"
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cboProducto_Click()
  Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  'If aqui <> Product Then
  '   LimpiarControles Me
  'End If
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub
Private Sub Form_Load()
Dim Server As String
Set mclsAniform = New clsAnimated
LimpiarControles Me
CboProducto.Clear
Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
CboProducto.Text = UCase("Winko Mart")
CargaBodegasServ cboBodegas
Call STAT
'Call CargaBodegas(cboBodegas)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtFactura) <> "" Then
            
            sqls = "sp_FactBEVarios Null," & cboBodegas.ItemData(cboBodegas.ListIndex) & ",Null," & Val(txtFactura) & _
            " ," & Product & ",'BuscaCancela'"
            
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                
                If rsBD!bon_Fac_Status = 2 Then
                    MsgBox "Esta factura ya esta cancelada", vbInformation, "Factura cancelada"
                    txtFactura = ""
                    cmdCancelar.Enabled = False
                    Exit Sub
                End If
                
'                If gstrUsuario <> "1" Then  'solo yo
'                   If CDate(Format(rsBD!bon_fac_fechaemi, "dd/mm/yyyy")) < CDate(Format(rsBD!Fechadia, "dd/mm/yyyy")) Then
'                    MsgBox "No se puede cancelar la factura por la fecha, solo se pueden cancelar facturas del día de hoy", vbCritical, "Error en fecha"
'                    txtFactura = ""
'                    cmdCancelar.Enabled = False
'                    Exit Sub
'                   End If
'                End If
                
                cmdCancelar.Enabled = True
                lblCliente = Trim(rsBD!Nombre)
                lblImporte = Format(rsBD!bon_fac_valor, "$###,###,###.00")
                lblPedido = rsBD!bon_fac_pedido
                txtSerie = rsBD!bon_fac_serie
                Cte = rsBD!bon_fac_cliente
                
            Else
                MsgBox "La factura no existe!!", vbInformation, "Factura no existe"
                cmdCancelar.Enabled = False
            End If
                
        Else
            MsgBox "Primero capture el numero de factura a reimprimir", vbInformation, "Capture No. de factura"
            cmdCancelar.Enabled = False
        End If
    End If
End Sub

Private Sub txtFactura_LostFocus()
    txtFactura_KeyPress (13)
End Sub

Sub STAT()
    sql = "SELECT * FROM CONFIGBODEGAS WHERE Bodega=" & cboBodegas.ItemData(cboBodegas.ListIndex)
    Set rsBD = New ADODB.Recordset
    rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    If Not rsBD.EOF Then
       If Month(Date) <> Month(rsBD!FechaCierre) Then
          cmdCancelar.Enabled = False
       Else
          cmdCancelar.Enabled = True
       End If
    Else
       MsgBox "Error no hay fecha de cierre actual para esa Sucursal", vbCritical, "Sin fecha"
    End If
End Sub

