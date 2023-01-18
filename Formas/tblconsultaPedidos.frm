VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form tblConsultaPedidos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Pedidos de Bono Electrónico"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   Icon            =   "tblconsultaPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   120
      TabIndex        =   27
      Top             =   480
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
         ItemData        =   "tblconsultaPedidos.frx":1CFA
         Left            =   1800
         List            =   "tblconsultaPedidos.frx":1D04
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   160
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
         TabIndex        =   28
         Top             =   230
         Width           =   1545
      End
   End
   Begin VB.Frame fraPedido 
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   12615
      Begin VB.OptionButton optSta 
         Caption         =   "Status"
         Height          =   195
         Left            =   9960
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFac 
         Caption         =   "Factura"
         Height          =   195
         Left            =   6960
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optPed 
         Caption         =   "Pedido"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optCte 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid tblPedido 
         Height          =   3255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   3
         BackColor       =   16777215
         ForeColor       =   0
         BackColorBkg    =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblPie 
         Alignment       =   1  'Right Justify
         Caption         =   "T O T A L E S  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   7320
         TabIndex        =   19
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   11880
         X2              =   9480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblCant 
         Alignment       =   1  'Right Justify
         Caption         =   " 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   9360
         TabIndex        =   20
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   " 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   10680
         TabIndex        =   21
         Top             =   3960
         Width           =   1215
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   12615
      Begin VB.CheckBox chkTodos 
         Caption         =   "Ver Pedidos de Desp y Gasolina."
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
         Left            =   8880
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtAAM 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   8
         Top             =   630
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarCte 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2160
         Picture         =   "tblconsultaPedidos.frx":1D16
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   585
         Width           =   375
      End
      Begin VB.TextBox txtClte 
         DataField       =   "bon_zon_numero"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.ComboBox cboSucursal 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6000
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   18
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "tblconsultaPedidos.frx":1E18
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "tblconsultaPedidos.frx":1F2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "tblconsultaPedidos.frx":203C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   4560
         TabIndex        =   23
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   345
         Left            =   7200
         TabIndex        =   24
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
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
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
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
         Left            =   6840
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblAM 
         Caption         =   "Año Mes :"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   660
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre Cliente"
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   630
         Width           =   5175
      End
      Begin VB.Label lblClte 
         Caption         =   "Cliente : "
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
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   855
      End
      Begin VB.Label lblSucursal 
         Caption         =   "Sucursal :"
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
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   714
      BandCount       =   2
      _CBWidth        =   11370
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinWidth1       =   1305
      MinHeight1      =   345
      Width1          =   4650
      NewRow1         =   0   'False
      Child2          =   "Toolbar2"
      MinWidth2       =   495
      MinHeight2      =   345
      Width2          =   1455
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   345
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   609
         ButtonWidth     =   661
         ButtonHeight    =   609
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Filtrar"
               Object.ToolTipText     =   "Filtrar Denominación"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   345
         Left            =   10020
         TabIndex        =   3
         Top             =   30
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         ButtonWidth     =   661
         ButtonHeight    =   609
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Salir"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "tblConsultaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub CboSucursal_Click()
   Call doLlenaTablaPed
End Sub

Private Sub chkTodos_Click()
        Call doLlenaTablaPed
End Sub

Private Sub cmdBuscarCte_Click()
Dim frmConsulta As New frmBusca_Cliente
    
   frmConsulta.Bodega = CboSucursal.ItemData(CboSucursal.ListIndex)
   frmConsulta.Show vbModal
    
   If frmConsulta.cliente >= 0 Then
      txtClte.Text = frmConsulta.cliente
      lblNombre.Caption = frmConsulta.Nombre
   End If
   Set frmConsulta = Nothing
   MsgBar "", False

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
Dim strsql As String, rstTmp As ADODB.Recordset
Dim mes, fechaini As Date, fechafin As Date, ano

   optPed.value = True  'Ordenado por Pedido
   lblNombre.Caption = ""
   If inicia_consulped <> 0 Then
      Call doLlenaTablaPed
   End If
   
   strsql = "SELECT ltrim(rtrim(ltrim(rtrim(str(datepart(yy,getdate()))))+ right('00' + (ltrim(rtrim(str( datepart(mm,getdate()))))),2) )) as AnioMes "
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly
   txtAAM.Text = rstTmp!AnioMes
   rstTmp.Close
   Set rstTmp = Nothing
   ano = Year(Date)
   mes = Month(Date) + 1
   If mes > 12 Then
      mes = 1
      ano = ano + 1
   End If
   fechaini = "01/" & Month(Date) & "/" & Year(Date)
   fechafin = ("01/" & mes & "/" & ano)
   fechafin = fechafin - 1
   mskFechaIni = Format(fechaini, "dd/mm/yyyy")
   mskFechaFin = Format(fechafin, "dd/mm/yyyy")
   
   Call doCreatePed
End Sub


Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
 InicializaForma
End Sub

Private Sub Form_Load()
Dim strsql As String, rstTmp As ADODB.Recordset
Dim mes, fechaini As Date, fechafin As Date, ano
   Set mclsAniform = New clsAnimated
   TipoBusqueda = "ClienteBE"
   'Call CargaBodegas(cboSucursal)
   'Call CboPosiciona(cboSucursal, gnBodega)

   CboProducto.Clear
   Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
   CboProducto.Text = UCase("Winko Mart")
   If user_master = False Then
       CargaBodegasS2 CboSucursal 'CargaBodegasS2
   Else
       CargaBodegas CboSucursal
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
  inicia_consulped = 0
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case LCase(Button.Key)
   Case "filtrar"
      inicia_consulped = 1
      Call doLlenaTablaPed
   End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case LCase(Button.Key)
   Case "salir"
      Unload Me
   End Select
End Sub

Private Sub txtClte_KeyPress(KeyAscii As Integer)
   If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii >= vbKeySpace Then
      KeyAscii = 0
   End If
   entertab KeyAscii
End Sub

Private Sub txtClte_LostFocus()
Dim strsql As String
Dim rstClte As ADODB.Recordset
Dim strNombre As String

   If Len(txtClte.Text) > 0 Then
      strsql = "SELECT Nombre, Bodega " & _
               "FROM Clientes " & _
               "WHERE Cliente = " & txtClte.Text
      Set rstClte = New ADODB.Recordset
      rstClte.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly
      If rstClte.EOF = False Then
         strNombre = rstClte!Nombre & ""
         strNombre = Trim(Mid(strNombre, 1, 30)) & " " & Trim(Mid(strNombre, 31, 30))
         lblNombre.Caption = strNombre
         
         If rstClte!Bodega <> CboSucursal.ItemData(CboSucursal.ListIndex) Then
            MsgBox " Cliente no Pertenece a Sucursal Seleccionada", vbOKOnly, "Verifique"
            txtClte.SetFocus
            txtClte.Text = ""
            lblNombre.Caption = ""
         End If
      Else
         MsgBox "No Existe Cliente Capturado", vbOKOnly, "Verifique"
         txtClte.SetFocus
         txtClte.Text = ""
      End If
      rstClte.Close
      Set rstClte = Nothing
   Else
      lblNombre.Caption = "TODOS LOS CLIENTES"
   End If
   MsgBar "", False
End Sub

Private Sub optCte_Click()
   Call doLlenaTablaPed
End Sub

Private Sub optFac_Click()
   Call doLlenaTablaPed
End Sub

Private Sub optPed_Click()
   Call doLlenaTablaPed
End Sub

Private Sub optSta_Click()
   Call doLlenaTablaPed
End Sub

Private Sub doCreatePed()

   tblPedido.Clear
   tblPedido.Rows = 2
   tblPedido.Cols = 14
   
   tblPedido.Row = 0
   tblPedido.Col = 1
   tblPedido.Text = "Pedido"
   tblPedido.Col = 2
   tblPedido.Text = "Factura"
   tblPedido.Col = 3
   tblPedido.Text = "Fec.Ent."
   tblPedido.Col = 4
   tblPedido.Text = "Clave"
   tblPedido.Col = 5
   tblPedido.Text = "Cliente"
   tblPedido.Col = 6
   tblPedido.Text = "Producto"
   tblPedido.Col = 7
   tblPedido.Text = "Valor Ped."
   tblPedido.Col = 8
   tblPedido.Text = "Valor Fac."
   tblPedido.Col = 9
   tblPedido.Text = "Tipo Ped."
   tblPedido.Col = 10
   tblPedido.Text = "Pasan"
   tblPedido.Col = 11
   tblPedido.Text = "Status"
   tblPedido.Col = 12
   tblPedido.Text = "IBR"
   tblPedido.Col = 13
   tblPedido.Text = "SAC"

   tblPedido.ColWidth(0) = 150
   tblPedido.ColWidth(1) = 650
   tblPedido.ColWidth(2) = 650
   tblPedido.ColWidth(3) = 1000
   tblPedido.ColWidth(4) = 650
   tblPedido.ColWidth(5) = 2000
   tblPedido.ColWidth(6) = 1000
   tblPedido.ColWidth(7) = 1100
   tblPedido.ColWidth(8) = 1100
   tblPedido.ColWidth(9) = 850
   tblPedido.ColWidth(10) = 850
   tblPedido.ColWidth(11) = 900
   tblPedido.ColWidth(12) = 600
   tblPedido.ColWidth(13) = 600
   
   tblPedido.ColAlignment(7) = 6
   tblPedido.ColAlignment(8) = 6

End Sub

Private Sub doLlenaTablaPed()
Dim strsql As String, sOrderBy As String, sWhere As String, rstTmp As ADODB.Recordset
Dim i As Integer, sac As String
Dim nCant As Long, nTotal As Double
If inicia_consulped <> 0 Then
   
   If txtAAM.Text = "" Then
'      MsgBox "Capture un Mes "
      Exit Sub
   End If
   
   Call doCreatePed
   
   lblCant.Caption = ""
   lblTotal.Caption = ""

   strsql = "SELECT a.bon_ped_numero pedido,ISNULL(bon_ped_viaent,0) bon_ped_viaent," & _
                     "ISNULL((SELECT Max(ISNULL(J.BON_FAC_NUMERO,0)) FROM BON_FACTURA J Where J.BON_FAC_SUCURSAL = a.bon_ped_sucursal And year(BON_FAC_FECHAEMI)>=2011 And J.BON_FAC_PEDIDO = a.bon_ped_numero), 0) FACTURA, " & _
                     "convert(varchar,a.bon_ped_fechaent,101) fec_ped, " & _
                     "a.bon_ped_cliente cte, " & _
                     "ltrim(rtrim(c.nombre)) nom_cte, " & _
                     "case a.bon_ped_producto when 1 then 'Despensa' when 2 then 'Comb Nal' when 3 then 'Comb Control' end nom_pro, " & _
                     "ISNULL(a.bon_ped_valor,0) valor, " & _
                     "ISNULL((SELECT Max(ISNULL(w.BON_FAC_VALOR,0)) FROM BON_FACTURA w Where w.BON_FAC_SUCURSAL = a.bon_ped_sucursal And year(BON_FAC_FECHAEMI)>=2011 And w.BON_FAC_PEDIDO = a.bon_ped_numero), 0) fac_valor, " & _
                     "Case a.bon_ped_tipo when 9 then 'Granel' when 11 then 'N.Consumo' when 12 then 'Ensobretado' when 13 then 'Stock' else '******' end TIPOPED, " & _
                     "CASE a.bon_ped_statpas WHEN 0 THEN 'Se Envian' ELSE 'Pasan' END Pasan, " & _
                     "Case a.bon_ped_status when 0 then 'Capturado' when 1 then 'Facturado' when 2 then 'Cancelado' when 4 then 'Impreso' when 5 then 'Re-Impreso' when 10 then 'Liberado' when 12 then 'Dispersado' else '******' end status, " & _
                     "case a.bon_ped_usuario when 80 then 'IBR' ELSE '' END IBR " & _
            "FROM bon_pedido a WITH (NOLOCK), clientes c WITH (NOLOCK)" & _
            "WHERE a.bon_ped_sucursal = " & CboSucursal.ItemData(CboSucursal.ListIndex) & " " & _
            "AND a.bon_ped_fechaent between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'" & _
            "AND a.bon_ped_sucursal = c.Bodega " & _
            "AND a.bon_ped_cliente = c.cliente"

   If txtClte.Text <> "" Then
      sWhere = " AND a.bon_ped_cliente = " & txtClte.Text
   End If
   
   If chkTodos.value = False Then
      sWhere = sWhere & " AND a.bon_ped_producto =" & Product
   End If
   
   
   If optPed.value Then
      sOrderBy = " ORDER BY a.bon_ped_numero"
   ElseIf optFac.value Then
      sOrderBy = " ORDER BY Factura"
   ElseIf optCte.value Then
      sOrderBy = " ORDER BY a.bon_ped_cliente"
   ElseIf optSta.value Then
      sOrderBy = " ORDER BY Status"
   Else
      sOrderBy = " ORDER BY a.bon_ped_cliente"
   End If
   
   strsql = strsql & sWhere & sOrderBy

   
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockOptimistic, adCmdText
   i = 1
   Do Until rstTmp.EOF
      With tblPedido
         .Rows = .Rows + 1
         .TextMatrix(i, 1) = rstTmp!pedido
         .TextMatrix(i, 2) = rstTmp!Factura & ""
         .TextMatrix(i, 3) = rstTmp!fec_ped
         .TextMatrix(i, 4) = Pad(rstTmp!Cte & "", 6, "0", "L")
         .TextMatrix(i, 5) = rstTmp!nom_cte
         .TextMatrix(i, 6) = rstTmp!nom_pro
         .TextMatrix(i, 7) = Format(rstTmp!valor, "##,###,###.00")
         .TextMatrix(i, 8) = Format(rstTmp!fac_valor, "##,###,###.00")
         .TextMatrix(i, 9) = rstTmp!TipoPed
         .TextMatrix(i, 10) = rstTmp!Pasan
         .TextMatrix(i, 11) = rstTmp!Status
         .TextMatrix(i, 12) = rstTmp!IBR
         sac = IIf(rstTmp!bon_ped_viaent = 0, "", "X")
         .TextMatrix(i, 13) = sac
         nCant = nCant + 1
         nTotal = nTotal + rstTmp!valor
      End With
      i = i + 1
      rstTmp.MoveNext
      tblPedido.Rows = i
   Loop
   With tblPedido
      .ForeColorSel = &H8000000E
      .BackColorSel = &H8000000D
   End With
   rstTmp.Close
   Set rstTmp = Nothing
    
   lblTotal.Caption = CStr(Format(CDbl(nTotal), "##,###,##0.00"))
   lblCant.Caption = CStr(Format(CDbl(nCant), "###,##0"))
    
 End If
End Sub
