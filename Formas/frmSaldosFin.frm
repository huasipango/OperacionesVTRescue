VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmSaldosFin 
   Caption         =   "Insercion de Saldos Finales Manuales"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spSaldos 
      Height          =   2655
      Left            =   240
      OleObjectBlob   =   "frmSaldosFin.frx":0000
      TabIndex        =   10
      Top             =   2880
      Width           =   7815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Signo (+/-)"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Especifique datos "
      Height          =   2295
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   6495
      Begin MSComDlg.CommonDialog cmd 
         Left            =   240
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   4800
         Picture         =   "frmSaldosFin.frx":02FA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar"
         Top             =   1560
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   5400
         Picture         =   "frmSaldosFin.frx":03FC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   1560
         Width           =   450
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   4200
         Picture         =   "frmSaldosFin.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Subir desde archivo"
         Top             =   1560
         Width           =   450
      End
      Begin VB.ComboBox cbomes 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboano 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
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
         ItemData        =   "frmSaldosFin.frx":0B18
         Left            =   1800
         List            =   "frmSaldosFin.frx":0B22
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes:"
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
         Left            =   3500
         TabIndex        =   9
         Top             =   1005
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   1000
         Width           =   405
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
         TabIndex        =   7
         Top             =   420
         Width           =   1545
      End
   End
   Begin VB.Label lblAgrega 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmSaldosFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim elmes As Byte, FechaCorte As String, maximo As Integer
Dim sal As Boolean
Dim prod As Byte

Private Sub Check1_Click()
Dim x As Integer
If Check1.value = 1 And spSaldos.MaxRows > 0 Then
   For x = 1 To maximo
       spSaldos.Col = 5
       spSaldos.Row = x
       spSaldos.Text = Val(spSaldos.Text) * (-1)
   Next
End If
If Check1.value = 0 And spSaldos.MaxRows > 0 Then
   For x = 1 To maximo
       spSaldos.Col = 5
       spSaldos.Row = x
       spSaldos.Text = Val(spSaldos.Text) * (-1)
   Next
End If
End Sub

Private Sub cmdAbrir_Click()
Dim n As Integer
Dim arreglo() As String, clinea As String
Dim Cuenta As String, diferencia As Double, i As Integer
Dim sqls As String
On Error GoTo ERR:
   Call lee_mes
   sal = False
   busca_fecha
   If sal = True Then
      Exit Sub
   End If
   cmd.Filter = "Archivos delimitados por comas (*.csv)|*.csv"
   cmd.InitDir = "C:\"
   cmd.ShowOpen
   With spSaldos
   .Col = -1
   .Row = -1
   .Action = 12
   If cmd.Filename <> "" Then
      n = FreeFile
      Open cmd.Filename For Input Access Read As #n
           i = 0
           Do While Not EOF(n)
              Line Input #n, clinea
              arreglo = Split(clinea, ",")
              i = i + 1
              .Row = i
              .MaxRows = i
              .Col = 1
              .Text = FechaCorte
              .Col = 2
              Cuenta = arreglo(1)
              .Text = Cuenta
              '---Busca el nombre y tarjeta
              'prod = IIf(Product = 8, 6, Product)
              producto_cual
              sqls = "select top 1 convert(varchar(16),dbo.DesEncriptar(NoTarjeta)) as NoTarjeta,Nombre from tarjetasbe where nocuenta=" & Cuenta
              sqls = sqls & " and Status=1 and Tipo in ('T','RT') and Producto=" & Product
              Set rsBD = New ADODB.Recordset
              rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
              If Not rsBD.EOF Then
                 .Col = 3
                 .Text = rsBD!NoTarjeta
                 .Col = 4
                 .Text = rsBD!Nombre
              Else
                 MsgBox "La Cuenta " & Cuenta & " no existe o no es titular", vbCritical, "Error"
                 Exit Sub
              End If
              '----------------------------
              .Col = 5
              diferencia = arreglo(5)
              .Text = Val(diferencia)
           Loop
      Close #n
      Check1.Visible = True
      maximo = i
      lblAgrega.Visible = True
      lblAgrega.Caption = "Estan listos " & i & " registro(s) para agregarlos a saldosfinales"
   End If
   End With
Exit Sub
ERR:
   MsgBox ERR.Description, vbCritical, "Errores generados"
   Exit Sub
End Sub

Private Sub cmdGrabar_Click()
Dim sql As String, nreg As Integer
Dim fechac As String, Cuenta As String, Tarjeta As String, Nombre As String, importe As Double
On Error GoTo ERR:
If spSaldos.MaxRows > 0 Then
 If MsgBox("¿Esta seguro de insertar estos registros en saldos finales?", vbQuestion + vbYesNo + vbDefaultButton2, "Continuar?...") = vbYes Then
   With spSaldos
   i = 1
   Do While i <= .MaxRows
      .Row = i
      .Col = 1
      fechac = .Text
      .Col = 2
      Cuenta = .Text
      .Col = 3
      Tarjeta = .Text
      .Col = 4
      Nombre = .Text
      .Col = 5
      importe = Val(.Text)
      sql = "sp_SaldosFinalesBE '" & fechac & "'," & Cuenta & ",'" & Tarjeta & "','" & Nombre & "'," & importe & "," & Product
      cnxBD.Execute sql
      i = i + 1
   Loop
   End With
   MsgBox "Se han insertado correctamente " & i & " registro(s) en saldos finales del mes de " & cboMes.Text & " registro(s)", vbInformation, "Insercion correcta"
   Exit Sub
 Else
   Exit Sub
 End If
Else
  MsgBox "No hay resgistros que añadir a la tabla de saldos finales", vbCritical, "Sin registros"
  Exit Sub
End If
Exit Sub
ERR:
  MsgBox ERR.Description & " Posiblemente ya exista el registro que desea insertar en el mes señalado", vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
   Set mclsAniform = New clsAnimated
   cboProducto.Clear
   Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
   cboProducto.Text = UCase("Despensa Total")
   Call carga_meses
   cboano.Text = Year(Date)
   cboMes.Text = "Enero"
   'Check1.Visible = False
   InicializaForma
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
     'CargaSpread
  End If
End Sub

Sub InicializaForma()
    lblAgrega.Visible = False
    spSaldos.MaxRows = 0
    Check1.Visible = False
    With spSaldos
      .Col = -1
      .Row = -1
      .Action = 12
   End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Sub carga_meses()
Dim x As Byte
    For x = 0 To 3
        cboano.AddItem Year(Date) - x
    Next
    cboMes.AddItem "Enero"
    cboMes.AddItem "Febrero"
    cboMes.AddItem "Marzo"
    cboMes.AddItem "Abril"
    cboMes.AddItem "Mayo"
    cboMes.AddItem "Junio"
    cboMes.AddItem "Julio"
    cboMes.AddItem "Agosto"
    cboMes.AddItem "Septiembre"
    cboMes.AddItem "Octubre"
    cboMes.AddItem "Noviembre"
    cboMes.AddItem "Diciembre"
End Sub

Sub lee_mes()
    Select Case cboMes.Text
    Case "Enero"
          elmes = 1
    Case "Febrero"
          elmes = 2
    Case "Marzo"
          elmes = 3
    Case "Abril"
          elmes = 4
    Case "Mayo"
          elmes = 5
    Case "Junio"
          elmes = 6
    Case "Julio"
          elmes = 7
    Case "Agosto"
          elmes = 8
    Case "Septiembre"
          elmes = 9
    Case "Octubre"
          elmes = 10
    Case "Noviembre"
          elmes = 11
    Case "Diciembre"
          elmes = 12
    End Select
End Sub

Sub busca_fecha()
Dim sql As String
Dim fec As String, fechaf As String
fec = Format(elmes, "00") & "/01/" & cboano.Text
sql = "sp_SaldosFin_Varios '" & fec & "'"
cnxBD.Execute sql
Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
fechaf = rsBD!Fecha
fechaf = Format(fechaf, "mm/dd/yyyy")
'prod = IIf(Product = 8, 6, Product)
producto_cual
sql = "Select top 1 fechacorte from saldosfinalesbe"
sql = sql & " where (fechacorte>='" & fechaf & "' and fechacorte<='" & fechaf & " 23:59:00')"
sql = sql & " and Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
If Not rsBD.EOF Then
   FechaCorte = Format(rsBD!FechaCorte, "yyyy-mm-dd hh:mm:ss")
Else
   MsgBox "No existen registros de saldos finales en el periodo que desea insertar", vbCritical, "No debe continuar"
   sal = True
   Exit Sub
End If
End Sub
