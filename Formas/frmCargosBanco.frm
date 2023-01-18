VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCargosBanco 
   Caption         =   "Captura de Cargos del Banco"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   ScaleHeight     =   6480
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   240
      TabIndex        =   12
      Top             =   360
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
         ItemData        =   "frmCargosBanco.frx":0000
         Left            =   1800
         List            =   "frmCargosBanco.frx":000A
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
         TabIndex        =   13
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo (mm/dd/yyyy) "
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   7215
      Begin FPSpread.vaSpread spdCargos 
         Height          =   3495
         Left            =   360
         OleObjectBlob   =   "frmCargosBanco.frx":001C
         TabIndex        =   6
         Top             =   960
         Width           =   6375
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   6240
         Picture         =   "frmCargosBanco.frx":0376
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   5040
         Picture         =   "frmCargosBanco.frx":0478
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Subir desde archivo"
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   5640
         Picture         =   "frmCargosBanco.frx":057A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grabar"
         Top             =   360
         Width           =   450
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   960
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
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
         Left            =   3240
         TabIndex        =   3
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "TOTAL:"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   4605
         Width           =   855
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   435
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmCargosBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte

Private Sub cmdAbrir_Click()
    CargaDatos
End Sub

Private Sub cmdGrabar_Click()
Screen.MousePointer = 11
i = 0
With spdCargos
    For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    FC = Format(.Text, "mm/dd/yyyy")
    .Col = 2
    total = CDbl(.Text)
    .Col = 3
    FP = Format(.Text, "mm/dd/yyyy")
    .Col = 4
    If .Text = "" Then
        TotalBanco = 0
    Else
        TotalBanco = CDbl(.Text)
    End If
    
    If TotalBanco <> 0 Then
        'prod = IIf(Product = 8, 6, Product)
        producto_cual
        sqls = "Exec sp_CargoBancosIns @FechaConciliacion = '" & FC & "'" & _
               ",@Total = " & total & " ,@FechaCargo = '" & IIf(FP = "", Null, FP) & "'" & _
               ",@TotalBanco = " & TotalBanco & ",@Producto=" & Product
               
        cnxBD.Execute sqls, intRegistros
    End If
    
    Next i
End With
Screen.MousePointer = 1
MsgBox "Cargos Actualizados!", vbInformation, "Cargos actualizados"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
     CargaDatos
  End If
  'Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
End Sub
Sub InicializaForma()
 mskFechaIni = Format((Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Trim(Str(Year(Date)))), "mm/dd/yyyy")
    mskFechaFin = Format((FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), Year(Date))), "mm/dd/yyyy")
    With spdCargos
        .Col = -1
        .Row = -1
        .Action = 12
    End With
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Sub valida_fecha_cierre()
  sqls = "SELECT Fecha_CierreBE FROM INFORMACION_GENERAL"
  Set rsBD = New ADODB.Recordset
  rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
  If Format(rsBD!Fecha_CierreBE, "mm/dd/yyyy") >= Format(Date - 12, "mm/dd/yyyy") Then
     sqls = "UPDATE INFORMACION_GENERAL SET Fecha_CierreBE='" & Format(Date - 20, "mm/dd/yyyy") & "'"
     cnxBD.Execute sqls
  End If
End Sub
Private Sub Form_Load()
    Set mclsAniform = New clsAnimated

    mskFechaIni = Format((Format(IIf(Month(Date) = 1, 12, Month(Date)), "00") + "/01/" + Trim(Str(Year(Date)))), "mm/dd/yyyy")
    mskFechaFin = Format((FechaFinMes(IIf(Month(Date) = 1, 12, Month(Date)), Year(Date))), "mm/dd/yyyy")
    With spdCargos
        .Col = -1
        .Row = -1
        .Action = 12
    End With
    cboProducto.Clear
    Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
    cboProducto.Text = UCase("Despensa Total")
    valida_fecha_cierre
End Sub

Sub CargaDatos()
Dim total As Double

If Product = 2 Then
   Call busca_inserta
End If

With spdCargos

'prod = IIf(Product = 8, 6, Product)
producto_cual

sqls = " select Fechaconciliacion FC, Importe Total,FechaCargo FP, Importebanco TotalBanco" & _
       " From cargosBanco" & _
       " where fechaconciliacion between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'" & _
       " and Producto=" & Product & _
       " Union" & _
       " select fechaconciliacion, sum(importe), Fechapago, 0" & _
       " From liquidacionesbe" & _
       " where  fechaconciliacion between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'" & _
       " and comercio in (select comercio from comercios where tipocomercio = 1 and Producto=" & Product & ")" & _
       " and fechaconciliacion not in (select fechaconciliacion from cargosbanco Where Producto=" & Product & ")" & _
       " And Producto=" & Product & _
       " group by Fechaconciliacion, Fechapago" & _
       " order by 1,3"

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

total = 0
i = 0
Do While Not rsBD.EOF
    i = i + 1
    .MaxRows = i
    .Row = i
    .Col = 1
    .Text = Format(rsBD!FC, "DD/MM/YY")
    .Col = 2
    .Text = Format(CDbl(rsBD!total), "###,###,000.00")
    total = total + CDbl(rsBD!total)
    .Col = 3
    .Text = Format(rsBD!FP, "DD/MM/YY")
    .Col = 4
    .Text = Format(CDbl(rsBD!TotalBanco), "###,###,000.00")
    rsBD.MoveNext
Loop

txtTotal.Text = Format(total, "###,###,000.00")

End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Sub busca_inserta() 'exclusivo para pago gas los inserta automatico
Dim sq As String
sqls = "select fechaconciliacion,2,sum(importe),fechaconciliacion,sum(importe) From liquidacionesbe "
sqls = sqls & " where  fechaconciliacion between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'"
sqls = sqls & "and comercio in (select comercio from comercios where tipocomercio = 1 and Producto=2)"
sqls = sqls & " and fechaconciliacion not in (select fechaconciliacion from cargosbanco Where Producto=2)"
sqls = sqls & "And Producto=2 group by Fechaconciliacion, Fechapago order by 1,3"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsBD.EOF Then
   Exit Sub
Else
   sq = "INSERT INTO CargosBanco "
   sqls = sq & sqls
   cnxBD.Execute sqls, intRegistros
End If
End Sub
