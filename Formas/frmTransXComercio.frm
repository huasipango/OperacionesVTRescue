VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransXComercio 
   Caption         =   "Transacciones por comercio"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   10695
      Begin FPSpread.vaSpread spdTrans 
         Height          =   3135
         Left            =   120
         OleObjectBlob   =   "frmTransXComercio.frx":0000
         TabIndex        =   15
         Top             =   240
         Width           =   10335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   10695
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   2
         Left            =   7320
         TabIndex        =   18
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   17
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtImpTot 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   2
         Left            =   8280
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtImpTot 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtImpTot 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Pos-Débito"
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
         Left            =   7320
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Débito"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Interredes"
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
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   10695
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
         ItemData        =   "frmTransXComercio.frx":049E
         Left            =   4560
         List            =   "frmTransXComercio.frx":04A8
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3120
      End
      Begin VB.CommandButton cmdSubir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   8760
         Picture         =   "frmTransXComercio.frx":04BA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cargar Archivo"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H00C0C0C0&
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
         Height          =   450
         Left            =   9360
         Picture         =   "frmTransXComercio.frx":0ADC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Actualizar"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   9960
         Picture         =   "frmTransXComercio.frx":0BDE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   450
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de transaccion:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog cmnAbrir 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo de Comercio diferente en el Catálogo de Comercios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   6240
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   6240
      Width           =   375
   End
End
Attribute VB_Name = "frmTransXComercio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub cmdActualizar_Click()
   
If spdTrans.MaxRows <> 0 Then
   Screen.MousePointer = 11
   
   GrabaTrans

   Screen.MousePointer = 1
End If
End Sub
Sub GrabaTransDet()
Dim i, tipocom, comercio
Dim numvta, impvta, numdev, impdev, Total As Double
On Error GoTo ERR:
With spdTrans
   For i = 1 To .MaxRows
      .Row = i
      .Col = 1
      tipocom = .TypeComboBoxCurSel
      .Col = 2
      comercio = .Text
      .Col = 4
      numvta = CInt(.Text)
      .Col = 5
      impvta = CDbl(.Text)
      .Col = 6
      numdev = CInt(.Text)
      .Col = 7
      impdev = CDbl(.Text)
      
      
      If CDbl(impvta) <> 0 Or CDbl(impdev) <> 0 Then
         Total = impvta - impdev
      
         sqls = " EXEC sp_Insupd_trans_x_comdet "
         sqls = sqls & vbCr & "  @Fecha       = '" & Format(mskFecha, "mm/dd/yyyy") & "'"
         sqls = sqls & vbCr & ", @TipoComercio      = " & tipocom
         sqls = sqls & vbCr & ", @comercio      = '" & Format(comercio, "0000000") & "'"
         sqls = sqls & vbCr & ", @Cantvta       =    " & CInt(numvta) & ""
         sqls = sqls & vbCr & ", @ImporteVta    = " & CDbl(impvta)
         sqls = sqls & vbCr & ", @Cantdev       =    " & CInt(numdev) & ""
         sqls = sqls & vbCr & ", @Importedev    = " & CDbl(impdev)
         sqls = sqls & vbCr & ", @status    = 0"
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         sqls = sqls & vbCr & ", @Producto=" & Product
         cnxBD.Execute sqls, intRegistros
      End If
   Next i
End With
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores..."
  Exit Sub
End Sub

Sub GrabaTrans()
On Error GoTo ERR:
'prod = IIf(Product = 8, 6, Product)
producto_cual
sqls = "select * from trans_x_com where fecha = '" & Format(mskFecha, "mm/dd/yyyy") & "'"
sqls = sqls & " and Producto=" & Product
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    MsgBox "Su archivo no puede ser procesado , ya que existe un archivo de liquidaciones con esa fecha.", vbCritical, "Error en archivo"
    Exit Sub
End If

For i = 0 To 2
sqls = " EXEC sp_Insupd_trans_x_com "
         sqls = sqls & vbCr & "  @Fecha       = '" & Format(mskFecha, "mm/dd/yyyy") & "'"
         sqls = sqls & vbCr & ", @TipoComercio      = " & i
         sqls = sqls & vbCr & ", @Cantidad        =    " & txtNum(i)
         sqls = sqls & vbCr & ", @Importe    = " & CDbl(txtImpTot(i))
         'prod = IIf(Product = 8, 6, Product)
         producto_cual
         sqls = sqls & vbCr & ", @Producto=" & Product
         cnxBD.Execute sqls, intRegistros
Next i

GrabaTransDet
  
MsgBox "Datos actualizados!!!", vbInformation, "Datos actualizados"
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores..."
  Exit Sub
End Sub
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub cmdsubir_Click()
    
    For i = 0 To 2
      txtImpTot(i) = 0
      txtNum(i) = 0
   Next i
   
   Subearchivo
   
End Sub

Private Function Subearchivo() As String
Dim nArchivo, clinea As String, i As Long
Dim ImpVentas, impdev, ImpTotal As Double
Dim TotInter, TotDeb, TotPos As Double
Dim NoInter, NoDeb, NoPos, tipocom As Integer
Dim nombre As String
Dim ArrVal() As String
Dim ArrImp() As Double
Dim ArrCant() As Integer

With spdTrans

    Subearchivo = ""
    .Col = -1
    .Row = -1
    .Action = 12
        
    On Error GoTo ErrorImport
    cmnAbrir.ShowOpen
    If cmnAbrir.Filename <> "" Then
        nArchivo = FreeFile
        Open cmnAbrir.Filename For Input Access Read As #nArchivo
        i = 0
        Do While Not EOF(nArchivo)
            Line Input #nArchivo, clinea
            ArrVal = Split(clinea, ",")
            i = i + 1
            .Row = i
            .MaxRows = i
            .Col = 1
               
            .TypeComboBoxCurSel = ArrVal(0)
            tipocom = ArrVal(0)
            
            .Col = 2
            .Text = Format(ArrVal(1), "0000000")
            'prod = IIf(Product = 8, 6, Product)
            producto_cual
            sqls = "select descripcion, tipocomercio from comercios where comercio = '" & Format(Trim(ArrVal(1)), "0000000") & "'"
            sqls = sqls & " AND Producto=" & Product
                    Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If Not rsBD.EOF Then
                 If rsBD!tipocomercio <> ArrVal(0) Then 'viene el tipo de comercio diferente al catálogio de comercios
                     .Col = 1
                     .BackColor = RGB(250, 250, 0)
                 End If
                 
                .Col = 3
                .Text = rsBD!Descripcion
                .Col = 4
                .Text = CInt(ArrVal(2))
                .Col = 5
                .Text = CDbl(ArrVal(3))
                .Col = 6
                .Text = CInt(ArrVal(4))
                .Col = 7
                .Text = CDbl(ArrVal(5))
                .Col = 8
                .Text = CDbl(ArrVal(6))
                
                
                txtImpTot(tipocom) = Val(txtImpTot(tipocom)) + CDbl(ArrVal(6))
                txtNum(tipocom) = Val(txtNum(tipocom)) + CInt(ArrVal(2))
                
            Else
                .Col = 3
                .Text = "Grupo no existe en el catalogo"
                .Col = 4
                .Text = 0
                .Col = 5
                .Text = 0
                .Col = 6
                .Text = 0
                .Col = 7
                .Text = 0
                .Col = 8
                .Text = 0
                .BackColor = RGB(255, 0, 0)

            End If
          DoEvents
        Loop
        Close #nArchivo
        
        Subearchivo = True
        cmnAbrir.Filename = ""
        Exit Function
    Else
        Exit Function
    End If
    
End With
ErrorImport:
    Beep
  
    MsgBox "Hubo un error al actualizar! Favor de avisar a sistemas! Error: " & ERR.Number & vbCrLf & ERR.Description, vbCritical + vbOKOnly, Me.Caption
    'aniAvi.Close
    'fraAni.Visible = False
    Resume Next
End Function

Private Sub Form_Load()
   Set mclsAniform = New clsAnimated

   mskFecha = Format(Date, "mm/dd/yyyy")
   For i = 0 To 2
      txtImpTot(i) = 0
      txtNum(i) = 0
   Next i
   spdTrans.MaxRows = 0
   cboProducto.Clear
   Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
   cboProducto.Text = UCase("Despensa Total")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
