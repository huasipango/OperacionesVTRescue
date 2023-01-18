VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmEnvios_Caratulas 
   Caption         =   "Datos de envio"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   Begin FPSpread.vaSpread spdPlazas 
      Height          =   1935
      Left            =   9000
      OleObjectBlob   =   "frmEnvios_Caratulas.frx":0000
      TabIndex        =   41
      Top             =   1440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdDeshacer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reiniciar"
      CausesValidation=   0   'False
      Height          =   1050
      Left            =   11160
      Picture         =   "frmEnvios_Caratulas.frx":02C7
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Grabar"
      Top             =   7200
      Width           =   780
   End
   Begin VB.TextBox lblPlaza 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      MaxLength       =   100
      TabIndex        =   40
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtPlaza 
      Height          =   375
      Left            =   10200
      MaxLength       =   8
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscaPlaza 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9700
      Picture         =   "frmEnvios_Caratulas.frx":48B51
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdImprime 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      CausesValidation=   0   'False
      Height          =   1050
      Left            =   10200
      Picture         =   "frmEnvios_Caratulas.frx":48C53
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Grabar"
      Top             =   7200
      Width           =   780
   End
   Begin VB.CommandButton cmdBuscarC 
      BackColor       =   &H00C0C0C0&
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1080
      Picture         =   "frmEnvios_Caratulas.frx":914DD
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   1050
      Left            =   12120
      Picture         =   "frmEnvios_Caratulas.frx":915DF
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   7200
      Width           =   780
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grabar"
      CausesValidation=   0   'False
      Height          =   1050
      Left            =   9240
      Picture         =   "frmEnvios_Caratulas.frx":D9E69
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Grabar"
      Top             =   7200
      Width           =   780
   End
   Begin VB.TextBox txtObs 
      Height          =   3975
      Left            =   9240
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox txtruta 
      Height          =   375
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   12
      Top             =   8520
      Width           =   855
   End
   Begin VB.TextBox txtNomRuta 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   33
      Top             =   8520
      Width           =   5895
   End
   Begin VB.TextBox txtmail3 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   11
      Top             =   7920
      Width           =   5655
   End
   Begin VB.TextBox txtcontac3 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   10
      Top             =   7320
      Width           =   5655
   End
   Begin VB.TextBox txtmail2 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   9
      Top             =   6720
      Width           =   5655
   End
   Begin VB.TextBox txtcontac2 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   8
      Top             =   6120
      Width           =   5655
   End
   Begin VB.TextBox txtmail1 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   7
      Top             =   5520
      Width           =   5655
   End
   Begin VB.TextBox txtcontac1 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   6
      Top             =   4920
      Width           =   5655
   End
   Begin VB.TextBox txttel 
      Height          =   375
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   5
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox txtestado 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   25
      Top             =   3720
      Width           =   6855
   End
   Begin VB.TextBox txtPoblacion 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   23
      Top             =   3165
      Width           =   6855
   End
   Begin VB.ComboBox cboColonia 
      Height          =   315
      ItemData        =   "frmEnvios_Caratulas.frx":1226F3
      Left            =   1560
      List            =   "frmEnvios_Caratulas.frx":1226F5
      TabIndex        =   4
      Top             =   2640
      Width           =   6855
   End
   Begin VB.TextBox txtCP 
      Height          =   375
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtcalle 
      Height          =   375
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      MaxLength       =   100
      TabIndex        =   18
      Top             =   840
      Width           =   5895
   End
   Begin VB.TextBox txtcliente 
      Height          =   375
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Plaza:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   9000
      TabIndex        =   39
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   9240
      TabIndex        =   35
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      X1              =   8880
      X2              =   8880
      Y1              =   840
      Y2              =   8880
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Ruta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   34
      Top             =   8640
      Width           =   405
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "E-Mail 3° Cont:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   32
      Top             =   8040
      Width           =   1155
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "3° Contacto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Top             =   7440
      Width           =   975
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "E-Mail 2° Cont:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   30
      Top             =   6840
      Width           =   1155
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "2° Contacto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   29
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "E-Mail 1° Cont:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   28
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "1° Contacto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   27
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Telefono:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Width           =   780
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Codigo Postal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Poblacion:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   21
      Top             =   3285
      Width           =   840
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Colonia:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   660
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Calle:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   11400
      Picture         =   "frmEnvios_Caratulas.frx":1226F7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1425
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Capture y Valide los datos donde el Cliente recibirá las Tarjetas."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmEnvios_Caratulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim boded As Byte, continua As Boolean
Dim asi As Boolean

Private Sub cmdBuscaPlaza_Click()
Dim i As Integer
If Val(txtcliente) = 0 Then
   Exit Sub
End If
    
    spdPlazas.Visible = True
    sqls = "select * from plazasbe"
    If Trim(txtcliente) <> "" Then
        sqls = sqls & " where cliente = " & Val(txtcliente) & ""
    End If
    sqls = sqls & " order by cliente, plaza"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    i = 0
    With spdPlazas
    .Col = -1
    .Row = -1
    .Action = 12
    
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Plaza
        .Col = 3
        .Text = rsBD!Descripcion
        
        rsBD.MoveNext
        
    Loop

    rsBD.Close
    Set rsBD = Nothing
    End With
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    TipoBusqueda = "ClienteBE"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtcliente = frmConsulta.cliente
       Call BUSCA_Cliente
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub

Private Sub cmdDeshacer_Click()
  Call iniciar
  txtcliente.SetFocus
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ERR:
  Call valida_datos
  If continua = True Then
     sqls = "sp_Captura_CEnviosBE @Cliente=" & Val(txtcliente.Text)
     sqls = sqls & " ,@Calle='" & Trim(txtcalle) & "'"
     sqls = sqls & " ,@CP='" & Format(txtcp, "00000") & "'"
     sqls = sqls & " ,@Colonia='" & Trim(cboColonia.Text) & "'"
     sqls = sqls & " ,@Poblacion='" & Trim(txtPoblacion.Text) & "'"
     sqls = sqls & " ,@Estado='" & Trim(txtestado.Text) & "'"
     sqls = sqls & " ,@Telefono='" & Trim(txttel.Text) & "'"
     sqls = sqls & " ,@Cont1='" & Trim(txtcontac1.Text) & "'"
     sqls = sqls & " ,@Mail1='" & Trim(txtmail1.Text) & "'"
     sqls = sqls & " ,@Cont2='" & Trim(txtcontac2.Text) & "'"
     sqls = sqls & " ,@Mail2='" & Trim(txtmail2.Text) & "'"
     sqls = sqls & " ,@Cont3='" & Trim(txtcontac3.Text) & "'"
     sqls = sqls & " ,@Mail3='" & Trim(txtmail3.Text) & "'"
     sqls = sqls & " ,@Ruta=" & Val(txtruta.Text)
     sqls = sqls & " ,@Obser='" & Trim(txtObs.Text) & "'"
     sqls = sqls & " ,@Plaza=" & Val(txtPlaza.Text)
     cnxbdMty.Execute sqls
     MsgBox "Datos actualizados", vbInformation, "Ok..."
     LimpiarControles Me
  End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Private Sub cmdImprime_Click()
  Imprime crptToWindow
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.connect = "driver={SQL Server};server=SBMTY;uid=Operaciones;pwd=" & gpwdDataBase & ";database=Bonos"

    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCaratulas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtcliente.Text)
    mdiMain.cryReport.StoredProcParam(1) = Val(Usuario)
    mdiMain.cryReport.StoredProcParam(2) = Val(txtPlaza)
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub


Private Sub cmdSalir_Click()
  Unload Me
End Sub

Sub BUSCA_Cliente()
   sqls = " EXEC sp_BuscaCliente_Datos '" & txtcliente.Text & "','ClienteBE'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
   If rsBD.EOF Then
       MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
       txtcliente.SetFocus
       rsBD.Close
       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
   Else
      txtNombre.Text = rsBD!b
      boded = rsBD!c
      Call Muestra_datos
   End If
End Sub

Sub BUSCA_CP()
Dim sqls As String, sq As String
   sqls = "SELECT * FROM Codigos_Postales WHERE CP=" & Val(txtcp.Text)
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
   If rsBD.EOF Then
       MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
       txtPoblacion.Text = ""
       txtestado.Text = ""
       cboColonia.Clear
       txtcp.SetFocus
       rsBD.Close
       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
   Else
      txtPoblacion.Text = Trim(UCase(rsBD!poblacion))
      txtestado.Text = Trim(UCase(rsBD!estado))
      Call CargaColonias(cboColonia, "sp_sel_productobe 'BE','" & Val(txtcp.Text) & "','CP'")
      cboColonia.ListIndex = 0
      txttel.SetFocus
   End If
End Sub

Private Sub Form_Load()
  iniciar
End Sub

Private Sub spdPlazas_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim Plaza As Integer, cliente As Integer

With spdPlazas
    .Row = Row
    .Col = 1
    cliente = Val(.Text)
    .Col = 2
    Plaza = Val(.Text)
    Call BuscaPlaza(cliente, Plaza)
End With
spdPlazas.Visible = False
txtcalle.SetFocus
End Sub

Sub BuscaPlaza(cliente As Integer, Plaza As Integer)
    sqls = "select * from plazasbe where cliente = " & cliente & _
           " and plaza = " & Plaza
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        txtPlaza.Text = rsBD!Plaza
        lblPlaza.Text = rsBD!Descripcion
    End If
End Sub

Private Sub txtcalle_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 And txtcalle.Text <> "" Then
     txtcp.SetFocus
  End If
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtcliente <> "" Then
   Call BUSCA_Cliente
End If
End Sub

Private Sub txtcontac1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
     txtmail1.SetFocus
  End If
End Sub

Private Sub txtcontac2_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
     txtmail2.SetFocus
  End If
End Sub

Private Sub txtcontac3_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
     txtmail3.SetFocus
  End If
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtcp <> "" Then
   Call BUSCA_CP
End If
End Sub

Private Sub txtmail1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      txtcontac2.SetFocus
   End If
End Sub

Private Sub txtmail2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtcontac3.SetFocus
   End If
End Sub

Private Sub txtmail3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtruta.SetFocus
   End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdGrabar.SetFocus
  End If
End Sub


Private Sub txtruta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And txtruta.Text <> "" Then
      txtruta.Text = Val(txtruta.Text)
      Call BUSCA_Distribuidor
   End If
End Sub

Private Sub txttel_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtcontac1.SetFocus
   End If
End Sub

Sub BUSCA_Distribuidor()
   If boded = 0 Then
      boded = 1
   End If
   
   sqls = "SELECT * FROM DISTRIBUIDORES WHERE BODEGA=" & boded
   sqls = sqls & " AND Ruta=" & Val(txtruta.Text)
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
   If rsBD.EOF Then
       If asi = False Then
          MsgBox "No hubo resultados en su busqueda", vbCritical, "Sin resultados"
          txtruta.SetFocus
       End If
       rsBD.Close
       Set rsBD = Nothing
       MsgBar "", False
       Exit Sub
   Else
      txtNomRuta.Text = UCase(Trim(rsBD!nombre))
      txtObs.SetFocus
   End If
End Sub

Sub valida_datos()
continua = True
If Val(txtcliente.Text) <= 0 Then
   continua = False
   MsgBox "El cliente no es válido", vbCritical, "Error en cliente"
   Exit Sub
End If
'If txtcalle.Text = "" Then
'   continua = False
'   MsgBox "La calle no es válida", vbCritical, "Error en el nombre de calle"
'   Exit Sub
'End If
'If Val(txtCP.Text) = 0 Then
'   continua = False
'   MsgBox "El Codigo Postal no es válido", vbCritical, "Error en el CP"
'   Exit Sub
'End If
'If cboColonia.Text = "" Then
'   continua = False
'   MsgBox "La colonia no es válida", vbCritical, "Error en la Colonia"
'   Exit Sub
'End If
End Sub

Sub Muestra_datos()
Dim tmpplaza As Integer
On Error GoTo ERR:
sqls = "SELECT * FROM Envios_CtesBE "
sqls = sqls & "WHERE Cliente=" & Val(txtcliente)
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
        
If rsBD.EOF Then
   txtPlaza.SetFocus
   asi = False
Else
   txtcalle.Text = Trim(UCase(rsBD!Calle))
   txtcp.Text = rsBD!CP
   cboColonia.Text = Trim(UCase(rsBD!Colonia))
   txtPoblacion.Text = Trim(UCase(rsBD!poblacion))
   txtestado.Text = Trim(UCase(rsBD!estado))
   txttel.Text = Trim(rsBD!Telefono)
   txtcontac1.Text = Trim(UCase(rsBD!Contacto1))
   txtmail1.Text = Trim(UCase(rsBD!mail1))
   txtcontac2.Text = Trim(UCase(rsBD!Contacto2))
   txtmail2.Text = Trim(UCase(rsBD!mail2))
   txtcontac3.Text = Trim(UCase(rsBD!Contacto3))
   txtmail3.Text = Trim(UCase(rsBD!mail3))
   txtruta.Text = Val(rsBD!ruta)
   txtObs.Text = Trim(rsBD!Observacion)
   If IsNull(rsBD!Plaza) = True Then
     tmpplaza = 0
    Else
      tmpplaza = Val(rsBD!Plaza)
    End If
   txtPlaza.Text = tmpplaza
   asi = True
   Call BUSCA_Distribuidor
   Call BuscaPlaza(txtcliente, txtPlaza)
End If
   
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub

Sub iniciar()
    txtcliente.Text = ""
    txtNombre.Text = ""
    txtcalle.Text = ""
    txtcp.Text = ""
    cboColonia.Text = ""
    txtPoblacion.Text = ""
    txtestado.Text = ""
    txttel.Text = ""
    txtcontac1 = ""
    txtmail1 = ""
    txtcontac2 = ""
    txtmail2 = ""
    txtcontac3 = ""
    txtmail3 = ""
    txtruta = ""
    txtNomRuta.Text = ""
    txtPlaza.Text = ""
    spdPlazas.Visible = False
    lblPlaza.Text = ""
    txtObs.Text = ""
End Sub
