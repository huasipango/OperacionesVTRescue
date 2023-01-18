VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEnviaSyC 
   Caption         =   "Validación de Productos Electrónicos (S & C)"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "frmEnviaSyC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "spdClientes"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tarjetas"
      TabPicture(1)   =   "frmEnviaSyC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "spdTarjetas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dispersiones"
      TabPicture(2)   =   "frmEnviaSyC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "spdDispersiones"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ajustes"
      TabPicture(3)   =   "frmEnviaSyC.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "spdAjustes"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin FPSpread.vaSpread spdDispersiones 
         Height          =   6855
         Left            =   -74880
         OleObjectBlob   =   "frmEnviaSyC.frx":0070
         TabIndex        =   12
         Top             =   480
         Width           =   11055
      End
      Begin FPSpread.vaSpread spdClientes 
         Height          =   6975
         Left            =   -74760
         OleObjectBlob   =   "frmEnviaSyC.frx":11CF
         TabIndex        =   10
         Top             =   480
         Width           =   10815
      End
      Begin FPSpread.vaSpread spdTarjetas 
         Height          =   7095
         Left            =   -74760
         OleObjectBlob   =   "frmEnviaSyC.frx":2260
         TabIndex        =   11
         Top             =   480
         Width           =   10815
      End
      Begin FPSpread.vaSpread spdAjustes 
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "frmEnviaSyC.frx":3398
         TabIndex        =   13
         Top             =   480
         Width           =   11055
      End
   End
   Begin VB.Frame frBotones 
      Height          =   855
      Left            =   9840
      TabIndex        =   5
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   960
         Picture         =   "frmEnviaSyC.frx":44F9
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   400
         Left            =   240
         Picture         =   "frmEnviaSyC.frx":45FB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   400
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Tipos de Archivos a generar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.CheckBox chkCombustibles 
         Caption         =   "Estados de Cuenta"
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkClientes 
         Caption         =   "Clientes"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkTarjetas 
         Caption         =   "Tarjetas"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkAjustes 
         Caption         =   "Ajustes"
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkDispersiones 
         Caption         =   "Dispersiones"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmEnviaSyC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cadenafija As String
Dim ArchivoTarjetasPersonalizadas As String
Dim ArchivoTarjetasStock(17) As String
Dim ArchivoDispersiones As String
Dim ArchivoAjustes As String
Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ERR:
    cadenafija = "+00000+00000+00000+00000+00000"
    CargaDatos
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores generados"
  Exit Sub
End Sub
Sub CargaDatos()
    Screen.MousePointer = vbHourglass
    Call CargaClientes
    Call CargaTarjetas
    Call CargaDispersiones
    Call CargaAjustes
    Screen.MousePointer = vbNormal
End Sub
Sub CargaClientes()
On Error GoTo ERRO:
    sqls = "sp_CargaDatosAutorizador 'Clientes'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdClientes
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        .Col = 3
        .Text = rsBD!Plaza
        .Col = 4
        .Text = rsBD!descripcion
        .Col = 5
        .value = 1
        rsBD.MoveNext
    Loop
    End With
    If i > 0 Then
        chkClientes.value = Checked
    End If
    rsBD.Close
    Set rsBD = Nothing
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub
Sub CargaTarjetas()
On Error GoTo ERRO:
    sqls = "sp_CargaDatosAutorizador 'Tarjetas'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdTarjetas
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = rsBD!cliente
        .Col = 2
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        .Col = 3
        .Text = rsBD!Folio
        .Col = 4
        .Text = rsBD!Cantidad
        .Col = 5
        .Text = rsBD!Producto
        .Col = 6
        .Text = rsBD!Plaza
        .Col = 7
        .value = 1
        rsBD.MoveNext
    Loop
    End With
    If i > 0 Then
        chkTarjetas.value = Checked
    End If
    rsBD.Close
    Set rsBD = Nothing
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub
Sub CargaDispersiones()
On Error GoTo ERRO:
    sqls = "sp_CargaDatosAutorizador 'Dispersiones'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdDispersiones
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = CStr(rsBD!Folio)
        .Col = 2
        .Text = rsBD!Producto
        .Col = 3
        .Text = rsBD!cliente
        .Col = 4
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        .Col = 5
        .Text = Format(rsBD!valor, "Currency")
        .Col = 6
        .Text = rsBD!fechaEntrega
        .Col = 7
        If rsBD!fechaEntrega >= Now Then
            .value = 0
        Else
            .value = 1
        End If
        rsBD.MoveNext
    Loop
    End With
    If i > 0 Then
        chkDispersiones.value = Checked
    End If
    rsBD.Close
    Set rsBD = Nothing
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub

Sub CargaAjustes()
On Error GoTo ERRO:
    sqls = "sp_CargaDatosAutorizador 'Ajustes'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    i = 0
    With spdAjustes
    .Row = -1
    .Col = -1
    .Action = 12
    .MaxRows = 0
    Do While Not rsBD.EOF
        i = i + 1
        .MaxRows = i
        .Row = i
        .Col = 1
        .Text = CStr(rsBD!Folio)
        .Col = 2
        .Text = rsBD!Producto
        .Col = 3
        .Text = rsBD!cliente
        .Col = 4
        .Text = rsBD!Nombre
        If Len(.Text) > 100 Then
           .BackColor = vbYellow
        End If
        .Col = 5
        .Text = Format(rsBD!ValorCargo, "Currency")
        .Col = 6
        .Text = Format(rsBD!ValorAbono, "Currency")
        .Col = 7
        .value = 1
        rsBD.MoveNext
    Loop
    End With
    If i > 0 Then
        chkAjustes.value = Checked
    End If
    rsBD.Close
    Set rsBD = Nothing
Exit Sub
ERRO:
   MsgBox ERR.Description, vbCritical, "Errores encontrados"
   Exit Sub
End Sub

Private Sub cmdGrabar_Click()
Dim resp, okas As Boolean
On Error GoTo ERR:
okas = False
resp = MsgBox("¿Desea generar los archivos para S&C ? Una vez procesado el día, ya no se podra generar informacion de tarjetas y clientes, solo dispersiones", vbYesNo + vbQuestion + vbDefaultButton2, "Generacion de archivos electrónicos")
If resp = vbYes Then
    cmdGrabar.Enabled = False
    Screen.MousePointer = vbHourglass
    okas = True
       If chkClientes.value = 1 And spdClientes.MaxRows > 0 And okas Then
            okas = ActualizaStatusClientes
            If okas Then
                okas = GeneraArchivoClientes
            End If
       End If
       If chkTarjetas.value = 1 And spdTarjetas.MaxRows > 0 And okas Then
            okas = SiguienteEnvioTarjetas
            If okas Then
                okas = GuardaSBITarjetas
            End If
            If okas Then
                okas = GeneraArchivoTarjetas
            End If
       End If
       If (chkDispersiones.value = 1 And spdDispersiones.MaxRows > 0) Or (chkAjustes.value = 1 And spdAjustes.MaxRows > 0) And okas Then
            okas = SiguienteEnvioDispAj
            If okas And chkDispersiones.value = 1 Then
                okas = GuardaSBIDispersiones
            End If
            If okas And chkAjustes.value = 1 Then
                okas = GuardaSBIAjustes
            End If
            If okas Then
                okas = GeneraArchivoDispersiones
            End If
            If okas Then
               okas = GeneraArchivoAjustes
            End If
       End If
       If chkCombustibles.value = 1 And okas Then
            okas = GeneraArchivoDomicilios
            If okas Then
             GrabaCombustibles
             okas = GeneraArchivoCombustibles
            End If
       End If
       If okas = True Then
          MsgBox "Archivos generados " & dir_prueba, vbInformation, "Sistema Bono Electronico"
       Else
          MsgBox "No se generaron archivos", vbExclamation, "Sistema Bono Electronico"
       End If
    Screen.MousePointer = vbNormal
End If
Exit Sub
ERR:
  MsgBox ERR.Description, vbCritical, "Errores genereados"
End Sub

Function ActualizaStatusClientes() As Boolean
Dim cliente, Plaza
Dim i As Integer
ActualizaStatusClientes = False
On Error GoTo err_gral
If chkClientes.value = 1 Then
    With spdClientes
            For i = 1 To .MaxRows
               .Row = i
               .Col = 5
               If .value = 1 Then
                    .Col = 1
                    cliente = Val(.Text)
                    .Col = 3
                    Plaza = Val(.Text)
                    sqls = "sp_status_upd 'PlazasBE', 2," & cliente & "," & Plaza
                    cnxbdMty.Execute sqls, intRegistros
                End If
            Next i
    End With
Else
  Exit Function
End If
ActualizaStatusClientes = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de empleadoras"
   Exit Function
End Function

Function GeneraArchivoClientes() As Boolean
Dim clinea As String, i As Long
Dim cliente, Nombre, Plaza, NombrePlaza
Dim reg90 As Integer, reg91 As Integer
Dim elgrupo As String
Dim Archivo As String
Dim cont As Integer
Dim bPrimero As Boolean

GeneraArchivoClientes = False
On Error GoTo err_gral

cont = 0
'nfile = FreeFile
If chkClientes.value = 1 Then
    bPrimero = True
    Archivo = gstrPath & "AEMP" & Format(CDate(Date), "mmdd") & "_SC.vlt"
    Do While Dir(Archivo) <> ""
        cont = cont + 1
        Archivo = gstrPath & "AEMP" & Format(CDate(Date), "mmdd") & "_" & Format(cont, "00") & "_SC.vlt"
    Loop
    Open Archivo For Output As #1
    'Header A
    clinea = "10COMPANY04" & Format(CDate(Date), "YYYYMMDD") & rellena("0", 251, "")
    Print #1, clinea
    'Header B
    reg91 = 0
    With spdClientes
        reg90 = 0
        sqls = "sp_Producto_All"
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
        Do While Not rsBD.EOF
            bPrimero = True
            For i = 1 To .MaxRows
               .Row = i
               .Col = 5
               If .value = 1 Then
                    .Col = 2
                    elgrupo = Mid(UCase(.Text), 1, 100)
                    .Col = 3
                    Plaza = Val(.Text)
                    .Col = 4
                    NombrePlaza = Mid(UCase(.Text), 1, 60)
                   .Col = 1
                    cliente = Val(.Text)
                    clinea = "2001"
                    clinea = clinea & Format(cliente, "000000") & Format(Plaza, "0000") & rellena(elgrupo, 100, " ", "I")
                    clinea = clinea & Format(rsBD!Bon_Pro_Tipo, "00") & rellena(rsBD!DescCorta, 10, " ")
                    clinea = clinea & cadenafija & rellena("0", 114, "")
                    Print #1, clinea
                    sqls = "WebEmpresas..sp_sl_PlazasBE " & cliente & "," & Plaza
                    Set rsBD2 = New ADODB.Recordset
                    rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
                     Do While Not rsBD2.EOF
                        reg90 = reg90 + 1
                        clinea = "2101"
                        clinea = clinea & Format(cliente, "000000") & Format(Plaza, "0000")
                        clinea = clinea & rellena(CStr(NombrePlaza), 60, " ", "I")
                        clinea = clinea & rellena(Left(rsBD2!Direccion & " " & rsBD2!Colonia, 68), 68, " ", "I")
                        clinea = clinea & rellena(Left(rsBD2!Ciudad, 20), 20, " ", "I")
                        clinea = clinea & rellena(Left(rsBD2!EstadoCorto, 3), 3, " ", "I") & Format(rsBD2!CodigoPostal, "00000")
                        clinea = clinea & rellena(Left(rsBD2!Telefono, 12), 12, " ", "I")
                        clinea = clinea & rellena(Left(rsBD2!contacto, 26), 26, " ", "I")
                        clinea = clinea & rellena(" ", 50, "") & rellena("0", 12, "")
                        Print #1, clinea
                        rsBD2.MoveNext
                    Loop
                    rsBD2.Close
                    Set rsBD2 = Nothing
                    reg91 = reg91 + reg90
                    clinea = "90" & Format(reg90, "000000") & rellena("0", 262, "")
                    Print #1, clinea
                    reg90 = 0
                End If
            Next i
            rsBD.MoveNext
        Loop
    End With
    rsBD.Close
    Set rsBD = Nothing
Else
  Exit Function
End If
'Trailer A
clinea = "91" & Format(reg91, "000000") & rellena("0", 262, "")
Print #1, clinea

Close #1
GeneraArchivoClientes = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de empleadoras"
   Exit Function
End Function
Function SiguienteEnvioTarjetas() As Boolean
Dim cont
    cont = 0
    SiguienteEnvioTarjetas = False
    ArchivoTarjetasPersonalizadas = "TAT" & Format(CDate(Date), "yymmdd") & "_SC.vlt"
    Do While Dir(gstrPath & ArchivoTarjetasPersonalizadas) <> ""
        cont = cont + 1
        ArchivoTarjetasPersonalizadas = "TAT" & Format(CDate(Date), "yymmdd") & "_" & Format(cont, "00") & "_SC.vlt"
    Loop
    cont = 1
    ArchivoTarjetasStock(1) = "CAF50640601" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(1)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(1) = "CAF50640601" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(2) = "CAF50640501" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(2)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(2) = "CAF50640501" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(3) = "CAF50640602" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(3)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(3) = "CAF50640602" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(6) = "CAF50640611" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(6)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(6) = "CAF50640611" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(7) = "CAF50640511" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(7)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(7) = "CAF50640511" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(8) = "CAF50640612" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(8)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(8) = "CAF50640612" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(10) = "CAF50640502" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(10)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(10) = "CAF50640502" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(11) = "CAF50640503" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(11)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(11) = "CAF50640503" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(16) = "CAF50640512" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(16)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(16) = "CAF50640512" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    cont = 1
    ArchivoTarjetasStock(17) = "CAF50640513" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Do While Dir(gstrPath & ArchivoTarjetasStock(17)) <> ""
        cont = cont + 1
        ArchivoTarjetasStock(17) = "CAF50640513" & Format(CDate(Date), "yymmdd") & Format(cont, "00") & "EMBSYC-01.txt"
    Loop
    SiguienteEnvioTarjetas = True
End Function
Function GuardaSBITarjetas() As Boolean
Dim cliente, Pedido, Plaza
Dim i As Integer
Dim Archivo As String
On Error GoTo err_gral
GuardaSBITarjetas = False
If chkTarjetas.value = 1 Then
    sqls = "sp_SBIAltaTarjetas @Cliente = 0, @Pedido = 0, @Accion = 'Borrar'"
    cnxbdMty.Execute sqls, intRegistros
    With spdTarjetas
            For i = 1 To .MaxRows
               .Row = i
               .Col = 7
               If .value = 1 Then
                    .Col = 1
                    cliente = Val(.Text)
                    .Col = 3
                    Pedido = Val(.Text)
                    .Col = 6
                    Plaza = Val(.Text)
                    sqls = "sp_SBIAltaTarjetas @Cliente = " & cliente & ", @Pedido = " & Pedido & ", @Accion = 'Buscar'"
                    Set rsBD = New ADODB.Recordset
                    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
                    Do While Not rsBD.EOF
                        If rsBD!Stock = 0 Then
                            Archivo = ArchivoTarjetasPersonalizadas
                        Else
                            Archivo = ArchivoTarjetasStock(rsBD!Producto)
                        End If
                        If rsBD!tipo = "T" Or rsBD!tipo = "A" Then
                            sqls = "sp_SBIAltaTarjetas @Cliente = " & cliente & ", @Plaza = " & Plaza & ", @Pedido = " & Pedido & ",@Accion = 'Insertar', @Archivo = '" & Archivo & "', @Tipo = '" & rsBD!tipo & "', @Empleado = '" & rsBD!empleado & "', @Producto = " & rsBD!Producto & ", @Stock = " & rsBD!Stock
                            cnxbdMty.Execute sqls, intRegistros
                        End If
                        rsBD.MoveNext
                    Loop
                    sqls = "sp_SBIAltaTarjetas @Cliente = " & cliente & ", @Pedido = " & Pedido & ", @Accion = 'Guardar'"
                    cnxbdMty.Execute sqls, intRegistros
                    rsBD.Close
                    Set rsBD = Nothing
                End If
            Next i
    End With
Else
  Exit Function
End If
GuardaSBITarjetas = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Guarda Tarjetas"
   Exit Function
End Function


Function GeneraArchivoTarjetas() As Boolean
Dim clinea As String, i As Long
Dim cliente, Nombre, Plaza, NombrePlaza, Producto
Dim reg91 As Integer, reg92 As Integer, reg93 As Integer, regClientes As Integer
Dim elgrupo As String
Dim Archivo As String
Dim cont As Integer
Dim Cuenta As Long
Dim bPrimerArchivo As Boolean, bPrimerProducto As Boolean, bPrimerCliente As Boolean, bStockAbierto As Boolean


On Error GoTo err_gral
GeneraArchivoTarjetas = False
cont = 0
reg91 = 0
reg92 = 0
reg93 = 0
regClientes = 0
'nfile = FreeFile
If chkTarjetas.value = 1 Then
    bPrimerArchivo = True
    bPrimerProducto = True
    bPrimerCliente = True
    bStockAbierto = False
    
    sqls = "sp_SBIAltaTarjetas @Cliente = 0, @Pedido = 0, @Accion = 'Generar'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        If (cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza) And Not bPrimerCliente Then ' Footer B
            reg93 = reg93 + reg91
            If Left(Archivo, 1) = "T" Then 'Personalizado
                clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
                Print #1, clinea
            Else
'                If bStockAbierto Then
'                    Close #1
'                    bStockAbierto = False
'                End If
            End If
            reg91 = 0
            bPrimerCliente = True
        End If
        If (Producto <> rsBD!Producto And Not bPrimerProducto) Or (Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo) Then  ' Footer A
            reg92 = reg92 + reg93
            If Left(Archivo, 1) = "T" Then 'Personalizado
                clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
                Print #1, clinea
            Else
                If bStockAbierto Then
                    Close #1
                    bStockAbierto = False
                End If
            End If
            reg93 = 0
            bPrimerProducto = True
        End If
        If Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo Then ' Cerrar archivo
            If Left(Archivo, 1) = "T" Then 'Personalizado
                clinea = "92" & Format(regClientes, "000000") & Format(reg92, "000000") & rellena("0", 58, "")
                Print #1, clinea
                Close #1
            Else
                If bStockAbierto Then
                    Close #1
                    bStockAbierto = False
                End If
            End If
            reg92 = 0
            bPrimerArchivo = True
        End If
        If Archivo <> rsBD!ID_Archivo Then ' Crear archivo
            bPrimerArchivo = False
            Archivo = rsBD!ID_Archivo
            Producto = 0
            Open gstrPath & Archivo For Output As #1
            If Left(Archivo, 1) = "T" Then 'Personalizado
                'Header A
                clinea = "11EMP-REG04" & Format(CDate(Date), "YYYYMMDD") & rellena("0", 53, "")
                Print #1, clinea
            Else
                bStockAbierto = True
            End If
        End If
        If Producto <> rsBD!Producto Then ' Header B
            bPrimerProducto = False
            Producto = rsBD!Producto
            'Header B
            reg93 = 0
            If Left(Archivo, 1) = "T" Then 'Personalizado
                clinea = "13" & Format(rsBD!Producto, "00") & rellena("0", 68, "")
                Print #1, clinea
            End If
        End If
        If cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza Then 'Header C
            bPrimerCliente = False
            cliente = rsBD!cliente
            Plaza = rsBD!Plaza
            'Header C
            reg91 = 0
            regClientes = regClientes + 1
            If Left(Archivo, 1) = "T" Then 'Personalizado
                clinea = "12" & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 60, "")
                Print #1, clinea
            End If
        End If
        'Detail
        If Left(Archivo, 1) = "T" Then 'Personalizado
            Cuenta = 0
            If rsBD!tipo = "A" Then
                sqls = "sp_SBIAltaTarjetas @Cliente = " & rsBD!cliente & ", @Pedido = 0, @Plaza = " & rsBD!Plaza & ", @Tipo = '" & rsBD!tipo & "', @Empleado = '" & rsBD!empleado & "', @Producto = " & rsBD!Producto & ", @Accion = 'Adicional'"
                Set rsBD2 = New ADODB.Recordset
                rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
                If Not rsBD2.EOF Then
                    Cuenta = rsBD2!Cuenta
                End If
                rsBD2.Close
                Set rsBD2 = Nothing
            End If
            sqls = "sp_SBIAltaTarjetas @Cliente = " & rsBD!cliente & ", @Pedido = 0, @Plaza = " & rsBD!Plaza & ", @Tipo = '" & rsBD!tipo & "', @Empleado = '" & rsBD!empleado & "', @Producto = " & rsBD!Producto & ", @Accion = 'Detalle'"
            Set rsBD2 = New ADODB.Recordset
            rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
            If Not rsBD2.EOF Then
                clinea = "21" & IIf(rsBD!tipo = "T", "01", "02")
                clinea = clinea & rellena(rsBD!empleado, 10, " ", "I")
                reg91 = reg91 + 1
                clinea = clinea & rellena(Left(rsBD2!Nombre, 26), 26, " ", "I")
                clinea = clinea & Format(Cuenta, "00000000") & rellena("0", 24, "")
            End If
        Else
            sqls = "sp_SBIAltaTarjetas @Cliente = " & rsBD!cliente & ", @Pedido = 0, @Plaza = " & rsBD!Plaza & ", @Tipo = '" & rsBD!tipo & "', @Empleado = '" & rsBD!empleado & "', @Producto = " & rsBD!Producto & ", @Accion = 'DetalleSto'"
            Set rsBD2 = New ADODB.Recordset
            rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
            If Not rsBD2.EOF Then
                clinea = "1" & rsBD2!Bin & rellena(" ", 11, "")
                clinea = clinea & "51" & rellena(" ", 23, "") & "0450"
                clinea = clinea & rellena(Left(rsBD2!Nombre, 26), 35, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!Nombre, 24), 24, " ", "I")
                clinea = clinea & IIf(rsBD!tipo = "T", "0", "1")
                clinea = clinea & rellena(Left(rsBD2!Direccion, 45), 45, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!Colonia, 45), 45, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!Ciudad, 45), 45, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!estado, 45), 45, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!CodigoPostal, 5), 5, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!Telefono, 10), 10, " ", "I")
                clinea = clinea & rellena(Left(rsBD2!Telefono, 10), 10, " ", "I")
                clinea = clinea & rellena(Left(rsBD!Plaza, 3), 3, " ", "I")
                clinea = clinea & "01" & rellena("0", 12, "")
                clinea = clinea & rellena(Left(rsBD2!Rfc, 13), 13, " ", "I")
                clinea = clinea & rellena(" ", 85, "") & rellena(Left(rsBD2!NombreCliente, 30), 30, " ", "I")
            End If
        End If
        Print #1, clinea
        rsBD2.Close
        Set rsBD2 = Nothing
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    reg93 = reg93 + reg91
    reg92 = reg92 + reg93
    If Left(Archivo, 1) = "T" Then 'Personalizado
        clinea = "91" & Format(reg91, "000000") & rellena("0", 64, "")
        Print #1, clinea
        clinea = "93" & Format(reg93, "000000") & rellena("0", 64, "")
        Print #1, clinea
        clinea = "92" & Format(regClientes, "000000") & Format(reg92, "000000") & rellena("0", 58, "")
        Print #1, clinea
        Close #1
    Else
        If bStockAbierto Then
            Close #1
        End If
    End If
    sqls = "sp_SBIAltaTarjetas @Cliente = 0, @Pedido = 0, @Accion = 'Actualizar'"
    cnxbdMty.Execute sqls, intRegistros

Else
  Exit Function
End If
GeneraArchivoTarjetas = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Genera Archivo Tarjetas"
   Exit Function
End Function

Function SiguienteEnvioDispAj() As Boolean
Dim cont
    cont = 0
    SiguienteEnvioDispAj = False
    ArchivoDispersiones = "DS" & Format(CDate(Date), "yymmdd") & "_SC.vlt"
    Do While Dir(gstrPath & ArchivoDispersiones) <> ""
        cont = cont + 1
        ArchivoDispersiones = "DS" & Format(CDate(Date), "yymmdd") & "_" & Format(cont, "00") & "_SC.vlt"
    Loop
    ArchivoAjustes = "AJ" & Format(CDate(Date), "yymmdd") & "_SC.vlt"
    Do While Dir(gstrPath & ArchivoAjustes) <> ""
        cont = cont + 1
        ArchivoAjustes = "AJ" & Format(CDate(Date), "yymmdd") & "_" & Format(cont, "00") & "_SC.vlt"
    Loop
    SiguienteEnvioDispAj = True
End Function

Function GuardaSBIDispersiones() As Boolean
Dim cliente, Pedido, Producto
Dim i As Integer
Dim Archivo As String
On Error GoTo err_gral
GuardaSBIDispersiones = False
If chkDispersiones.value = 1 Then
    
    sqls = "  exec  sp_SBIDispersionesBE " & _
                " @Pedido = 0" & _
               ", @Cliente = 0" & _
                ", @FechaProc = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                ", @Accion = 'Borrar'"
    cnxBD.Execute sqls, intRegistros
    
    With spdDispersiones
            For i = 1 To .MaxRows
               .Row = i
               .Col = 7
               If .value = 1 Then
                    .Col = 3
                    cliente = Val(.Text)
                    .Col = 1
                    Pedido = Val(.Text)
                    .Col = 2
                    Producto = Val(.Text)
                    sqls = "  exec  sp_SBIDispersionesBE " & _
                                "  @Id_Archivo= '" & ArchivoDispersiones & "'" & _
                                ", @Accion = 'Insertar'" & _
                               ", @Pedido = " & Pedido & _
                               ", @Cliente = " & cliente & _
                               ", @ClienteDisp = " & cliente & _
                               ", @FechaProc = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                               ", @Producto=" & Producto
                    cnxBD.Execute sqls, intRegistros
                End If
            Next i
    End With
End If
GuardaSBIDispersiones = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Guarda Dispersiones"
   Exit Function
End Function


Function GuardaSBIAjustes() As Boolean
Dim cliente, Pedido, Producto
Dim i As Integer
Dim Archivo As String
On Error GoTo err_gral
GuardaSBIAjustes = False
If chkAjustes.value = 1 Then
    
    sqls = "  exec  sp_SBIAjustesBE " & _
                " @Folio = 0" & _
               ", @Cliente = 0" & _
                ", @FechaProc = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                ", @Accion = 'Borrar'"
    cnxBD.Execute sqls, intRegistros
    
    With spdAjustes
            For i = 1 To .MaxRows
               .Row = i
               .Col = 7
               If .value = 1 Then
                    .Col = 3
                    cliente = Val(.Text)
                    .Col = 1
                    Pedido = Val(.Text)
                    .Col = 2
                    Producto = Val(.Text)
                    sqls = "  exec  sp_SBIAjustesBE " & _
                               "  @Id_Archivo= '" & ArchivoAjustes & "'" & _
                               ",  @Id_ArchivoDisp= '" & ArchivoDispersiones & "'" & _
                               ", @Folio = " & Pedido & _
                               ", @Cliente = " & cliente & _
                               ", @FechaProc = '" & Format(Now, "mm/dd/yyyy") & "'" & _
                               ", @Accion = 'Insertar'"
                    cnxBD.Execute sqls, intRegistros
                End If
            Next i
    End With
Else
  Exit Function
End If
GuardaSBIAjustes = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Guarda Ajustes"
   Exit Function
End Function

Function GeneraArchivoDispersiones() As Boolean
Dim clinea As String, i As Long
Dim cliente, Nombre, Plaza, NombrePlaza, Producto
Dim reg95 As Integer, reg96 As Integer, reg97 As Integer, regClientes As Integer
Dim val95 As Double, valr96 As Double, val97 As Double
Dim elgrupo As String
Dim Archivo As String
Dim cont As Integer
Dim bPrimerArchivo As Boolean, bPrimerProducto As Boolean, bPrimerCliente As Boolean, bStockAbierto As Boolean


On Error GoTo err_gral
GeneraArchivoDispersiones = False
cont = 0
reg95 = 0
reg96 = 0
reg97 = 0
val95 = 0
val96 = 0
val97 = 0
regClientes = 0
'nfile = FreeFile
If chkDispersiones.value = 1 Or chkAjustes.value = 1 Then
    bPrimerArchivo = True
    bPrimerProducto = True
    bPrimerCliente = True
    bStockAbierto = False
        
    sqls = "sp_SBIDispersionesBE @Cliente = 0, @Pedido = 0, @Accion = 'Generar'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        If (cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza) And Not bPrimerCliente Then ' Footer B
            reg97 = reg97 + reg95
            val97 = val97 + val95
            clinea = "95" & Format(reg95, "000000") & Format(val95 * 100, "00000000000")
            Print #1, clinea
            reg95 = 0
            val95 = 0
            bPrimerCliente = True
        End If
        If (Producto <> rsBD!Producto And Not bPrimerProducto) Or (Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo) Then  ' Footer A
            reg96 = reg96 + reg97
            val96 = val96 + val97
            clinea = "97" & Format(reg97, "000000") & Format(val97 * 100, "00000000000")
            Print #1, clinea
            reg97 = 0
            val97 = 0
            bPrimerProducto = True
        End If
        If Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo Then ' Cerrar archivo
            clinea = "96" & Format(reg96, "000000") & Format(val96 * 100, "00000000000")
            Print #1, clinea
            Close #1
            reg96 = 0
            val96 = 0
            bPrimerArchivo = True
        End If
        If Archivo <> rsBD!ID_Archivo Then ' Crear archivo
            bPrimerArchivo = False
            Archivo = rsBD!ID_Archivo
            Open gstrPath & Archivo For Output As #1
            'Header A
            clinea = "15SALDOS 04" & Format(CDate(Date), "YYYYMMDD")
            Print #1, clinea
        End If
        If Producto <> rsBD!Producto Then ' Header B
            bPrimerProducto = False
            Producto = rsBD!Producto
            'Header B
            reg97 = 0
            val97 = 0
            clinea = "17" & Format(rsBD!Producto, "00") & rellena("0", 15, "")
            Print #1, clinea
        End If
        If cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza Then 'Header C
            bPrimerCliente = False
            cliente = rsBD!cliente
            Plaza = rsBD!Plaza
            'Header C
            reg95 = 0
            val95 = 0
            regClientes = regClientes + 1
            clinea = "16" & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
            Print #1, clinea
        End If
        'Detail
        clinea = "23" & Format(rsBD!Cuenta, "00000000")
        clinea = clinea & Format(rsBD!importe * 100, "000000000")
        reg95 = reg95 + 1
        val95 = val95 + rsBD!importe
        Print #1, clinea
        
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    If Not bPrimerArchivo Then
        reg97 = reg97 + reg95
        val97 = val97 + val95
        reg96 = reg96 + reg97
        val96 = val96 + val97
        clinea = "95" & Format(reg95, "000000") & Format(val95 * 100, "00000000000")
        Print #1, clinea
        clinea = "97" & Format(reg97, "000000") & Format(val97 * 100, "00000000000")
        Print #1, clinea
        clinea = "96" & Format(reg96, "000000") & Format(val96 * 100, "00000000000")
        Print #1, clinea
        Close #1
    End If
    sqls = "sp_SBIDispersionesBE @Cliente = 0, @Pedido = 0, @Accion = 'Actualizar'"
    cnxbdMty.Execute sqls, intRegistros

Else
  Exit Function
End If
GeneraArchivoDispersiones = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Genera Archivo Dispersiones"
   Exit Function
End Function


Function GeneraArchivoAjustes() As Boolean
Dim clinea As String, i As Long
Dim cliente, Nombre, Plaza, NombrePlaza, Producto
Dim reg95 As Integer, reg96 As Integer, reg97 As Integer, regClientes As Integer
Dim val95 As Double, valr96 As Double, val97 As Double
Dim elgrupo As String
Dim Archivo As String
Dim cont As Integer
Dim bPrimerArchivo As Boolean, bPrimerProducto As Boolean, bPrimerCliente As Boolean, bStockAbierto As Boolean


On Error GoTo err_gral
GeneraArchivoAjustes = False
cont = 0
reg95 = 0
reg96 = 0
reg97 = 0
val95 = 0
val96 = 0
val97 = 0
regClientes = 0
'nfile = FreeFile
If chkAjustes.value = 1 Then
    bPrimerArchivo = True
    bPrimerProducto = True
    bPrimerCliente = True
    bStockAbierto = False
        
    sqls = "sp_SBIAjustesBE @Cliente = 0, @Folio = 0, @Accion = 'Generar'"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    Do While Not rsBD.EOF
        If (cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza) And Not bPrimerCliente Then ' Footer B
            reg97 = reg97 + reg95
            val97 = val97 + val95
            clinea = "95" & Format(reg95, "000000") & Format(val95 * 100, "00000000000")
            Print #1, clinea
            reg95 = 0
            val95 = 0
            bPrimerCliente = True
        End If
        If (Producto <> rsBD!Producto And Not bPrimerProducto) Or (Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo) Then  ' Footer A
            reg96 = reg96 + reg97
            val96 = val96 + val97
            clinea = "97" & Format(reg97, "000000") & Format(val97 * 100, "00000000000")
            Print #1, clinea
            reg97 = 0
            val97 = 0
            bPrimerProducto = True
        End If
        If Archivo <> rsBD!ID_Archivo And Not bPrimerArchivo Then ' Cerrar archivo
            clinea = "96" & Format(regClientes, "000000") & Format(reg96, "000000") & Format(val96 * 100, "00000000000")
            Print #1, clinea
            Close #1
            reg96 = 0
            val96 = 0
            bPrimerArchivo = True
        End If
        If Archivo <> rsBD!ID_Archivo Then ' Crear archivo
            bPrimerArchivo = False
            Archivo = rsBD!ID_Archivo
            Open gstrPath & Archivo For Output As #1
            'Header A
            clinea = "15AJUSTES04" & Format(CDate(Date), "YYYYMMDD")
            Print #1, clinea
        End If
        If Producto <> rsBD!Producto Then ' Header B
            bPrimerProducto = False
            Producto = rsBD!Producto
            'Header B
            reg97 = 0
            val97 = 0
            clinea = "17" & Format(rsBD!Producto, "00") & rellena("0", 15, "")
            Print #1, clinea
        End If
        If cliente <> rsBD!cliente Or Plaza <> rsBD!Plaza Then 'Header C
            bPrimerCliente = False
            cliente = rsBD!cliente
            Plaza = rsBD!Plaza
            'Header C
            reg95 = 0
            val95 = 0
            regClientes = regClientes + 1
            clinea = "16" & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000") & rellena("0", 7, "")
            Print #1, clinea
        End If
        'Detail
        clinea = "23" & Format(rsBD!Cuenta, "00000000")
        clinea = clinea & Format(rsBD!importe * 100, "000000000")
        reg95 = reg95 + 1
        val95 = val95 + rsBD!importe
        Print #1, clinea
        
        rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    If Not bPrimerArchivo Then
        reg97 = reg97 + reg95
        val97 = val97 + val95
        reg96 = reg96 + reg97
        val96 = val96 + val97
        clinea = "95" & Format(reg95, "000000") & Format(val95 * 100, "00000000000")
        Print #1, clinea
        clinea = "97" & Format(reg97, "000000") & Format(val97 * 100, "00000000000")
        Print #1, clinea
        clinea = "96" & Format(reg96, "000000") & Format(val96 * 100, "00000000000")
        Print #1, clinea
        Close #1
    End If
    sqls = "sp_SBIAjustesBE @Cliente = 0, @Folio = 0, @Accion = 'Actualizar'"
    cnxbdMty.Execute sqls, intRegistros
End If
GeneraArchivoAjustes = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Genera Archivo Ajustes"
   Exit Function
End Function

Function GeneraArchivoDomicilios() As Boolean
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal, reg90 As Integer, reg91 As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
Dim CP
   GeneraArchivoDomicilios = False
'nfile = FreeFile
'If chkUni_Domicilios.value = 1 Or chkAli_Domicilios.value = 1 Or chkReg_Domicilios.value = 1 Then
    ' Creamos Domicilios de Entrega
    Open gstrPath & "MDF" & Format(CDate(Date), "yymmdd") & "_SC.vlt" For Output As #1
    'Header A
    clinea = "11" & Format(CDate(Date), "YYYYMMDD") & "03" & rellena("0", 521, "")
    Print #1, clinea
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(Date), "mm/dd/yyyy") & "', 0, 'DetDomEnt2',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
    'Detalle
    reg91 = 0
    
     Do While Not rsBD.EOF
         reg91 = reg91 + 1
        clinea = "22" & Format(rsBD!Clave, "00")
        clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
        clinea = clinea & rellena(CStr(rsBD!Nombre), 150, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!Rfc), 13, " ", "I")
        clinea = clinea & Format(rsBD!cliente, "000000") & Format(rsBD!Plaza, "0000")
        clinea = clinea & rellena(CStr(rsBD!Direccion), 90, " ", "I") & rellena(CStr(rsBD!Colonia), 90, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!Ciudad), 50, " ", "I") & rellena(CStr(rsBD!Ciudad), 50, " ", "I")
        clinea = clinea & rellena(CStr(rsBD!estado), 50, " ", "I") & Format(rsBD!CodigoPostal, "00000")
        clinea = clinea & "01" & rellena("0", 9, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
           
    'Trailer A
    clinea = "91" & Format(reg91, "00000000") & rellena("0", 523, "")
    Print #1, clinea
    
    Close #1
    rsBD.Close
    Set rsBD = Nothing
'End If
    GeneraArchivoDomicilios = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Alta de Archivo Domicilios"
   Exit Function
End Function

Function GeneraArchivoCombustibles() As Boolean
Dim nArchivo, clinea As String, i As Long, NumTar As Long
Dim cliente, Nombre
Dim Empleadora As String
Dim RetVal, reg90 As Integer, reg91 As Integer
Dim ImporteCliente As Double, EmpleadosCliente As Integer
Dim ImporteClienteGlobal As Double, EmpleadosClienteGlobal As Integer
Dim CEROS As String, elgrupo As String, COP As Integer
Dim transa As String, ntransa As Integer
Dim Plaza, Sucursal, Cuenta As String, telef As String, elcontact As String
On Error GoTo err_gral
EmpleadosClienteGlobal = 0
Dim CP

    GeneraArchivoCombustibles = False
    ' Creamos Domicilios de Entrega
    Open gstrPath & "IEDOCTA" & Format(CDate(Date), "yymmdd") & "_SC.vlt" For Output As #1
    'Header A
    clinea = "11" & Format(CDate(Date), "YYYYMMDD") & "03" & rellena("0", 50, "")
    Print #1, clinea
    
    reg91 = 0
    
    'Detalle Magna
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(Date), "mm/dd/yyyy") & "', 0, 'Combustible',2"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Detalle Premium
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(Date), "mm/dd/yyyy") & "', 0, 'Combustible',10"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Detalle Diesel
    
    sqls = "sp_Vistas_AltasSBI '" & Format(CDate(Date), "mm/dd/yyyy") & "', 0, 'Combustible',11"
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
     Do While Not rsBD.EOF
        reg91 = reg91 + 1
        clinea = "22"
        clinea = clinea & Format(rsBD!Tarjeta, "0000000000000000")
        clinea = clinea & Format(rsBD!TipoCombustible, "0")
        clinea = clinea & rellena("0", 43, "")
        Print #1, clinea
        rsBD.MoveNext
     Loop
    rsBD.Close
    Set rsBD = Nothing
           
    'Trailer A
    clinea = "91" & Format(reg91, "00000000") & rellena("0", 52, "")
    Print #1, clinea
    
    Close #1

    GeneraArchivoCombustibles = True
Exit Function
err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Archivo Combustibles"
   Exit Function
End Function
Sub GrabaCombustibles()
    sqls = "sp_Vistas_PagoGas '" & Format(Now, "mm/dd/yyyy") & "',NULL,'Combustible'"
    cnxBD.Execute sqls, intRegistros
End Sub
