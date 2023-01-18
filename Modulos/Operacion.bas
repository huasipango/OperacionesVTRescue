Attribute VB_Name = "Module1"
Global cnxBD As ADODB.Connection
Global cnxBD2 As ADODB.Connection
Global cnxbdMty As ADODB.Connection
Global cnxbdMty2 As ADODB.Connection
Global cnxbdMatriz As ADODB.Connection
Global cnxbdMatriz2 As ADODB.Connection
Global cnxBDSuc As ADODB.Connection
Global rsBD As ADODB.Recordset
Global rsBD2 As ADODB.Recordset
Global rsBD3 As ADODB.Recordset
Global rsNombre As ADODB.Recordset
Global consulta As New ADODB.Recordset
Global Consultaint As New ADODB.Recordset
Global rsfol As ADODB.Recordset
Global sqls As String
Global gstrServidor As String, gstrDataBase As String, gstrPC     As String
Global gTipoServidor As String, gpwdDataBase As String
Global gstrImpFact As String
Global strPuerto As String
Global gstrPath As String
Global nResp As Integer
Global PathSBI As String
Global gsPathFE As String
Global Impresora As String
Global Product As Byte 'variable para determinar el producto
Global Reimp_FCNC As Byte
Global Producto_factura As Byte 'variable aux para saber ke producto estoy facturando
Global actual_manejo As Boolean 'variable que uso para distinguir si aun manejo el producto junto o separado
Global Prodof As Byte 'variable que uso para distinguir si es pago bono o pago uniforme
Global serverc As String, based As String
Global RespMsg As Integer                 ' Variable que toma el valor de la respuesta de los mensajes

Public nombre_com As String
Public ape_pat As String
Public ape_mat As String
Public Const ivagral As Double = 0.16 'IVA GENERAL
Public cliente_busca As String
Public latarjeta As String
Public elgrupo As Integer
Public prod_anterior As Byte
Public tipo_estad As Byte
Public inicia_consulped As Byte

'Pase de datos para cambio de cliente en pedidos
Public Bodegp As Byte
Public pedidop As Long
Public clientepp As Integer
Public palabra_ok As Boolean
Public polizaBE As Byte
Public elserver As String

'DATOS DE ENTREGA CTES_PAGOGAS
Public GCliente As String
Public GNombre As String
Public GCalle As String
Public GColonia As String
Public GCd As String
Public GEstado As String
Public GCP As String
Public GTel As String
Public Gprod As Byte
Public GprodText As String
Public plazasuc As String
'----------------------------
Public user_externo As Boolean
Public user_master As Boolean
Public ajuste_o_cancel As Byte

Public sEmail As String, sAsunto As String, sTexto As String
Public Const glngFolxCajaFact As Long = 1000
Public blnreimp As Boolean
Public ImpUnaVez As Boolean
   
Global FormatoFactura As Integer
Global PRIMERA As Boolean
Global letrero As String
Global Usuario As String, nomUsuario As String
Global Reporte As String, subReporte As String
Global rango_periodo As String
Global mov_o_stok As Byte

Global gPath As String
Global TipoBusqueda As String
Global TipoRep As String

Public directorio_pagogas As String
Public gstrSO As String
Global busc_clientes As String 'para ejecutar directa la buskeda de cliente
 
Public Const gstrSistema As String = "FACTBE"
Public Const gstrConfigSist As String = "Configuracion"
Public Const gstrKeyLastUser As String = "LastUser"
'Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Global gbytBodega As Byte
Global gArrDataBase() As String
Public gnBodega As Integer, gsTipoBodega As String, gnMaquina As Integer, gsOS As String, sUbicacion As Integer
Public gnFormatoFactura As Integer
Public gstrSerieFactura As String
Public gnMultiBodega As String
Public gnMultiVend As String
Public gnUEN, i As Long, j As Long
Public gnMultiUEN As String
Public gnAdministrador As String, gstrUsuario As String
Public bExiste As Boolean
Public swEntrada As Boolean
Public Nuevo As Boolean
Public intRegistros As Integer
Public TipoAcceso As String
Public TipoEntrada As String ' Para cuando entra sin archivo de configuracion facturacion.ini  es "SC", si si, es "CC"
Public gblnSendMail As Boolean

'Variables para envio de correo
Public Const gstrMailSMTPHost As String = "mail.valetotal.com"
Public Const gstrMailSMTPPort As String = "2525"
Public Const gstrMailFrom As String = "webmaster@valetotal.com"
Public Const gstrMailUser As String = "webmaster@valetotal.com"
Public Const gstrMailPassword As String = "Mateo2208@"


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Public Const SS_ACTION_SMARTPRINT = 32

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' DAutoSize property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' SortBy property settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1

Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103
Private Declare Function OpenProcess Lib "kernel32" _
(ByVal dwDesiredAccess&, ByVal bInheritHandle&, ByVal dwProcessId&) _
As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) _
As Long

Sub Main()
Dim fLogin As New frmLogin
Dim VersionBD As String
Dim VersionApp As String
    gPath = App.Path
    Screen.MousePointer = vbHourglass
    DoEvents
    frmSplash.Show
    frmSplash.Refresh
    Checaini
    
    If gstrPC = "" Then
      gstrPC = UCase(GetPCName())
    End If
        
    PathSBI = gstrPath
  
       
    On Error GoTo NoConeccion
    
    Set cnxBD = New ADODB.Connection
    cnxBD.CommandTimeout = 60000
    cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase & ";Column Encryption Setting = Enabled"
    
    Set cnxbdMty = New ADODB.Connection
    cnxbdMty.CommandTimeout = 60000
    cnxbdMty.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase & ";Column Encryption Setting = Enabled"
    
   ' Set cnxbdMatriz = New ADODB.Connection
   ' cnxbdMatriz.CommandTimeout = 60000
   ' cnxbdMatriz.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase & ";Column Encryption Setting = Enabled"
       
    'Busco versión actual del sistema
      VersionApp = CStr(VB.App.Major) & "." & CStr(VB.App.Minor) & "." & CStr(VB.App.Revision)
      Set rstTmp = New ADODB.Recordset
      strsql = "spr_Version"
      rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not rstTmp.EOF Then
         VersionBD = rstTmp!VersionOpe
      End If
      rstTmp.Close
      Set rstTmp = Nothing
     
      If VersionApp <> VersionBD And Not InIDE() Then
        MsgBox "Hay una versión más nueva del sistema, favor de actualizarla!"
        End
      End If
    'Obtener impresora de Facturas
      strRuta = ""
      gstrImpFact = gstrPC
      Set rstTmp = New ADODB.Recordset
      strsql = "sp_Impresoras_Sel '" & gstrPC & "', " & _
                                    "'FACTURAS'"
      rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not rstTmp.EOF Then
         strRuta = UCase(Trim(rstTmp!Impresora & ""))
      End If
      If strRuta <> "" And _
            strRuta <> "LPT1" Then
         strsql = "sp_Maquinas_SelxRuta 3, '" & strRuta & "'"
         Set rstTmp = New ADODB.Recordset
         rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
         If Not rstTmp.EOF Then
            gstrImpFact = rstTmp!Maquina
         End If
      End If
      Set rstTmp = Nothing
      
    
    On Error GoTo 0
    Unload frmSplash
    ApagaMenus
    Screen.MousePointer = vbNormal
    DoEvents
    
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Fallo al iniciar la sesión, se sale de la aplicación
        cnxBD.Close
        End
    End If
    Unload fLogin

    sUbicacion = Pad(CStr(Val(getDatoBodega("BodegaImprime", gnBodega))), 2, "0")
    gnFormatoFactura = Val(getDatoBodega("FormatoFactura", gnBodega))
    gstrSerieFactura = Trim(getDatoBodega("Serie_Factura", gnBodega))
    gblnSendMail = False
    
    Call ActualizaMenus("FBE")
    
    '---Segmento que uso para saber si aun se usa pago bono junto o separado de pago gas
    actual_manejo = False   '=FALSE AUN ESTAN JUNTOS =TRUE CADA UNO SEPARADO
    '------------------
    
    mdiMain.Show
    Exit Sub
NoConeccion:
    MsgBox "No se pudo conectar al servidor. Favor de Avisar a Soporte Técnico y Reintentar mas tarde!" & vbCr & "Error " & Str(ERR.Number) & ", " & ERR.Description, vbCritical, "Error al conectar..."
    End
    
End Sub

Sub Checaini()

    Dim vntLinea As Variant
    Dim strPath As String
    Dim strFile As String
    Dim Pos As Integer, CadenaDB As String
    Dim blnDB As Boolean
    
    gstrServidor = ""
    gstrDataBase = ""
    gnMaquina = 0
    
    If UCase(Dir(gPath & "\OperacionesVT.ini")) = "OPERACIONESVT.INI" Then

        Open gPath & "\OperacionesVT.ini" For Input As #1
        Input #1, gstrServidor, gstrDataBase, gnBodega, gnMaquina, gstrPath, gsOS, gTipoServidor
        Close #1
        TipoEntrada = "CC"
        If gTipoServidor = "Pruebas" Then
            gstrServidor = gstrServidor + ",5004"
            gpwdDataBase = "OpVt2016"
            MsgBox ("Trabajando en Base de Datos de Pruebas!")
        ElseIf gTipoServidor = "Calidad" Then
            gstrServidor = gstrServidor + ",5005"
            gpwdDataBase = "OpVt2016.2"
            MsgBox ("Trabajando en Base de Datos de QA!")
        ElseIf gTipoServidor = "Produccion" Then
            gstrServidor = gstrServidor + ",5004"
            gpwdDataBase = "OpVT.2017@3"
        End If
        
        serverc = gstrServidor
        based = gstrDataBase
    Else
        MsgBox ("No se encontro el archivo OperacionesVT.ini, favor de avisar a sisteamas!")
        End
    End If
Exit Sub

err_gral:
    MsgBox ERR.Description
    Exit Sub
End Sub

Function GetPCName() As String
Dim sPCName As String
Dim nLength As Long

    sPCName = Space$(256)
    nLength = Len(sPCName)
    GetComputerName sPCName, nLength
    sPCName = Left$(sPCName, nLength)

    GetPCName = sPCName

End Function

Sub ApagaMenus()
    Dim vntMenus As Variant
    
    On Error Resume Next
    
    For Each vntMenus In mdiMain.Controls
       If TypeOf vntMenus Is Menu Then
          Select Case vntMenus.Name
            Case Is = "mnuArchivo"
            Case Is = "mnuAbrirBD"
            Case Is = "mnuSalir"
            Case Is = "mnuAyuda"
            Case Is = "mnuAyudaTemAyu"
            Case Is = "mnuAcercade"
            Case Else
                 vntMenus.Visible = False
          End Select
       End If
    Next
End Sub

Function getDatoBodega(ByVal sCampo As String, ByVal iBodega As Integer) As String
Dim oRS As ADODB.Recordset, strsql As String

   strsql = "SELECT " & sCampo & " " & _
                        "FROM Bodegas " & _
                        "WHERE Bodega = " & iBodega
   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open strsql, cnxBD, adOpenForwardOnly, adLockPessimistic, adCmdText
   DoEvents
   
' Create and open an ADO Recordset using the command object
   If Not oRS.EOF Then
      getDatoBodega = oRS.Fields(sCampo).value & ""
   Else
      getDatoBodega = ""
   End If
   
   oRS.Close
   Set oRS = Nothing
   DoEvents
End Function


Public Function Pad(ByVal sText As String, ByVal iLength As Integer, Optional ByVal sCharacter As String, Optional ByVal sType As String) As String
Dim iLng1 As Integer, iLng2 As Integer, iSpaces As Integer

   If IsMissing(sCharacter) Then
      sCharacter = " "
   End If
   If IsMissing(sType) Or sType = "" Then
      sType = "L"
   Else
      sType = IIf(InStr(1, "LRC", UCase(sType), vbBinaryCompare) <> 0, sType, "L")
   End If
   If Len(sText) >= iLength Then
      Pad = Left(sText, iLength)
   Else
      If sType = "L" Then
         Pad = String$(iLength - Len(sText), sCharacter) & sText
      ElseIf sType = "R" Then
         Pad = sText & String$(iLength - Len(sText), sCharacter)
      ElseIf sType = "C" Then
         iSpaces = iLength - Len(sText)
         iLng1 = Int(iSpaces / 2)
         iLng2 = iSpaces - iLng1
         Pad = String$(iLng1, sCharacter) & sText & String$(iLng2, sCharacter)
      End If
   End If
End Function

Sub ActualizaMenus(ByVal strModulo As String)
Dim vntMenus As Variant, bytCount As Byte, strsql As String
   
   On Error GoTo RollBack
    
    sqls = "Exec Sp_OpcionesxModulo_Upd 1,'" & strModulo & "'"
    cnxBD.Execute sqls

    For Each vntMenus In mdiMain.Controls
        If TypeOf vntMenus Is Menu Then
            bytCount = bytCount + 1
    
            sqls = "Exec Sp_OpcionesxModulo_Ins '" & strModulo & "','" & Left(vntMenus.Name, 30) & "','" & Left(vntMenus.Caption, 40) & "'," & bytCount & ""
            cnxBD.Execute sqls
        End If
    Next
    
    sqls = "Exec Sp_OpcionesxModulo_Upd 2, '" & strModulo & "'"
    cnxBD.Execute sqls
    
    Exit Sub
    
RollBack:
 '  MsgBox "Actualización No Concretada", vbCritical
    
End Sub

Sub ChecaRegWin()
   Usuario = GetSetting(gstrSistema, gstrConfigSist, gstrKeyLastUser, "SUPERVISOR")
End Sub
Public Sub CargaBodegas(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Bodega, Descripcion"
    sqls = sqls & vbCr & " FROM Bodegas "
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    
    intCount = -1
    intCount = intCount + 1
    cbo.AddItem "<< TODAS >>"
    cbo.ItemData(intCount) = 0
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![Bodega])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, gnBodega)
End Sub
Sub CboPosiciona(cboCombo As ComboBox, intValor As Variant)
 Dim IntPos As Long
 For IntPos = 0 To cboCombo.ListCount - 1
      If cboCombo.ItemData(IntPos) = intValor Then
          cboCombo.ListIndex = IntPos
          Exit For
      End If
  Next IntPos
  If cboCombo.ListIndex <> -1 Then
    If cboCombo.ItemData(cboCombo.ListIndex) <> intValor Then
      cboCombo.ListIndex = -1
    End If
  End If
End Sub
Sub ActivaMenus()
    Dim vntMenus As Variant
    
    MsgBar "Activado Menus", True
    
    For Each vntMenus In mdiMain.Controls
       If TypeOf vntMenus Is Menu Then
            vntMenus.Visible = True
            vntMenus.Enabled = True
       End If
    Next
    
    MsgBar "", False
    
End Sub
Function FileExist(Filename As String) As Boolean
   If Dir$(Filename) = "" Then
      FileExist = False
   Else
      FileExist = True
   End If
End Function
Sub UserMenu(ByVal strModulo As String, Perfil As Integer)
Dim vntMenus As Variant, DerechoValido As Boolean, i As Integer
Dim ContDerechos, DerechosUsuario() As String
    
sqls = "Exec Sp_OpcionesxPerfil_Sel '" & strModulo & "'," & Perfil & ""
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

ContDerechos = 0
Do While rsBD.EOF = False
    ContDerechos = ContDerechos + 1
    ReDim Preserve DerechosUsuario(ContDerechos)
    DerechosUsuario(ContDerechos) = UCase(Trim(rsBD!CveOpcion))
    rsBD.MoveNext
Loop
rsBD.Close
Set rsBD = Nothing

For Each vntMenus In mdiMain.Controls
    If TypeOf vntMenus Is Menu Then
        If vntMenus.Caption = "&Archivo" Or _
            vntMenus.Caption = "&Abrir Base de Datos" Or _
            vntMenus.Caption = "&Ayuda" Or _
            vntMenus.Caption = "&Acerca de.." Or _
            vntMenus.Caption = "&Salir" Then
            vntMenus.Enabled = True
        Else
            If ContDerechos > 0 Then
                DerechoValido = False
                For i = 1 To ContDerechos
                    If DerechosUsuario(i) = UCase(Trim(vntMenus.Name)) Then
                        DerechoValido = True
                        Exit For
                    End If
                Next
                If DerechoValido Then
                    vntMenus.Visible = True
                    vntMenus.Enabled = True
                Else
                    vntMenus.Visible = True
                    If vntMenus.Caption <> "-" Then
                        vntMenus.Enabled = False
                    End If
                End If
            End If
        End If
    End If
Next
End Sub
Sub CargaComboBE(cbo As ComboBox, sSql As String)
Dim oRS As ADODB.Recordset
'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase

   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open sSql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   cbo.Clear
   Do While Not oRS.EOF
      If sSql = "sp_Bancos_Sel" Then
        cbo.AddItem UCase(Trim(oRS!Nombre))
      ElseIf Left(sSql, 13) = "sp_Claves_Sel" Then
        cbo.AddItem UCase(Trim(oRS!descripcion))
      Else
        cbo.AddItem UCase(Trim(oRS!Bon_Pro_Descripcion))
      End If
      oRS.MoveNext
   Loop
   oRS.Close
   Set oRS = Nothing
End Sub
Sub CargaComboBE2(cbo As ComboBox, sSql As String)
Dim oRS As ADODB.Recordset
Dim intCount As Integer

'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase

   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open sSql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   cbo.Clear
   intCount = -1
   Do While Not oRS.EOF
      intCount = intCount + 1
      cbo.AddItem UCase(Trim(oRS!b))
      cbo.ItemData(intCount) = oRS!a
      oRS.MoveNext
   Loop
   oRS.Close
   Set oRS = Nothing
   Call CboPosiciona(cbo, 0)
End Sub
Sub LeeproductoBE(cbo As ComboBox, sSql As String)
Dim oRS As ADODB.Recordset
'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
   
   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open sSql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   If Not oRS.EOF Then
      Product = oRS!Bon_Pro_Tipo
   End If
   oRS.Close
   Set oRS = Nothing
End Sub
Public Sub CargaBodegasS2(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Bodega, Descripcion"
    sqls = sqls & vbCr & " FROM Bodegas "
    If user_master = False Then
       sqls = sqls & vbCr & " where Bodega = " & gnBodega
    End If
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
        
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![Bodega])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, gnBodega)
End Sub
Public Sub CargaBodegasServ(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Bodega, Descripcion"
    sqls = sqls & vbCr & " FROM Bodegas "
'    SQLS = SQLS & vbCr & " where Bodega = " & gnBodega
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![Bodega])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    
    rsBD.Close
    Set rsBD = Nothing
    
'    IntCount = IntCount + 1
'    intBodega = 16
'    strBodega = "D.F."
'    cbo.AddItem Trim(strBodega)
'    cbo.ItemData(IntCount) = intBodega
    Call CboPosiciona(cbo, gnBodega)
End Sub
Sub producto_cual()
'   If actual_manejo = False Then
'       Product = IIf(Product = 8, 6, Product)
'   Else
       Product = Product
'   End If
End Sub
Public Sub LimpiarControles(frmForm As Form)

      Dim ctlControl As Object

      On Error Resume Next
      For Each ctlControl In frmForm.Controls
         ctlControl.Text = ""
'         ctlControl.ListIndex = -1
         ctlControl.value = False
         DoEvents
      Next ctlControl

End Sub
Public Sub CargaPoblaciones(cbo As Control, estado As Integer)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Poblacion, Descripcion"
    sqls = sqls & vbCr & " FROM POBLACIONES "
    sqls = sqls & vbCr & " WHERE ESTADO = " & Val(estado)
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD2 = New ADODB.Recordset
    rsBD2.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD2.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD2![poblacion])
       strBodega = Trim("" & rsBD2![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD2.MoveNext
    Loop
    rsBD2.Close
    Set rsBD2 = Nothing
    
End Sub
Public Sub CargaEstados(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Estado, Descripcion"
    sqls = sqls & vbCr & " FROM Estados "
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![estado])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    
End Sub
Public Sub CargaGiros(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves "
    sqls = sqls & vbCr & "  WHERE Tabla = 'Grupos' "
    sqls = sqls & vbCr & " Order By Descripcion"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, -1)
End Sub
Sub MsgBar(ByVal value As String, ByVal Status As Boolean)

    If value = "" Then
       mdiMain.sbStatusBar.Panels(1).Text = "Listo..." & gstrServidor
       Screen.MousePointer = vbDefault
    Else
      If Status Then
         mdiMain.sbStatusBar.Panels(1).Text = value & ", Espere..."
         Screen.MousePointer = vbHourglass
         DoEvents
      Else
         mdiMain.sbStatusBar.Panels(1).Text = value
      End If
    End If
        
End Sub
Sub CentraForma(FRM As Form, Optional ByVal mdiChild As Boolean)
Dim nDifV As Integer, nDifH As Integer

    If Not mdiChild Then
      FRM.Top = (Screen.Height - FRM.Height) / 2
      FRM.Left = (Screen.Width - FRM.Width) / 2
   Else
      nDifV = mdiMain.Height - mdiMain.ScaleHeight
      nDifH = mdiMain.Width - mdiMain.ScaleWidth
      FRM.Top = (mdiMain.Height - (FRM.Height + nDifV)) / 2
      FRM.Left = (mdiMain.Width - (FRM.Width + nDifH)) / 2
   End If
End Sub
Public Sub CargaVendedores(cbo As Control)
    Dim intVendedor As Integer
    Dim strVendedor As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Vendedor, nombre"
    sqls = sqls & vbCr & " FROM Vendedores "
    sqls = sqls & vbCr & " Order By Bodega, vendedor"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    intCount = intCount + 1
    intBodega = 0
    strBodega = "Todos"
    cbo.AddItem Trim(strBodega)
    cbo.ItemData(intCount) = intBodega
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![vendedor])
       strBodega = Trim("" & rsBD![Nombre])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
End Sub
Sub Limpia_CryReport()
    Dim intCount As Integer
        
    For intCount = 0 To 10
        mdiMain.cryReport.StoredProcParam(intCount) = ""
        mdiMain.cryReport.Formulas(intCount) = ""
    Next intCount
    
    mdiMain.cryReport.ReportFileName = ""
    mdiMain.cryReport.SelectionFormula = ""
    mdiMain.cryReport.WindowTitle = ""
End Sub
Sub doErrorLog(ByVal Sucursal As Integer, ByVal Modulo As String, ByVal Number As Long, ByVal Description As String, ByVal User As String, ByVal Source As String, Optional ByVal Line As Long, Optional ByVal sql As String)
Dim sPCName As String, sLine As String
Dim strsql As String

   sLine = IIf(IsMissing(Line), "", Trim(CStr(Line)))
   sLine = IIf(sLine = "0", "", sLine)
   sql = IIf(IsMissing(sql), "", sql)
   
   sPCName = GetPCName()
   
   strsql = "spb_GrabaErrorLog " & Sucursal & ", " & _
                                 "'" & Modulo & "', " & _
                                 "'" & Format$(Date, "YYYY-MM-DD") & "', " & _
                                 "'" & Format(Time(), "HH:nn") & "', " & _
                                 "'" & sPCName & "', " & _
                                 "'" & User & "', " & _
                                 "'" & Source & "', " & _
                                 "'" & sLine & "', " & _
                                 "'" & Trim(CStr(Number)) & "', " & _
                                 "'" & Replace(Description, "'", "''") & "', " & _
                                 "'" & Replace(sql, "'", "''") & "'"
   cnxbdMty.Execute strsql
End Sub
Sub Producto_actual()
    Producto_factura = Product
End Sub
Public Function doGenArchFE(ByVal nBodega As Integer, _
                         ByVal sserie As String, _
                         ByVal nFactIni As Long, _
                         ByVal nFactFin As Long, _
                         ByVal nTipoArch As Integer) As Boolean
Dim nfile As Long, SFILE As String, blnFileOpen As Boolean
Dim blnOk As Boolean, Cob_dob As Double, cliente_dob As Integer
Dim strsql As String, rstTmp As ADODB.Recordset, rstTmp2 As ADODB.Recordset
Dim sRutina As String
'Variables para el Do While
Dim nFactAnt As Long, nNR As Long, nCantFact As Long
Dim nLineas As Long, nUltReg As Long
Dim blnAgrupa As Boolean, blnAdd As Boolean
'nRegAct As Long,
Dim nTipoPed As Long
Dim lngI As Long
'Variables para denominaciones raras
Dim SCodigo As String, sDescripcion As String, _
   nUnidades As Long, nIvaProd As Double, _
   nImporte As Double, nIva As Double, nTotal As Double
Dim sFolioIni As String, sFolioFin As String
Dim nPrecio As Double, nPrecioSug As Double
Dim sDesglose As String
'Variable para Agrupar Prefacturas
Dim sBodega As String, nPrefSig As Double, nFolSig As Double
'Documento
Dim sSerieDocto As String, sFactura As String, sFecha As String, _
      sNoAprob As String, sAnoAprob As String, sNoCert As String, _
      sCveBod As String, sDescBodega As String, _
      other_fecha As Date
'Empresa
Dim sEmpRazonSocial As String, sEmpRFC As String, sEmpCalle As String, _
   sEmpNoExt As String, sEmpNoInt As String, sEmpCol As String, _
   sEmpCP As String, sEmpLocalidad As String, sEmpMun As String, _
   sEmpEdo As String, sEmpPais As String, sEmpTel As String, _
   sEmpPagWeb As String, sMetodoDePago As String
'AdendaSAT
Dim sVersionCMPV As String, sTipoOperacionCMPV As String, _
    sRegistroPatronalCMPV As String, sNumeroCuentaCMPV As String, _
    sTotalCMPV As String
'Expide
Dim sExpCalle As String, sExpNoExt As String, sExpNoInt As String, _
   sExpCol As String, sExpCP As String, sExpLocalidad As String, _
   sExpMun As String, sExpEdo As String, sExpPais As String, _
   sExpTel As String
'Cliente a Facturar
Dim sRCliente As String, sRRFC As String, _
   sRNombre As String, _
   sRCalle As String, sRNoExt As String, sRNoInt As String, _
   sREntreCalles As String, sRLocalidad As String, sRCol As String, _
   sRMun As String, sREdo As String, sRPais As String, _
   sRCP As String
', sRApPat As String, sRAPMat As String
'Cliente a Enviar
Dim sECliente As String, sENombre As String, _
   sECalle As String, sENoExt As String, sENoInt As String, _
   sEEntreCalles As String, sELocalidad As String, sECol As String, _
   sEMun As String, sEEdo As String, sEPais As String, _
   sECP As String, sEGuia As String, SETel As String
'Encabezado
Dim sPedido As String, sRuta As String, sCondiciones As String, sFechaVence As String
'Variables de la Factura
Dim blnGravIVA As Boolean, versioncfd As String, metodop As String, cuentap As String, UsoCFDI As String, DeduceDesp As Integer
Dim sRef As String, sRefBonos As String, sPedRef
Dim sComentario As String, sImporteLetras As String
Dim sLeyendaPagare As String, sPagareDatosCliente As String, sPagareCiudadContrato As String
'Detalle de la Factura
Dim aCodigo() As String, aDescripcion() As String, _
   aUnidades() As Long, aPrecio() As Double, aPrecioSug() As Double, _
   aFolIni() As String, aFolFin() As String, _
   aImporte() As Double, aIVA() As Double, aTotal() As Double, elgrav As String
'Detalle de la Factura (Prefacturas)
Dim aSerie() As String, aDescBod() As String, _
      aPrefIni() As Double, aPrefFin() As Double, aAnoMes() As String
'Datos de la Comision
Dim nImpComision As Double, nImpIVAComision As Double
Dim nPrcCom As Double, nPrcComCalc As Double, nPrcIVA As Double
'Totales
Dim nTotalUnidades As Long, _
   ntotalPedido, nTotalImporte As Double, nTotalIVA As Double, nTotalGen As Double
Dim sImpresora As String, nofiscal As String
Dim minutos As String, Unimed As String
'  1.- Factura Normal o de Stock
'  2.- Factura de Prefacturas

   On Error GoTo hdErr
   
  
   nofiscal = "ESTIMADO CLIENTE, LE RECORDAMOS QUE SU GASTO DE COMBUSTIBLE LO PODRA DEDUCIR EN BASE AL CONSUMO REAL QUE SEÑALE EL ESTADO DE CUENTA PROPORCIONADO POR EL BANCO SCOTIABANK."
   sImpresora = doFindPrinter(gstrPC, "FACTURAS")
   '-blnImprime = True
   blnImprime = False
   
   blnOk = False
   
   Call Producto_actual
   strsql = "sp_GenArchFE " & nBodega & ", " & _
                           "'" & sserie & "', " & _
                           nFactIni & ", " & _
                           nFactFin & ", " & _
                           nTipoArch
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   If nTipoArch = 7 Then
      nTipoArch = 1
   End If
   Do While Not rstTmp.EOF
      Unimed = UCase(rstTmp!Unimed)
      Cob_dob = rstTmp!Comision
      cliente_dob = 0
      sqls = "SELECT * FROM CLIENTESCONFIG WITH (NOLOCK) WHERE FactComision='S'"
      sqls = sqls & "  AND CLIENTE = " & Val(rstTmp!RCliente)
      'sqls = "sp_ConsultasBE_Varios @Accion='Ctes_config_fact',@Cliente=" & Val(rstTmp!RCLIENTE)
      Set Consultaint = New ADODB.Recordset
      Consultaint.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
      If Not Consultaint.EOF Then
         cliente_dob = Val(rstTmp!RCliente)
      Else
         cliente_dob = 0
      End If
      
      blnOk = True
      If nFactAnt <> rstTmp!Folio Then
         If nFactAnt <> 0 Then
            If nNR > 0 Or sRef <> "" Then
               nCantFact = nCantFact + 1
               If nTipoArch = 1 Then
'                 Factura Normal
'                 Si tiene fracciones agregar registro
                  If nUnidades > 0 Then
                     ReDim Preserve aCodigo(nLineas)
                     ReDim Preserve aDescripcion(nLineas)
                     ReDim Preserve aUnidades(nLineas)
                     ReDim Preserve aPrecio(nLineas)
                     ReDim Preserve aPrecioSug(nLineas)
                     ReDim Preserve aFolIni(nLineas)
                     ReDim Preserve aFolFin(nLineas)
                     ReDim Preserve aImporte(nLineas)
                     ReDim Preserve aIVA(nLineas)
                     ReDim Preserve aTotal(nLineas)
                     
                     
                     aCodigo(nLineas) = SCodigo
                     aDescripcion(nLineas) = sDescripcion
                     aUnidades(nLineas) = nUnidades
                     aPrecio(nLineas) = 0
                     aPrecioSug(nLineas) = 0
                     aFolIni(nLineas) = ""
                     aFolFin(nLineas) = ""
                     aImporte(nLineas) = nImporte
                     aIVA(nLineas) = nIva
                     aTotal(nLineas) = nTotal
                     
                     nUltReg = nLineas - 1
                  Else
                     nLineas = nLineas - 1
                     nUltReg = nLineas
                  End If
               ElseIf nTipoArch = 2 Then
                  nLineas = nLineas - 1
               End If
               
               GoSub doGenFile
            End If
         End If
         nFactAnt = rstTmp!Folio
'         nRegAct = 1
         nLineas = 1
         
         If nTipoArch = 1 Then
'           Reinicia valores para agrupar denominaciones raras
            SCodigo = ""
            sDescripcion = ""
            nUnidades = 0
            nImporte = 0
            nIva = 0
            nTotal = 0
         Else
'           Reinicia valores para agrupar prefacturas
            sBodega = ""
            nPrefSig = 0
            nFolSig = 0
            
'           Reinicia Totales
            nTotalUnidades = 0
         End If
         
         GoSub doFillEnc
      End If
      
      GoSub doFillDet
      
'      nRegAct = nRegAct + 1
      rstTmp.MoveNext
   Loop
   
   If nNR > 0 Or sRef <> "" Then
      nCantFact = nCantFact + 1
      If nTipoArch = 1 Then
'        Si tiene fracciones agregar registro
         If nUnidades > 0 Then
            ReDim Preserve aCodigo(nLineas)
            ReDim Preserve aDescripcion(nLineas)
            ReDim Preserve aUnidades(nLineas)
            ReDim Preserve aPrecio(nLineas)
            ReDim Preserve aPrecioSug(nLineas)
            ReDim Preserve aFolIni(nLineas)
            ReDim Preserve aFolFin(nLineas)
            ReDim Preserve aImporte(nLineas)
            ReDim Preserve aIVA(nLineas)
            ReDim Preserve aTotal(nLineas)
         
            aCodigo(nLineas) = SCodigo
            aDescripcion(nLineas) = sDescripcion
            aUnidades(nLineas) = nUnidades
            aPrecio(nLineas) = 0
            aPrecioSug(nLineas) = 0
            aFolIni(nLineas) = ""
            aFolFin(nLineas) = ""
            aImporte(nLineas) = nImporte
            aIVA(nLineas) = nIva
            aTotal(nLineas) = nTotal
            
            nUltReg = nLineas - 1
         Else
            nLineas = nLineas - 1
            nUltReg = nLineas
         End If
      ElseIf nTipoArch = 2 Then
         nLineas = nLineas - 1
      End If
      
      'GeneraArchivo
      GoSub doGenFile
   End If
   
   On Error GoTo 0
   doGenArchFE = blnOk
   Exit Function

doFillEnc:
   If nTipoArch = 1 Then
      nTipoPed = rstTmp!TipoPedido  'TIPO DE PRODUCTO
   End If
   blnAgrupa = (rstTmp!NR > 20)
         
   sRutina = "doFillEnc"
   sSerieDocto = Trim(rstTmp!Serie)
   sFactura = rstTmp!Folio
   
   minutos = Format(rstTmp!Fecha_Factura, "hh:mm:ss")
   sFecha = Format(rstTmp!Fecha_Factura, "YYYY-MM-DD")
   other_fecha = Format(rstTmp!Fecha_Factura, "mm/dd/yyyy")
   sFecha = sFecha & "T" & minutos
   sNoAprob = rstTmp!NoAprob
   sAnoAprob = rstTmp!AnoAprob
   sNoCert = IIf(IsNull(rstTmp!NoCert), 0, rstTmp!NoCert)
   sCveBod = rstTmp!BodegaNumero
   sDescBodega = rstTmp!BodegaNombre
   metodop = Format(rstTmp!MetodoPago, "00")
   cuentap = Trim(rstTmp!CuentaPago)
   UsoCFDI = Trim(rstTmp!UsoCFDI)
   versioncfd = rstTmp!versioncfd
   DeduceDesp = rstTmp!DeduceDespensa
   sMetodoDePago = rstTmp!MetodoDePago
   
   sEmpRazonSocial = rstTmp!EmpRazonSocial
   sEmpRFC = rstTmp!EmpRFC
   sEmpCalle = rstTmp!EmpCalle
   sEmpNoExt = rstTmp!EmpNoExt
   sEmpNoInt = rstTmp!EmpNoInt
   sEmpCol = rstTmp!EmpColonia
   sEmpCP = rstTmp!EmpCP
   sEmpLocalidad = rstTmp!EmpLocalidad
   sEmpMun = rstTmp!EmpMunicipio
   sEmpEdo = rstTmp!EmpEstado
   sEmpPais = rstTmp!EmpPais
   sEmpTel = rstTmp!EmpTel
   sEmpPagWeb = rstTmp!EmpPagWeb
   
   sVersionCMPV = rstTmp!VersionCMPV
   sTipoOperacionCMPV = rstTmp!TipoOperacionCMPV
   sRegistroPatronalCMPV = IIf(IsNull(rstTmp!RegistroPatronalCMPV), 0, rstTmp!RegistroPatronalCMPV)
   sNumeroCuentaCMPV = rstTmp!NumeroCuentaCMPV
   sTotalCMPV = rstTmp!TotalCMPV
   
   sExpCalle = rstTmp!ExpCalle
   sExpNoExt = rstTmp!ExpNoExt
   sExpNoInt = rstTmp!ExpNoInt
   sExpCol = rstTmp!ExpColonia
   sExpCP = rstTmp!ExpCP
   sExpLocalidad = rstTmp!EmpLocalidad
   sExpMun = rstTmp!ExpMunicipio
   sExpEdo = rstTmp!ExpEstado
   sExpPais = rstTmp!ExpPais
   sExpTel = rstTmp!ExpTel
   
   sRCliente = rstTmp!RCliente
   sRRFC = DeleteChar("-", DeleteChar(" ", rstTmp!RRFC))
   sRNombre = rstTmp!RNombre
'   sRApPat = rstTmp!RApPat
'   sRAPMat = rstTmp!RAPMat
   sRCalle = rstTmp!RCalle
   sRNoExt = rstTmp!RNoExt
   sRNoInt = rstTmp!RNoInt
   sREntreCalles = rstTmp!REntreCalles
   sRLocalidad = rstTmp!RLocalidad
   sRCol = rstTmp!Rcolonia
   sRMun = rstTmp!RMunicipio
   sREdo = rstTmp!REstado
   sRPais = rstTmp!RPais
   sRCP = rstTmp!RCP
   
   sECliente = rstTmp!Ecliente
   sENombre = rstTmp!ENombre
   sECalle = rstTmp!ECalle
   sENoExt = rstTmp!ENoExt
   sENoInt = rstTmp!ENoInt
   sEEntreCalles = rstTmp!EEntreCalles
   sELocalidad = rstTmp!ELocalidad
   sECol = rstTmp!EColonia
   sEMun = rstTmp!EMunicipio
   sEEdo = rstTmp!EEstado
   sEPais = rstTmp!EPais
   sECP = rstTmp!ECP
   If nTipoArch = 1 Then
      sEGuia = rstTmp!EGuiaRoji
      SETel = rstTmp!ETel
      sPedido = rstTmp!Pedido
   End If
   sRuta = rstTmp!ruta
   sCondiciones = IIf(rstTmp!Condiciones = 1, "CONTADO", Trim(Str(rstTmp!Condiciones)) & " DIAS")
   sFechaVence = doFormatDate(Format(rstTmp!FechaVence, "YYYY-MM-DD"), 2)
   
   blnGravIVA = IIf(rstTmp!GravIVA = "S", True, False)
   sRef = rstTmp!Referencia
   sRefBonos = rstTmp!RefBonos
   sPedRef = rstTmp!PedidoRef
   
   nImpComision = rstTmp!Comision
   nImpIVAComision = Round(rstTmp!ivacomision, 2)
   nPrcCom = rstTmp!PorcComision
   If rstTmp!ValorPedido = 0 Then
      nPrcComCalc = 0
   Else
      nPrcComCalc = Round(rstTmp!Comision / rstTmp!ValorPedido * 100, 2)
   End If
   nPrcCom = IIf(nPrcCom <> nPrcComCalc, nPrcComCalc, nPrcCom)
   
   nPrcIVA = rstTmp!PrcIVA
   
   nIvaProd = IIf(blnGravIVA, (rstTmp!PrcIVA / 100), 0)
   
   If blnGravIVA = False Then  'si es BE despensa
      ntotalPedido = rstTmp!ValorPedido
      nTotalImporte = rstTmp!TotalImporte
      nTotalIVA = rstTmp!TotalIva
      nTotalGen = rstTmp!TotalGeneral
   Else   'SI ES BE GASOLINA O UNIFORMES
      ntotalPedido = rstTmp!ValorPedido '/ (1 + nIvaProd) Se lo quite el 4 Abril porque el total es con iva
      nTotalImporte = rstTmp!ValorPedido
      nTotalIVA = rstTmp!TotalIva ' Era nTotalImporte - ntotalPedido
      nTotalGen = rstTmp!TotalGeneral ' Era ntotalPedido + nTotalIVA vsp 24/08/2017
   End If
   
   If nTipoArch = 1 Then
      nTotalUnidades = 1
      sComentario = rstTmp!Comentario
   End If
   'sFecha = Format(sfecha, "YYYY-MM-DD") & "T00:00:00"
   sImporteLetras = Leyenda(ntotalPedido + nImpComision + nImpIVAComision)
   sLeyendaPagare = "Por el presente PAGARE me(nos) obligo(amos) incondicionalmente a pagar " & _
                    "a la orden de " & sEmpRazonSocial & " en esta plaza en moneda nacional el dia " & _
                    Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4) & _
                    " la cantidad de " & Format(ntotalPedido + nImpComision + nImpIVAComision, "###,###,##0.00") & " (" & sImporteLetras & ") " & _
                    "valor en efectivo. Si no fuere pagado satisfactoriamente este pagaré me(nos) obligo(amos) " & _
                    "a pagar durante todo el tiempo que permaneciera total o parcialmente insoluto, " & _
                    "un interés legal de tasa TIIE + 5 puntos, sin que por esto se considere prorrogado " & _
                    "el plazo fijado para el cumplimiento de esta obligación."
   sPagareDatosCliente = sRNombre & " con Domicilio en: " & sRCalle & " " & sRNoExt & _
                         " Colonia: " & sRCol & " en " & sRMun & ", " & sREdo
   sPagareCiudadContrato = "Lugar y Fecha de Expedición: " & sExpMun & ", " & sExpEdo & ", a " & _
                           Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4)
                           
   sMailFE = IIf(rstTmp!MailFe = 1, "MAIL", "")
   sMailFETo = rstTmp!MailFETo
   Return

doFillDet:
   sRutina = "doFillDet"
   If nTipoArch = 1 Then
      If blnAgrupa _
               And (rstTmp!Precio <> 1 _
               And rstTmp!Precio <> 2 _
               And rstTmp!Precio <> 3 _
               And rstTmp!Precio <> 4 _
               And rstTmp!Precio <> 5 _
               And rstTmp!Precio <> 6 _
               And rstTmp!Precio <> 7 _
               And rstTmp!Precio <> 8 _
               And rstTmp!Precio <> 9 _
               And rstTmp!Precio <> 10 _
               And rstTmp!Precio <> 20 _
               And rstTmp!Precio <> 30 _
               And rstTmp!Precio <> 50 _
               And rstTmp!Precio <> 100 _
               And rstTmp!Precio <> 200) Then
   '     Agregarlo al final a un arreglo para al final imprimir
         SCodigo = rstTmp!codigo
         sDescripcion = rstTmp!descripcion
         nUnidades = nUnidades + rstTmp!Unidades
         nImporte = nImporte + Round((rstTmp!total / (1 + nIvaProd)), 2)
         nIva = nIva + (rstTmp!total - Round((rstTmp!total / (1 + nIvaProd)), 2))
         nTotal = nTotal + rstTmp!total
      Else
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aPrecio(nLineas)
         ReDim Preserve aPrecioSug(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = rstTmp!codigo
         aDescripcion(nLineas) = rstTmp!descripcion
         aUnidades(nLineas) = rstTmp!Unidades
         aPrecio(nLineas) = Round((rstTmp!Precio / (1 + nIvaProd)), 2)
         aPrecioSug(nLineas) = rstTmp!Precio
         aFolIni(nLineas) = IIf(rstTmp!TipoPedido = 12, rstTmp!FolioIniGral & "", rstTmp!FolioIni & "")
         aFolFin(nLineas) = IIf(rstTmp!TipoPedido = 12, rstTmp!FolioFinGral & "", rstTmp!FolioFin & "")
         aImporte(nLineas) = Round((rstTmp!total / (1 + nIvaProd)), 2)
         aIVA(nLineas) = (rstTmp!total - Round((rstTmp!total / (1 + nIvaProd)), 2))
         aTotal(nLineas) = rstTmp!total
         
         nLineas = nLineas + 1
      End If
   ElseIf nTipoArch = 2 Then
      If blnAgrupa Then
         If sBodega = rstTmp!descBodega & "" And _
               nPrefSig = rstTmp!NoPreFactura And _
               nFolSig = Val(Right(Trim(rstTmp!FolioIni & ""), 7)) Then
            aPrefFin(nLineas - 1) = rstTmp!NoPreFactura
            aFolFin(nLineas - 1) = Right(Trim(rstTmp!FolioFin & ""), 7)
            aUnidades(nLineas - 1) = aUnidades(nLineas - 1) + rstTmp!Unidades
            aImporte(nLineas - 1) = aImporte(nLineas - 1) + rstTmp!importe
            aIVA(nLineas - 1) = aIVA(nLineas - 1) + rstTmp!iva
            aTotal(nLineas - 1) = aTotal(nLineas - 1) + rstTmp!total
            blnAdd = False
         Else
            blnAdd = True
'            nRegAct = nRegAct + 1
         End If
         
         sBodega = rstTmp!descBodega & ""
         nPrefSig = rstTmp!NoPreFactura + 1
         nFolSig = Val(Right(Trim(rstTmp!FolioFin & ""), 7)) + 1
      Else
         blnAdd = True
      End If
      
      If blnAdd Then
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         
         ReDim Preserve aSerie(nLineas)
         ReDim Preserve aPrefIni(nLineas)
         ReDim Preserve aPrefFin(nLineas)
         ReDim Preserve aDescBod(nLineas)
         ReDim Preserve aAnoMes(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = rstTmp!codigo
         aDescripcion(nLineas) = rstTmp!descripcion
         
         aSerie(nLineas) = Trim(rstTmp!SeriePreFactura & "")
         aPrefIni(nLineas) = rstTmp!NoPreFactura
         aPrefFin(nLineas) = rstTmp!NoPreFactura
         aDescBod(nLineas) = rstTmp!descBodega & ""
         aAnoMes(nLineas) = Left(Trim(rstTmp!FolioIni & ""), 6)
         aFolIni(nLineas) = Right(Trim(rstTmp!FolioIni & ""), 7)
         aFolFin(nLineas) = Right(Trim(rstTmp!FolioFin & ""), 7)
         
         aUnidades(nLineas) = rstTmp!Unidades
         aImporte(nLineas) = rstTmp!importe
         aIVA(nLineas) = rstTmp!iva
         aTotal(nLineas) = rstTmp!total
      
         nLineas = nLineas + 1
      End If
      
      nTotalUnidades = nTotalUnidades + rstTmp!Unidades
   End If
   Return
   
doGenFile:
   sRutina = "doGenFile"
   GoSub doFileOpen
   GoSub doGenEnc
   GoSub doGenDet
   GoSub doGenPie
   GoSub doFileClose
   Return
   
doFileOpen:
   sRutina = "doFileOpen"
   nfile = FreeFile()
   gsPathFE = "c:\Facturacion"
   SFILE = gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT"
   Open SFILE For Output As #nfile
   blnFileOpen = True
   Return

doFileClose:
   sRutina = "doFileClose"
   Close #nfile
   blnFileOpen = False
   GoSub doFileMove
   Return

doFileMove:
   sRutina = "doFileMove"
   Call doWaitShell("MOVE " & gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT" & " C:\GoDir\WsHome\datos\salida")
   Return
   
doGenEnc:
   sRutina = "doGenEnc"
   If Producto_factura = 1 Then 'Or Producto_factura = 8 Or Producto_factura >= 9 Then
      Print #nfile, "TFormato;" & "VTOFNOR"
      Print #nfile, "TIPO_DOCTO;" & "FACTURA"
   ElseIf Producto_factura = 2 Then
      Print #nfile, "TFormato;" & "VTOFNOR"
      Print #nfile, "TIPO_DOCTO;" & "FACTURA"
   ElseIf Producto_factura = 3 Then
      Print #nfile, "TFormato;" & "VTONOTA"
      Print #nfile, "TIPO_DOCTO;" & "NOTA"
   Else
      Print #nfile, "TFormato;" & "VTOFNOR"
      Print #nfile, "TIPO_DOCTO;" & "FACTURA"
   End If
   If Producto_factura = 1 And DeduceDesp = 1 Then
       Print #nfile, "Ndistrib;" & "FE," & sMailFE & ",DESPEN,NO," & IIf(blnImprime, "SI", "NO")
   Else
       Print #nfile, "Ndistrib;" & "FE," & sMailFE & ",,NO," & IIf(blnImprime, "SI", "NO")
   End If
   '--------------
   If sMailFE = "MAIL" Then
      Print #nfile, "COPIAS;1,,"
   Else
      Print #nfile, "COPIAS;2,,"
   End If
   '--------------
   Print #nfile, "Impresora;" & sImpresora
   '--------------
   
   If sMailFE = "MAIL" Then
      Print #nfile, "CorreoTo;" & sMailFETo
      Print #nfile, "CorreoSub;" & "Envio de Factura Electronica " & sSerieDocto & "-" & sFactura & IIf(sPedido = "0" Or sPedido = "", "", " del Pedido " & sPedido)
      Print #nfile, "CorreoAtt;PDF,XML"
   End If
   
   '--------------
   Print #nfile, "Serie;" & sSerieDocto '2
   Print #nfile, "Folio;" & sFactura '3
   Print #nfile, "Folio1;" & Trim(sSerieDocto) & "-" & sFactura
   Print #nfile, "Fecha_FACTURA;" & sFecha  '& Format(Now, "YYYY-MM-DD") & "T" & minutos  '
   Print #nfile, "Fecha_FACTURA1;" & doFormatDate(Format(sFecha, "YYYY-MM-DD")) & " " & minutos
'   Print #nfile, "NO. DE APROBACION;" & sNoAprob
'   Print #nfile, "ANNO;" & sAnoAprob
   Print #nfile, "Certificado;" & sNoCert '6
'   Print #nfile, "BodegaNumero;" & sCveBod
'   Print #nfile, "BodegaNombre;" & sDescBodega
   Print #nfile, "CodBar_Encabezado;" & Trim(sSerieDocto) & sFactura
   Print #nfile, "VersionCFD;" & versioncfd '1
   Print #nfile, "TipoCambio;1" '12
   Print #nfile, "moneda;MXN" '11
   Print #nfile, "FormaPago;" & metodop '5
   If cuentap = "" Or cuentap = "0" Or cuentap = "0000" Then
      Print #nfile, "NumeroCuentaCMPV;    " '17
   Else
      Print #nfile, "NumeroCuentaCMPV;" & cuentap '17
   End If
   
   Print #nfile, "EmpresaRazonSocial;" & sEmpRazonSocial '26
   Print #nfile, "EmpresaRFC;" & sEmpRFC   '25
   Print #nfile, "EmpresaCalle;" & sEmpCalle '28
   Print #nfile, "EmpresaNumExterior;" & sEmpNoExt '29
   Print #nfile, "EmpresaNumInterior;" & sEmpNoInt '30
   Print #nfile, "EmpresaColonia;" & sEmpCol '31
   Print #nfile, "EmpresaCodigoPostal;" & sEmpCP '37
   Print #nfile, "EmpresaLocalidad;" & sEmpLocalidad '32
   Print #nfile, "EmpresaMunicipio;" & sEmpMun '34
   Print #nfile, "EmpresaEstado;" & sEmpEdo '35
   Print #nfile, "EmpresaPais;" & sEmpPais '36
   Print #nfile, "EmpresaTel;" & sEmpTel
   Print #nfile, "EmpresaPaginaWeb;" & sEmpPagWeb
   Print #nfile, "EmpresaDomicilio1;" & sEmpCalle & " " & sEmpNoExt & IIf(sEmpNoInt <> "", " INT ", "") & sEmpNoInt
   Print #nfile, "EmpresaColonia1;" & sEmpCol & " CP." & sEmpCP
   Print #nfile, "EmpresaCiudad1;" & sEmpMun & ", " & sEmpEdo
   Print #nfile, "RegimenFiscal;601" '27
   
   Print #nfile, "EmpresaExpCalle;" & sExpCalle '38
   Print #nfile, "EmpresaExpNumExterior;" & sExpNoExt '39
   Print #nfile, "EmpresaExpNumInterior;" & sExpNoInt '40
   Print #nfile, "EmpresaExpColonia;" & sExpCol '41
   Print #nfile, "EmpresaExpCodigoPostal;" & sExpCP '47
   Print #nfile, "EmpresaExpLocalidad;" & sExpLocalidad '42
   Print #nfile, "EmpresaExpMunicipio;" & sExpMun '44
   Print #nfile, "EmpresaExpEstado;" & sExpEdo '45
   Print #nfile, "EmpresaExpPais;" & sExpPais '46
   Print #nfile, "EmpresaExpTel;" & sExpTel
   Print #nfile, "EmpresaExpDomicilio1;" & sExpCalle & " " & sExpNoExt & " " & sExpNoInt
   Print #nfile, "EmpresaExpColonia1;" & sExpCol & " CP." & sExpCP
   Print #nfile, "EmpresaExpCiudad1;" & sExpCP
   
   Print #nfile, "RCliente;" & sRCliente
   Print #nfile, "RFC A QUIEN SE EXPIDE;" & sRRFC '48
   Print #nfile, "Rnombre;" & sRNombre '49
'   Print #nFile, "RApPaterno;" & sRApPat
'   Print #nFile, "RApMaterno;" & sRAPMat
   Print #nfile, "Rcalle;" & sRCalle '53
   Print #nfile, "Rnumero exterior;" & sRNoExt '54
   Print #nfile, "Rnumero interior;" & sRNoInt '55
   Print #nfile, "REntreCalles;" & sREntreCalles '58
   Print #nfile, "Rlocalidad;" & sRLocalidad '57
   Print #nfile, "Rcolonia;" & sRCol '56
   Print #nfile, "Rmunicipio;" & sRMun '59
   Print #nfile, "Restado;" & sREdo '60
   Print #nfile, "Rpais;" & sRPais '61
   Print #nfile, "Rcp;" & sRCP '62
   Print #nfile, "RDomicilio1;" & sRCalle & " " & sRNoExt
   Print #nfile, "RColonia1;" & sRCol & " CP." & sRCP
   Print #nfile, "RTelefonoNCR;" & SETel
   Print #nfile, "UsoCFDI;" & UsoCFDI '52
   
   Print #nfile, "ECliente;" & sECliente
   Print #nfile, "ENVIAR a;" & sENombre
   Print #nfile, "Ecalle;" & sECalle
   Print #nfile, "Enumero exterior;" & sENoExt
   Print #nfile, "Enumero interior;" & sENoInt
   Print #nfile, "EEntreCalles;" & sEEntreCalles
   Print #nfile, "Elocalidad;" & sELocalidad
   Print #nfile, "Ecolonia;" & sECol
   Print #nfile, "Emunicipio;" & sEMun
   Print #nfile, "Eestado;" & sEEdo
   Print #nfile, "Epais;" & sEPais
   Print #nfile, "Ecp Receptor;" & sECP
   Print #nfile, "EguiaRoji;" & sEGuia
   Print #nfile, "EDomicilio1;" & sECalle & " " & sENoExt
   Print #nfile, "EColonia1;" & sECol & " CP." & sECP
   
'   Print #nfile, "Ruta;" & sRuta
   Print #nfile, "Pedido;" & sPedido
   Print #nfile, "Condiciones;CONTADO" '7
   Print #nfile, "Vence;" & sFechaVence
   Print #nfile, "Importe;" & nTotalImporte '8
   Print #nfile, "TIPO DE COMPROBANTE;I" '14
   Print #nfile, "MetodoPago;" & sMetodoDePago '15
   Print #nfile, "LugarExpedicion;" & sExpCP '16
   
   
   If aDescripcion(1) = "Winko Mart" Then
        Print #nfile, "VersionCMPV;" & sVersionCMPV
        Print #nfile, "TipoOperacionCMPV;" & sTipoOperacionCMPV
        Print #nfile, "RegistroPatronalCMPV;" & sRegistroPatronalCMPV
'        Print #nfile, "NumeroCuentaCMPV;" & sNumeroCuentaCMPV
        Print #nfile, "TotalCMPV;" & sTotalCMPV
   End If
   
   Return
   
doGenDet:
   sRutina = "doGenDet"
   If nTipoArch = 1 Then
      For lngI = 1 To nLineas
         If nTipoPed = 12 Then
            If lngI = 1 Then
               sFolioIni = aFolIni(lngI)
            Else
               sFolioIni = Space(13)
            End If
            If lngI = nUltReg Then
               sFolioFin = aFolFin(lngI)
            Else
               sFolioFin = Space(13)
            End If
         Else
            sFolioIni = 0
            sFolioFin = 0
         End If
         
         sDesglose = ""
        Print #nfile, "CANTIDAD;1" '65
        Print #nfile, "ClaveProdServ;84141602" '63
        Print #nfile, "NoIdentificacion;7500000000000" '64
        Print #nfile, "ClaveUnidad;E48" '66
        Print #nfile, "U.M.;Unidad de servicio" '67
        Print #nfile, "DESCRIPCION;" & UCase(aDescripcion(lngI)) & " CARGA DE SALDOS" '68
        
        Print #nfile, "ValorUnitarioDisp;" & Round((ntotalPedido / 1.16), 2)
        Print #nfile, "ValorIVADisp;" & Round(((ntotalPedido / 116) * 16), 2)
        
        Print #nfile, "PRECIO;" & ntotalPedido '69
        Print #nfile, "IMPORTE BRUTO;" & ntotalPedido '70
        
      Next
   ElseIf nTipoArch = 2 Then
      For lngI = 1 To nLineas
         
         sDesglose = ""
        Print #nfile, "CANTIDAD;1" '65
        Print #nfile, "ClaveProdServ;84141602" '63
        Print #nfile, "NoIdentificacion;7500000000000" '64
        Print #nfile, "ClaveUnidad;E48" '66
        Print #nfile, "U.M.;Unidad de servicio" '67
        Print #nfile, "DESCRIPCION;" & aDescripcion(lngI) '68
        Print #nfile, "PRECIO;" & ntotalPedido '69
        Print #nfile, "IMPORTE BRUTO;" & ntotalPedido '70
             
      Next
   End If
   
   If nImpComision > 0 Then
         
        Print #nfile, "CANTIDAD;1" '65
        Print #nfile, "ClaveProdServ;84141602" '63
        Print #nfile, "NoIdentificacion;7500000000000" '64
        Print #nfile, "ClaveUnidad;E48" '66
        Print #nfile, "U.M.;Unidad de servicio" '67
        Print #nfile, "DESCRIPCION;CARGO ADMINISTRATIVO"  '68
        Print #nfile, "PRECIO;" & nImpComision '69
        Print #nfile, "IMPORTE BRUTO;" & nImpComision '70
       
   End If
        
   If nTotalIVA > 0 Then
       
        Print #nfile, "BaseImpTras;" & nImpComision '72
        Print #nfile, "ClaveImpTras;002" '73
        Print #nfile, "TipoFactorImpTras;Tasa" '74
        Print #nfile, "TasaCuotaImpTras;" & nPrcIVA / 100 '75
        Print #nfile, "ImporteImpTras;" & nImpIVAComision '76
        'Total Impuestos Trasladados
        Print #nfile, "TotalImpuestosTrasladados;" & nImpIVAComision '97
        Print #nfile, "TotalServ;" & nImpComision + nImpIVAComision

        
   End If
   
   ' Genero detalle CMPV
    If aDescripcion(1) = "Winko Mart" Then
       strsql = "sp_GenArchFE_Det " & nBodega & ", " & _
                           "'" & sserie & "', " & _
                           nFactIni
        Set rstTmp2 = New ADODB.Recordset
        rstTmp2.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rstTmp2.EOF
           Print #nfile, "IdentificadorCMPV;" & rstTmp2!IdentificadorCMPV
           Print #nfile, "FechaCMPV;" & Format(rstTmp2!FechaCMPV, "yyyy-mm-ddThh:mm:ss")
           Print #nfile, "RfcCMPV;" & rstTmp2!RfcCMPV
           Print #nfile, "CurpCMPV;" & rstTmp2!CurpCMPV
           Print #nfile, "NombreCMPV;" & rstTmp2!NombreCMPV
           Print #nfile, "NumSeguridadSocialCMPV;" & rstTmp2!NumSeguridadSocialCMPV
           Print #nfile, "ImporteCMPV;" & rstTmp2!ImporteCMPV
           rstTmp2.MoveNext
        Loop
    End If
   Return
   
doGenPie:
   sRutina = "doGenPie"
    Print #nfile, "Producto;" & aDescripcion(1)
    Print #nfile, "ImporteTotPer;" & sTotalCMPV
    Print #nfile, "TotalImporte;" & ntotalPedido + nImpComision
    If nImpIVAComision > 0 Then
        Print #nfile, "TotalIva;" & nImpIVAComision
        Print #nfile, "TipoImpuesto;002"
        Print #nfile, "TipoFactor;Tasa"
        Print #nfile, "TasaIva;" & nPrcIVA / 100
    End If
    Print #nfile, "TotalGeneral;" & ntotalPedido + nImpComision + nImpIVAComision

    
      
   If Producto_factura >= 1 Then
      Print #nfile, "LeyendaObservaciones;" & sComentario
   Else
      Print #nfile, "LeyendaObservaciones;" & nofiscal
   End If
   Print #nfile, "LeyendaPagare;" & sLeyendaPagare
   Print #nfile, "IMPORTE CON LETRA;" & sImporteLetras
   Print #nfile, "PagareDatosCliente;" & sPagareDatosCliente
   Print #nfile, "PagareCiudadContrato;" & sPagareCiudadContrato

   Return
   
hdErr:
   Call doErrorLog(gnSucursal, "OPE", ERR.Number, ERR.Description, gstrUsuario, "Module1.doGenArchFE" & IIf(sRutina <> "", "." & sRutina, ""))
   If blnFileOpen Then
      Close #nfile
   End If
   MsgBox ERR.Description
'   Resume
End Function

Public Function doFindPrinter(ByVal sMaquina As String, ByVal sTipo As String) As String
Dim strsql As String, rstTmp As ADODB.Recordset, strPuerto As String

   strsql = "sp_Impresoras_Sel '" & sMaquina & "', " & _
                              "'" & sTipo & "'"
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   If Not rstTmp.EOF Then
      strPuerto = Trim(rstTmp!Impresora & "")
      Impresora = Trim(rstTmp!Impresora & "")
      If strPuerto <> "LPT1" Then
        Impresora = Mid(Left(Impresora, InStr(3, Impresora, "\") - 1), 3)
        strPuerto = IIf(strPuerto = "", "LPT1", strPuerto)
      End If
   Else
      strPuerto = "LPT1"
   End If
   rstTmp.Close
   Set rstTmp = Nothing
   
   doFindPrinter = strPuerto
   
End Function
Function DeleteChar(ByVal mCh As String, ByVal mStr As String)
Dim i       As Integer
Dim mStrNew As String
   mStrNew = ""
   For i = 1 To Len(mStr)
      If Mid$(mStr, i, 1) <> mCh Then mStrNew = mStrNew & Mid$(mStr, i, 1)
   Next i
   DeleteChar = mStrNew
End Function
Public Function doGenArchFE_OI(ByVal nBodega As Integer, _
                         ByVal sserie As String, _
                         ByVal nFactIni As Long, _
                         ByVal nFactFin As Long) As Boolean
Dim nfile As Long, SFILE As String, blnFileOpen As Boolean
Dim blnOk As Boolean
Dim strsql As String, rstTmp As ADODB.Recordset
Dim sRutina As String
'Variables para el Do While
Dim nFactAnt As Long, nNR As Long, nCantFact As Long
Dim nLineas As Long, nUltReg As Long
Dim blnAgrupa As Boolean, blnAdd As Boolean
'nRegAct As Long,
Dim nTipoPed As Long
Dim lngI As Long
'Variables para denominaciones raras
Dim SCodigo As String, sDescripcion As String, _
   nUnidades As Long, nIvaProd As Double, _
   nImporte As Double, nIva As Double, nTotal As Double
Dim sFolioIni As String, sFolioFin As String
Dim nPrecio As Double, nPrecioSug As Double
Dim sDesglose As String
'Variable para Agrupar Prefacturas
Dim sBodega As String, nPrefSig As Double, nFolSig As Double
'Documento
Dim sSerieDocto As String, sFactura As String, sFecha As String, _
      sNoAprob As String, sAnoAprob As String, sNoCert As String, _
      sCveBod As String, sDescBodega As String
      
'Empresa
Dim sEmpRazonSocial As String, sEmpRFC As String, sEmpCalle As String, _
   sEmpNoExt As String, sEmpNoInt As String, sEmpCol As String, _
   sEmpCP As String, sEmpLocalidad As String, sEmpMun As String, _
   sEmpEdo As String, sEmpPais As String, sEmpTel As String, _
   sEmpPagWeb As String, metodop As String, cuentap As String, other_fecha As Date, _
   sMetodoDePago As String

Dim UsoCFDI As String

'Expide
Dim sExpCalle As String, sExpNoExt As String, sExpNoInt As String, _
   sExpCol As String, sExpCP As String, sExpLocalidad As String, _
   sExpMun As String, sExpEdo As String, sExpPais As String, _
   sExpTel As String
'Cliente a Facturar
Dim sRCliente As String, sRRFC As String, _
   sRNombre As String, _
   sRCalle As String, sRNoExt As String, sRNoInt As String, _
   sREntreCalles As String, sRLocalidad As String, sRCol As String, _
   sRMun As String, sREdo As String, sRPais As String, _
   sRCP As String
', sRApPat As String, sRAPMat As String
'Cliente a Enviar
Dim sECliente As String, sENombre As String, _
   sECalle As String, sENoExt As String, sENoInt As String, _
   sEEntreCalles As String, sELocalidad As String, sECol As String, _
   sEMun As String, sEEdo As String, sEPais As String, _
   sECP As String, sEGuia As String, SETel As String
'Encabezado
Dim sPedido As String, sRuta As String, sCondiciones As String, sFechaVence As String
'Variables de la Factura
Dim blnGravIVA As Boolean
Dim sRef As String, sRefBonos As String, sPedRef
Dim sComentario As String, sImporteLetras As String
Dim sComprobante As String
Dim sLeyendaPagare As String, sPagareDatosCliente As String, sPagareCiudadContrato As String
'Detalle de la Factura
Dim aCodigo() As String, aDescripcion() As String, _
   aUnidades() As Long, aPrecio() As Double, aPrecioSug() As Double, _
   aFolIni() As String, aFolFin() As String, _
   aImporte() As Double, aIVA() As Double, aTotal() As Double
'Detalle de la Factura (Prefacturas)
Dim aSerie() As String, aDescBod() As String, _
      aPrefIni() As Double, aPrefFin() As Double, aAnoMes() As String
'Datos de la Comision
Dim nImpComision As Double, nImpIVAComision As Double
Dim nPrcCom As Double, nPrcComCalc As Double, nPrcIVA As Double
'Totales
Dim nTotalUnidades As Long, _
   ntotalPedido, nTotalImporte As Double, nTotalIVA As Double, nTotalGen As Double
Dim sImpresora As String, minutos As String
'  1.- Factura Normal o de Stock
'  2.- Factura de Prefacturas

   On Error GoTo hdErr
   blnImprime = True
   If blnImprime Then
      sImpresora = doFindPrinter(gstrPC, "FACTURAS")
   Else
      sImpresora = ""
   End If
   
   blnOk = False
   
   strsql = "sp_GenArchFE_OI " & nBodega & ", " & _
                           "'" & sserie & "', " & _
                           nFactIni & ", " & _
                           nFactFin
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   Do While Not rstTmp.EOF
      blnOk = True
      If nFactAnt <> rstTmp!Folio Then
         If nFactAnt <> 0 Then
            If nNR > 0 Or sRef <> "" Then
               nCantFact = nCantFact + 1
               If nTipoArch = 1 Then
'                 Factura Normal
'                 Si tiene fracciones agregar registro
                  If nUnidades > 0 Then
                     ReDim Preserve aCodigo(nLineas)
                     ReDim Preserve aDescripcion(nLineas)
                     ReDim Preserve aUnidades(nLineas)
                     ReDim Preserve aPrecio(nLineas)
                     ReDim Preserve aPrecioSug(nLineas)
                     ReDim Preserve aFolIni(nLineas)
                     ReDim Preserve aFolFin(nLineas)
                     ReDim Preserve aImporte(nLineas)
                     ReDim Preserve aIVA(nLineas)
                     ReDim Preserve aTotal(nLineas)
                  
                     aCodigo(nLineas) = SCodigo
                     aDescripcion(nLineas) = sDescripcion
                     aUnidades(nLineas) = nUnidades
                     aPrecio(nLineas) = 0
                     aPrecioSug(nLineas) = 0
                     aFolIni(nLineas) = ""
                     aFolFin(nLineas) = ""
                     aImporte(nLineas) = nImporte
                     aIVA(nLineas) = nIva
                     aTotal(nLineas) = nTotal
                     
                     nUltReg = nLineas - 1
                  Else
                     nLineas = nLineas - 1
                     nUltReg = nLineas
                  End If
               ElseIf nTipoArch = 2 Then
                  nLineas = nLineas - 1
               End If
               
               GoSub doGenFile
            End If
         End If
         nFactAnt = rstTmp!Folio
'         nRegAct = 1
         nLineas = 1
         nTipoArch = 1
         If nTipoArch = 1 Then
'           Reinicia valores para agrupar denominaciones raras
            SCodigo = ""
            sDescripcion = ""
            nUnidades = 0
            nImporte = 0
            nIva = 0
            nTotal = 0
         Else
'           Reinicia valores para agrupar prefacturas
            sBodega = ""
            nPrefSig = 0
            nFolSig = 0
            
'           Reinicia Totales
            nTotalUnidades = 0
         End If
         
         GoSub doFillEnc
      End If
      
      GoSub doFillDet
      
'      nRegAct = nRegAct + 1
      rstTmp.MoveNext
   Loop
   
   nNR = 1
   If nNR > 0 Or sRef <> "" Then
      nCantFact = nCantFact + 1
      If nTipoArch = 1 Then
'        Si tiene fracciones agregar registro
         If nUnidades > 0 Then
            ReDim Preserve aCodigo(nLineas)
            ReDim Preserve aDescripcion(nLineas)
            ReDim Preserve aUnidades(nLineas)
            ReDim Preserve aPrecio(nLineas)
            ReDim Preserve aPrecioSug(nLineas)
            ReDim Preserve aFolIni(nLineas)
            ReDim Preserve aFolFin(nLineas)
            ReDim Preserve aImporte(nLineas)
            ReDim Preserve aIVA(nLineas)
            ReDim Preserve aTotal(nLineas)
         
            aCodigo(nLineas) = SCodigo
            aDescripcion(nLineas) = sDescripcion
            aUnidades(nLineas) = nUnidades
            aPrecio(nLineas) = 0
            aPrecioSug(nLineas) = 0
            aFolIni(nLineas) = ""
            aFolFin(nLineas) = ""
            aImporte(nLineas) = nImporte
            aIVA(nLineas) = nIva
            aTotal(nLineas) = nTotal
            
            nUltReg = nLineas - 1
         Else
            nLineas = nLineas - 1
            nUltReg = nLineas
         End If
      ElseIf nTipoArch = 2 Then
         nLineas = nLineas - 1
      End If
      
      'GeneraArchivo
      GoSub doGenFile
   End If
   
   On Error GoTo 0
   
   doGenArchFE_OI = blnOk
   Exit Function

doFillEnc:
   
   blnAgrupa = (rstTmp!NR > 20)
         
   sRutina = "doFillEnc"
   sSerieDocto = Trim(rstTmp!Serie)
   sFactura = rstTmp!Folio
   
   minutos = Format(rstTmp!Fecha_Factura, "hh:mm:ss")
   sFecha = Format(rstTmp!Fecha_Factura, "YYYY-MM-DD")
   sFecha = sFecha & "T" & minutos
   sNoAprob = rstTmp!NoAprob
   sAnoAprob = rstTmp!AnoAprob
   sNoCert = rstTmp!NoCert
   sCveBod = rstTmp!BodegaNumero
   sDescBodega = rstTmp!BodegaNombre
   versioncfd = rstTmp!versioncfd
   metodop = Format(rstTmp!MetodoPago, "00")
   cuentap = rstTmp!CuentaPago
   UsoCFDI = rstTmp!UsoCFDI
   
   sEmpRazonSocial = rstTmp!EmpRazonSocial
   sEmpRFC = rstTmp!EmpRFC
   sEmpCalle = rstTmp!EmpCalle
   sEmpNoExt = rstTmp!EmpNoExt
   sEmpNoInt = rstTmp!EmpNoInt
   sEmpCol = rstTmp!EmpColonia
   sEmpCP = rstTmp!EmpCP
   sEmpLocalidad = rstTmp!EmpLocalidad
   sEmpMun = rstTmp!EmpMunicipio
   sEmpEdo = rstTmp!EmpEstado
   sEmpPais = rstTmp!EmpPais
   sEmpTel = rstTmp!EmpTel
   sEmpPagWeb = rstTmp!EmpPagWeb
   
   sExpCalle = rstTmp!ExpCalle
   sExpNoExt = rstTmp!ExpNoExt
   sExpNoInt = rstTmp!ExpNoInt
   sExpCol = rstTmp!ExpColonia
   sExpCP = rstTmp!ExpCP
   sExpLocalidad = rstTmp!EmpLocalidad
   sExpMun = rstTmp!ExpMunicipio
   sExpEdo = rstTmp!ExpEstado
   sExpPais = rstTmp!ExpPais
   sExpTel = rstTmp!ExpTel
   
   sRCliente = rstTmp!RCliente
   sRRFC = DeleteChar("-", DeleteChar(" ", rstTmp!RRFC))
   sRNombre = rstTmp!RNombre & " " & rstTmp!RApPat & " " & rstTmp!RAPMat
  'sRAPPat = rstTmp!RApPat
   'sRAPMat = rstTmp!RAPMat
   sRCalle = rstTmp!RCalle
   sRNoExt = IIf(IsNull(rstTmp!RNoExt), "", rstTmp!RNoExt)
   sRNoInt = rstTmp!RNoInt
   sREntreCalles = rstTmp!REntreCalles
   sRLocalidad = rstTmp!RLocalidad
   sRCol = rstTmp!Rcolonia
   sRMun = rstTmp!RMunicipio
   sREdo = rstTmp!REstado
   sRPais = rstTmp!RPais
   sRCP = rstTmp!RCP
   
   sECliente = rstTmp!Ecliente
   sENombre = rstTmp!ENombre
   sECalle = rstTmp!ECalle
   sENoExt = rstTmp!ENoExt
   sENoInt = rstTmp!ENoInt
   sEEntreCalles = rstTmp!EEntreCalles
   sELocalidad = rstTmp!ELocalidad
   sECol = rstTmp!EColonia
   sEMun = rstTmp!EMunicipio
   sEEdo = rstTmp!EEstado
   sEPais = rstTmp!EPais
   sECP = rstTmp!ECP
   If nTipoArch = 1 Then
            SETel = rstTmp!ETel
   End If
   sRuta = rstTmp!ruta
   sCondiciones = ""
   sFechaVence = ""
   
   nPrcIVA = rstTmp!PrcIVA
   blnGravIVA = False ' IIf(rstTmp!GravIVA = "S", True, False)
   
   nImpComision = rstTmp!Comision
   nImpIVAComision = Round(rstTmp!ivacomision, 2)
   nPrcCom = rstTmp!PorcComision
   nPrcComCalc = 0
   nTotalImporte = rstTmp!TotalImporte
   nTotalIVA = rstTmp!TotalIva
   nTotalGen = rstTmp!TotalGeneral
   
   If nTipoArch = 1 Then
      nTotalUnidades = 1
      sComprobante = rstTmp!Comprobante
   End If
   sImporteLetras = Leyenda(rstTmp!TotalGeneral)
   sLeyendaPagare = "Por el presente PAGARE me(nos) obligo(amos) incondicionalmente a pagar " & _
                    "a la orden de " & sEmpRazonSocial & " en esta plaza en moneda nacional el dia " & _
                    Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4) & _
                    " la cantidad de " & Format(nTotalGen, "###,###,##0.00") & " (" & sImporteLetras & ") " & _
                    "valor en efectivo. Si no fuere pagado satisfactoriamente este pagaré me(nos) obligo(amos) " & _
                    "a pagar durante todo el tiempo que permaneciera total o parcialmente insoluto, " & _
                    "un interés legal de tasa TIIE + 5 puntos, sin que por esto se considere prorrogado " & _
                    "el plazo fijado para el cumplimiento de esta obligación."
   sPagareDatosCliente = sRNombre & " con Domicilio en: " & sRCalle & " " & sRNoExt & _
                         " Colonia: " & sRCol & " en " & sRMun & ", " & sREdo
   sPagareCiudadContrato = "Lugar y Fecha de Expedición: " & sExpMun & ", " & sExpEdo & ", a " & _
                           Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4)
   
   sMailFE = IIf(rstTmp!MailFe = 1, "MAIL", "")
   sMailFETo = rstTmp!MailFETo
   sMailFETo = Replace(sMailFETo, ";", ",")
   Return

doFillDet:
   sRutina = "doFillDet"
   If nTipoArch = 1 Then
      If blnAgrupa _
               And (rstTmp!Precio <> 1 _
               And rstTmp!Precio <> 2 _
               And rstTmp!Precio <> 3 _
               And rstTmp!Precio <> 4 _
               And rstTmp!Precio <> 5 _
               And rstTmp!Precio <> 6 _
               And rstTmp!Precio <> 7 _
               And rstTmp!Precio <> 8 _
               And rstTmp!Precio <> 9 _
               And rstTmp!Precio <> 10 _
               And rstTmp!Precio <> 20 _
               And rstTmp!Precio <> 30 _
               And rstTmp!Precio <> 50 _
               And rstTmp!Precio <> 100 _
               And rstTmp!Precio <> 200) Then
   '     Agregarlo al final a un arreglo para al final imprimir
         SCodigo = rstTmp!codigo
         sDescripcion = rstTmp!descripcion
         nUnidades = nUnidades + rstTmp!Unidades
         nImporte = nImporte + Round((rstTmp!total / (1 + nIvaProd)), 2)
         nIva = nIva + (rstTmp!total - Round((rstTmp!total / (1 + nIvaProd)), 2))
         nTotal = nTotal + rstTmp!total
      Else
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aPrecio(nLineas)
         ReDim Preserve aPrecioSug(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = Trim(rstTmp!codigo)
         aDescripcion(nLineas) = rstTmp!Comprobante ' era rstTmp!Descripcion
         aUnidades(nLineas) = rstTmp!Unidades
         aPrecio(nLineas) = Round((rstTmp!Precio / (1 + nIvaProd)), 2)
         aPrecioSug(nLineas) = rstTmp!Precio
         aImporte(nLineas) = rstTmp!total
         aIVA(nLineas) = rstTmp!ivacomision
         aTotal(nLineas) = rstTmp!total
         
         nLineas = nLineas + 1
      End If
   ElseIf nTipoArch = 2 Then
      If blnAgrupa Then
         If sBodega = rstTmp!descBodega & "" And _
               nPrefSig = rstTmp!NoPreFactura And _
               nFolSig = Val(Right(Trim(rstTmp!FolioIni & ""), 7)) Then
            aPrefFin(nLineas - 1) = rstTmp!NoPreFactura
            aFolFin(nLineas - 1) = Right(Trim(rstTmp!FolioFin & ""), 7)
            aUnidades(nLineas - 1) = aUnidades(nLineas - 1) + rstTmp!Unidades
            aImporte(nLineas - 1) = aImporte(nLineas - 1) + rstTmp!importe
            aIVA(nLineas - 1) = aIVA(nLineas - 1) + rstTmp!iva
            aTotal(nLineas - 1) = aTotal(nLineas - 1) + rstTmp!total
            blnAdd = False
         Else
            blnAdd = True
'            nRegAct = nRegAct + 1
         End If
         
         sBodega = rstTmp!descBodega & ""
         nPrefSig = rstTmp!NoPreFactura + 1
         nFolSig = Val(Right(Trim(rstTmp!FolioFin & ""), 7)) + 1
      Else
         blnAdd = True
      End If
      
      If blnAdd Then
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         
         ReDim Preserve aSerie(nLineas)
         ReDim Preserve aPrefIni(nLineas)
         ReDim Preserve aPrefFin(nLineas)
         ReDim Preserve aDescBod(nLineas)
         ReDim Preserve aAnoMes(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = rstTmp!codigo
         aDescripcion(nLineas) = rstTmp!descripcion
         
         aSerie(nLineas) = Trim(rstTmp!SeriePreFactura & "")
         aPrefIni(nLineas) = rstTmp!NoPreFactura
         aPrefFin(nLineas) = rstTmp!NoPreFactura
         aDescBod(nLineas) = rstTmp!descBodega & ""
         aAnoMes(nLineas) = Left(Trim(rstTmp!FolioIni & ""), 6)
         aFolIni(nLineas) = Right(Trim(rstTmp!FolioIni & ""), 7)
         aFolFin(nLineas) = Right(Trim(rstTmp!FolioFin & ""), 7)
         
         aUnidades(nLineas) = rstTmp!Unidades
         aImporte(nLineas) = rstTmp!importe 'rstTmp!Precio
         aIVA(nLineas) = rstTmp!iva
         aTotal(nLineas) = rstTmp!total
      
         nLineas = nLineas + 1
      End If
      
      nTotalUnidades = nTotalUnidades + rstTmp!Unidades
   End If
   Return
   
doGenFile:
   sRutina = "doGenFile"
   GoSub doFileOpen
   GoSub doGenEnc
   GoSub doGenDet
   GoSub doGenPie
   GoSub doFileClose
   Return
   
doFileOpen:
   sRutina = "doFileOpen"
   nfile = FreeFile()
   SFILE = gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT"
   Open SFILE For Output As #nfile
   blnFileOpen = True
   Return
   

doFileClose:
   sRutina = "doFileClose"
   Close #nfile
   blnFileOpen = False
   GoSub doFileMove
   Return
   
doFileMove:
   sRutina = "doFileMove"
   Call doWaitShell("MOVE " & gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT" & " C:\GoDir\WsHome\datos\salida")
   Return
   
   
doGenEnc:
   sRutina = "doGenEnc"
   Print #nfile, "TFormato;" & "VTOFAC"
   Print #nfile, "TIPO_DOCTO;" & "FACTURA"
   Print #nfile, "Ndistrib;" & "FE,MAIL,,NO," & IIf(blnImprime, "SI", "NO")
   '---------------------------
   If sMailFE = "MAIL" Then
      Print #nfile, "COPIAS;1,,"
   Else
      Print #nfile, "COPIAS;2,,"
   End If
   '--------------
   Print #nfile, "Impresora;" & sImpresora
   '--------------
   
   If sMailFE = "MAIL" Then
      Print #nfile, "CorreoTo;" & sMailFETo
      Print #nfile, "CorreoSub;" & "Envio de Factura Electronica " & sSerieDocto & "-" & sFactura & IIf(sPedido = "0" Or sPedido = "", "", " del Pedido " & sPedido)
      Print #nfile, "CorreoAtt;PDF,XML"
   End If
       
   Print #nfile, "Serie;" & sSerieDocto
   Print #nfile, "Folio;" & sFactura
   Print #nfile, "Folio1;" & Trim(sSerieDocto) & "-" & sFactura
   Print #nfile, "Fecha_FACTURA;" & sFecha  '& Format(Now, "YYYY-MM-DD") & "T" & minutos '
   Print #nfile, "Fecha_FACTURA1;" & doFormatDate(sFecha, 0) & " " & minutos
'   Print #nfile, "NO. DE APROBACION;" & sNoAprob
'   Print #nfile, "ANNO;" & sAnoAprob
   Print #nfile, "Certificado;" & sNoCert
'   Print #nfile, "BodegaNumero;" & sCveBod
'   Print #nfile, "BodegaNombre;" & sDescBodega
   Print #nfile, "CodBar_Encabezado;" & Trim(sSerieDocto) & sFactura
   Print #nfile, "VersionCFD;" & versioncfd
   Print #nfile, "TipoCambio;1" '12
   Print #nfile, "moneda;MXN"
   Print #nfile, "FormaPago;" & metodop
   If cuentap = "" Or cuentap = "0" Or cuentap = "0000" Then
       Print #nfile, "NumeroCuentaCMPV;    "
   Else
       Print #nfile, "NumeroCuentaCMPV;" & Trim(cuentap)
   End If
   
   Print #nfile, "EmpresaRazonSocial;" & sEmpRazonSocial
   Print #nfile, "EmpresaRFC;" & sEmpRFC
   Print #nfile, "EmpresaCalle;" & sEmpCalle
   Print #nfile, "EmpresaNumExterior;" & sEmpNoExt
   Print #nfile, "EmpresaNumInterior;" & sEmpNoInt
   Print #nfile, "EmpresaColonia;" & sEmpCol
   Print #nfile, "EmpresaCodigoPostal;" & sEmpCP
   Print #nfile, "EmpresaLocalidad;" & sEmpLocalidad
   Print #nfile, "EmpresaMunicipio;" & sEmpMun
   Print #nfile, "EmpresaEstado;" & sEmpEdo
   Print #nfile, "EmpresaPais;" & sEmpPais
   Print #nfile, "EmpresaTel;" & sEmpTel
   Print #nfile, "EmpresaPaginaWeb;" & sEmpPagWeb
   Print #nfile, "EmpresaDomicilio1;" & sEmpCalle & " " & sEmpNoExt & IIf(sEmpNoInt <> "", " INT ", "") & sEmpNoInt
   Print #nfile, "EmpresaColonia1;" & sEmpCol & " CP." & sEmpCP
   Print #nfile, "EmpresaCiudad1;" & sEmpMun & ", " & sEmpEdo
   Print #nfile, "RegimenFiscal;601"
   
   Print #nfile, "EmpresaExpCalle;" & sExpCalle
   Print #nfile, "EmpresaExpNumExterior;" & sExpNoExt
   Print #nfile, "EmpresaExpNumInterior;" & sExpNoInt
   Print #nfile, "EmpresaExpColonia;" & sExpCol
   Print #nfile, "EmpresaExpCodigoPostal;" & sExpCP
   Print #nfile, "EmpresaExpLocalidad;" & sExpLocalidad
   Print #nfile, "EmpresaExpMunicipio;" & sExpMun
   Print #nfile, "EmpresaExpEstado;" & sExpEdo
   Print #nfile, "EmpresaExpPais;" & sExpPais
   Print #nfile, "EmpresaExpTel;" & sExpTel
   Print #nfile, "EmpresaExpDomicilio1;" & sExpCalle & " " & sExpNoExt & " " & sExpNoInt
   Print #nfile, "EmpresaExpColonia1;" & sExpCol & " CP." & sExpCP
   Print #nfile, "EmpresaExpCiudad1;" & sExpCP
   
   Print #nfile, "RCliente;" & sRCliente
   Print #nfile, "RFC A QUIEN SE EXPIDE;" & sRRFC
   Print #nfile, "Rnombre;" & sRNombre
   'Print #nFile, "RApPaterno;" & sRAPPat
   'Print #nFile, "RApMaterno;" & sRAPMat
   Print #nfile, "Rcalle;" & sRCalle
   Print #nfile, "Rnumero exterior;" & sRNoExt
   Print #nfile, "Rnumero interior;" & sRNoInt
   Print #nfile, "REntreCalles;" & sREntreCalles
   Print #nfile, "Rlocalidad;" & sRLocalidad
   Print #nfile, "Rcolonia;" & sRCol
   Print #nfile, "Rmunicipio;" & sRMun
   Print #nfile, "Restado;" & sREdo
   Print #nfile, "Rpais;" & sRPais
   Print #nfile, "Rcp;" & sRCP
   Print #nfile, "RDomicilio1;" & sRCalle & " " & sRNoExt
   Print #nfile, "RColonia1;" & sRCol & " CP." & sRCP
   Print #nfile, "RTelefonoNCR;" & SETel
   Print #nfile, "UsoCFDI;" & UsoCFDI '52
   
   
   Print #nfile, "ECliente;" & sECliente
   Print #nfile, "ENVIAR a;" & sENombre
   Print #nfile, "Ecalle;" & sECalle
   Print #nfile, "Enumero exterior;" & sENoExt
   Print #nfile, "Enumero interior;" & sENoInt
   Print #nfile, "EEntreCalles;" & sEEntreCalles
   Print #nfile, "Elocalidad;" & sELocalidad
   Print #nfile, "Ecolonia;" & sECol
   Print #nfile, "Emunicipio;" & sEMun
   Print #nfile, "Eestado;" & sEEdo
   Print #nfile, "Epais;" & sEPais
   Print #nfile, "Ecp Receptor;" & sECP
   Print #nfile, "EguiaRoji;" & sEGuia
   Print #nfile, "EDomicilio1;" & sECalle & " " & sENoExt
   Print #nfile, "EColonia1;" & sECol & " CP." & sECP
   
'   Print #nfile, "Ruta;" & sRuta
   Print #nfile, "Pedido;" & sPedido
   Print #nfile, "Condiciones;" & sCondiciones
   Print #nfile, "Vence;" & sFechaVence
   Print #nfile, "Importe;" & nTotalImporte '8
   Print #nfile, "TIPO DE COMPROBANTE;I" '14
        Print #nfile, "MetodoPago;PUE" '15
   Print #nfile, "LugarExpedicion;" & sExpCP '16

   Return
   
doGenDet:
   sRutina = "doGenDet"
   If nTipoArch = 1 Then
      For lngI = 1 To nLineas
         If nTipoPed = 12 Then
            If lngI = 1 Then
               sFolioIni = aFolIni(lngI)
            Else
               sFolioIni = Space(13)
            End If
            If lngI = nUltReg Then
               sFolioFin = aFolFin(lngI)
            Else
               sFolioFin = Space(13)
            End If
         Else
            sFolioIni = 0
            sFolioFin = 0
         End If
         
         sComprobante = ""
         sDesglose = ""
         
         Print #nfile, "CANTIDAD;" & aUnidades(lngI)
         Print #nfile, "ClaveProdServ;84141602" '63
         Print #nfile, "NoIdentificacion;7500000000000" '64
         Print #nfile, "ClaveUnidad;E48" '66
         'Print #nfile, "CODIGO;" & aCodigo(lngI)
         Print #nfile, "U.M.;Unidad de servicio" '67
         If Trim(aCodigo(lngI)) = "FACPRO" Then
            sComprobante = "COBRO DE COMISIONES SEGUN COMPROBANTE ANEXO NO. " & sComprobante
            Print #nfile, "DESCRIPCION;" & sComprobante
         ElseIf Trim(aCodigo(lngI)) = "COMBE" Then
            sComprobante = "COBRO DE COMISION DE LAS TRANSACCIONES " & sComprobante
            Print #nfile, "DESCRIPCION;" & sComprobante
         Else 'If Trim(aCodigo(lngI)) = "OI15" Or Trim(aCodigo(lngI)) = "OI10" Or Trim(aCodigo(lngI)) = "INTADM" Then
            sComprobante = "TARJETAS " & aDescripcion(lngI)
            Print #nfile, "DESCRIPCION;" & sComprobante
         End If
         Print #nfile, "PRECIO;" & aImporte(lngI) / aUnidades(lngI)
         Print #nfile, "IMPORTE BRUTO;" & aImporte(lngI)
                 
                If aIVA(lngI) > 0 Then
                                  
                        Print #nfile, "BaseImpTras;" & aImporte(lngI) '72
                        Print #nfile, "ClaveImpTras;002" '73
                        Print #nfile, "TipoFactorImpTras;Tasa" '74
                        Print #nfile, "TasaCuotaImpTras;" & nPrcIVA / 100 '75
                        Print #nfile, "ImporteImpTras;" & aIVA(lngI) '76
                        'Total Impuestos Trasladados
                        Print #nfile, "TotalImpuestosTrasladados;" & nImpIVAComision '97
                        Print #nfile, "TotalServ;" & aImporte(lngI) + nImpIVAComision
                
                End If
        
      Next
   ElseIf nTipoArch = 2 Then
      For lngI = 1 To nLineas
         sDesglose = aDescBod(lngI) & " " & _
                     Trim(aSerie(lngI)) & "-" & Trim(aPrefIni(lngI)) & IIf(aPrefIni(lngI) <> aPrefFin(lngI), " A " & Trim(aPrefFin(lngI)), "") & _
                     " " & aAnoMes(lngI) & _
                     " DEL " & aFolIni(lngI) & " AL " & aFolFin(lngI)
         
         nPrecio = Round(aImporte(lngI) / aUnidades(lngI), 2)
         Print #nfile, "CANTIDAD;" & aUnidades(lngI)
         Print #nfile, "ClaveProdServ;84141602" '63
         Print #nfile, "NoIdentificacion;7500000000000" '64
         Print #nfile, "CODIGO;" & aCodigo(lngI)
         Print #nfile, "U.M.;Unidad de servicio" '67

         Print #nfile, "ClaveUnidad;E48" '66
         Print #nfile, "DESCRIPCION;" & aDescripcion(lngI)
         Print #nfile, "PRECIO;" & nPrecio
         Print #nfile, "IMPORTE BRUTO;" & aImporte(lngI)

                 If aIVA(lngI) > 0 Then
                                  
                        Print #nfile, "BaseImpTras;" & aImporte(lngI) '72
                        Print #nfile, "ClaveImpTras;002" '73
                        Print #nfile, "TipoFactorImpTras;Tasa" '74
                        Print #nfile, "TasaCuotaImpTras;" & nPrcIVA / 100 '75
                        Print #nfile, "ImporteImpTras;" & aTotal(lngI) '76
                        'Total Impuestos Trasladados
                        Print #nfile, "TotalImpuestosTrasladados;" & nImpIVAComision '97
                        Print #nfile, "TotalServ;" & aImporte(lngI) + nImpIVAComision
                End If
    
      Next
   End If
'   If nTotalIVA > 0 Then
'      Print #nfile, "LeyendaDet;" & "Iva Tasa " & nPrcIVA & "%: " & Format(nTotalIVA, "###,###,##0.00")
'   End If
   
   Return
   
doGenPie:
   sRutina = "doGenPie"
 '  Print #nfile, "TotalUNIDADES;" & nTotalUnidades
   Print #nfile, "TotalImporte;" & nTotalImporte
   If nTotalIVA > 0 Then
       Print #nfile, "TotalIva;" & nTotalIVA
       Print #nfile, "TipoImpuesto;002"
       Print #nfile, "TipoFactor;Tasa"
       Print #nfile, "TasaIva;" & nPrcIVA / 100
   End If
   Print #nfile, "TotalGeneral;" & nTotalGen
   Print #nfile, "LeyendaObservaciones;" & sComentario
   Print #nfile, "LeyendaPagare;" & sLeyendaPagare
   Print #nfile, "IMPORTE CON LETRA;" & sImporteLetras
   Print #nfile, "PagareDatosCliente;" & sPagareDatosCliente
   Print #nfile, "PagareCiudadContrato;" & sPagareCiudadContrato

   Return
   
   
  
hdErr:
Call doErrorLog(gnSucursal, "OPE", ERR.Number, ERR.Description, gstrUsuario, "Module1.doGenArchFE_OI" & IIf(sRutina <> "", "." & sRutina, ""))
   If blnFileOpen Then
      Close #nfile
   End If
   Resume
   MsgBox ERR.Description
End Function
Public Function doFormatDate(ByVal sFecha As String, Optional ByVal nTipo As Integer = 0) As String
' 0 - Formato Corto (dd-MMM-yyyy)
' 1 - Formato Largo (dd-mmmmm-yyyy)
' 2 - Formato Minimo (dd-MM-yy)
' sFecha = 'YYYY-MM-DD'
Dim sDia As String, sMes As String, sAno As String
   sDia = Mid(sFecha, 9, 2)
   If nTipo = 2 Then
      sMes = Mid(sFecha, 6, 2)
      sAno = Mid(sFecha, 3, 2)
   Else
      sMes = getNombreMes(Mid(sFecha, 6, 2))
      sMes = IIf(nTipo = 0, Left(sMes, 3), sMes)
      sAno = Mid(sFecha, 1, 4)
   End If
   doFormatDate = sDia & "-" & sMes & "-" & sAno
End Function
Function Leyenda(ByVal valor As Variant, Optional Divisa As Variant) As String
Static v_uni(9) As String
Static v_dec(9) As String
Static v_cen(9) As String
Static v_die(9) As String
    
Dim xDivisa As Variant, ValLetras As Variant
Dim Deci As String
Dim valor2 As Variant
Dim Num As Double
Dim die, Can, divisor, Pos
    
   If IsMissing(Divisa) Then
      xDivisa = "N"
   Else
      xDivisa = UCase(IIf(Divisa <> "S", "N", Divisa))
   End If
   v_uni(1) = "UN"
   v_uni(2) = "DOS"
   v_uni(3) = "TRES"
   v_uni(4) = "CUATRO"
   v_uni(5) = "CINCO"
   v_uni(6) = "SEIS"
   v_uni(7) = "SIETE"
   v_uni(8) = "OCHO"
   v_uni(9) = "NUEVE"

   v_dec(1) = "DIEZ"
   v_dec(2) = "VEINTE"
   v_dec(3) = "TREINTA"
   v_dec(4) = "CUARENTA"
   v_dec(5) = "CINCUENTA"
   v_dec(6) = "SESENTA"
   v_dec(7) = "SETENTA"
   v_dec(8) = "OCHENTA"
   v_dec(9) = "NOVENTA"

   v_cen(1) = "CIENTO"
   v_cen(2) = "DOSCIENTOS"
   v_cen(3) = "TRESCIENTOS"
   v_cen(4) = "CUATROCIENTOS"
   v_cen(5) = "QUINIENTOS"
   v_cen(6) = "SEISCIENTOS"
   v_cen(7) = "SETECIENTOS"
   v_cen(8) = "OCHOCIENTOS"
   v_cen(9) = "NOVECIENTOS"

   v_die(1) = "ONCE"
   v_die(2) = "DOCE"
   v_die(3) = "TRECE"
   v_die(4) = "CATORCE"
   v_die(5) = "QUINCE"

   Deci = Right(Format$(valor, "Standard"), 2) & IIf(UCase(xDivisa) <> "S", "/100 M.N.", "/100 USCY")
   If Left(Deci, 1) = "/" Then Deci = "00" & Deci

   valor2 = valor
   valor = Abs(Int(valor))
   Num = valor

   If Num >= 10 ^ 9 Then
      divisor = 10 ^ 9
   ElseIf Num >= 10 ^ 6 Then
      divisor = 10 ^ 6
   ElseIf Num >= 1000 Then
      divisor = 1000
   Else
      divisor = 1
   End If

   ValLetras = ""

   Do While True
      die = False
      Can = Int(Num / divisor)

      If Can > 99 Then
         If Can = 100 Then
            ValLetras = ValLetras & "CIEN "
            Can = 0
         Else
            Pos = Int(Can / 100)
            ValLetras = ValLetras & v_cen(Pos) & " "
            Can = Can - (Pos * 100)
         End If
      End If

      If Can > 9 Then
         Pos = 1
         If (Can = 10 Or Can > 15) And (Can < 21 Or Can > 29) Then
            Pos = Int(Can / 10)
            ValLetras = ValLetras & v_dec(Pos) & " "
            If (Pos > 2) And ((Can / 10) - Int(Can / 10) <> 0) Then
               ValLetras = ValLetras & "Y "
            End If
            If (Can > 15 And Can < 20) Then
               ValLetras = Mid(ValLetras, 1, Len(ValLetras) - 2) & "CI"
            End If
         Else
            If (Can > 20 And Can <= 29) Then
               ValLetras = ValLetras & "VEINTI"
               Pos = 2
            Else
               ValLetras = ValLetras & v_die(Can - 10) & " "
               die = True
            End If
         End If
         Can = Can - (Pos * 10)
      End If

      If (Not die) And (Can <> 0) Then
         ValLetras = ValLetras & v_uni(Can) & " "
      End If

      If (Num >= 10 ^ 6 And Num < 10 ^ 9) Then
         If (Int(Num / 10 ^ 6) = 1 And valor < 10 ^ 9) Then
            ValLetras = ValLetras & "MILLON "
         Else
            ValLetras = ValLetras & "MILLONES "
         End If
      Else
         If Num >= 1000 Then
            If (Int(valor / 1000) = 1 Or Int(Num / 10 ^ 9) = 1) Then
               ValLetras = ""
            End If
            ValLetras = ValLetras & "MIL "
         Else
            If valor = 1 Or (Mid(Str(valor), 6, 3) = "000" And Mid(Str(valor), 9, 2) = "00" And Can = 1) Then
               ValLetras = ValLetras & IIf(UCase(xDivisa) <> "S", "PESO", "DOLAR AMERICANO")
            Else
               ValLetras = IIf(Sgn(valor2) = -1, "MENOS ", "") & ValLetras & IIf(UCase(xDivisa) <> "S", "PESOS", "DOLARES AMERICANOS")
            End If
         End If
      End If

      If Num < 1000 Then
         Exit Do
      End If

      Num = Num - (Int(Num / divisor) * divisor)

      If (divisor = 10 ^ 9 And Num < 10 ^ 6) Then
         ValLetras = ValLetras & "MILLONES "
      End If

      If Num >= 10 ^ 6 Then
         divisor = 10 ^ 6
      ElseIf Num >= 1000 Then
         divisor = 1000
      Else
         divisor = 1
      End If
   Loop
   
   ValLetras = ValLetras & " " & Deci
   Leyenda = ValLetras
   valor = valor2

End Function
Public Function getNombreMes(ByVal nMes As Integer) As String
   getNombreMes = Choose(nMes, "Enero", "Febrero", "Marzo", _
                                       "Abril", "Mayo", "Junio", _
                                       "Julio", "Agosto", "Septiembre", _
                                       "Octubre", "Noviembre", "Diciembre")
End Function
Sub doWaitShell(sCmd As String)
Dim hShell As Long
Dim hProc As Long
Dim codExit As Long
'  Ejecutar comando
   hShell = Shell(Environ$("Comspec") & " /c " & sCmd, vbMinimizedFocus)
'  Esperar a que se complete el proceso
   hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)
   Do
      GetExitCodeProcess hProc, codExit
      DoEvents
   Loop While codExit = STILL_ACTIVE
End Sub
Function rellena(ByRef Cadena As String, ByRef nceros As Integer, ByRef caracter As String, Optional ByRef lado As String) As String
Dim longitud As Integer, X As Integer
Dim Texto As String, cadcero As String
Cadena = CStr(Cadena)
If caracter <> "" Then
   caracter = Mid(caracter, 1, 1)
   cadcero = ""
   For X = 1 To nceros
       cadcero = cadcero & caracter
   Next
   longitud = Len(CStr(Cadena))
   If lado = "D" Then
       Texto = Mid(cadcero, 1, nceros - longitud) & Cadena
   Else
       Texto = Cadena & Mid(cadcero, 1, nceros - longitud)
   End If
   rellena = Texto
Else
   cadcero = ""
   For X = 1 To nceros
       cadcero = cadcero & Cadena
   Next
   rellena = Mid(cadcero, 1, nceros)
End If
End Function
Public Sub entertab(Tecla As Integer)
    If Tecla = vbKeyReturn Then
        Tecla = 0
        SendKeys "{TAB}"
    End If
End Sub
Public Sub CargaConcep(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
       
    sqls = "select * from fm_tipos_mov_cartera" & _
           " where subtipo = 'O'"
           
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenStatic, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD!TIPO_MOV)
       strBodega = Trim("" & rsBD!descripcion)
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, -1)
End Sub
Function BuscaCliente(cliente As Long, NombreC As String)
Dim Nombre As String
Dim rsNombre As Recordset

  
        
   sqls = " select nombre from clientes" & _
          " where cliente = " & cliente
   Set rsNombre = New ADODB.Recordset
   rsNombre.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
   If Not rsNombre.EOF Then
      Nombre = rsNombre!Nombre
   Else
      Nombre = "No existe cliente"
   End If
   
   
   BuscaCliente = Nombre
   
rsNombre.Close
Set rsNombre = Nothing


End Function
Public Function Comprobar_Contraseña(Contraseña As String) As Boolean
On Error GoTo algomalopasa
Dim oReg As RegExp
' Crea un Nuevo objeto RegExp
    Set oReg = New RegExp
    ' Expresión regular
    oReg.Pattern = "^(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{8,10}$"
' Comprueba y Retorna TRue o false
    Comprobar_Contraseña = oReg.Test(Contraseña)
    Set oReg = Nothing
    If Comprobar_Contraseña Then
      If InStr(Contraseña, "@") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "!") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "#") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "$") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "%") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "&") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "/") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "(") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, ")") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "=") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "?") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "¿") > 0 Then
        Exit Function
      ElseIf InStr(Contraseña, "¡") > 0 Then
        Exit Function
      Else
        Comprobar_Contraseña = False
      End If
    End If
    Exit Function
algomalopasa:
MsgBox ("Error al crear la contraseña comprobando longitudes"), vbCritical, "Error"
End Function
Sub Switch(ByRef frmDestino As Form, ByVal strTag As String, ByVal blnValue As Boolean)

    Dim i As Integer
    
    On Error Resume Next
    
    strTag = UCase(strTag)
    
    For i = 0 To frmDestino.Count - 1
        If UCase(frmDestino(i).Tag) = strTag Then
            If blnValue = False And UCase(frmDestino(i).Tag) = "DET" Then
                If TypeName(frmDestino(i)) = "TextBox" Then
                       frmDestino(i).Text = ""
                End If
            End If
            frmDestino(i).Enabled = blnValue
        End If
    Next i
    
    On Error GoTo 0
    
End Sub

Sub CargaComboBE_All(cbo As ComboBox, sSql As String)
Dim oRS As ADODB.Recordset
'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
 
   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open sSql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   cbo.Clear
   intCount = -1
   intCount = intCount + 1
   cbo.AddItem "<< TODOS >>"
   cbo.ItemData(intCount) = 0
   Do While Not oRS.EOF
      cbo.AddItem UCase(Trim(oRS!Bon_Pro_Descripcion))
      oRS.MoveNext
   Loop
   oRS.Close
   Set oRS = Nothing
End Sub
Public Sub CargaClaves(cbo As Control, tabla As String, campo As String)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves"
    sqls = sqls & vbCr & " where tabla = '" & tabla & "'  and campo = '" & campo & "' and status = 1 Order By nocve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, -1)
End Sub
Function FechaFinMes(mes, Anio)
  'Determina el Ultimo Dia del Mes y lo devuelve con fecha dd/mm/yyyy
  If IsDate(Format(mes, "00") + "/" + "31/" + Trim(Str(Anio))) Then
    FechaFinMes = Format(mes, "00") + "/" + "31/" + Trim(Str(Anio))
  ElseIf IsDate(Format(mes, "00") + "/" + "30/" + Trim(Str(Anio))) Then
    FechaFinMes = Format(mes, "00") + "/" + "30/" + Trim(Str(Anio))
  ElseIf IsDate(Format(mes, "00") + "/" + "29/" + Trim(Str(Anio))) Then
    FechaFinMes = Format(mes, "00") + "/" + "29/" + Trim(Str(Anio))
  ElseIf IsDate(Format(mes, "00") + "/" + "28/" + Trim(Str(Anio))) Then
    FechaFinMes = Format(mes, "00") + "/" + "28/" + Trim(Str(Anio))
  End If
End Function

Public Function doGenArchCanc(ByVal nBodega As Integer, _
                              ByVal sserie As String, _
                              ByVal sFolio As Long, _
                              ByVal nTipoArch As Long)

Dim nfile As Long, SFILE As String, blnFileOpen As Boolean, sExt As String
Dim blnOk As Boolean
Dim strsql As String, rstTmp As ADODB.Recordset
   blnOk = False
   strsql = "sp_GenArchCancFE " & nBodega & ", " & _
                           "'" & sserie & "', " & _
                           sFolio & ", " & _
                           nTipoArch

   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   If Not rstTmp.EOF Then
      nfile = FreeFile()
      sExt = IIf(nTipoArch = 1, "CFA", "CFA")
      SFILE = gsPathFE & "\Paso\03" & sserie & sFolio & "." & sExt
      Open SFILE For Output As #nfile
      Print #nfile, rstTmp!CveDoc & "|" & rstTmp!RFCEmisor & "|" & DeleteChar("-", DeleteChar(" ", rstTmp!RFCReceptor)) & "|" & _
                     rstTmp!Serie & "|" & _
                     rstTmp!Folio & "|" & _
                     Format$(rstTmp!FechaFact, "DD/MM/YY") & "|" & _
                     Format$(rstTmp!FechaCanc, "DD/MM/YY") & "|" & _
                     rstTmp!TotalIva & "|" & _
                     rstTmp!TotalGeneral & "|" & _
                     rstTmp!TipoCambio

      Close #nfile
      Call doWaitShell("MOVE " & gsPathFE & "\Paso\03" & sserie & sFolio & "." & sExt & " C:\GoDir\WsHome\datos\salida")
   End If
   blnOk = True
End Function


Function QuitaCeros(Cadena As String)
    For j = 1 To Len(Cadena)
        If Mid(Cadena, j, 1) <> "0" Then
            QuitaCeros = Mid(Cadena, j, Len(Cadena) + 1 - j)
            Exit For
        End If
    Next j
End Function
Public Sub CargaMensajerias(cbo As Control)
    Dim intID As Integer
    Dim strDescripcion As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT CveMensajeria IdMens, Descripcion"
    sqls = sqls & vbCr & " FROM Mensajerias "
    sqls = sqls & vbCr & " Order By IDMens"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intID = Val("" & rsBD![idMens])
       strDescripcion = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strDescripcion)
       cbo.ItemData(intCount) = intID
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, 3)
End Sub

Function ValidaNumericos(Tecla As Integer, Texto As String, tipo As Integer) As Integer

' Funcion que permite solo la entada de los valores correctos para un Numerico
' tecla ... Contiene el valor ascii de la tecla presionada
' Texto ... Contiene el valor de la propiedad text del cuadro de texto que se este validando
' Tipo  ... Contiene el valor de:
'           0 ... Para valores enteros  +
'           1 ...                       + y -
'           2 ...             decimales +
'           3 ...                       + y -

    Dim Switch As Integer

    If Tecla >= 48 And Tecla <= 57 Then
        ValidaNumericos = Tecla
        Exit Function
    End If
    
    If Tecla = 8 Or Tecla = 9 Or Tecla = 10 Or Tecla = 13 Or Tecla = 27 Then
        ValidaNumericos = Tecla
        Exit Function
    End If
    
    
    If (tipo = 1 Or tipo = 3) And Tecla = 45 Then
        If InStr(Texto, "-") = 0 Then
            ValidaNumericos = Tecla
            Switch = 1
        End If
    End If
    
    If (tipo = 2 Or tipo = 3) And Tecla = 46 Then
        If InStr(Texto, ".") = 0 Then
            ValidaNumericos = Tecla
            Switch = 1
        End If
    End If

    
    If Switch <> 1 Then ValidaNumericos = 0 Else ValidaNumericos = Tecla

End Function

Sub CargaColonias(cbo As ComboBox, sSql As String)
Dim oRS As ADODB.Recordset
'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=OpVt2016;database=" & gstrDataBase
 
   Set oRS = New ADODB.Recordset
   oRS.CursorLocation = adUseClient
   oRS.Open sSql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   cbo.Clear
   Do While Not oRS.EOF
      cbo.AddItem UCase(Trim(oRS!Colonia))
      oRS.MoveNext
   Loop
   oRS.Close
   Set oRS = Nothing
End Sub

Public Sub CargaUsuarios(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
'    Set cnxBD = New ADODB.Connection
'    cnxBD.CommandTimeout = 2000
'    cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=OpVt2016;database=" & gstrDataBase
    
    sqls = " SELECT Usuario, nombre"
    sqls = sqls & vbCr & " FROM usuarios "
    sqls = sqls & vbCr & " where puesto like '%Operaciones%' and Status=1 Order By Usuario"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       'intBodega = Val("" & rsBD![Bodega])
       'strBodega = Trim("" & rsBD![descripcion])
       
       cbo.AddItem Trim(rsBD!Usuario) & " - " & UCase(Trim(rsBD!Nombre))
       cbo.ItemData(intCount) = intCount
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, -1)
End Sub

Public Sub CargaProblemas(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves"
    sqls = sqls & vbCr & " where tabla = 'Aclaraciones'  and campo = 'Problema' Order By nocve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, -1)
End Sub

Public Sub CargaStatus(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves"
    sqls = sqls & vbCr & " where tabla = 'Aclaraciones'  and campo = 'Status' Order By nocve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, 0)
End Sub

Public Sub CargaStatusCliente(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves"
    sqls = sqls & vbCr & " where tabla = 'Clientes'  and campo = 'Status' Order By nocve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    intCount = -1
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, 0)
End Sub
Public Sub CargaBodegas2(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT Bodega, Descripcion"
    sqls = sqls & vbCr & " FROM Bodegas "
    sqls = sqls & vbCr & " Order By Bodega"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    
    intCount = -1
    intCount = intCount + 1
    cbo.AddItem "<< TODAS >>"
    cbo.ItemData(intCount) = 0
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![Bodega])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, gnBodega)
End Sub
Public Sub Carga_Status_Solicitud(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves where Tabla='Solicitudesbe' "
    sqls = sqls & vbCr & " Order By NoCve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    
    intCount = -1
    intCount = intCount + 1
    cbo.AddItem "<< TODAS >>"
    cbo.ItemData(intCount) = 0
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, gnBodega)
End Sub


Public Sub Carga_Status_Solicitud2(cbo As Control)
    Dim intBodega As Integer
    Dim strBodega As String
    Dim intCount As Integer
    cbo.Clear
    sqls = " SELECT NoCve, Descripcion"
    sqls = sqls & vbCr & " FROM Claves where Tabla='Solicitudesbe' And Status=1"
    sqls = sqls & vbCr & " Order By NoCve"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxbdMty, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbDefault
    
    intCount = -1
'    IntCount = IntCount + 1
'    cbo.AddItem "<< TODAS >>"
'    cbo.ItemData(IntCount) = 0
    
    Do While Not rsBD.EOF
       intCount = intCount + 1
       intBodega = Val("" & rsBD![nocve])
       strBodega = Trim("" & rsBD![descripcion])
       cbo.AddItem Trim(strBodega)
       cbo.ItemData(intCount) = intBodega
       rsBD.MoveNext
    Loop
    rsBD.Close
    Set rsBD = Nothing
    Call CboPosiciona(cbo, gnBodega)
End Sub

Function Valida_RFC(ByVal Rfc As String, Optional ByVal blnShowMsg As Boolean) As Boolean
   Dim Pos As Integer, i As Integer
   Dim Pg1 As Integer, Pg2 As Integer, Pgn As Integer
   Dim Ano_pas As String, Mes_pas As String, Dia_pas As String    ', Fecha
   
   Pos = 1: Pg1 = 0: Pg2 = 0: Pgn = 0
   
   blnShowMsg = IIf(IsMissing(blnShowMsg), False, blnShowMsg)
   
   ' Validación de Caracteres "-"
   For Pos = 1 To Len(Rfc)
      If Mid(Rfc, Pos, 1) = "-" Then
         If Pos = 1 Then
            ' En caso de que el Primer caracter sea " - "
            If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
            Valida_RFC = False
            Exit Function
         End If
         Pgn = Pgn + 1
         If Pgn > 2 Then
            ' En caso de que Existan mas de dos "-"
            If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
            Valida_RFC = False
            Exit Function
         End If

         ' Asigna Posici¢n de guion
         If Pg1 = 0 Then
            Pg1 = Pos
         Else
            Pg2 = Pos
         End If
      End If
   Next

   ' Valida que el Primer Guión este en las posiciones 4 o 5
   If Pg1 <> 4 And Pg1 <> 5 Then
      If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
      Valida_RFC = False
      Exit Function
   End If
   If Pg1 = 4 Then
      If Pg2 <> 11 And Len(Rfc) > 10 Then
         If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
         Valida_RFC = False
         Exit Function
      End If
   End If
   If Pg1 = 5 Then
      If Pg2 <> 12 And Len(Rfc) > 11 Then
         If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
         Valida_RFC = False
         Exit Function
      End If
   End If
      
   ' Valida que los primeros caracteres antes del primer guión sean alfabético
   For i = 1 To Pg1 - 1
      If IsNumeric(Mid(Rfc, i, 1)) Then
         If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
         Valida_RFC = False
         Exit Function
      End If
   Next
   
   If Pg1 = 4 Then
      Ano_pas = Mid(Rfc, 5, 2)
      Mes_pas = Mid(Rfc, 7, 2)
      Dia_pas = Mid(Rfc, 9, 2)
   Else
      Ano_pas = Mid(Rfc, 6, 2)
      Mes_pas = Mid(Rfc, 8, 2)
      Dia_pas = Mid(Rfc, 10, 2)
   End If
   
   If Not IsDate(FormatoFecha(Dia_pas + Mes_pas + Ano_pas)) Then
      ' En caso de que sea Fecha Inv lida (dia o mes no existentes)
      If blnShowMsg Then MsgBox "Formato de RFC Inválido.. ,Utilizar : XXXX-AAMMDD-AAA", vbExclamation, "Error de RFC"
      Valida_RFC = False
      Exit Function
   End If
   
   Valida_RFC = True

End Function

Public Function FormatoFecha(ByVal dText As String) As String
   Dim i, a, FormF As String, Car As String, CantCar As Long
   a = 1
   CantCar = Len(dText)
   dText = Pad(dText, IIf(CantCar = 7 Or CantCar = 8, 8, _
                            IIf(CantCar = 5 Or CantCar = 6, 6, _
                            IIf(CantCar = 9 Or CantCar = 10, 10, 10))), "0", "R")
   If IsNumeric(dText) Then
      For i = 1 To Len(dText) + 2
         If i = 3 Or i = 6 Then
            FormF = FormF & "/"
         Else
            FormF = FormF & Mid(dText, a, 1)
            a = a + 1
         End If
      Next
   Else
      FormF = dText
   End If
   FormatoFecha = FormF
End Function


Public Function doGenArchFE_NCA(ByVal nBodega As Integer, _
                         ByVal sserie As String, _
                         ByVal nFactIni As Long, _
                         ByVal nFactFin As Long, _
                         ByVal nTipoArch As Integer) As Boolean
Dim nfile As Long, SFILE As String, blnFileOpen As Boolean
Dim blnOk As Boolean
Dim strsql As String, rstTmp As ADODB.Recordset
Dim sRutina As String
'Variables para el Do While
Dim nFactAnt As Long, nNR As Long, nCantFact As Long
Dim nLineas As Long, nUltReg As Long
Dim blnAgrupa As Boolean, blnAdd As Boolean
'nRegAct As Long,
Dim nTipoPed As Long
Dim lngI As Long
'Variables para denominaciones raras
Dim SCodigo As String, sDescripcion As String, _
   nUnidades As Long, nIvaProd As Double, _
   nImporte As Double, nIva As Double, nTotal As Double
Dim sFolioIni As String, sFolioFin As String
Dim nPrecio As Double, nPrecioSug As Double
Dim sDesglose As String
'Variable para Agrupar Prefacturas
Dim sBodega As String, nPrefSig As Double, nFolSig As Double
'Documento
Dim sSerieDocto As String, sSerieFac As String, sFactura As String, sFacturaRef As String, sFecha As String, sFechaFac As String, _
      sConcepto As String, sNoAprob As String, sAnoAprob As String, sNoCert As String, _
      sCveBod As String, sDescBodega As String
'Empresa
Dim sEmpRazonSocial As String, sEmpRFC As String, sEmpCalle As String, _
   sEmpNoExt As String, sEmpNoInt As String, sEmpCol As String, _
   sEmpCP As String, sEmpLocalidad As String, sEmpMun As String, _
   sEmpEdo As String, sEmpPais As String, sEmpTel As String, _
   sEmpPagWeb As String
'Expide
Dim sExpCalle As String, sExpNoExt As String, sExpNoInt As String, _
   sExpCol As String, sExpCP As String, sExpLocalidad As String, _
   sExpMun As String, sExpEdo As String, sExpPais As String, _
   sExpTel As String
'Cliente a Facturar
Dim sRCliente As String, sRRFC As String, _
   sRNombre As String, _
   sRCalle As String, sRNoExt As String, sRNoInt As String, _
   sREntreCalles As String, sRLocalidad As String, sRCol As String, _
   sRMun As String, sREdo As String, sRPais As String, _
   sRCP As String, Splaza As String
', sRApPat As String, sRAPMat As String
'Cliente a Enviar
Dim sECliente As String, sENombre As String, _
   sECalle As String, sENoExt As String, sENoInt As String, _
   sEEntreCalles As String, sELocalidad As String, sECol As String, _
   sEMun As String, sEEdo As String, sEPais As String, _
   sECP As String, sEGuia As String, SETel As String
'Encabezado
Dim sPedido As String, sRuta As String, sCondiciones As String, sFechaVence As String
'Variables de la Factura
Dim blnGravIVA As Boolean
Dim sRef As String, sRefBonos As String, sPedRef
Dim sComentario As String, sImporteLetras As String
Dim sLeyendaPagare As String, sPagareDatosCliente As String, sPagareCiudadContrato As String
'Detalle de la Factura
Dim aCodigo() As String, aDescripcion() As String, _
   aUnidades() As Long, aPrecio() As Double, aPrecioSug() As Double, _
   aFolIni() As String, aFolFin() As String, _
   aImporte() As Double, aIVA() As Double, aTotal() As Double
'Detalle de la Factura (Prefacturas)
Dim aSerie() As String, aDescBod() As String, _
      aPrefIni() As Double, aPrefFin() As Double, aAnoMes() As String
'Datos de la Comision
Dim nImpComision As Double, nImpIVAComision As Double
Dim nPrcCom As Double, nPrcComCalc As Double, nPrcIVA As Double
'Totales
Dim nTotalUnidades As Long, _
   ntotalPedido, nTotalImporte As Double, nTotalIVA As Double, nTotalGen As Double
Dim sImpresora As String
Dim minutos As String, minutos2 As String
'  1.- Factura Normal o de Stock
'  2.- Factura de Prefacturas

   On Error GoTo hdErr
   blnImprime = True
   
   
    blnImprime = True
   
   
    If blnImprime Then
       sImpresora = doFindPrinter(gstrPC, "NOTAS")
    Else
       sImpresora = ""
    End If
     
   
   blnOk = False
   
   strsql = "sp_GenArchFE " & nBodega & ", " & _
                           "'" & sserie & "', " & _
                           nFactIni & ", " & _
                           nFactFin & ", " & _
                           nTipoArch
   Set rstTmp = New ADODB.Recordset
   rstTmp.Open strsql, cnxBD, adOpenForwardOnly, adLockReadOnly, adCmdText
   
   Do While Not rstTmp.EOF
      blnOk = True
      If nFactAnt <> rstTmp!Folio Then
         If nFactAnt <> 0 Then
            If nNR > 0 Or sRef <> "" Then
               nCantFact = nCantFact + 1
               If nTipoArch = 1 Then
'                 Factura Normal
'                 Si tiene fracciones agregar registro
                  If nUnidades > 0 Then
                     ReDim Preserve aCodigo(nLineas)
                     ReDim Preserve aDescripcion(nLineas)
                     ReDim Preserve aUnidades(nLineas)
                     ReDim Preserve aPrecio(nLineas)
                     ReDim Preserve aPrecioSug(nLineas)
                     ReDim Preserve aFolIni(nLineas)
                     ReDim Preserve aFolFin(nLineas)
                     ReDim Preserve aImporte(nLineas)
                     ReDim Preserve aIVA(nLineas)
                     ReDim Preserve aTotal(nLineas)
                  
                     aCodigo(nLineas) = SCodigo
                     aDescripcion(nLineas) = sDescripcion
                     aUnidades(nLineas) = nUnidades
                     aPrecio(nLineas) = 0
                     aPrecioSug(nLineas) = 0
                     aFolIni(nLineas) = ""
                     aFolFin(nLineas) = ""
                     aImporte(nLineas) = nImporte
                     aIVA(nLineas) = nIva
                     aTotal(nLineas) = nTotal
                     
                     nUltReg = nLineas - 1
                  Else
                     nLineas = nLineas - 1
                     nUltReg = nLineas
                  End If
               ElseIf nTipoArch = 2 Then
                  nLineas = nLineas - 1
               End If
               
               GoSub doGenFile
            End If
         End If
         nFactAnt = rstTmp!Folio
'         nRegAct = 1
         nLineas = 1
         
         If nTipoArch = 3 Then
'           Reinicia valores para agrupar denominaciones raras
            SCodigo = ""
            sDescripcion = ""
            nUnidades = 0
            nImporte = 0
            nIva = 0
            nTotal = 0
         Else
'           Reinicia valores para agrupar prefacturas
            sBodega = ""
            nPrefSig = 0
            nFolSig = 0
            
'           Reinicia Totales
            nTotalUnidades = 0
         End If
         
         GoSub doFillEnc
      End If
      
      GoSub doFillDet
      
'      nRegAct = nRegAct + 1
      rstTmp.MoveNext
   Loop
   
   
   nNR = 1
   If nNR > 0 Or sRef <> "" Then
      nCantFact = nCantFact + 1
      If nTipoArch = 4 Or nTipoArch = 3 Then
'        Si tiene fracciones agregar registro
         If nUnidades > 0 Then
            ReDim Preserve aCodigo(nLineas)
            ReDim Preserve aDescripcion(nLineas)
            ReDim Preserve aUnidades(nLineas)
            ReDim Preserve aPrecio(nLineas)
            ReDim Preserve aPrecioSug(nLineas)
            ReDim Preserve aFolIni(nLineas)
            ReDim Preserve aFolFin(nLineas)
            ReDim Preserve aImporte(nLineas)
            ReDim Preserve aIVA(nLineas)
            ReDim Preserve aTotal(nLineas)
         
            aCodigo(nLineas) = SCodigo
            aDescripcion(nLineas) = sDescripcion
            aUnidades(nLineas) = nUnidades
            aPrecio(nLineas) = 0
            aPrecioSug(nLineas) = 0
            aFolIni(nLineas) = ""
            aFolFin(nLineas) = ""
            aImporte(nLineas) = nImporte
            aIVA(nLineas) = nIva
            aTotal(nLineas) = nTotal
            
            nUltReg = nLineas - 1
         Else
            nLineas = nLineas - 1
            nUltReg = nLineas
         End If
      'ElseIf nTipoArch = 2 Then
      '   nLineas = nLineas - 1
      End If
      
      'GeneraArchivo
      GoSub doGenFile
   End If
   
   On Error GoTo 0
   
   doGenArchFE_NCA = blnOk
   Exit Function

doFillEnc:
  
   blnAgrupa = False
         
   sRutina = "doFillEnc"
   sSerieDocto = Trim(rstTmp!serieNota)
   sFactura = rstTmp!Folio
   minutos = Format(rstTmp!Fecha_Nota, "hh:mm:ss")
   sFecha = Format(rstTmp!Fecha_Nota, "YYYY-MM-DD") 'THH:MM:SS")
   sFecha = sFecha & "T" & minutos
    
   sSerieFac = Trim(rstTmp!serieFac)
   sFacturaRef = rstTmp!Factura
   minutos2 = Format(rstTmp!FechaFac, "hh:mm:ss")
   sFechaFac = Format(rstTmp!FechaFac, "YYYY-MM-DD") ' & "T00:00:00"
   sConcepto = rstTmp!Concepto
 
   
   sNoAprob = rstTmp!NoAprob
   sAnoAprob = rstTmp!AnoAprob
   sNoCert = rstTmp!NoCert
   sCveBod = rstTmp!Bodega
   sDescBodega = Trim(rstTmp!BodegaDescripcion)
      metodop = Format(rstTmp!MetodoPago, "00")
   cuentap = Trim(rstTmp!CuentaPago)
   versioncfd = rstTmp!versioncfd

   sEmpRazonSocial = rstTmp!EmpRazonSocial
   sEmpRFC = rstTmp!EmpRFC
   sEmpCalle = rstTmp!EmpCalle
   sEmpNoExt = rstTmp!EmpNoExt
   sEmpNoInt = rstTmp!EmpNoInt
   sEmpCol = rstTmp!EmpColonia
   sEmpCP = rstTmp!EmpCP
   sEmpLocalidad = rstTmp!EmpLocalidad
   sEmpMun = rstTmp!EmpMunicipio
   sEmpEdo = rstTmp!EmpEstado
   sEmpPais = rstTmp!EmpPais
   sEmpTel = rstTmp!EmpTel
   sEmpPagWeb = rstTmp!EmpPagWeb
   
   sExpCalle = rstTmp!ExpCalle
   sExpNoExt = rstTmp!ExpNoExt
   sExpNoInt = rstTmp!ExpNoInt
   sExpCol = rstTmp!ExpColonia
   sExpCP = rstTmp!ExpCP
   sExpLocalidad = rstTmp!EmpLocalidad
   sExpMun = rstTmp!ExpMunicipio
   sExpEdo = rstTmp!ExpEstado
   sExpPais = rstTmp!ExpPais
   sExpTel = rstTmp!ExpTel
   
   sRCliente = rstTmp!RCliente
   sRRFC = DeleteChar("-", DeleteChar(" ", rstTmp!RRFC))
   sRNombre = rstTmp!RNombre
'   sRApPat = rstTmp!RApPat
'   sRAPMat = rstTmp!RAPMat
   sRCalle = rstTmp!RCalle
   sRNoExt = rstTmp!RNoExt
   sRNoInt = rstTmp!RNoInt
   sREntreCalles = rstTmp!REntreCalles
   sRLocalidad = rstTmp!RLocalidad
   sRCol = rstTmp!Rcolonia
   sRMun = rstTmp!RMunicipio
   sREdo = rstTmp!REstado
   sRPais = rstTmp!RPais
   sRCP = rstTmp!RCP
   Splaza = rstTmp!Plaza
   
'   sECliente = rstTmp!ECliente
'   sENombre = rstTmp!ENombre
'   sECalle = rstTmp!ECalle
'   sENoExt = rstTmp!ENoExt
'   sENoInt = rstTmp!ENoInt
'   sEEntreCalles = rstTmp!EEntreCalles
'   sELocalidad = rstTmp!ELocalidad
'   sECol = rstTmp!EColonia
'   sEMun = rstTmp!EMunicipio
'   sEEdo = rstTmp!EEstado
'   sEPais = rstTmp!EPais
'   sECP = rstTmp!ECP
 '  If nTipoArch = 1 Then
 '     sEGuia = rstTmp!EGuiaRoji
 '     SETel = rstTmp!ETel
 '
 '     sPedido = rstTmp!pedido
 '  End If
  ' sRuta = rstTmp!Ruta
   'sCondiciones = IIf(rstTmp!Condiciones = 1, "CONTADO", Trim(Str(rstTmp!Condiciones)) & " DIAS")
   'sFechaVence = doFormatDate(Format(rstTmp!FechaVence, "YYYY-MM-DD"), 2)
   
   blnGravIVA = IIf(rstTmp!GravIVA = "S", True, False)
'   sRef = rstTmp!Referencia
'   sRefBonos = rstTmp!RefBonos
'   sPedRef = rstTmp!PedidoRef
   
   nImpComision = rstTmp!ComisionDet
   nImpIVAComision = Round(rstTmp!IvaComisionDet, 2)
 '  nPrcCom = rstTmp!PorcComision
 '  If rstTmp!ValorPedido = 0 Then
 '     nPrcComCalc = 0
 '  Else
 '     nPrcComCalc = Round(rstTmp!comision / rstTmp!ValorPedido * 100, 2)
 '  End If
  ' nPrcCom = IIf(nPrcCom <> nPrcComCalc, nPrcComCalc, nPrcCom)
   
   nPrcIVA = rstTmp!PrcIVA
   
   nIvaProd = IIf(blnGravIVA, (rstTmp!PrcIVA / 100), 0)
   
   ntotalPedido = rstTmp!TotalDet
   nImpComision = rstTmp!ComisionDet
   nImpIVAComision = Round(rstTmp!IvaComisionDet, 2)
   
   nTotalImporte = rstTmp!TotalImporte
   nTotalIVA = rstTmp!TotalIva
   nTotalGen = rstTmp!TotalGeneral
   
   If nTipoArch = 4 Then
      nTotalUnidades = 1
      sComentario = rstTmp!Comentario
   End If
   sImporteLetras = Leyenda(rstTmp!TotalGeneral)
   sMailFE = IIf(rstTmp!MailFe = 1, "MAIL", "")
   sMailFETo = rstTmp!MailFETo
'   sLeyendaPagare = "Por el presente PAGARE me(nos) obligo(amos) incondicionalmente a pagar " & _
                    "a la orden de " & sEmpRazonSocial & " en esta plaza en moneda nacional el dia " & _
                    Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4) & _
                    " la cantidad de " & Format(nTotalGen, "###,###,##0.00") & " (" & sImporteLetras & ") " & _
                    "valor en efectivo. Si no fuere pagado satisfactoriamente este pagaré me(nos) obligo(amos) " & _
                    "a pagar durante todo el tiempo que permaneciera total o parcialmente insoluto, " & _
                    "intereses moratorios a razón del 2% mensual, sin que por esto se considere prorrogado " & _
                    "el plazo fijado para el cumplimiento de esta obligación."
 '  sPagareDatosCliente = sRNombre & " con Domicilio en: " & sRCalle & " " & sRNoExt & _
                         " Colonia: " & sRCol & " en " & sRMun & ", " & sREdo
  ' sPagareCiudadContrato = sExpMun & ", " & sExpEdo & ", a " & _
                           Mid(sFecha, 9, 2) & " de " & UCase(getNombreMes(Mid(sFecha, 6, 2))) & " del " & Mid(sFecha, 1, 4)
   Return

doFillDet:
   sRutina = "doFillDet"
   If nTipoArch = 3 Or nTipoArch = 4 Then
      If blnAgrupa Then
   '     Agregarlo al final a un arreglo para al final imprimir
         SCodigo = rstTmp!codigo
         sDescripcion = rstTmp!descripcion
         nUnidades = nUnidades + rstTmp!Unidades
         nImporte = nImporte + Round((rstTmp!total / (1 + nIvaProd)), 2)
         nIva = nIva + (rstTmp!total - Round((rstTmp!total / (1 + nIvaProd)), 2))
         nTotal = nTotal + rstTmp!total
      Else
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aPrecio(nLineas)
         ReDim Preserve aPrecioSug(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = rstTmp!codigo
         aDescripcion(nLineas) = rstTmp!descripcion
         aUnidades(nLineas) = 0
         aPrecio(nLineas) = 0
         aPrecioSug(nLineas) = 0
         aFolIni(nLineas) = ""
         aFolFin(nLineas) = ""
         aImporte(nLineas) = Round((rstTmp!importedet), 2)
         aIVA(nLineas) = Round(rstTmp!Ivagradet, 2)
         aTotal(nLineas) = Round(rstTmp!TotalDet, 2)
         nImpComision = Round(rstTmp!ComisionDet, 2)
         nImpIVAComision = Round(rstTmp!IvaComisionDet, 2)
   
         nTotalIVA = rstTmp!TotalIva
         nTotalGen = rstTmp!TotalGeneral
         
         nLineas = nLineas + 1
      End If
   ElseIf nTipoArch = 2 Then
      If blnAgrupa Then
         If sBodega = rstTmp!descBodega & "" And _
               nPrefSig = rstTmp!NoPreFactura And _
               nFolSig = Val(Right(Trim(rstTmp!FolioIni & ""), 7)) Then
            aPrefFin(nLineas - 1) = rstTmp!NoPreFactura
            aFolFin(nLineas - 1) = Right(Trim(rstTmp!FolioFin & ""), 7)
            aUnidades(nLineas - 1) = aUnidades(nLineas - 1) + rstTmp!Unidades
            aImporte(nLineas - 1) = aImporte(nLineas - 1) + rstTmp!importe
            aIVA(nLineas - 1) = aIVA(nLineas - 1) + rstTmp!iva
            aTotal(nLineas - 1) = aTotal(nLineas - 1) + rstTmp!total
            blnAdd = False
         Else
            blnAdd = True
'            nRegAct = nRegAct + 1
         End If
         
         sBodega = rstTmp!descBodega & ""
         nPrefSig = rstTmp!NoPreFactura + 1
         nFolSig = Val(Right(Trim(rstTmp!FolioFin & ""), 7)) + 1
      Else
         blnAdd = True
      End If
      
      If blnAdd Then
         ReDim Preserve aCodigo(nLineas)
         ReDim Preserve aDescripcion(nLineas)
         
         ReDim Preserve aSerie(nLineas)
         ReDim Preserve aPrefIni(nLineas)
         ReDim Preserve aPrefFin(nLineas)
         ReDim Preserve aDescBod(nLineas)
         ReDim Preserve aAnoMes(nLineas)
         ReDim Preserve aFolIni(nLineas)
         ReDim Preserve aFolFin(nLineas)
         
         ReDim Preserve aUnidades(nLineas)
         ReDim Preserve aImporte(nLineas)
         ReDim Preserve aIVA(nLineas)
         ReDim Preserve aTotal(nLineas)
      
         aCodigo(nLineas) = rstTmp!codigo
         aDescripcion(nLineas) = rstTmp!descripcion
         
         aSerie(nLineas) = Trim(rstTmp!SeriePreFactura & "")
         aPrefIni(nLineas) = rstTmp!NoPreFactura
         aPrefFin(nLineas) = rstTmp!NoPreFactura
         aDescBod(nLineas) = rstTmp!descBodega & ""
         aAnoMes(nLineas) = Left(Trim(rstTmp!FolioIni & ""), 6)
         aFolIni(nLineas) = Right(Trim(rstTmp!FolioIni & ""), 7)
         aFolFin(nLineas) = Right(Trim(rstTmp!FolioFin & ""), 7)
         
         aUnidades(nLineas) = rstTmp!Unidades
         aImporte(nLineas) = rstTmp!importe
         aIVA(nLineas) = rstTmp!iva
         aTotal(nLineas) = rstTmp!total
      
         nLineas = nLineas + 1
      End If
      
      nTotalUnidades = nTotalUnidades + rstTmp!Unidades
   End If
   Return
   
doGenFile:
   sRutina = "doGenFile"
   GoSub doFileOpen
   GoSub doGenEnc
   GoSub doGenDet
   GoSub doGenPie
   GoSub doFileClose
    If Trim(sMailFETo) <> "" Then MsgBox "La nota fue enviada por correo electronico", vbInformation
   Return
   
doFileOpen:

  sRutina = "doFileOpen"
   nfile = FreeFile()
   SFILE = gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT"
   Open SFILE For Output As #nfile
   blnFileOpen = True
   Return

doFileClose:
   sRutina = "doFileClose"
   Close #nfile
   blnFileOpen = False
   GoSub doFileMove
   Return
 
doFileMove:
   sRutina = "doFileMove"
   Call doWaitShell("MOVE " & gsPathFE & "\Paso\03" & sSerieDocto & sFactura & ".TXT C:\GoDir\WsHome\Datos\Salida\")
   Return

doGenEnc:
   sRutina = "doGenEnc"
   
    If Trim(sMailFETo) <> "" Then
        blnImprime = False
   Else
        blnImprime = True
   End If
   
   If nTipoArch = 3 Then
    Print #nfile, "TFormato;" & "VTONCRE"
    Print #nfile, "TIPO_DOCTO;" & "NOTA DE CREDITO"
   Else
    Print #nfile, "TFormato;" & "VTONCAR"
    Print #nfile, "TIPO_DOCTO;" & "NOTA DE CARGO"
   End If
   Print #nfile, "Ndistrib;" & "FE,MAIL,,NO," & IIf(blnImprime, "SI", "NO")
   '-------------------
   If sMailFE = "MAIL" Then
      Print #nfile, "COPIAS;1,,"
   Else
      Print #nfile, "COPIAS;2,,"
   End If
   '-------------------
   Print #nfile, "Impresora;" & sImpresora
   
   
   '--------------
   
   If sMailFE = "MAIL" Then
      Print #nfile, "CorreoTo;" & sMailFETo
      If nTipoArch = 3 Then
        Print #nfile, "CorreoSub;" & "Envio de Nota de Credito " & sSerieDocto & "-" & sFactura
      Else
        Print #nfile, "CorreoSub;" & "Envio de Nota de Cargo " & sSerieDocto & "-" & sFactura
      End If
      Print #nfile, "CorreoAtt;PDF,XML"
   End If
   
   
   '--------------
   
 
   
   Print #nfile, "Serie;" & sSerieDocto
   Print #nfile, "Folio;" & sFactura
   Print #nfile, "Folio1;" & Trim(sSerieDocto) & "-" & sFactura
   Print #nfile, "Folio2;" & Trim(sSerieFac) & "-" & sFacturaRef
   
   Print #nfile, "Fecha_FACTURA;" & sFecha '& Format(Now, "YYYY-MM-DD") & "T" & minutos '
   Print #nfile, "Fecha_FACTURA1;" & doFormatDate(sFecha, 0) & " " & minutos
   Print #nfile, "Fecha_FACTURA_NCR;" & doFormatDate(sFechaFac, 0) & " " & minutos2
   Print #nfile, "Concepto_Emp_NCR;" & sConcepto


   Print #nfile, "NO. DE APROBACION;" & sNoAprob
   Print #nfile, "ANNO;" & sAnoAprob
   Print #nfile, "Certificado;" & sNoCert
   Print #nfile, "BodegaNumero;" & sCveBod
   Print #nfile, "BodegaNombre;" & sDescBodega
   Print #nfile, "CodBar_Encabezado;" & Trim(sSerieDocto) & sFactura
   Print #nfile, "VersionCFD;" & versioncfd '1
   Print #nfile, "TipoCambio;1" '12
   Print #nfile, "moneda;MXN" '11
   Print #nfile, "FormaPago;" & metodop '5
   If cuentap = "" Or cuentap = "0" Or cuentap = "0000" Then
      Print #nfile, "NumeroCuentaCMPV;    " '17
   Else
      Print #nfile, "NumeroCuentaCMPV;" & cuentap '17
   End If

   Print #nfile, "EmpresaRazonSocial;" & sEmpRazonSocial
   Print #nfile, "EmpresaRFC;" & sEmpRFC
   Print #nfile, "EmpresaCalle;" & sEmpCalle
   Print #nfile, "EmpresaNumExterior;" & sEmpNoExt
   Print #nfile, "EmpresaNumInterior;" & sEmpNoInt
   Print #nfile, "EmpresaColonia;" & sEmpCol
   Print #nfile, "EmpresaCodigoPostal;" & sEmpCP
   Print #nfile, "EmpresaLocalidad;" & sEmpLocalidad
   Print #nfile, "EmpresaMunicipio;" & sEmpMun
   Print #nfile, "EmpresaEstado;" & sEmpEdo
   Print #nfile, "EmpresaPais;" & sEmpPais
   Print #nfile, "EmpresaTel;" & sEmpTel
   Print #nfile, "EmpresaPaginaWeb;" & sEmpPagWeb
   Print #nfile, "EmpresaDomicilio1;" & sEmpCalle & " " & sEmpNoExt & IIf(sEmpNoInt <> "", " INT ", "") & sEmpNoInt
   Print #nfile, "EmpresaColonia1;" & sEmpCol & " CP." & sEmpCP
   Print #nfile, "EmpresaCiudad1;" & sEmpMun & ", " & sEmpEdo
   Print #nfile, "RegimenFiscal;601" '27
   
   Print #nfile, "EmpresaExpCalle;" & sExpCalle
   Print #nfile, "EmpresaExpNumExterior;" & sExpNoExt
   Print #nfile, "EmpresaExpNumInterior;" & sExpNoInt
   Print #nfile, "EmpresaExpColonia;" & sExpCol
   Print #nfile, "EmpresaExpCodigoPostal;" & sExpCP
   Print #nfile, "EmpresaExpLocalidad;" & sExpLocalidad
   Print #nfile, "EmpresaExpMunicipio;" & sExpMun
   Print #nfile, "EmpresaExpEstado;" & sExpEdo
   Print #nfile, "EmpresaExpPais;" & sExpPais
   Print #nfile, "EmpresaExpTel;" & sExpTel
   Print #nfile, "EmpresaExpDomicilio1;" & sExpCalle & " " & sExpNoExt
   Print #nfile, "EmpresaExpColonia1;" & sExpCol & " CP." & sExpCP
   Print #nfile, "EmpresaExpCiudad1;" & sExpCP
   
   Print #nfile, "RCliente;" & sRCliente
   Print #nfile, "RFC A QUIEN SE EXPIDE;" & sRRFC
   Print #nfile, "Rnombre;" & sRNombre
'   Print #nFile, "RApPaterno;" & sRApPat
'   Print #nFile, "RApMaterno;" & sRAPMat
   Print #nfile, "Rcalle;" & sRCalle
   Print #nfile, "Rnumero exterior;" & sRNoExt
   Print #nfile, "Rnumero interior;" & sRNoInt
   Print #nfile, "REntreCalles;" & sREntreCalles
   Print #nfile, "Rlocalidad;" & sRLocalidad
   Print #nfile, "Rcolonia;" & sRCol
   Print #nfile, "Rmunicipio;" & sRMun
   Print #nfile, "Restado;" & sREdo
   Print #nfile, "Rpais;" & sRPais
   Print #nfile, "Rcp;" & sRCP
   Print #nfile, "RDomicilio1;" & sRCalle & " " & sRNoExt
   Print #nfile, "RColonia1;" & sRCol & " CP." & sRCP
   Print #nfile, "RTelefonoNCR;" & SETel
   Print #nfile, "UsoCFDI;" & UsoCFDI '52
   
  ' Print #nfile, "Importe;" & nTotalImporte '8
   If nTipoArch = 4 Then
        Print #nfile, "TIPO DE COMPROBANTE;I" '14
   Else
        Print #nfile, "TIPO DE COMPROBANTE;E" '14
    End If
    Print #nfile, "LugarExpedicion;" & sExpCP '16
 
   Return
   
doGenDet:
   sRutina = "doGenDet"
   If nTipoArch = 3 Or nTipoArch = 4 Then
      For lngI = 1 To nLineas
         If nTipoPed = 12 Then
            If lngI = 1 Then
               sFolioIni = aFolIni(lngI)
            Else
               sFolioIni = Space(13)
            End If
            If lngI = nUltReg Then
               sFolioFin = aFolFin(lngI)
            Else
               sFolioFin = Space(13)
            End If
         Else
            sFolioIni = 0
            sFolioFin = 0
         End If
         
'         aPrecio(nLineas) = Round((rstTmp!Precio / (1 + nIvaProd)), 2)
'         aPrecioSug(nLineas) = rstTmp!Precio
        
       
         sDesglose = ""
        If ntotalPedido > 0 Then
            Print #nfile, "CANTIDAD;1" '65
            Print #nfile, "ClaveProdServ;84141602" '63
            Print #nfile, "NoIdentificacion;7500000000000" '64
            Print #nfile, "ClaveUnidad;E48" '66
            Print #nfile, "U.M.;Unidad de servicio" '67
            Print #nfile, "DESCRIPCION;" & UCase(aDescripcion(lngI)) & " AJUSTE DE SALDOS" '68
            Print #nfile, "PRECIO;" & ntotalPedido '69
            Print #nfile, "IMPORTE BRUTO;" & ntotalPedido '70
        End If
        
        If nImpComision > 0 Then
            Print #nfile, "CANTIDAD;1" '65
            Print #nfile, "ClaveProdServ;84141602" '63
            Print #nfile, "NoIdentificacion;7500000000000" '64
            Print #nfile, "ClaveUnidad;E48" '66
            Print #nfile, "U.M.;Unidad de servicio" '67
            Print #nfile, "DESCRIPCION; AJUSTE CARGO ADMINISTRATIVO" '68
            Print #nfile, "PRECIO;" & nImpComision '69
            Print #nfile, "IMPORTE BRUTO;" & nImpComision '70
        End If
        
        If nImpIVAComision > 0 Then
            Print #nfile, "BaseImpTras;" & nImpComision '72
            Print #nfile, "ClaveImpTras;002" '73
            Print #nfile, "TipoFactorImpTras;Tasa" '74
            Print #nfile, "TasaCuotaImpTras;" & nPrcIVA / 100 '75
            Print #nfile, "ImporteImpTras;" & nImpIVAComision '76
            'Total Impuestos Trasladados
            Print #nfile, "TotalImpuestosTrasladados;" & nImpIVAComision '97
            Print #nfile, "TotalServ;" & nImpComision + nImpIVAComision
        End If
      Next
   End If
   
   Return
   
doGenPie:
   sRutina = "doGenPie"
   
    Print #nfile, "TotalImporte;" & ntotalPedido + nImpComision
    If nImpIVAComision > 0 Then
        Print #nfile, "TotalIva;" & nImpIVAComision
        Print #nfile, "TipoImpuesto;002"
        Print #nfile, "TipoFactor;Tasa"
        Print #nfile, "TasaIva;" & nPrcIVA / 100
    End If
    Print #nfile, "TotalGeneral;" & ntotalPedido + nImpComision + nImpIVAComision
   
   Print #nfile, "LeyendaObservaciones;" & sComentario
   Print #nfile, "IMPORTE CON LETRA;" & sImporteLetras

   Return
   
hdErr:
   Call doErrorLog(gnSucursal, "OPE", ERR.Number, ERR.Description, gstrUsuario, "Module1.doGenArchFE_NCA" & IIf(sRutina <> "", "." & sRutina, ""))
   If blnFileOpen Then
      Close #nfile
   End If
   Resume
End Function
Sub CentraFormaMDI(pfHija As Form)
Dim AltoMenuStatus As Double
Dim l As Integer, t As Integer
Dim i As Integer, j As Integer

For j = 0 To Forms.Count - 1
  If TypeOf Forms(j) Is MDIForm Then
    Exit For
  End If
Next j

For i = 0 To Forms(j).Controls.Count - 1
  If TypeOf Forms(j).Controls(i) Is StatusBar Then
      AltoMenuStatus = Forms(j).Controls(i).Height
      Exit For
  End If
Next i

  ' obtener el offset izquierdo
  l = Forms(j).Left + ((Forms(j).ScaleWidth - pfHija.Width) / 2)
  If (l + pfHija.Width > Screen.Width) Then
    l = Screen.Width - pfHija.Width
  End If

  ' obtener el offset superior
  t = (Forms(j).Top + ((Forms(j).Height - pfHija.Height) / 2)) - AltoMenuStatus
  If (t + pfHija.Height > Screen.Height) Then
    t = Screen.Height - pfHija.Height
  End If

  ' centrar forma hija
  pfHija.Move l, t

End Sub


Public Function Mensajes(tipo As Integer)
Select Case tipo
   Case 0
      RespMsg = MsgBox("¿Desea guardar la información?", vbQuestion + vbYesNo + vbDefaultButton1, "Vale Total")
   Case 1
      RespMsg = MsgBox("¿Desea borrar la información?", vbQuestion + vbYesNo + vbDefaultButton2, "Vale Total")
   Case 2
      RespMsg = MsgBox("¿Desea actualizar la información?", vbQuestion + vbYesNo + vbDefaultButton1, "Vale Total")
   Case 3
      MsgBox "La información está en un estado que no puede ser borrada", vbInformation, "Vale Total"
   Case 4
      MsgBox "La información está en un estado que no puede ser mofdificada", vbInformation, "Vale Total"
   Case 5
      MsgBox "Faltan datos por capturar, verifique", vbInformation, "Vale Total"
   Case 6
      MsgBox "Ocurrio un error, Favor de avisar a sistemas", vbCritical, "Vale Total"
   Case 7
      MsgBox "Favor de seleccionar información para borrar", vbInformation, "Vale Total"
   Case 8
      MsgBox "Datos grabados!", vbInformation, "Vale Total"
End Select

End Function
Sub PosicionaComboEnItemData(pCombo As ComboBox, pValorItemData$)
Dim i%
For i% = 0 To pCombo.ListCount - 1
  If pCombo.ItemData(i%) = pValorItemData Then
    pCombo.ListIndex = i%
    Exit Sub
  End If
Next i%
pCombo.ListIndex = -1
End Sub
Public Function CapImporte(Digito As Integer) As Boolean
  If (Val(Digito) >= 48 And Val(Digito) <= 57) Or Val(Digito) = 46 Or Val(Digito) = 8 Or Val(Digito) = 13 Or Val(Digito) = 45 Then
     CapImporte = True
  Else
     CapImporte = False
  End If
End Function
Public Function CapNumerica(Digito As Integer) As Boolean
  If Val(Digito) >= 48 And Val(Digito) <= 57 Or Val(Digito) = 8 Or Val(Digito) = 13 Then
     CapNumerica = True
  Else
     CapNumerica = False
  End If
End Function
Public Function InIDE() As Boolean
  On Error Resume Next
  Debug.Print 0 / 0
  InIDE = ERR.Number <> 0
End Function
