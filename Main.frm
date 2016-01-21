VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fmrMain 
   Caption         =   "GMImport"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Procesamiento de Deuda"
      Height          =   1005
      Left            =   0
      TabIndex        =   13
      Top             =   7740
      Width           =   3645
   End
   Begin VB.CommandButton cmdPedidosTraspaso 
      Caption         =   "Proceso Pedidos traspaso (CLIENTE OUT)"
      Height          =   1005
      Left            =   180
      TabIndex        =   12
      Top             =   6120
      Width           =   3645
   End
   Begin VB.CommandButton cmdGestionDeuda 
      Caption         =   "Gestion de Cobros (generacion de pendientes)"
      Height          =   1005
      Left            =   180
      TabIndex        =   11
      Top             =   4860
      Width           =   3645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Tablas temporales"
      Height          =   465
      Left            =   10530
      TabIndex        =   10
      Top             =   540
      Width           =   1545
   End
   Begin VB.TextBox txtLog2 
      Height          =   3165
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4680
      Width           =   8115
   End
   Begin VB.CommandButton cmdActualizaPoliza 
      Caption         =   "Actualizacion de datos de Poliza (sin limpieza)"
      Height          =   1005
      Left            =   180
      TabIndex        =   8
      Top             =   3600
      Width           =   3645
   End
   Begin VB.CommandButton cmdLimpiezaPoliza 
      Caption         =   "Actualizacion completa de Poliza (Mensual)"
      Height          =   1005
      Left            =   180
      TabIndex        =   7
      Top             =   2340
      Width           =   3645
   End
   Begin VB.CommandButton cmdProcesamientoDeuda 
      Caption         =   "Procesamiento de Deuda"
      Height          =   1005
      Left            =   180
      TabIndex        =   6
      Top             =   1080
      Width           =   3645
   End
   Begin VB.CommandButton cmdSolicitudes 
      Caption         =   "Solicitudes"
      Height          =   285
      Left            =   5490
      TabIndex        =   5
      Top             =   990
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar gm6 Hist notas, etc"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.TextBox txtLog 
      Height          =   3075
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1440
      Width           =   8115
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   630
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStart1 
      Caption         =   "Procesar altas y bajas de contratos"
      Height          =   465
      Left            =   8370
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblRPS 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   9255
   End
End
Attribute VB_Name = "fmrMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngLastRecord As Long
Dim lngCurrentRecord As Long
Dim blnProcesaLimpieza As Boolean
Private Sub cmdLookupFile_Click()
    CommonDialog1.ShowOpen
    txtSourceFile.Text = CommonDialog1.FileName
End Sub

Private Sub cmdStart_Click()
End Sub

Private Sub cmdDetalle1CC_Click()
    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=conejo;Persist Security Info=True;User ID=sa;Initial Catalog=GM;Data Source=notebook-pipi2"
    p_ADODBConnection.Open
    lngCurrentRecord = 0
    lngLastRecord = 0
    ProcessDetalle1CC
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub cmdDetalle2_Click()
    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=conejo;Persist Security Info=True;User ID=sa;Initial Catalog=GM;Data Source=notebook-pipi2"
    p_ADODBConnection.Open
    lngCurrentRecord = 0
    lngLastRecord = 0
    ProcessDetalle2
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub cmdDetalleCC_Click()
    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=conejo;Persist Security Info=True;User ID=sa;Initial Catalog=GM;Data Source=notebook-pipi2"
    p_ADODBConnection.Open
    lngCurrentRecord = 0
    lngLastRecord = 0
    ProcessDetalleCCLlamado
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub cmdReferenciaOS_Click()
    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=conejo;Persist Security Info=True;User ID=sa;Initial Catalog=GM;Data Source=notebook-pipi2"
    p_ADODBConnection.Open
    lngCurrentRecord = 0
    lngLastRecord = 0
    ProcessReferenciasOS "OC"
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub cmdReferencias_Click()
    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = "Provider=SQLOLEDB.1;Password=conejo;Persist Security Info=True;User ID=sa;Initial Catalog=GM;Data Source=notebook-pipi2"
    p_ADODBConnection.Open
    lngCurrentRecord = 0
    lngLastRecord = 0
    ProcessReferencias "OC"
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub cmdActualizaPoliza_Click()
    blnProcesaLimpieza = False
    cmdLimpiezaPoliza_Click
End Sub

Private Sub cmdGestionDeuda_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recPending As New ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim strUser As String
Dim i As Integer
Dim intContador As Integer

    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
        'strUser = frmSelectUser.getuser()
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = Val(InputBox("Desde que linea del XLS empezamos?", App.Title, "3"))
        txtLog2.Text = "" & vbCrLf
        intContador = 0
        Do While Trim(xlSht.Cells(i, 2)) <> ""
            intContador = intContador + 1
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            'si no tengo cuit busco por poliza en la columna 4
            If Trim(xlSht.Cells(i, 1)) <> "" Then
                recContact1.Open "SELECT * from contact1 WHERE Phone1='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            Else
                recContact1.Open "SELECT contact1.key4, contact1.contact,contact1.accountno from contact1 inner join contact2 on contact1.accountno=contact2.accountno WHERE Userdef05='" & Trim(xlSht.Cells(i, 4)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            End If
            If Not recContact1.EOF Then
                'Busco el user
                strUser = GetUserByCanal(recContact1("Key4").Value & "")
                lblInfo.Caption = "Contacto: " & recContact1("Contact").Value
                DoEvents
                'genero pending de reclamar deuda
                Set recNewPending = New ADODB.Recordset
                recNewPending.Open "SELECT * FROM CAL WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                recNewPending.AddNew
                recNewPending("RECID").Value = GenerateRecID
                recNewPending("Accountno").Value = recContact1("Accountno").Value & ""
                recNewPending("USERID").Value = strUser
                recNewPending("ONDATE").Value = Format(Now, "DD/MM/YYYY")
                recNewPending("ENDDATE").Value = Format(Now, "DD/MM/YYYY")
                recNewPending("ACTVCODE").Value = Trim(xlSht.Cells(i, 3))
                recNewPending("REF").Value = Trim(xlSht.Cells(i, 2))
                recNewPending("Notes").Value = ""
                recNewPending("CREATEBY").Value = "MASTER"
                recNewPending("LASTUSER").Value = "MASTER"
                recNewPending("CREATEON").Value = Format(Now, "DD/MM/YYYY")
                recNewPending("Company").Value = recContact1("Contact").Value & ""
                'rellenos
                recNewPending("ONTIME").Value = "00:00"
                recNewPending("ALARMFLAG").Value = "N"
                recNewPending("ALARMTIME").Value = ""
                recNewPending("RECTYPE").Value = "T" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                recNewPending("LINKRECID").Value = ""
                recNewPending("LOPRECID").Value = ""
                recNewPending.Update
                'txtLog2.Text = txtLog2.Text & Trim(xlSht.Cells(i, 1)) & ";" & Trim(xlSht.Cells(i, 3)) & ";" & recContact1("Contact").Value
                recNewPending.Close
            Else
                'si no encontre el cuit no pasa nada genero log
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & ";" & Trim(xlSht.Cells(i, 4)) & ";No existe" & vbCrLf
            End If
            recContact1.Close
            
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    MsgBox "Proceso finalizado! Procesados: " & intContador, vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub cmdLimpiezaPoliza_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recContact2 As New ADODB.Recordset
Dim recPending As New ADODB.Recordset
Dim recTemp As New ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim i As Integer
Dim strARTAnterior As String
Dim strCLAnterior As String

    'On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
    
        Screen.MousePointer = vbHourglass
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = Val(InputBox("Desde que linea del XLS empezamos?", App.Title, "3"))
        txtLog2.Text = ""
        Screen.MousePointer = vbDefault
        
        'preparo la tabla de cuits procesados
        p_ADODBConnection.Execute "TRUNCATE TABLE TEMPCUITS"
        recTemp.Open "SELECT * FROM TEMPCUITS", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText

        Do While Trim(xlSht.Cells(i, 1)) <> ""
            ' lo primero inserto en la tabla temporal para luego hacer el query de eliminacion
            recTemp.AddNew
            recTemp("CUIT").Value = Trim(xlSht.Cells(i, 1))
            recTemp.Update
            
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE Phone1='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            If Not recContact1.EOF Then
                recContact2.Open "SELECT * FROM CONTACT2 WHERE Accountno='" & ReplaceQuote(recContact1("Accountno").Value) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                lblInfo.Caption = "Contacto: " & recContact1("Contact").Value
                DoEvents
                'guardo ART y cliente anterior y todos los datos anteriores
                strARTAnterior = recContact2("Userdef03").Value & ";" & recContact2("Userdef04").Value & ";" & recContact2("Userdef05").Value & ";" & recContact2("Userdef06").Value & ";" & recContact2("uactcapita").Value & ";" & recContact2("uactmasa").Value & ";" & recContact2("uactaf").Value & ";" & recContact2("uactav").Value & ";" & recContact2("uactprima").Value
                'lo guardo en comments que es mas facil
                recContact2("comments").Value = Left(strARTAnterior, 64)
                
                'actualizo contacto
                'comparo si tiene masa salarial o no
                recContact2("userdef04").Value = Trim(xlSht.Cells(i, 2)) 'ART ACTUAL
                recContact2("userdef05").Value = Trim(xlSht.Cells(i, 3)) 'POLIZA
                recContact2("userdef06").Value = Trim(xlSht.Cells(i, 4)) 'vigencia
                recContact2("uactcapita").Value = IIf(xlSht.Cells(i, 7) = "", 0, xlSht.Cells(i, 6)) ' capitas
                recContact2("uactmasa").Value = IIf(xlSht.Cells(i, 7) = "", 0, xlSht.Cells(i, 7)) 'masa salarial
                recContact2("uactaf").Value = IIf(xlSht.Cells(i, 7) = "", 0, xlSht.Cells(i, 9)) 'AF
                recContact2("uactav").Value = IIf(xlSht.Cells(i, 7) = "", 0, xlSht.Cells(i, 8)) 'AV
                recContact2("uactprima").Value = IIf(xlSht.Cells(i, 7) = "", 0, xlSht.Cells(i, 10)) 'PRIMA
                recContact2("uactFecha").Value = Now
                
                recContact2("userdef03").Value = "CLIENTE"
                recContact2.Update
                'genero un log con los cambios
                recContact2.Close
            Else
                'si no encontre el cuit no pasa nada genero log
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & "No existe en goldmine" & vbCrLf
            End If
            recContact1.Close
            
            i = i + 1
        Loop
        
        Set recTemp = Nothing
        'xlTmp.Workbooks.Close
        xlTmp.Quit
        
        txtLog2.Text = ""
        If blnProcesaLimpieza Then
            'ahora empiezo la limpieza
            'Busco los CUITS QUE NO ESTAN EN EL XLS y que son tipo cliente
            recContact1.Open "SELECT * from contact1 WHERE KEY1='RAZON SOCIAL' and source in('NET BROKER','BROKER ASSET SA') and Phone1 not in (SELECT CUIT FROM TEMPCUITS)  ", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            Do While Not recContact1.EOF
                
                recContact2.Open "SELECT * FROM CONTACT2 WHERE Accountno='" & recContact1("Accountno").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                lblInfo.Caption = "ELIMINADO : " & recContact1("Contact").Value
                DoEvents
                'guardo ART y cliente anterior y todos los datos anteriores
                strARTAnterior = recContact2("Userdef03").Value & ";" & recContact2("Userdef04").Value & ";" & recContact2("Userdef05").Value & ";" & recContact2("Userdef06").Value & ";" & recContact2("uactcapita").Value & ";" & recContact2("uactmasa").Value & ";" & recContact2("uactaf").Value & ";" & recContact2("uactav").Value & ";" & recContact2("uactprima").Value
                txtLog2.Text = txtLog2.Text & "BAJA: " & recContact1("Contact").Value & ";" & recContact1("phone1").Value & ";" & strARTAnterior & vbCrLf
                'lo guardo en comments que es mas facil
                recContact2("comments").Value = Left(strARTAnterior, 64)
                
                'actualizo contacto
                recContact2("userdef04").Value = ""
                recContact2("userdef05").Value = ""
                recContact2("userdef06").Value = ""
                recContact2("uactcapita").Value = 0
                recContact2("uactmasa").Value = 0
                recContact2("uactaf").Value = 0
                recContact2("uactav").Value = 0
                recContact2("uactprima").Value = 0
                
                recContact2("userdef03").Value = "CLIENTE OUT"
                recContact2.Update
                
                'inserto log en el conthist
                recContact2.Close
                recContact1.MoveNext
            Loop
            Set recContact1 = Nothing
        End If
    End If
    
    MsgBox "Proceso finalizado!", vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub cmdPedidosTraspaso_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recContact2 As New ADODB.Recordset
Dim recHist As New ADODB.Recordset
Dim strUser As String
Dim i As Integer
Dim intContador As Integer

    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
        strUser = frmSelectUser.getuser()
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = Val(InputBox("Desde que linea del XLS empezamos?", App.Title, "3"))
        txtLog2.Text = "" & vbCrLf
        intContador = 0
        recHist.Open "SELECT * FROM CONTHIST WHERE Accountno='AA'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        
        Do While Trim(xlSht.Cells(i, 2)) <> ""
            intContador = intContador + 1
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE Phone1='" & Trim(xlSht.Cells(i, 2)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contacto: " & recContact1("Contact").Value
                DoEvents
                recContact2.Open "SELECT * FROM contact2 where accountno='" & recContact1("Accountno").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                'Tomo la ARTActual y genero el HIST
                'Actualizo los campos de contact2 CONTADOR y ARTFUTURA
                recContact2("UCONTDORO").Value = Val(0 & recContact2("UCONTDORO").Value) + 1
                recContact2("UARTFUTURA").Value = Trim(xlSht.Cells(i, 3))
                recContact2.Update
                'creo el historial
                recHist.AddNew
                    recHist("Accountno").Value = recContact1("Accountno").Value
                    recHist("recid").Value = GenerateRecID
                    recHist("actvcode").Value = Left(recContact2("USERDEF04").Value & "", 3) 'aca la ART ACTUAL
                    ' uso un execute para convertir
                    Set recTemp = New ADODB.Recordset
                    recTemp.Open "SELECT CONVERT(binary(60), '" & "Fecha: " & Format(Now, "dd/mm/yyyy") & " pedido por " & Trim(xlSht.Cells(i, 2)) & "') AS Binario", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                    If Not recTemp.EOF Then
                        'recHist("NOTES").Value = Trim(xlSht.Cells(i, 5))
                        recHist("NOTES").Value = recTemp("Binario").Value
                    End If
                    Set recTemp = Nothing
                    recHist("resultcode").Value = ""
                    recHist("REF").Value = "PEDIDO TRASPASO OUT”"
                    recHist("USERID").Value = strUser
                    recHist("SRECTYPE").Value = "T"
                    recHist("RECTYPE").Value = "T"
                    recHist("ONDATE").Value = Format(Now, "yyyy/mm/dd")
                    'recNewHist("ONTIME").Value = recOldHist("ONTIME").Value & ""
                    recHist("loprecid").Value = ""
                    recHist("completedid").Value = ""
                
                recHist.Update
                
            Else
                'si no encontre el cuit no pasa nada genero log
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 2)) & ";" & Trim(xlSht.Cells(i, 3)) & ";No existe" & vbCrLf
            End If
            recContact1.Close
            If recContact2.State = adStateOpen Then
                recContact2.Close
            End If
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recContact2 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    MsgBox "Proceso finalizado! Procesados: " & intContador, vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName



End Sub

Private Sub cmdProcesamientoDeuda_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recPending As New ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim strLastUser As String
Dim i As Integer
Dim intContador As Integer

    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
    
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = 2
        txtLog2.Text = "Detalle Deuda Pendiente" & vbCrLf
        intContador = 0
        Do While Trim(xlSht.Cells(i, 1)) <> ""
            intContador = intContador + 1
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE Phone1='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contacto: " & recContact1("Contact").Value
                DoEvents
                'Busco si tiene un pending
                recPending.Open "SELECT * FROM CAL WHERE AccountNo='" & recContact1("Accountno").Value & "' AND (REF ='SOLICITUD PRESENTADA ART' OR REF ='RECLAMO DEUDA')", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                If Not recPending.EOF Then
                    'guardo el user
                    strLastUser = recPending("USERID").Value
                    'Si encontre la actividad entonces sigo
                    If UCase(Trim(xlSht.Cells(i, 2))) = "OK" Then
                        CompleteCalendarActivity recPending("RECID").Value, "OK", ""
                    Else
                        If UCase(Trim(xlSht.Cells(i, 2))) = "DEU" Then
                            CompleteCalendarActivity recPending("RECID").Value, Trim(xlSht.Cells(i, 2)), "Valor Deuda: " & Trim(xlSht.Cells(i, 3))
                            'genero pending de reclamar deuda
                            Set recNewPending = New ADODB.Recordset
                            recNewPending.Open "SELECT * FROM CAL WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                            recNewPending.AddNew
                            recNewPending("RECID").Value = GenerateRecID
                            recNewPending("Accountno").Value = recContact1("Accountno").Value & ""
                            recNewPending("Company").Value = recContact1("Contact").Value & ""
                            recNewPending("USERID").Value = strLastUser
                            'recNewPending("ONDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                            'recNewPending("ENDDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                            recNewPending("ONDATE").Value = Format(Now, "DD/MM/YYYY")
                            recNewPending("ENDDATE").Value = Format(Now, "DD/MM/YYYY")
                            
                            recNewPending("ACTVCODE").Value = Trim(xlSht.Cells(i, 4)) ' PONGO la ART
                            recNewPending("REF").Value = "RECLAMO DEUDA"
                            recNewPending("Notes").Value = ConvertStringToBinary("Deuda al " & Format(Now, "dd/mm/yyyy") & " : $ " & Trim(xlSht.Cells(i, 3)) & "")
                            recNewPending("CREATEBY").Value = strLastUser
                            recNewPending("LASTUSER").Value = strLastUser
                            recNewPending("CREATEON").Value = Format(Now, "DD/MM/YYYY")
                            'rellenos
                            recNewPending("ONTIME").Value = "00:00"
                            recNewPending("ALARMFLAG").Value = "N"
                            recNewPending("ALARMTIME").Value = ""
                            recNewPending("RECTYPE").Value = "T" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                            recNewPending("LINKRECID").Value = ""
                            recNewPending("LOPRECID").Value = ""
                            
                            recNewPending.Update
                            txtLog2.Text = txtLog2.Text & Trim(xlSht.Cells(i, 1)) & ";" & Trim(xlSht.Cells(i, 3)) & ";" & recContact1("Contact").Value & vbCrLf
                            recNewPending.Close
                        Else
                            CompleteCalendarActivity recPending("RECID").Value, Left(Trim(xlSht.Cells(i, 2)), 3), ""
                        End If
                    End If
                Else
                    txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & ";No tiene Solicitud presentada" & vbCrLf
                End If
                recPending.Close
            Else
                'si no encontre el cuit no pasa nada genero log
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & ";No existe" & vbCrLf
            End If
            recContact1.Close
            
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    MsgBox "Proceso finalizado! : " & intContador, vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub cmdSolicitudes_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recHist As ADODB.Recordset
Dim i As Integer
Dim strAccountNo As String
Dim recTemp As ADODB.Recordset
    'On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
    
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = 2
        Set recHist = New ADODB.Recordset
        recHist.Open "SELECT * FROM CONTHIST WHERE Accountno='AA'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        
        Do While Trim(xlSht.Cells(i, 1)) <> ""
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE PHONE1='" & Trim(xlSht.Cells(i, 3)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contacto: " & recContact1("Company").Value
                strAccountNo = recContact1("Accountno").Value
                DoEvents
                'Edito el contact2
                Set recContact2 = New ADODB.Recordset
                recContact2.Open "SELECT * FROM CONTACT2 WHERE Accountno='" & strAccountNo & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                If Not recContact2.EOF Then
                    recContact2("USERDEF03").Value = Trim(xlSht.Cells(i, 7))
                    recContact2.Update
                End If
                Set recContact2 = Nothing
                'creo el historial
                recHist.AddNew
                    recHist("Accountno").Value = strAccountNo
                    recHist("recid").Value = GenerateRecID
                    recHist("actvcode").Value = Trim(xlSht.Cells(i, 4))
                    ' uso un execute para convertir
                    Set recTemp = New ADODB.Recordset
                    recTemp.Open "SELECT CONVERT(binary(60), '" & Trim(xlSht.Cells(i, 5)) & "') AS Binario", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                    If Not recTemp.EOF Then
                        'recHist("NOTES").Value = Trim(xlSht.Cells(i, 5))
                        recHist("NOTES").Value = recTemp("Binario").Value
                    End If
                    Set recTemp = Nothing
                    recHist("resultcode").Value = Trim(xlSht.Cells(i, 6))
                    recHist("REF").Value = Trim(xlSht.Cells(i, 8))
                    recHist("USERID").Value = "MASTER"
                    recHist("SRECTYPE").Value = "T"
                    recHist("RECTYPE").Value = "T"
                    'recNewHist("ONDATE").Value = recOldHist("ONDATE").Value & ""
                    'recNewHist("ONTIME").Value = recOldHist("ONTIME").Value & ""
                    recHist("loprecid").Value = ""
                    recHist("completedid").Value = ""
                
                recHist.Update
            End If
            recContact1.Close
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub cmdStart1_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recTemp As New ADODB.Recordset
Dim i As Integer

    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
    
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = 2
        recTemp.Open "SELECT * FROM TEMPCUITS", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        Do While Trim(xlSht.Cells(i, 1)) <> ""
            ' lo primero inserto en la tabla temporal para luego hacer el query de eliminacion
            recTemp.AddNew
            recTemp("CUIT").Value = Trim(xlSht.Cells(i, 1))
            recTemp.Update
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE KEY1='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contacto: " & recContact1("Company").Value
                DoEvents
                'Piso el contrato
                recContact1("KEY5").Value = xlSht.Cells(i, 2) & xlSht.Cells(i, 3)
                recContact1("u_KEY5").Value = xlSht.Cells(i, 2) & xlSht.Cells(i, 3)
                recContact1.Update
            Else
                'si no encontre el cuit logueo
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & vbCrLf
            End If
            recContact1.Close
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub Command1_Click()
Dim recContact1 As New ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recContact2Search As New ADODB.Recordset
Dim recNewHist As ADODB.Recordset
Dim recOldHist As ADODB.Recordset
Dim recOldContact1 As ADODB.Recordset
Dim recOldContact2 As ADODB.Recordset
Dim recContSuppNew As ADODB.Recordset
Dim strNewAccountno As String
Dim strOLDAccountNo As String

    On Error GoTo ERRORHANDLER
    
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
    p_ADODBConnection.Open
    
    Set p_ADODBConnectionOLDGM = New ADODB.Connection
    p_ADODBConnectionOLDGM.ConnectionString = GetIni("DataSourceName", "CSOLD")
    p_ADODBConnectionOLDGM.Open
        
    'recorro todo la contact1
    recContact1.Open "SELECT * FROM Contact1 where ext4='10'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    'Ahora tomo solo los nuevos
    'recContact1.Open "SELECT * FROM Contact1", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    'ahora tomo solo los que no tienen aun historial
    'recContact1.Open "select * from CONTACT1 where ACCOUNTNO not in (select distinct ACCOUNTNO from CONTHIST) and ACCOUNTNO in (select accountno from CONTACT2 where uoldacc is NOT null)", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    recContact2Search.Open "SELECT * FROM Contact2", p_ADODBConnection, adOpenDynamic, adLockReadOnly, adCmdText
    Do While Not recContact1.EOF
        lblInfo.Caption = recContact1("contact").Value & ""
        DoEvents
        'strOLDAccountNo = recContact1("").Value & "" ' BUSCO POR CAMPO DE COMBINACION
        strNewAccountno = recContact1("AccountNo").Value & ""
        'Busco el viejo ACCID
        recContact2Search.MoveFirst
        recContact2Search.Find "Accountno='" & strNewAccountno & "'"
        'solo entro si tiene registro viejo
        strOLDAccountNo = recContact2Search("UOLDACC").Value & ""
        If Trim(strOLDAccountNo) <> "" And Len(strOLDAccountNo) > 11 Then
            
            'Ahora Abro los viejos
            Set recOldContact1 = New ADODB.Recordset
            Set recOldContact2 = New ADODB.Recordset
            Set recOldHist = New ADODB.Recordset
            recOldContact1.Open "SELECT * From Contact1 where Accountno='" & strOLDAccountNo & "'", p_ADODBConnectionOLDGM, adOpenDynamic, adLockReadOnly, adCmdText
            recOldContact2.Open "SELECT * From Contact2 where Accountno='" & strOLDAccountNo & "'", p_ADODBConnectionOLDGM, adOpenDynamic, adLockReadOnly, adCmdText
            recOldHist.Open "SELECT * From ContHist where Accountno='" & strOLDAccountNo & "'", p_ADODBConnectionOLDGM, adOpenDynamic, adLockReadOnly, adCmdText
            'abro el contact2 nuevo
            If Not recOldContact1.EOF Then
                Set recContact2 = New ADODB.Recordset
                recContact2.Open "SELECT * From Contact2 where Accountno='" & strNewAccountno & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                'Copio los campos del contact1 y contact2
                recContact1("Notes").Value = recOldContact1("notes").Value & ""
                recContact1.Update
                recContact2("prevresult").Value = recOldContact2("prevresult").Value & ""
                recContact2.Update
                'Aca tengo que crear el contsupp con los campos del contact2
                Set recContSuppNew = New ADODB.Recordset
                recContSuppNew.Open "SELECT * FROM CONTSUPP WHERE ACCOUNTNO='AA'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                recContSuppNew.AddNew
                    recContSuppNew("ACCOUNTNO").Value = strNewAccountno
                    recContSuppNew("RECID").Value = GenerateRecID
                    recContSuppNew("rectype").Value = "P"
                    recContSuppNew("contact").Value = "SOLICITUD PRESENTADA"
                    recContSuppNew("u_contact").Value = "SOLICITUD PRESENTADA"
                    recContSuppNew("phone").Value = "xH4WFPYQ$;"
                    recContSuppNew("TITLE").Value = recOldContact2("UCOTARTDES").Value & ""
                    recContSuppNew("LINKACCT").Value = recOldContact2("UCOTCAPIT").Value & ""
                    recContSuppNew("FAX").Value = recOldContact2("UCOTFIJO").Value & ""
                    recContSuppNew("COUNTRY").Value = recOldContact2("UCOTMASA").Value & ""
                    recContSuppNew("DEAR").Value = recOldContact2("UCOTVAR").Value & ""
                    If Not IsNull(recOldContact2("UFECHACOT").Value) Then
                        recContSuppNew("LASTDATE").Value = recOldContact2("UFECHACOT").Value & ""
                    End If
                    recContSuppNew("zip").Value = Left(Mid(recOldContact2("prevresult").Value & "", 1, InStr(1, recOldContact2("prevresult").Value & "", " ")), 10)
                    recContSuppNew("u_contsupref").Value = ""
                    recContSuppNew("u_address1").Value = ""
                recContSuppNew.Update
                Set recContSuppNew = Nothing
                'Ahora transfiero los history
                If Not recOldHist.EOF Then
                    Set recNewHist = New ADODB.Recordset
                    recNewHist.Open "SELECT * from ContHist where Accountno='PIPI'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                    Do While Not recOldHist.EOF
                        lblRPS.Caption = "HIST: " & recOldHist("REF").Value & ""
                        DoEvents
                        recNewHist.AddNew
                            recNewHist("USERID").Value = "MASTER"
                            recNewHist("ACCOUNTNO").Value = strNewAccountno
                            recNewHist("SRECTYPE").Value = recOldHist("SRECTYPE").Value & ""
                            recNewHist("RECTYPE").Value = recOldHist("RECTYPE").Value & ""
                            recNewHist("ONDATE").Value = recOldHist("ONDATE").Value & ""
                            recNewHist("ONTIME").Value = recOldHist("ONTIME").Value & ""
                            recNewHist("ACTVCODE").Value = recOldHist("ACTVCODE").Value & ""
                            recNewHist("RESULTCODE").Value = recOldHist("RESULTCODE").Value & ""
                            recNewHist("STATUS").Value = recOldHist("STATUS").Value & ""
                            recNewHist("DURATION").Value = recOldHist("DURATION").Value & ""
                            recNewHist("UNITS").Value = recOldHist("UNITS").Value & ""
                            recNewHist("REF").Value = recOldHist("REF").Value & ""
                            'recNewHist("NOTES").Value = recOldHist("NOTES").Value & ""
                            recNewHist("CREATEBY").Value = recOldHist("CREATEBY").Value & ""
                            recNewHist("CREATEON").Value = recOldHist("CREATEON").Value & ""
                            recNewHist("CREATEAT").Value = recOldHist("CREATEAT").Value & ""
                            recNewHist("LASTUSER").Value = recOldHist("LASTUSER").Value & ""
                            recNewHist("LASTDATE").Value = recOldHist("LASTDATE").Value & ""
                            recNewHist("LASTTIME").Value = recOldHist("LASTTIME").Value & ""
                            'recNewHist("EXT").Value = recOldHist("EXT").Value & ""
                            recNewHist("recid").Value = GenerateRecID
                            recNewHist("loprecid").Value = ""
                            recNewHist("completedid").Value = ""
                        recNewHist.Update
                        recOldHist.MoveNext
                    Loop
                    recNewHist.Close
                    Set recNewHist = Nothing
                End If
                recOldHist.Close
                Set recOldHist = Nothing
                Set recOldContact1 = Nothing
                Set recOldContact2 = Nothing
                recContact2.Close
                Set recContact2 = Nothing
            End If
        End If
        WriteLog strOLDAccountNo & ""
        recContact1.MoveNext
    Loop
    
    
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName

End Sub

Private Sub Command2_Click()
    On Error Resume Next
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        p_ADODBConnection.Execute "CREATE TABLE TEMPCUITS(CUIT varchar(25) NULL) "
        p_ADODBConnection.Close
        MsgBox "Tabla TEMPCUITS creada!", vbExclamation

End Sub

Private Sub Command3_Click()
Dim xlTmp As Excel.Application
Dim strFile As String
Dim xlSht As Excel.Worksheet
Dim recContact1 As New ADODB.Recordset
Dim recPending As New ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim strLastUser As String
Dim i As Integer
Dim intContador As Integer

    On Error GoTo ERRORHANDLER
    
    CommonDialog1.ShowOpen
    strFile = CommonDialog1.FileName
    If strFile <> "" Then
        Set p_ADODBConnection = New ADODB.Connection
        p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
        p_ADODBConnection.Open
        lngCurrentRecord = 0
        lngLastRecord = 0
        
    
        Set xlTmp = New Excel.Application
        xlTmp.Workbooks.Open strFile
        Set xlSht = xlTmp.Sheets(1)
        i = 2
        txtLog2.Text = "Detalle Deuda Pendiente" & vbCrLf
        intContador = 0
        Do While Trim(xlSht.Cells(i, 1)) <> ""
            intContador = intContador + 1
            'Debug.Print xlSht.Cells(i, 2) & " - " & i
            recContact1.Open "SELECT * from contact1 WHERE Phone1='" & Trim(xlSht.Cells(i, 1)) & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recContact1.EOF Then
                lblInfo.Caption = "Contacto: " & recContact1("Contact").Value
                DoEvents
                If Trim(xlSht.Cells(i, 2)) = "DEU" Then
                    'Tomo el usuario del canal de ventas
                    Select Case Trim(recContact1("KEY4").Value)
                        Case "AREA DE SERVICIOS"
                            strLastUser = "ASERV"
                        Case "GABRIEL PEDEVILLA"
                            strLastUser = "GPEDEVIL"
                        Case "LEANDRO CASADO"
                            strLastUser = "LCASADO"
                        Case "LEONARDO MONTES"
                            strLastUser = "LMONTES"
                        Case "MARCOS SARA"
                            strLastUser = "MSARA"
                        Case "MARIANO URRERE PON"
                            strLastUser = "MURRERE"
                        Case "MATIAS DE LA TORRE"
                            strLastUser = "MTORRE"
                        Case "SOL DUBOIS"
                            strLastUser = "SDUBOIS"
                        Case "LUIS VILLARREAL"
                            strLastUser = "LVILLARE"
                        Case Else
                            strLastUser = "CANTOLIN"
                    End Select
                    'genero pending de reclamar deuda
                    Set recNewPending = New ADODB.Recordset
                    recNewPending.Open "SELECT * FROM CAL WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                    recNewPending.AddNew
                    recNewPending("RECID").Value = GenerateRecID
                    recNewPending("Accountno").Value = recContact1("Accountno").Value & ""
                    recNewPending("Company").Value = recContact1("Contact").Value & ""
                    recNewPending("USERID").Value = strLastUser
                    'recNewPending("ONDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                    'recNewPending("ENDDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                    recNewPending("ONDATE").Value = Format(Now, "DD/MM/YYYY")
                    recNewPending("ENDDATE").Value = Format(Now, "DD/MM/YYYY")
                    
                    recNewPending("ACTVCODE").Value = Trim(xlSht.Cells(i, 4)) ' PONGO la ART
                    recNewPending("REF").Value = "RECLAMO DEUDA"
                    recNewPending("Notes").Value = ConvertStringToBinary("Deuda al " & Format(Now, "dd/mm/yyyy") & " : $ " & Trim(xlSht.Cells(i, 3)) & "")
                    recNewPending("CREATEBY").Value = strLastUser
                    recNewPending("LASTUSER").Value = strLastUser
                    recNewPending("CREATEON").Value = Format(Now, "DD/MM/YYYY")
                    'rellenos
                    recNewPending("ONTIME").Value = "00:00"
                    recNewPending("ALARMFLAG").Value = "N"
                    recNewPending("ALARMTIME").Value = ""
                    recNewPending("RECTYPE").Value = "T" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                    recNewPending("LINKRECID").Value = ""
                    recNewPending("LOPRECID").Value = ""
                    
                    recNewPending.Update
                    txtLog2.Text = txtLog2.Text & Trim(xlSht.Cells(i, 1)) & ";" & Trim(xlSht.Cells(i, 3)) & ";" & recContact1("Contact").Value & vbCrLf
                    recNewPending.Close
                End If
            Else
                'si no encontre el cuit no pasa nada genero log
                txtLog.Text = txtLog.Text & Trim(xlSht.Cells(i, 1)) & ";No existe" & vbCrLf
            End If
            recContact1.Close
            
            i = i + 1
        Loop
        Set recContact1 = Nothing
        Set recTemp = Nothing
        xlTmp.Workbooks.Close
        xlTmp.Quit
    End If
    
    MsgBox "Proceso finalizado! : " & intContador, vbInformation, App.Title
    p_ADODBConnection.Close
    Exit Sub
    
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName


End Sub

Private Sub Form_Load()
    'Levanto los parametros de la reg
    lblInfo.Caption = "Importacion aun no iniciada"
    blnProcesaLimpieza = True
    If Trim(Command$) = "/GMVC" Then
        'Corro el proceso de integracion con vocalcom
        GMVocalcommPreVenta
        End
    End If
End Sub

Private Sub ProcessDetalle1()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecrm where campo1 is not null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Cirugia OC"
                recDetalle("u_Contact").Value = "Cirugia OC"
                recDetalle("Contsupref") = recImportacion("Campo1").Value
                recDetalle("u_Contsupref") = recImportacion("Campo1").Value
                'Aca ahora segun lo que viene hacemos algo
                strCampo3 = ""
                strCampo2 = Left(recImportacion("Campo2").Value, 20)
                If InStr(1, UCase(strCampo2), "TRAN") Then
                    strCampo2 = "Transitoria"
                Else
                    If InStr(1, UCase(strCampo2), "DEF") Then
                        strCampo2 = "Definitiva"
                    Else
                        strCampo3 = strCampo2
                        strCampo2 = ""
                    End If
                End If
                recDetalle("Title") = strCampo2
                recDetalle.Update
                
                'Ahora chequeo si estos putos pusieron lo que va en el campo3 ademas en el 2
                If strCampo3 <> "" Then
                    'inserto un detalle adicional, que puede ser el unico
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Patologia OC"
                    recDetalle("u_Contact").Value = "Patologia OC"
                    recDetalle("Contsupref") = Left(strCampo3, 20)
                    recDetalle("u_Contsupref") = Left(strCampo3, 20)
                    recDetalle.Update
                End If
                    
                If Not IsNull(recImportacion("Campo3").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Patologia OC"
                    recDetalle("u_Contact").Value = "Patologia OC"
                    recDetalle("Contsupref") = Left(recImportacion("Campo3").Value, 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("Campo3").Value, 20)
                    recDetalle.Update
                End If
            recDetalle.Close
            Set recDetalle = Nothing
        End If
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub InsertOrUpdateContact(objContact As Contact)
Dim recContact1 As ADODB.Recordset
Dim recContsup As ADODB.Recordset

    On Error GoTo ERRORHANDLER
    'Primero busco el contacto si no esta lo creo
    Set recContact1 = New ADODB.Recordset
    recContact1.Open "SELECT * FROM Contact1 WHERE Key1='" & objContact.CUIT & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If recContact1.EOF Then
        'Inserto contacto
        recContact1.AddNew
        recContact1("ACCOUNTNO").Value = GenerateAccountno(objContact.RazonSocial)
        recContact1("COMPANY").Value = Left(Trim(objContact.RazonSocial), 40)
        recContact1("U_COMPANY").Value = Left(Trim(objContact.RazonSocial), 40)
        recContact1("U_CONTACT").Value = Left(Trim(objContact.RazonSocial), 40)
        'recContact1("Contact").Value
        'recContact1("LASTNAME").Value
        'recContact1("DEPARTMENT").Value
        'recContact1("Title").Value
        'recContact1("SECR").Value
        recContact1("PHONE1").Value = " "
        recContact1("ADDRESS1").Value = Left(objContact.DomicilioAddress, 40)
        'recContact1("ADDRESS2").Value
        'recContact1("ADDRESS3").Value
        recContact1("CITY").Value = objContact.DomicilioCity
        recContact1("U_CITY").Value = objContact.DomicilioCity
        'recContact1("State").Value
        recContact1("U_State").Value = ""
        recContact1("ZIP").Value = objContact.DomicilioZip
        recContact1("COUNTRY").Value = "Argentina"
        recContact1("U_COUNTRY").Value = "Argentina"
        recContact1("KEY1").Value = objContact.CUIT
        recContact1("KEY2").Value = ""
        recContact1("KEY3").Value = ""
        recContact1("KEY4").Value = ""
        recContact1("KEY5").Value = ""
        recContact1("U_KEY1").Value = objContact.CUIT
        recContact1("U_KEY2").Value = ""
        recContact1("U_KEY3").Value = ""
        recContact1("U_KEY4").Value = ""
        recContact1("U_KEY5").Value = ""
        recContact1("NOTES").Value = objContact.Domicilio
        'recContact1("CREATEBY").Value
        recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
        recContact1("RECID").Value = GenerateRecID
        recContact1("U_Lastname").Value = " "
        recContact1.Update
    Else
    End If
'Siempre inserto el detalle
    'Inserto presentacion
    Set recContsup = New ADODB.Recordset
    recContsup.Open "SELECT * FROM Contsupp  WHERE accountno='" & recContact1("ACCOUNTNO").Value & "' AND contsupref='" & objContact.Periodo & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If recContsup.EOF Then
        recContsup.AddNew
        recContsup("ACCOUNTNO").Value = recContact1("ACCOUNTNO").Value
        recContsup("REcID").Value = GenerateRecID
    End If
    
    recContsup("Contact").Value = "Presentacion"
    recContsup("u_Contact").Value = "Presentacion"
    recContsup("Contsupref").Value = objContact.Periodo 'Campo1
    recContsup("u_Contsupref").Value = objContact.Periodo
    recContsup("rectype").Value = "P"
    recContsup("linkacct").Value = TrimZeros(objContact.CodigoActividad) 'Campo2
    recContsup("country").Value = TrimZeros(objContact.CodigoART) 'Campo3
    recContsup("zip").Value = AddComma(TrimZeros(objContact.PagoTotal)) 'Campo4
    recContsup("ext").Value = TrimZeros(objContact.Alicuta) 'Campo5
    recContsup("state").Value = TrimZeros(objContact.Fechapresentacion) 'Campo6
    recContsup("Address1").Value = AddComma(TrimZeros(objContact.Fijo)) 'Campo7
    recContsup("Address2").Value = AddComma(TrimZeros(objContact.MasaSalarial)) 'Campo8
    
    recContsup.Update

    
    recContact1.Close
    recContsup.Close
    
    Set recContact1 = Nothing
    Set recContsup = Nothing
    Exit Sub
ERRORHANDLER:
    MsgBox GetADODBErrorMessage(p_ADODBConnection.Errors), vbCritical, App.EXEName
    On Error Resume Next
    Set recContact1 = Nothing
    Set recContsup = Nothing
End Sub

Private Function GetAccountno(strRegistro As String) As String
Dim recSearch As New ADODB.Recordset

    On Error GoTo ERRORHANDLER:
    GetAccountno = ""
    recSearch.Open "select Accountno from Contact2 where Userdef01='" & Trim(strRegistro) & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not recSearch.EOF Then
        GetAccountno = recSearch("Accountno").Value & ""
        
    End If
    
    Set recSearch = Nothing
    Exit Function
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName


End Function


Private Sub ProcessDetalle2()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecrm  where campo1 is not null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle que es el del primer llamado
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Primer Llamado OC"
                recDetalle("u_Contact").Value = "Primer Llamado OC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("SERVICIO_ASESORIA").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("CANT_VISITAS").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("VISITO_COMPETENCIA").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("EMPRESA_COMPETENCIA").Value & "", 10) 'Campo4
                'recDetalle("ext").Value = TrimZeros(objContact.Alicuta) 'Campo5
                'recDetalle("state").Value = Left(recImportacion("DONDE_COMPRA").Value & "", 20) 'Campo6
                recDetalle("Address1").Value = Left(recImportacion("OBSERVACIONES").Value & "", 40) 'Campo7
                'recDetalle("Address2").Value = Left(recImportacion("OBTIENE_PRODUCTO").Value & "", 40) 'Campo8
                
                recDetalle.Update
                
                If Not IsNull(recImportacion("MARCA_BOLSA").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = Left(recImportacion("MODELO_COLOPLAST").Value & "", 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("MODELO_COLOPLAST").Value & "", 20)
                    recDetalle("Address1") = Left(recImportacion("MARCA_BOLSA").Value, 40)
                    If Trim(recImportacion("OTRA_COMPETENCIA").Value & "") <> "" Then
                        recDetalle("Address1") = Left(recImportacion("OTRA_COMPETENCIA").Value, 40)
                    End If

                    recDetalle.Update
                End If
                'Para la marca de la competencia ahora lo ponemos en el mismo
'                If Not IsNull(recImportacion("OTRA_COMPETENCIA").Value) Then
'                    recDetalle.AddNew
'                    recDetalle("Accountno").Value = ReplaceQuote(strAccountno)
'                    recDetalle("recid").Value = GenerateRecID
'                    recDetalle("rectype").Value = "P"
'                    recDetalle("Contact").Value = "Productos Utilizados OC"
'                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
'                    recDetalle("Contsupref") = ""
'                    recDetalle("u_Contsupref") = ""
'                    recDetalle("Title") = Left(recImportacion("OTRA_COMPETENCIA").Value, 20)
'                    recDetalle.Update
'                End If
                'Producto recomendad
                If Not IsNull(recImportacion("PRODUCT").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Producto Recomendado OC"
                    recDetalle("u_Contact").Value = "Producto Recomendado OC"
                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value, 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value, 20)
                    
                    recDetalle.Update
                End If
            
            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessReferencias(Base As String)
Dim recImportacion As New ADODB.Recordset
Dim recReferencia As New ADODB.Recordset
Dim recContacto As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strRecID1 As String
Dim strRecID2 As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion con las entidades
    If Base = "OC" Then
        recImportacion.Open "select * from basecrm inner join instituciones on basecrm.Entity=instituciones.Institucion", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        recImportacion.Open "select * from basecc inner join instituciones on basecc.Entity=instituciones.Institucion", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Obtengo el contacto
            recContacto.Open "SELECT * from Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            'Primero obtengo la referencia, la referencia la vamos a buscar por el campo completo
            recReferencia.Open "SELECT * from contact1 inner join contact2 on contact1.accountno=contact2.accountno where uinstlargo='" & recImportacion("Institucion").Value & "'", p_ADODBConnection, adOpenKeyset, adLockReadOnly, adcmdtect
            'Inserto la referencia
            If Not recReferencia.EOF Then
                'Inserto la referencia de ida
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    On Error GoTo retry
                    recDetalle("Accountno").Value = recReferencia("Accountno").Value
                    strRecID1 = GenerateRecID
                    recDetalle("recid").Value = strRecID1
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("u_Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("Contsupref") = "Institucion donde se atendio" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Institucion donde se atendio"
                    recDetalle("Title").Value = recContacto("Accountno").Value 'aca el accountno del otro
                    strRecID2 = GenerateRecID ' Lo genero aca para tenerlo abajo
                    recDetalle("linkacct").Value = strRecID2 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "T"
                    recDetalle.Update
                recDetalle.Close
                'inserto la referencia de vuelta
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    On Error GoTo retry
                    recDetalle("Accountno").Value = recContacto("Accountno").Value
                    'strRecID2 = GenerateRecID
                    recDetalle("recid").Value = strRecID2
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("Para:" & recReferencia("Company").Value, 30)
                    recDetalle("u_Contact").Value = Left("Para:" & recContacto("Company").Value, 30)
                    recDetalle("Contsupref") = "Institucion donde se atendio" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Institucion donde se atendio"
                    recDetalle("Title").Value = recReferencia("Accountno").Value 'aca el accountno del otro
                    recDetalle("linkacct").Value = strRecID1 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "R"
                    recDetalle.Update
                recDetalle.Close
                                
                Set recDetalle = Nothing
                recContacto.Close
                recReferencia.Close
            End If
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub


Private Sub ProcessReferenciasOS(Base As String)
Dim recImportacion As New ADODB.Recordset
Dim recReferencia As New ADODB.Recordset
Dim recContacto As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strRecID1 As String
Dim strRecID2 As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion con las entidades
    If Base = "OC" Then
        recImportacion.Open "select * from basecrm inner join os on basecrm.Membership=os.obrasocialnl", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Else
        recImportacion.Open "select * from basecc inner join os on basecc.Membership=os.obrasocialnl", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    End If
    
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Obtengo el contacto
            recContacto.Open "SELECT * from Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            'Primero obtengo la referencia, la referencia la vamos a buscar por el campo completo
            recReferencia.Open "SELECT * from contact1 inner join contact2 on contact1.accountno=contact2.accountno where uinstlargo='" & recImportacion("obrasocialnl").Value & "'", p_ADODBConnection, adOpenKeyset, adLockReadOnly, adcmdtect
            'Inserto la referencia
            If Not recReferencia.EOF Then
                'Inserto la referencia de ida
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    'On Error GoTo retry
                    recDetalle("Accountno").Value = recReferencia("Accountno").Value
                    strRecID1 = GenerateRecID
                    recDetalle("recid").Value = strRecID1
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("u_Contact").Value = Left("A:" & recContacto("Contact").Value, 30)
                    recDetalle("Contsupref") = "Obra social de pertenencia" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Obra social de pertenencia"
                    recDetalle("Title").Value = recContacto("Accountno").Value 'aca el accountno del otro
                    strRecID2 = GenerateRecID ' Lo genero aca para tenerlo abajo
                    recDetalle("linkacct").Value = strRecID2 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "T"
                    recDetalle.Update
                recDetalle.Close
                'inserto la referencia de vuelta
                Set recDetalle = New ADODB.Recordset
                recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle.AddNew
                    'On Error GoTo retry
                    recDetalle("Accountno").Value = recContacto("Accountno").Value
                    'strRecID2 = GenerateRecID
                    recDetalle("recid").Value = strRecID2
                    recDetalle("rectype").Value = "R"
                    recDetalle("Contact").Value = Left("Para:" & recReferencia("Company").Value, 30)
                    recDetalle("u_Contact").Value = Left("Para:" & recContacto("Company").Value, 30)
                    recDetalle("Contsupref") = "Obra social de pertenencia" 'Aca va el detalle
                    recDetalle("u_Contsupref") = "Obra social de pertenencia"
                    recDetalle("Title").Value = recReferencia("Accountno").Value 'aca el accountno del otro
                    recDetalle("linkacct").Value = strRecID1 'Aca va a ir el recid del otro registro
                    recDetalle("ext").Value = "R"
                    recDetalle.Update
                recDetalle.Close
                                
                Set recDetalle = Nothing
                recContacto.Close
                recReferencia.Close
            Else
                recContacto.Close
                recReferencia.Close
            
            End If
        Else
            recContacto.Close
            recReferencia.Close
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub


Private Sub ProcessDetalleTMK2()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM TMK2", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            'Importo los campos del contact1 primero
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * FROM Contact1 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle("Notes").Value = recImportacion("observaciones").Value & ""
                recDetalle("Key2").Value = recImportacion("Estado").Value & ""
                recDetalle("Key5").Value = recImportacion("Comprador").Value & ""
            recDetalle.Update
            recDetalle.Close
            'Importo los campos del contact2 primero
            recDetalle.Open "SELECT * FROM Contact2 where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
                recDetalle("Ucono1").Value = recImportacion("Informacion").Value & ""
                recDetalle("UMedicorec").Value = recImportacion("MEDICO_RECETA").Value & ""
                recDetalle("UFechaVis2").Value = IIf(IsNull(recImportacion("Fecha_Visita").Value), Null, recImportacion("Fecha_Visita").Value)
            recDetalle.Update
            recDetalle.Close
            
            DoEvents
            'Inserto el primer detalle que es el del segundo llamado
            
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                'On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Segundo Llamado OC"
                recDetalle("u_Contact").Value = "Segundo Llamado OC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("OBTENCION_PRODUCTO").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("CANT_BOLSAS_MES_OS").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("CANT_BOLSAS_MES_USA").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("CONFORMIDAD_BOLSA").Value & "", 10) 'Campo4
                recDetalle("ext").Value = Left(recImportacion("PROBO_OTRA_MARCA").Value & "", 6) 'Campo5
                recDetalle("state").Value = Left(recImportacion("CUAL_OTRA_MARCA").Value & "", 20) 'Campo6
                recDetalle("Address1").Value = Left(recImportacion("MOTIVO_DECISION_MARCA").Value & "", 40) 'Campo7
                recDetalle("Address2").Value = Left(recImportacion("MOTIVO_DECISION_MARCA2").Value & "", 40) 'Campo8
                recDetalle("city").Value = Format(recImportacion("FECHA_LLAMADO").Value & "", "YYYYMMDD")
                
                recDetalle.Update
                
'                If Not IsNull(recImportacion("PRODUCT").Value) Then
'                    recDetalle.AddNew
'                    recDetalle("Accountno").Value = ReplaceQuote(strAccountno)
'                    recDetalle("recid").Value = GenerateRecID
'                    recDetalle("rectype").Value = "P"
'                    recDetalle("Contact").Value = "Producto Recomendado OC"
'                    recDetalle("u_Contact").Value = "Producto Recomendado OC"
'                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
'                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
'
'                    recDetalle.Update
'                End If
                
                'Producto Utilizado
                If Not IsNull(recImportacion("CODIGO_BOLSA").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = ""
                    recDetalle("u_Contsupref") = ""
                    recDetalle("ext").Value = Left(recImportacion("CODIGO_BOLSA").Value & "", 6) 'Campo5
                    recDetalle("state").Value = Left(recImportacion("MARCA_BOLSA").Value & "", 20) 'Campo6
                    recDetalle("Address1").Value = Left(recImportacion("MODELO_BOLSA").Value & " " & recImportacion("DESCR_BOLSA").Value, 40) 'Campo7
                    recDetalle.Update
                End If
            
                If Not IsNull(recImportacion("cual_otro_producto_cuidado").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Productos Utilizados OC"
                    recDetalle("u_Contact").Value = "Productos Utilizados OC"
                    recDetalle("Contsupref") = ""
                    recDetalle("u_Contsupref") = ""
                    recDetalle("ext").Value = Left(recImportacion("cual_otro_producto_cuidado").Value & "", 6) 'Campo5
                    recDetalle.Update
                End If
            
            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessDetalle1CC()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM basecc", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            DoEvents
            'Inserto el primer detalle
            Set recDetalle = New ADODB.Recordset
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
            recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
            recDetalle("recid").Value = GenerateRecID
            recDetalle("rectype").Value = "P"
            recDetalle("Contact").Value = "Patologia CC"
            recDetalle("u_Contact").Value = "Patologia CC"
            recDetalle("Contsupref") = recImportacion("Campo1").Value & ""
            recDetalle("u_Contsupref") = recImportacion("Campo1").Value & ""
            recDetalle("Title") = Left(recImportacion("Campo2").Value & "", 20)
            recDetalle.Update
            recDetalle.Close
            Set recDetalle = Nothing
        End If
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub ProcessDetalleCCLlamado()
Dim recImportacion As New ADODB.Recordset
Dim recDetalle As Recordset
Dim strAccountNo As String
Dim strCampo2 As String
Dim strCampo3 As String

    'On Error GoTo ERRORHANDLER:
    'Primero Abro el recordset de la importacion
    recImportacion.Open "SELECT * FROM BaseCC", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recImportacion.EOF
        strAccountNo = GetAccountno(Str(recImportacion("ID_Contact").Value))
        If strAccountNo <> "" Then
            lblInfo.Caption = Str(recImportacion("ID_Contact").Value)
            Set recDetalle = New ADODB.Recordset
            DoEvents
            'Inserto el primer detalle que es el del segundo llamado
            
            recDetalle.Open "SELECT * from Contsupp where Accountno='" & ReplaceQuote(strAccountNo) & "'", p_ADODBConnection, adOpenKeyset, adLockOptimistic, adCmdText
            recDetalle.AddNew
                'On Error GoTo retry
                recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                recDetalle("recid").Value = GenerateRecID
                recDetalle("rectype").Value = "P"
                recDetalle("Contact").Value = "Primer Llamado CC"
                recDetalle("u_Contact").Value = "Primer Llamado CC"
                recDetalle("Contsupref") = ""
                recDetalle("u_Contsupref") = ""
                
                recDetalle("Title").Value = Left(recImportacion("CATETERISMO_INTERMITENTE").Value & "", 20) 'Campo1
                recDetalle("linkacct").Value = Left(recImportacion("MARCA").Value & "", 20) 'Campo2
                recDetalle("country").Value = Left(recImportacion("CONFORMIDAD").Value & "", 20) 'Campo3
                recDetalle("zip").Value = Left(recImportacion("USUARIO_EASICATH").Value & "", 10) 'Campo4
                recDetalle("ext").Value = Left(recImportacion("ENTREGA_OS").Value & "", 6) 'Campo5
                recDetalle("state").Value = Left(recImportacion("RECONOCIMIENTO").Value & "", 20) 'Campo6
                'recDetalle("Address1").Value = Left(recImportacion("MOTIVO_DECISION_MARCA").Value & "", 40) 'Campo7
                'recDetalle("Address2").Value = Left(recImportacion("MOTIVO_DECISION_MARCA2").Value & "", 40) 'Campo8
                'recDetalle("city").Value = Format(recImportacion("FECHA_LLAMADO").Value & "", "YYYYMMDD")
                
                recDetalle.Update
                
                If Not IsNull(recImportacion("PRODUCT").Value) Then
                    recDetalle.AddNew
                    recDetalle("Accountno").Value = ReplaceQuote(strAccountNo)
                    recDetalle("recid").Value = GenerateRecID
                    recDetalle("rectype").Value = "P"
                    recDetalle("Contact").Value = "Producto Utilizado CC"
                    recDetalle("u_Contact").Value = "Producto Utilizado CC"
                    recDetalle("Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)
                    recDetalle("u_Contsupref") = Left(recImportacion("PRODUCT").Value & "", 20)

                    recDetalle.Update
                End If
                            
            recDetalle.Close
            Set recDetalle = Nothing
        End If
retry:
        recImportacion.MoveNext
    Loop
    

    Set p_ADODBConnection = Nothing
    Exit Sub
    
ERRORHANDLER:
    On Error Resume Next
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Public Sub CompleteCalendarActivity(CalRecID As String, RESCode As String, Notes As String)
Dim recCalendar As New ADODB.Recordset
Dim recNewHist As New ADODB.Recordset
    
    'Primero busco la actividad en el CAL
    'La borro y creo el conthist
    On Error GoTo ERRORHANDLER
    'p_ADODBConnection.BeginTrans
    recCalendar.Open "SELECT * from CAL WHERE RECID='" & CalRecID & "'", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If recCalendar.EOF Then
        'p_ADODBConnection.RollbackTrans
        Exit Sub
    Else
        recNewHist.Open "SELECT * from CONTHIST WHERE Accountno='" & recCalendar("Accountno").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        'inserto el registro nuevo
        recNewHist.AddNew
            'recNewHist("USERID").Value = "MASTER"
            recNewHist("USERID").Value = recCalendar("userid").Value
            recNewHist("ACCOUNTNO").Value = recCalendar("Accountno").Value
            recNewHist("SRECTYPE").Value = recCalendar("RECTYPE").Value & ""
            recNewHist("RECTYPE").Value = recCalendar("RECTYPE").Value & ""
            recNewHist("ONDATE").Value = recCalendar("ONDATE").Value & ""
            recNewHist("ONTIME").Value = recCalendar("ONTIME").Value & ""
            recNewHist("ACTVCODE").Value = recCalendar("ACTVCODE").Value & ""
            recNewHist("RESULTCODE").Value = RESCode
            recNewHist("STATUS").Value = recCalendar("STATUS").Value & ""
            recNewHist("DURATION").Value = recCalendar("DURATION").Value & ""
            recNewHist("UNITS").Value = ""
            recNewHist("REF").Value = recCalendar("REF").Value & ""
            'recNewHist("NOTES").Value = Notes
            'recNewHist("NOTES").Value = ConvertStringToBinary(Notes)
            recNewHist("NOTES").Value = recCalendar("NOTES").Value
            recNewHist("CREATEBY").Value = recCalendar("CREATEBY").Value & ""
            recNewHist("CREATEON").Value = recCalendar("CREATEON").Value & ""
            recNewHist("CREATEAT").Value = recCalendar("CREATEAT").Value & ""
            recNewHist("LASTUSER").Value = recCalendar("LASTUSER").Value & ""
            If Not IsNull(recCalendar("LASTDATE").Value) Then
                recNewHist("LASTDATE").Value = recCalendar("LASTDATE").Value & ""
                recNewHist("LASTTIME").Value = recCalendar("LASTTIME").Value & ""
            End If
            'recNewHist("EXT").Value = recOldHist("EXT").Value & ""
            recNewHist("recid").Value = GenerateRecID
            recNewHist("loprecid").Value = ""
            recNewHist("completedid").Value = ""
        recNewHist.Update
        recCalendar.Close
        p_ADODBConnection.Execute "DELETE FROM CAL where recid='" & CalRecID & "'"
        'p_ADODBConnection.CommitTrans
    End If
    
    
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    'p_ADODBConnection.RollbackTrans
    MsgBox "Error: " & Err.Description, vbCritical, App.EXEName

End Sub

Private Sub GMVocalcomm()
Dim recEntradas As New ADODB.Recordset
Dim recContact1 As ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recContSupp As ADODB.Recordset
Dim strNewAccountno As String
Dim strNewRecid As String
Dim strTemp As String

    On Error GoTo ERRORHANDLER
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
    p_ADODBConnection.Open
    
    recEntradas.Open "SELECT * FROM MIDWARE.DBO.GMVC where PasoGM is null", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recEntradas.EOF
        'Busco si existe el contact1
        Set recContact1 = New ADODB.Recordset
        recContact1.Open "SELECT * FROM Contact1 where phone1='" & recEntradas("CUIT").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        If Not recContact1.EOF Then
            'Existe el contacto
            WriteLog "GMVC: Existe CUIT " & recEntradas("CUIT").Value
        Else
            'No existe el contacto, lo creo
            strNewAccountno = GenerateAccountno(recEntradas("RazonSocial").Value & "")
            recContact1.AddNew
                recContact1("ACCOUNTNO").Value = strNewAccountno
                recContact1("RECID").Value = GenerateRecID
                recContact1("PHONE1").Value = recEntradas("CUIT").Value
                recContact1("contact").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("Address1").Value = Left(recEntradas("Domicilio").Value, 40)
                recContact1("City").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("Phone2").Value = Left(recEntradas("Telefono").Value, 25)
                recContact1("ZIP").Value = Left(recEntradas("CP").Value & "", 4)
                recContact1("Status").Value = "I0"
                recContact1("Source").Value = "NET BROKER"
                recContact1("OWNER").Value = ""
                recContact1("U_COMPANY").Value = ""
                recContact1("U_CONTACT").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("U_LASTNAME").Value = ""
                recContact1("u_CITY").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("U_STATE").Value = ""
                recContact1("u_COUNTRY").Value = ""
                recContact1("U_KEY1").Value = ""
                recContact1("U_KEY2").Value = ""
                recContact1("KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                recContact1("U_KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                strTemp = Trim(GetIni("Agentes", recEntradas("EjecutivoAsignado").Value & ""))
                If strTemp <> "" Then
                    recContact1("KEY4").Value = strTemp
                    recContact1("U_KEY4").Value = strTemp
                Else
                    recContact1("KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                    recContact1("U_KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                End If
                recContact1("U_KEY5").Value = ""
                recContact1("CREATEBY").Value = "VOC2GLMD"
                recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
            recContact1.Update
            'ahora agrego el email
            If Trim(recEntradas("mail").Value & "") <> "" Then
                AddEmailAddress strNewAccountno, Left(recEntradas("mail").Value & "", 35)
            End If
            Set recContact2 = New ADODB.Recordset
            recContact2.Open "SELECT * from contact2 where accountno='" & strNewAccountno & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            recContact2.AddNew
                recContact2("accountno").Value = strNewAccountno
                recContact2("recid").Value = GenerateRecID
                recContact2("ucapitas").Value = recEntradas("capitas").Value & ""
                ' si son mas de 50 capitas pongo el tipo de cliente en segmento
                If Val("0" & recEntradas("capitas").Value) >= 50 Then
                    recContact2("USEGMENTO").Value = "CLIENTE CORPORATIVO"
                End If
                recContact2("umasasrial").Value = recEntradas("masasalarial").Value & ""
                'art actual
                recContact2("userdef04").Value = Left(recEntradas("ART").Value & "", 12)
                recContact2("uav").Value = recEntradas("AV").Value & ""
                recContact2("uaf").Value = recEntradas("AF").Value & ""
                strTemp = Trim(GetIni("Agentes", recEntradas("agente").Value & ""))
                If strTemp <> "" Then
                    recContact2("utlmk").Value = strTemp
                Else
                    recContact2("utlmk").Value = recEntradas("agente").Value & ""
                End If
            recContact2.Update
            'Ahora cargo las cosas de contact2
'            Agente
'            EjecutivoAsignado
'            Campana
'            ART
'            CIIU
'            MasaSalarial
'            Capitas
'            AV
'            AF

            'Ahora marco como actualizado
            recEntradas("PasoGM").Value = Now
            recEntradas.Update
        End If
        recEntradas.MoveNext
    Loop
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    WriteLog "GMVC: " & GetADODBErrorMessage(p_ADODBConnection.Errors)
End Sub

Private Sub GMVocalcommPreVenta()
Dim recEntradas As New ADODB.Recordset
Dim recContact1 As ADODB.Recordset
Dim recContact2 As ADODB.Recordset
Dim recContSupp As ADODB.Recordset
Dim strNewAccountno As String
Dim strNewRecid As String
Dim strTemp As String
Dim strUserIDGM As String
Dim recPending As ADODB.Recordset
Dim recNewPending As ADODB.Recordset
Dim InsertarPendiente As Boolean

    On Error GoTo ERRORHANDLER
    Set p_ADODBConnection = New ADODB.Connection
    p_ADODBConnection.ConnectionString = GetIni("DataSourceName", "CS")
    p_ADODBConnection.Open
    
    recEntradas.CursorLocation = adUseClient
    recEntradas.Open "SELECT * FROM MIDWARE.DBO.GMVC where EstadoEntrevista in ('G','B','R') order by EstadoEntrevista DESC ", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not recEntradas.EOF
        'Busco si existe el contact1
        Set recContact1 = New ADODB.Recordset
        Me.lblInfo.Caption = "Proceso: " & recEntradas("CUIT").Value
        DoEvents
        recContact1.Open "SELECT * FROM Contact1 where phone1='" & recEntradas("CUIT").Value & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
        If recContact1.EOF And recEntradas("EstadoEntrevista").Value = "G" Then
            'No existe el contacto, lo creo
            'p_ADODBConnection.BeginTrans
            strNewAccountno = GenerateAccountno(recEntradas("RazonSocial").Value & "")
            recContact1.AddNew
                recContact1("ACCOUNTNO").Value = strNewAccountno
                recContact1("RECID").Value = GenerateRecID
                recContact1("PHONE1").Value = recEntradas("CUIT").Value
                recContact1("contact").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("Address1").Value = Left(recEntradas("Domicilio").Value, 40)
                recContact1("City").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("Phone2").Value = Left(recEntradas("Telefono").Value, 25)
                recContact1("ZIP").Value = Left(recEntradas("CP").Value & "", 4)
                recContact1("Status").Value = "I0"
                recContact1("Source").Value = "NET BROKER"
                recContact1("OWNER").Value = ""
                recContact1("U_COMPANY").Value = ""
                recContact1("U_CONTACT").Value = Left(recEntradas("RazonSocial").Value, 40)
                recContact1("U_LASTNAME").Value = ""
                recContact1("u_CITY").Value = Left(recEntradas("Localidad").Value, 30)
                recContact1("U_STATE").Value = ""
                recContact1("u_COUNTRY").Value = ""
                recContact1("U_KEY1").Value = ""
                recContact1("U_KEY2").Value = ""
                recContact1("KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                recContact1("U_KEY3").Value = Left(recEntradas("CIIU").Value & "", 20)
                strTemp = Trim(GetIni("Agentes", recEntradas("EjecutivoAsignado").Value & ""))
                If strTemp <> "" Then
                    recContact1("KEY4").Value = strTemp
                    recContact1("U_KEY4").Value = strTemp
                Else
                    recContact1("KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                    recContact1("U_KEY4").Value = recEntradas("EjecutivoAsignado").Value & ""
                End If
                recContact1("U_KEY5").Value = ""
                recContact1("CREATEBY").Value = "VOC2GLMD"
                recContact1("CREATEON").Value = Format(Now, "yyyy/mm/dd")
            recContact1.Update
            'ahora agrego el email
            If Trim(recEntradas("mail").Value & "") <> "" Then
                AddEmailAddress strNewAccountno, Left(recEntradas("mail").Value & "", 35)
            End If
            Set recContact2 = New ADODB.Recordset
            recContact2.Open "SELECT * from contact2 where accountno='" & strNewAccountno & "'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
            recContact2.AddNew
                recContact2("accountno").Value = strNewAccountno
                recContact2("recid").Value = GenerateRecID
                recContact2("ucapitas").Value = recEntradas("capitas").Value & ""
                ' si son mas de 50 capitas pongo el tipo de cliente en segmento
                If Val("0" & recEntradas("capitas").Value) >= 50 Then
                    recContact2("USEGMENTO").Value = "CLIENTE CORPORATIVO"
                End If
                recContact2("umasasrial").Value = recEntradas("masasalarial").Value & ""
                'art actual
                recContact2("userdef04").Value = Left(recEntradas("ART").Value & "", 12)
                recContact2("uav").Value = recEntradas("AV").Value & ""
                recContact2("uaf").Value = recEntradas("AF").Value & ""
                strTemp = Trim(GetIni("Agentes", recEntradas("agente").Value & ""))
                If strTemp <> "" Then
                    recContact2("utlmk").Value = strTemp
                Else
                    recContact2("utlmk").Value = recEntradas("agente").Value & ""
                End If
                'los campos nuevos
                recContact2("UEPERSONA").Value = recEntradas("PersonaEntrevista").Value & ""
                recContact2("UEFECHA").Value = recEntradas("FechaEntrevista").Value
                recContact2("UEDOMICILI").Value = recEntradas("DireccionEntrevista").Value & ""
                recContact2("UECP").Value = recEntradas("CP").Value & ""
            
            recContact2.Update
            'p_ADODBConnection.CommitTrans
        End If
        
        'Genero el pendiente
        'ya sea lo encontre o no
        'Busco si tiene un pending
        If Not recContact1.EOF Then
            Set recPending = New ADODB.Recordset
            recPending.Open "SELECT * FROM CAL WHERE AccountNo='" & recContact1("ACCOUNTNO").Value & "' AND (REF ='ENTREVISTA' OR REF ='COTIZACION POR MAIL')", p_ADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not recPending.EOF Then
                'Solo si estoy reprogramando hago algo
                If recEntradas("EstadoEntrevista").Value = "B" Then
                    p_ADODBConnection.Execute "DELETE FROM CAL WHERE recID=" & FormatToSQL(recPending("recID").Value, gsdtString)
                    InsertarPendiente = False
                Else
                    If recEntradas("EstadoEntrevista").Value = "R" Then
                        p_ADODBConnection.Execute "DELETE FROM CAL WHERE recID=" & FormatToSQL(recPending("recID").Value, gsdtString)
                        InsertarPendiente = True
                    End If
                End If
            Else
                If recEntradas("EstadoEntrevista").Value = "G" Then
                    InsertarPendiente = True
                End If
            End If
            If InsertarPendiente Then
                Set recNewPending = New ADODB.Recordset
                recNewPending.Open "SELECT * FROM CAL WHERE Accountno='aa'", p_ADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
                recNewPending.AddNew
                recNewPending("RECID").Value = GenerateRecID
                recNewPending("Accountno").Value = recContact1("Accountno").Value & ""
                recNewPending("Company").Value = recContact1("Contact").Value & ""
                strUserIDGM = Trim(GetIni("UsersGM", recEntradas("EjecutivoAsignado").Value & ""))
                recNewPending("USERID").Value = Left(strUserIDGM, 8)
                'recNewPending("ONDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                'recNewPending("ENDDATE").Value = Format(DateAdd("d", 2, Now), "DD/MM/YYYY")
                recNewPending("ONDATE").Value = recEntradas("FechaEntrevista").Value
                recNewPending("ENDDATE").Value = recEntradas("FechaEntrevista").Value
                
                If recEntradas("TipoContacto").Value & "" = "M" Then
                    recNewPending("ACTVCODE").Value = "ML"
                    recNewPending("REF").Value = "ENVIAR COT POR MAIL"
                    recNewPending("RECTYPE").Value = "D" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                Else
                    recNewPending("ACTVCODE").Value = "ENT"
                    recNewPending("REF").Value = "ENTREVISTA"
                    recNewPending("Notes").Value = ConvertStringToBinary("Direccion de entrevista" & recEntradas("DireccionEntrevista").Value & "  Persona contacto: " & recEntradas("PersonaEntrevista").Value)
                    recNewPending("RECTYPE").Value = "A" 'A: Appointment, C: Call Back, T: Next Action, D: To-Do M: Message, S: Forecasted Sale, O: Other, E: Event
                End If
                
                recNewPending("CREATEBY").Value = Left(strUserIDGM, 8)
                recNewPending("LASTUSER").Value = Left(strUserIDGM, 8)
                recNewPending("CREATEON").Value = Format(Now, "DD/MM/YYYY")
                'rellenos
                recNewPending("ONTIME").Value = Format(recEntradas("FechaEntrevista").Value, "HH:NN")
                recNewPending("ALARMFLAG").Value = "N"
                recNewPending("ALARMTIME").Value = ""
                
                recNewPending("LINKRECID").Value = ""
                recNewPending("LOPRECID").Value = ""
                
                recNewPending.Update
                recNewPending.Close
                recPending.Close
            End If
        End If
        'Ahora marco como actualizado
        recEntradas("PasoGM").Value = Now
        recEntradas("EstadoEntrevista").Value = "T"
        recEntradas.Update
    
        
        recEntradas.MoveNext
    Loop
    
    Exit Sub
ERRORHANDLER:

    On Error Resume Next
    p_ADODBConnection.RollbackTrans
    WriteLog "GMVC: " & GetADODBErrorMessage(p_ADODBConnection.Errors)
End Sub

