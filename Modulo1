'Application.Wait (Now + TimeValue("00:00:05"))
'MsgBox "Has esperado 5 segundos"

Public miConexion As New ADODB.Connection

Public Rs As New ADODB.Recordset
Public Rv As New ADODB.Recordset
Public Rc As New ADODB.Recordset
Public Rx As New ADODB.Recordset
Public Rg As New ADODB.Recordset
Public Re As New ADODB.Recordset
Public Rf As New ADODB.Recordset
Public Rq As New ADODB.Recordset
Public Pr As New ADODB.Recordset

Public Rn As New ADODB.Recordset
Public Ry As New ADODB.Recordset
Public Rp As New ADODB.Recordset
Public Rk As New ADODB.Recordset


Sub BD_Principal()
Set miConexion = New ADODB.Connection
    With miConexion
        .Provider = "Microsoft.ACE.OLEDB.15.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\BD_Ventas\BD_Principal.accdb"
        .Open
    End With
End Sub

Sub BD_Productividad()
Set miConexion = New ADODB.Connection
    With miConexion
        .Provider = "Microsoft.ACE.OLEDB.15.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\BD_Ventas\BD_Productividad.accdb"
        .Open
    End With
End Sub

Sub BD_Seguridad()
Set miConexion = New ADODB.Connection
    With miConexion
       .Provider = "Microsoft.ACE.OLEDB.15.0"
       .ConnectionString = "data source=" & ThisWorkbook.Path & "\BD_Ventas\BD_Seguridad_&_Control.accdb"
       .Open
    End With
End Sub

Sub BD_Gestionados()
Set miConexion = New ADODB.Connection
    With miConexion
       .Provider = "Microsoft.ACE.OLEDB.15.0"
       .ConnectionString = "data source=" & ThisWorkbook.Path & "\BD_Ventas\BD_Gestionados.accdb"
       .Open
    End With
End Sub

Sub BD_Tipificacion()
Set miConexion = New ADODB.Connection
    With miConexion
       .Provider = "Microsoft.ACE.OLEDB." & Application.Version
       .ConnectionString = "data source=" & ThisWorkbook.Path & "\BD_Ventas\BD_Tipificacion.accdb"
       .Open
    End With
End Sub


Sub Abrir_login_BO()

ActiveWindow.DisplayWorkbookTabs = False 'Oculta las fichas de las hohas
ActiveWindow.DisplayHeadings = False 'Oculta títulos
Application.DisplayFormulaBar = False 'Oculta la barra de formulas
ActiveWindow.DisplayGridlines = False 'Oculta las lineas de la cuadricula
Application.DisplayStatusBar = False 'Oculta la barra de estado
Application.DisplayFullScreen = True 'Ves pantalla completa

 Application.Visible = True
 Load UserForm1
 UserForm1.Show
End Sub

Sub Salir()
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub

Sub Abrir_login_ADMIN()

ActiveWindow.DisplayWorkbookTabs = False 'Oculta las fichas de las hohas
ActiveWindow.DisplayHeadings = False 'Oculta títulos
Application.DisplayFormulaBar = False 'Oculta la barra de formulas
ActiveWindow.DisplayGridlines = False 'Oculta las lineas de la cuadricula
Application.DisplayStatusBar = False 'Oculta la barra de estado
Application.DisplayFullScreen = True 'Ves pantalla completa

 Application.Visible = True
 Load UserForm2
 UserForm2.Show
End Sub
