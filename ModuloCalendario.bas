Option Explicit
Option Private Module
Private SenalCambioMes As Long

Public Sub RecibeLaFecha(Dia As Long, Mes As Long, Ano As Long)

    Dim FechaRecibida As Date
    FechaRecibida = VBA.DateSerial((VBA.CInt(Ano)), (VBA.CInt(Mes)), (VBA.CInt(Dia)))
    
    If UserForm1.Label425.Caption = "Fecha inicio" Then
    UserForm1.TextBox88.Text = FechaRecibida
    UserForm1.TextBox88 = Format(UserForm1.TextBox88, "mm/dd/yyyy")
    End If
    
    If UserForm1.Label425.Caption = "Fecha fin" Then
    UserForm1.TextBox89.Text = FechaRecibida
    UserForm1.TextBox89 = Format(UserForm1.TextBox89, "mm/dd/yyyy")
    
    UserForm1.Label428.Caption = FechaRecibida + 1
    UserForm1.Label428 = Format(UserForm1.Label428, "mm/dd/yyyy")
    End If
    
    '__________________________________________________________________________________________________
    
    If UserForm2.Label350.Caption = "Fecha inicio" Then
    UserForm2.TextBox80.Text = FechaRecibida
    UserForm2.TextBox80 = Format(UserForm2.TextBox80, "mm/dd/yyyy")
    End If
    
    If UserForm2.Label350.Caption = "Fecha fin" Then
    UserForm2.TextBox81.Text = FechaRecibida
    UserForm2.TextBox81 = Format(UserForm2.TextBox81, "mm/dd/yyyy")
    
    UserForm2.Label411.Caption = FechaRecibida + 1
    UserForm2.Label411 = Format(UserForm2.Label411, "mm/dd/yyyy")
    
    End If
    
    
     '___________________ Exportar ventas Exitosos ____________________
    
    If UserForm2.Label350.Caption = "Fecha inicio EX" Then
    UserForm2.TextBox82.Text = FechaRecibida
    UserForm2.TextBox82 = Format(UserForm2.TextBox82, "mm/dd/yyyy")
    End If
    
    If UserForm2.Label350.Caption = "Fecha fin EX" Then
    UserForm2.TextBox83.Text = FechaRecibida
    UserForm2.TextBox83 = Format(UserForm2.TextBox83, "mm/dd/yyyy")
    UserForm2.Label411.Caption = FechaRecibida + 1
    UserForm2.Label411 = Format(UserForm2.Label411, "mm/dd/yyyy")
    End If
    
    
    '___________________ Exportar ventas Fallidos _____________________
    
    If UserForm2.Label350.Caption = "Fecha inicio FL" Then
    UserForm2.TextBox84.Text = FechaRecibida
    UserForm2.TextBox84 = Format(UserForm2.TextBox84, "mm/dd/yyyy")
    End If
    
    

    If UserForm2.Label350.Caption = "Fecha fin FL" Then
    UserForm2.TextBox85.Text = FechaRecibida
    UserForm2.TextBox85 = Format(UserForm2.TextBox85, "mm/dd/yyyy")
    UserForm2.Label411.Caption = FechaRecibida + 1
    UserForm2.Label411 = Format(UserForm2.Label411, "mm/dd/yyyy")
    End If
    
     
    '_______________ Ver mis ventas gestionadas _______________
    
    If UserForm1.Label425.Caption = "Fecha inicio X" Then
    
    UserForm1.TextBox96.Text = FechaRecibida
    UserForm1.TextBox96 = Format(UserForm1.TextBox96, "mm/dd/yyyy")
    
    End If
    
    If UserForm1.Label425.Caption = "Fecha fin X" Then
    
    UserForm1.TextBox97.Text = FechaRecibida
    UserForm1.TextBox97 = Format(UserForm1.TextBox97, "mm/dd/yyyy")
    UserForm1.Label428.Caption = FechaRecibida + 1
    UserForm1.Label428 = Format(UserForm1.Label428, "mm/dd/yyyy")
    
    End If
    
    
    
    
    
    
End Sub

Public Sub InicializaFormularioCalendario()
    SenalCambioMes = 1
    
    With frmCalendario.cboMes
        .AddItem 1
        .List(0, 1) = "Enero"
        .AddItem 2
        .List(1, 1) = "Febrero"
        .AddItem 3
        .List(2, 1) = "Marzo"
        .AddItem 4
        .List(3, 1) = "Abril"
        .AddItem 5
        .List(4, 1) = "Mayo"
        .AddItem 6
        .List(5, 1) = "Junio"
        .AddItem 7
        .List(6, 1) = "Julio"
        .AddItem 8
        .List(7, 1) = "Agosto"
        .AddItem 9
        .List(8, 1) = "Septiembre"
        .AddItem 10
        .List(9, 1) = "Octubre"
        .AddItem 11
        .List(10, 1) = "Noviembre"
        .AddItem 12
        .List(11, 1) = "Diciembre"
    End With
    
    frmCalendario.cboMes.ListIndex = VBA.Month(VBA.Date) - 1
    
    frmCalendario.spbAño.Value = VBA.Year(VBA.Date)
    
    frmCalendario.lblAno.Caption = VBA.Year(VBA.Date)
    
    Dim Ano As Long, Mes As Long
    Ano = VBA.Year(VBA.Date)
    Mes = VBA.Month(VBA.Date)
    Call ModuloCalendario.CargarLosDias(Ano, Mes)
    
    frmCalendario.lblHoy.Caption = VBA.Date
End Sub

Public Sub CargarLosDias(Ano As Long, Mes As Long)
    Dim FechaDelPrimerDia As Date
    Dim FechaDelUltimoDia As Date
    Dim DiaSemanaPrimerDia As Long
    Dim VariableControl As Control
    Dim Contador As Long
    
    FechaDelPrimerDia = VBA.DateSerial(Ano, Mes, 1)
    FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, Mes, 1), 0)
    DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
    Contador = 1
    
    For Each VariableControl In frmCalendario.mrcDias.Controls
        VariableControl.Caption = "-"
        If VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia) Then
            VariableControl.Caption = Contador
            Contador = Contador + 1
        End If
    Next VariableControl
End Sub

Public Sub CambioDeMes()
    If SenalCambioMes > 1 Then
        Dim MesEnElCombo As Long, AnoEnElLabel As Long
        
        If Not (IsNull(frmCalendario.cboMes.Value)) And Not (IsNull(frmCalendario.lblAno.Caption)) Then
            MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
            AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
            Call ModuloCalendario.DesmarcarDias
            Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
        End If
    End If
    SenalCambioMes = SenalCambioMes + 1
End Sub

Public Sub CambioDeAno()
    Dim MesEnElCombo As Long, AnoEnElLabel As Long
    
    frmCalendario.lblAno.Caption = frmCalendario.spbAño.Value
    
    MesEnElCombo = VBA.CLng(frmCalendario.cboMes.Value)
    AnoEnElLabel = VBA.CLng(frmCalendario.lblAno.Caption)
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(AnoEnElLabel, MesEnElCombo)
    
End Sub

Public Sub UnClickEnHoyEs()
    Dim Mes As Long, Ano As Long
    Dim FechaActual As Date
    
    FechaActual = VBA.CDate(frmCalendario.lblHoy.Caption)
    Mes = VBA.CLng(VBA.Month(FechaActual))
    Ano = VBA.CLng(VBA.Year(FechaActual))
    
    frmCalendario.lblAno.Caption = Ano
    frmCalendario.cboMes.ListIndex = Mes - 1
    frmCalendario.spbAño.Value = Ano
    frmCalendario.spbAño.SetFocus
    
    Call ModuloCalendario.DesmarcarDias
    Call ModuloCalendario.CargarLosDias(Ano, Mes)
    
End Sub

Sub SalirConEscape()
    Unload frmCalendario
End Sub

Sub MarcarDia(ControlDeEtiqueta As Control)
    Call ModuloCalendario.DesmarcarDias
    ControlDeEtiqueta.Font.Bold = True
    ControlDeEtiqueta.ForeColor = VBA.RGB(255, 0, 0)
End Sub

Sub DesmarcarDias()
    Dim ControlEtiqueta As Control
    
    For Each ControlEtiqueta In frmCalendario.mrcDias.Controls
        ControlEtiqueta.Font.Bold = False
        ControlEtiqueta.ForeColor = VBA.RGB(0, 0, 0)
    Next ControlEtiqueta
End Sub

