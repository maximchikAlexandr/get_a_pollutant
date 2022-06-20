Attribute VB_Name = "SQL_get_pollutant"
Function get_pollutant(Code As Variant, Optional Parametr& = 1)
Attribute get_pollutant.VB_Description = "Returns the characteristic of a pollutant by its code"
Attribute get_pollutant.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim connectionString As String
    Dim command As Object
'_________________________________________________________________________________________________________________________________________________________________
        
    Set objectConnection = CreateObject("ADODB.Connection")
    Set rst = CreateObject("ADODB.Recordset")
    Set command = CreateObject("ADODB.Command")         '//create a command to send to the DB
    Set comParam = CreateObject("ADODB.Parameter")

'_________________________________________________________________________________________________________________________________________________________________
        
    connectionString = "Provider=sqloledb;Data Source=HOME-PC\SQLEXPRESS;Initial Catalog=Spravochnik;Trusted_Connection=yes"
    objectConnection.connectionString = connectionString    ' init connection
    objectConnection.Open 'open connection
    command.ActiveConnection = objectConnection 'init command
    command.CommandType = 4  'type of command - "run stored procedure"
    command.CommandText = "proc_GetSubstances" 'name of stored procedure
    command.NamedParameters = True 'named parametrs
'_________________________________________________________________________________________________________________________________________________________________
     
    Set comParam = command.CreateParameter("@ од", 202, 1, Len(Code), Code) 'set value to parametr
    command.Parameters.Append comParam
    Set rst = command.Execute 'run stored procedure
    get_pollutant = rst(Parametr + 1).Value
'_________________________________________________________________________________________________________________________________________________________________
    Set rst = Nothing
    Set command = Nothing
    objectConnection.Close
    Set objectConnection = Nothing
End Function


Sub Descriptions()
    Dim FName As String
    Dim FOpis As String '
    Dim FCat As String
    Dim Arg1 As String, Arg2 As String '
    
    FName = "get_pollutant"
    FOpis = "Returns the characteristic of a pollutant by its code"
    FCat = "Ecology"
    Arg1 = "Code of a pollutant"
    Arg2 = "Characteristic: " & Chr(10) & _
    "1 Ц name pollutant; 2 Ц PDKmr; 3 Ц PDKss; 4 Ц PDKsg; 5 Ц OBUV" & Chr(10) & _
    "6 Ц class of a danger.; 7 Ц agregat; 8 Ц PDV?; 9 Ц  VOC?; 10 Ц ch. formula"
   
    Application.MacroOptions Macro:=FName, Description:=FOpis, Category:=FCat, ArgumentDescriptions:=Array(Arg1, Arg2)
End Sub

