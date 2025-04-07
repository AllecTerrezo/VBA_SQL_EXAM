Option Explicit
Dim cnn As ADODB.Connection


Public Sub OpenConnection()
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" & ThisWorkbook.FullName & ";" & _
    "Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;"";"
    
    cnn.Open
End Sub
Sub Tarefa()
    OpenConnection
    Dim Sql As String
    
    
    Sql = "select T1.ID_CONTA, replace(T3.NOME, 'NULL', 'nao encontrado') As NOME, replace(T5.CD_FIRMA, 1, NULL) As CD_FIRMA, replace(T4.EMAIL, 'NULL', 'nao encontrado') As EMAIL from " & _
    "Tabela4 T4, Tabela5 T5, Tabela1 T1, Tabela2 T2, Tabela3 T3 " & _
    "where (T2.ID_PESSOA = T4.ID_PESSOA) and " & _
    "(T1.CONTA_PAI = T2.CONTA_PAI) and " & _
    "(T1.ID_CONTA = T5.ID_CONTA) and " & _
    "(T1.ID_CONTA = T3.ID_CONTA)"



    Dim rs As ADODB.Recordset
    Set rs = GetData(Sql)


    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planilha2")
    
    ws.Cells.Delete

    ws.Range("A2").CopyFromRecordset rs
    
    Dim iField As Long
    For iField = 0 To rs.Fields.Count - 1
        ws.Range("A1").Offset(, iField) = rs.Fields(iField).Name
    Next iField
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    


    rs.Close
    cnn.Close
End Sub

Function GetData(Sql As String) As ADODB.Recordset
    
    Dim TableName1 As String
    Dim TableName2 As String
    Dim TableName3 As String
    Dim TableName4 As String
    Dim TableName5 As String
    
    With ThisWorkbook.Worksheets("Planilha1")
        TableName5 = "[" & .Name & "$" & .ListObjects("Table_5").Range.Address(0, 0) & "]"
        TableName4 = "[" & .Name & "$" & .ListObjects("Table_4").Range.Address(0, 0) & "]"
        TableName3 = "[" & .Name & "$" & .ListObjects("Table_3").Range.Address(0, 0) & "]"
        TableName2 = "[" & .Name & "$" & .ListObjects("Table_2").Range.Address(0, 0) & "]"
        TableName1 = "[" & .Name & "$" & .ListObjects("Table_1").Range.Address(0, 0) & "]"
    End With
    
    
    Sql = Replace(Sql, "Tabela5", TableName5)
    Sql = Replace(Sql, "Tabela4", TableName4)
    Sql = Replace(Sql, "Tabela3", TableName3)
    Sql = Replace(Sql, "Tabela2", TableName2)
    Sql = Replace(Sql, "Tabela1", TableName1)
    
    Dim rs As ADODB.Recordset
    
    Set rs = cnn.Execute(Sql)

    Set GetData = rs

End Function

Sub enviar_email()

Set objeto_outlook = CreateObject("Outlook.Application")

Set Email = objeto_outlook.createitem(0)

'Email.display


Email.To = "xxxxx@email.com"

Email.Subject = "Subject"

Email.Body = "Prezada(o)," & vbNewLine & vbNewLine _
& "Segue select utilizado, bem como, arquivo da atividade contendo códigos em anexo." & vbNewLine & vbNewLine _
& "select T1.ID_CONTA, replace(T3.NOME, 'NULL', 'nao encontrado') As NOME, replace(T5.CD_FIRMA, 1, NULL) As CD_FIRMA, replace(T4.EMAIL, 'NULL', 'nao encontrado') As EMAIL from " & _
    "Tabela4 T4, Tabela5 T5, Tabela1 T1, Tabela2 T2, Tabela3 T3 " & _
    "where (T2.ID_PESSOA = T4.ID_PESSOA) and " & _
    "(T1.CONTA_PAI = T2.CONTA_PAI) and " & _
    "(T1.ID_CONTA = T5.ID_CONTA) and " & _
    "(T1.ID_CONTA = T3.ID_CONTA)" & vbNewLine _
    & "Atenciosamente," & vbNewLine _
    & "Allec Terrezo."
    
    
    
Email.Attachments.Add (ThisWorkbook.Path & "\prova_vba_e_sql_2020_Allec1.xltm")

Email.Send

End Sub



