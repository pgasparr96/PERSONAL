Public PATHMWSCHEDULE, PATHLOG, Cdir, GSEP, run ''''' Sub(?)
Public ArqINI, LogPath, IniModel, user_id, password, datasource, database ''''' Sub(?)
Dim param
Dim r, ArgItem, Args

run = 0

    Set Args = wscript.arguments ''''' Atribuindo variáveis 

    If Args.count > 0 Then 
              
        For Each ArgItem in Args
          param = param & ArgItem & " "
          'wscript.echo param
        Next

        param = Trim(param)
        param = replace(param,"'", """")
    Else 

       param = ""

    End if
    
    If param <> ""  Then

       wscript.echo Date & " " & Time
       'wscript.stderr.write Date & " " & Time & vbCrLf

       wscript.echo param
       'wscript.stderr.write param & vbCrLf

       Execute param
   'Wscript.echo run
    ElseIf param = "" Then 
    	 
       Wscript.echo "Falta argumentos"
       'wscript.stderr.write "Falta argumentos" & vbCrLf

    End If
    wscript.echo "Pronto."
    'wscript.stderr.write "Pronto." & vbCrlf 

Sub carga_credit()
Dim conn, conn_str, exec, cmd, qry_str, rs, res
Dim sgbd, owner, mw_schema, schedule_name

     owner="MW"
     sgbd="SQLSERVER"
     mw_schema="CRED_ADIC"
     schedule_name="CARGA_CRED_ADIC"

     conn_str="Provider=SQLOLEDB.1;Data Source=" + datasource + ";User ID=" + user_id + ";Extended Properties=;pwd="+ password + ";DataBase=" + database
     cmd=" " + schedule_name  + GSEP + conn_str + GSEP + database + GSEP + sgbd + GSEP + owner + GSEP + mw_schema 

'''' ATENÇÃO!!! ESTE TRECHO FOI COMENTADO PORQUE O INÍCIO DESTE PROCESSO NÃO ESTÁ MAIS VINCULADO AO FIM CONTÁBIL
''''            A CARGA AGORA É DIÁRIA -- Referência: OS369/2008.
''''     qry_str="select modelo from (select cast(left(ano,4)+right('0' + cast(cast(mes as numeric) as varchar),2) as numeric) anomes, modelo from custo..eex_ano_mes_carga where modelo = 'ORCAMENTO') a, (select cast(left(ano,4)+right('0' + cast(cast(mes as numeric) as varchar),2) as numeric) anomes from custo..eex_ano_mes_carga where modelo like '%FIM%CONTABIL%') b where a.ANOMES < b.ANOMES"
''''     Set conn = CreateObject("ADODB.Connection")
''''     conn.open conn_str 
''''     Set rs = conn.Execute(qry_str)
''''     qry_str="update a set a.ANO=b.ano, a.MES=b.mes from custo..eex_ano_mes_carga a, (select ano, mes, cast(left(ano,4)+right('0' + cast(cast(mes as numeric) as varchar),2) as numeric) anomes from custo..eex_ano_mes_carga where modelo like '%FIM%CONTABIL%') b where modelo = 'ORCAMENTO' and cast(left(a.ano,4)+right('0' + cast(cast(a.mes as numeric) as varchar),2) as numeric) < b.ANOMES"
''''     If not(rs.EOF) Then Set rs = conn.Execute(qry_str) : COMMANDSCH cmd : run=0
''''     If run=1 Then COMMANDSCH cmd 
''''     conn.Close
''''     Set rs = nothing     
     COMMANDSCH cmd : run=0' Esta linha foi incluída em substituição ao código acima.
End Sub

Sub COMMANDSCH(DBparam)
Dim WshShell_1
Dim cmd, lncmd

    Wscript.echo "Iniciando o proceso."
    lncmd = PATHMWSCHEDULE 
    cmd = lncmd & DBparam ': wscript.echo cmd
    Set WshShell_1 = CreateObject("WScript.Shell")
    Set oExec_1 = WshShell_1.Exec(cmd)

   'Faça loop enquanto o processo estiver ativo.
    Do While oExec_1.Status = 0 : Loop

End Sub

'''''/Chamada da função 'IniConnect'
'''''É passado como parâmetros as variáveis 'CARGA_CREDIT' e 'D:\SIG\scripts\MWScript\Connect.ini'.
'''''Respectivamente: 'model' e 'AbsPath'."/

Sub IniConnect(model, AbsPath)
Dim  fso '''''Dim: Declara e aloca espaço de armazenamento para uma ou até mais variáveis'''''
 Set fso = CreateObject("Scripting.FileSystemObject") '''''CreateObject: cria um objeto de um tipo especificado.'''''
 ArqINI = AbsPath
 If fso.FileExists(ArqINI) <> true Then '''''Log de erro(?).'''''
 	 wscript.echo "O arquivo " & ArqINI & " de configuração não foi encontrado."  & Chr(13) & Chr(10) & "Não é possível prosseguir." 
 	 wscript.stderr.write "O arquivo " & ArqINI & " de configuração não foi encontrado."  & Chr(13) & Chr(10) & "Não é possível prosseguir." & vbCrLf
 	 Exit Sub
 Else
   Execute Le_Ini(model) '''''Execução de outra função.'''''
   'dercrypt
   Execute model
 End If
End Sub

'''''Sub()
'''''Declare o nome, argumentos e código que formam o corpo de um procedimento Sub.

'''''Sintaxe
'''''      [ Público [Padrão] | Privado] Sub Nome [( arglist )]
'''''        [afirmações]
'''''         [ Sub Sair ]
'''''         [afirmações]
'''''      End Sub

'''''Público Indica que o procedimento Sub está acessível a todos os outros procedimentos em todos os scripts.
'''''Padrão Usado apenas com a palavra-chave Public em um bloco de Classes para indicar que o procedimento Sub é o método padrão para a classe. Um erro ocorre se mais de um procedimento padrão for especificado em uma classe.
'''''Privado Indica que o procedimento Sub está acessível apenas a outros procedimentos no script em que é declarado.
'''''nome Nome do Sub; segue convenções de nomenclatura de variáveis ​​padrão.
'''''arglist Lista de variáveis ​​que representam argumentos que são passados ​​para o procedimento Sub quando ele é chamado. Vírgulas separam várias variáveis.
'''''declarações Qualquer grupo de instruções a serem executadas dentro do corpo do procedimento Sub.

Sub dercrypt()
Dim dercrypt
 Set decrypt = CreateObject("Decrypt.cDecrypt")
 decrypt.decrypt = password : password = decrypt.decrypt
End Sub

'''' Escopo da função 
Function Le_Ini(model)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8, vbBinaryCompare = 0, vbTextCompare = 1
   Dim  fso, FileRead, LineText, flRcvr, sPosI, sPosF, flRcvrStr, StrConnect
On Error Resume Next
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set FileRead = fso.OpenTextFile(ArqINI, ForReading, True)
     fRecover = 0
     Do While not(FileRead.AtEndOfStream)
       LineText = FileRead.ReadLine & Chr(13)
       'wscript.echo LineText
       sPosI = InStr(LineText,"["): sPosF = InStr(LineText,"]")
       If sPosI > 0 and sPosF > 0 Then
       	flRcvrStr = Mid(LineText,sPosI + 1, sPosF - sPosI -1)
        If UCase(flRcvrStr) = UCase(model) Then
       	 flRcvr = 1
       	 StrConnect = ""
       	 LineText = ""
        ElseIf  flRcvrStr = "END" Then
       	 flRcvr = 0      	 
        End If 
       End If
       If flRcvr = 1 Then
        StrConnect = StrConnect & LineText
        'wscript.echo  LineText
       End If
     Loop
     FileRead.Close
     'Execute StrConnect 
     Le_Ini = StrConnect 
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''' <OpenTextFile> Line 133

'''''Abre um determinado arquivo e retorna um objeto TextStream que pode ser usado para ler, gravar ou anexar ao arquivo.

'''''Sintaxe
'''''objeto. OpenTextFile (nome de arquivo, [ argumentoiomode, [ criar, [ formato ]]])

'''''O método OpenTextFile possui as seguintes partes:

'''''Parte	Descrição
'''''objeto	Obrigatório. Sempre o nome de um FileSystemObject.
'''''nomes	Obrigatório. Expressão de cadeia de caracteres que identifica o arquivo a ser aberto.
'''''iomode	Opcional. Indica o modo de entrada e saída. Pode ser uma das três constantes: ParaLer, ParaEscrever, ou ParaAnexar.
'''''criar	Opcional. **** Valor booliano que indica se um novo arquivo pode ser criado se o nome de arquivo especificado não existir. O valor é true se um novo arquivo é criado; False se não for criado. O padrão é False.
'''''formato	Opcional. Um dos três valores Tristate usados para indicar o formato do arquivo aberto. Caso seja omitido, o arquivo é aberto como ASCII.

