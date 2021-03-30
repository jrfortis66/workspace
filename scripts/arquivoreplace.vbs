strArquivo_Original = "arquivo1.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strArquivo_Original, 1, false)

Text1="<Event>"
Text2="<Timestamp data_type=""4"">"
Text3="</Timestamp>"

If objFile.AtEndOfStream Then
	ReadAllTextFile = ""
Else
	ReadAllTextFile = objFile.ReadAll
End If
objFile.close

ReadAllTextFile = replace(ReadAllTextFile,Text1,"") 'substitui por ""
ReadAllTextFile = replace(ReadAllTextFile,Text2,"") 'substitui por ""
ReadAllTextFile = replace(ReadAllTextFile,Text3,",")'substitui por virgula

Set objFile = objFSO.OpenTextFile(strArquivo_Original, 2, true)

objFile.write ReadAllTextFile

objFile.close

.ps1

$Arquivo=".\arquivo1.txt"

#Carrega conteudo do arquivo na variavel $Novo
$Novo = (Get-Content $Arquivo)


#Faz substituições
$Novo = $Novo | Foreach-Object {$_ -replace "<Event>",""}
$Novo = $Novo | Foreach-Object {$_ -replace '<Timestamp data_type="4">',""}
$Novo = $Novo | Foreach-Object {$_ -replace "</Timestamp>",","}

#Grava substituições
$Novo | Set-Content ($Arquivo)