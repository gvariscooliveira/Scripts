#CREATE TABLE PETR# (
#Codigo varchar(10) NULL,
#Abertura FLOAT(5) NULL,
#Maximo FLOAT(5) NULL,
#Minimo FLOAT(5) NULL,
#Medio FLOAT(5) NULL,
#Ultimo FLOAT(5) NULL,
#Oscilacao FLOAT(5) NULL,
#Nome FLOAT(5) NULL,
#Data DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
#)

#Define caminho para execução
$Caminho = "C:\Temp\"
$UrlBovespa = "http://www.bmfbovespa.com.br"
$UrlSourceOpcoes = $UrlBovespa + "/pt_br/servicos/market-data/consultas/mercado-a-vista/opcoes/series-autorizadas/"
$FiltroInnexTextOpcoes = "Lista Completa de S*ries Autorizadas"
$linkrequest = "http://bvmf.bmfbovespa.com.br/cotacoes2000/FormConsultaCotacoes.asp?strListaCodigos="
$arquivoOpcoes = $Caminho + "ListaOpcoes.txt"

#Navega até caminho de execução
cd $Caminho

#Terceira segunda feira do mês
$vencimento = Get-Date
$terceirasegunda = ((1..[DateTime]::DaysInMonth((Get-Date).Year,(Get-Date).month) | ForEach-Object { (Get-Date -Day $_| ? {$_.dayofweek -eq 'sunday'}) })[3]).adddays(-6)
if ($vencimento -ge $terceirasegunda){
	$tsegunda = $vencimento.month
	} else {
	$tsegunda = $vencimento.month -1
	}

#Link para baixar todas as opções na Bovespa
$SourceOpcoes = Invoke-WebRequest($UrlSourceOpcoes)
$linkOpcoesParcial = $SourceOpcoes.Links | Where {$_.innertext -like $FiltroInnexTextOpcoes} | Select-Object href
$linkOpcoesFinal = $UrlBovespa + $linkOpcoesParcial.href 

#Download e extração do arquivo de opções
$arquivo = $Caminho + "arquivo.zip"
Invoke-WebRequest -Uri $linkOpcoesFinal -OutFile $arquivo
Expand-Archive -LiteralPath $arquivo
$arquivoTexto = $Caminho + "arquivo\SI_D_SEDE.txt"

Remove-Item $arquivoOpcoes -EA SilentlyContinue
Get-ChildItem -Path "C:\Temp" | Where {$_.Name -like "*.xml"} | Remove-Item

$ListaCompletaOpcoes = Get-Content $arquivoTexto | Select-Object -Skip 1
foreach ($opcao in $ListaCompletaOpcoes){
	$arrOpcao = $opcao -Split "\|"
	
	#Janeiro
	if (($vencimento.Month -eq 1) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRA*" -OR $arrOpcao[13] -like "PETRB*" -OR $arrOpcao[13] -like "PETRM*" -OR $arrOpcao[13] -like "PETRN*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 1) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRB*" -OR $arrOpcao[13] -like "PETRC*" -OR $arrOpcao[13] -like "PETRN*" -OR $arrOpcao[13] -like "PETRO*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Fevereiro
	if (($vencimento.Month -eq 2) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRB*" -OR $arrOpcao[13] -like "PETRC*" -OR $arrOpcao[13] -like "PETRN*" -OR $arrOpcao[13] -like "PETRO*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 2) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRC*" -OR $arrOpcao[13] -like "PETRD*" -OR $arrOpcao[13] -like "PETRO*" -OR $arrOpcao[13] -like "PETRP*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Março
	if (($vencimento.Month -eq 3) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRC*" -OR $arrOpcao[13] -like "PETRD*" -OR $arrOpcao[13] -like "PETRO*" -OR $arrOpcao[13] -like "PETRP*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 3) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRD*" -OR $arrOpcao[13] -like "PETRE*" -OR $arrOpcao[13] -like "PETRP*" -OR $arrOpcao[13] -like "PETRQ*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Abril
	if (($vencimento.Month -eq 4) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRD*" -OR $arrOpcao[13] -like "PETRE*" -OR $arrOpcao[13] -like "PETRP*" -OR $arrOpcao[13] -like "PETRQ*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 4) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRE*" -OR $arrOpcao[13] -like "PETRF*" -OR $arrOpcao[13] -like "PETRQ*" -OR $arrOpcao[13] -like "PETRR*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}

	#Maio
	if (($vencimento.Month -eq 5) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRE*" -OR $arrOpcao[13] -like "PETRF*" -OR $arrOpcao[13] -like "PETRQ*" -OR $arrOpcao[13] -like "PETRR*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 5) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRF*" -OR $arrOpcao[13] -like "PETRG*" -OR $arrOpcao[13] -like "PETRR*" -OR $arrOpcao[13] -like "PETRS*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}

	#Junho
	if (($vencimento.Month -eq 6) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRF*" -OR $arrOpcao[13] -like "PETRG*" -OR $arrOpcao[13] -like "PETRR*" -OR $arrOpcao[13] -like "PETRS*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 6) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRG*" -OR $arrOpcao[13] -like "PETRH*" -OR $arrOpcao[13] -like "PETRS*" -OR $arrOpcao[13] -like "PETRT*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}	
	
	#Julho
	if (($vencimento.Month -eq 7) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRG*" -OR $arrOpcao[13] -like "PETRH*" -OR $arrOpcao[13] -like "PETRS*" -OR $arrOpcao[13] -like "PETRT*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 7) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRH*" -OR $arrOpcao[13] -like "PETRI*" -OR $arrOpcao[13] -like "PETRT*" -OR $arrOpcao[13] -like "PETRU*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Agosto
	if (($vencimento.Month -eq 8) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRH*" -OR $arrOpcao[13] -like "PETRI*" -OR $arrOpcao[13] -like "PETRT*" -OR $arrOpcao[13] -like "PETRU*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 8) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRI*" -OR $arrOpcao[13] -like "PETRJ*" -OR $arrOpcao[13] -like "PETRU*" -OR $arrOpcao[13] -like "PETRV*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Setembro
	if (($vencimento.Month -eq 9) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRI*" -OR $arrOpcao[13] -like "PETRJ*" -OR $arrOpcao[13] -like "PETRU*" -OR $arrOpcao[13] -like "PETRV*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 9) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRJ*" -OR $arrOpcao[13] -like "PETRK*" -OR $arrOpcao[13] -like "PETRV*" -OR $arrOpcao[13] -like "PETRW*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Outubro
	if (($vencimento.Month -eq 10) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRJ*" -OR $arrOpcao[13] -like "PETRK*" -OR $arrOpcao[13] -like "PETRV*" -OR $arrOpcao[13] -like "PETRW*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 10) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRK*" -OR $arrOpcao[13] -like "PETRL*" -OR $arrOpcao[13] -like "PETRW*" -OR $arrOpcao[13] -like "PETRX*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Novembro
	if (($vencimento.Month -eq 11) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRK*" -OR $arrOpcao[13] -like "PETRL*" -OR $arrOpcao[13] -like "PETRW*" -OR $arrOpcao[13] -like "PETRX*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 11) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRL*" -OR $arrOpcao[13] -like "PETRA*" -OR $arrOpcao[13] -like "PETRX*" -OR $arrOpcao[13] -like "PETRM*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	
	#Dezembro
	if (($vencimento.Month -eq 12) -and ($tsegunda -eq 0)){
		if ($arrOpcao[13] -like "PETRL*" -OR $arrOpcao[13] -like "PETRA*" -OR $arrOpcao[13] -like "PETRX*" -OR $arrOpcao[13] -like "PETRM*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}
	if (($vencimento.Month -eq 12) -and ($tsegunda -eq 1)){
		if ($arrOpcao[13] -like "PETRA*" -OR $arrOpcao[13] -like "PETRB*" -OR $arrOpcao[13] -like "PETRM*" -OR $arrOpcao[13] -like "PETRN*"){
			$arrOpcao[13].Trim() >> $arquivoOpcoes
		}
	}

} 

$Opcoes = Get-Content $arquivoOpcoes
$TotalOpcoes = $Opcoes.Count
$seqreq = 0
$contTotOpcoes = 0

foreach ($linhaopcao in $Opcoes){
    $seqreq++
    $contTotOpcoes++
    $conjOpcoes += $linhaopcao + "|"
    if ($seqreq -eq 24){
        $linkrequestopcao = $linkrequest + $conjOpcoes
        $Request = Invoke-WebRequest $linkrequestopcao
		start-sleep -s 15
        if ($Request.StatusCode -eq "200" -AND $Request.StatusDescription -eq "OK" -AND $Request.Content -notlike "<ERROS>*"){
            $XML = $Request.Content
            $CaminhoXML = $Caminho + $contTotOpcoes + ".xml"
            $XML | Set-Content -Path $CaminhoXML -Encoding utf8
        }
    $conjOpcoes = ""
    $seqreq = 0
    }
    if ($contTotOpcoes -eq $TotalOpcoes){
        $linkrequestopcao = $linkrequest + $conjOpcoes
        $Request = Invoke-WebRequest $linkrequestopcao
        if ($Request.StatusCode -eq "200" -AND $Request.StatusDescription -eq "OK" -AND $Request.Content -notlike "<ERROS>*"){
            $XML = $Request.Content
            $CaminhoXML = $Caminho + $contTotOpcoes + ".xml"
            $XML | Set-Content -Path $CaminhoXML -Encoding utf8
        }
    }
}

# Get all XML files
$items = Get-ChildItem $Caminho\*.xml -Include *.xml

# Loop over them and append them to the document
foreach ($item in $items) {
    $inputFile = [XML](Get-Content $item) #load xml document    	
#export xml as csv
$inputFile.ComportamentoPapeis.papel | Export-Csv -Path $caminho\importdb.csv -Append -NoTypeInformation -Delimiter ";" -Encoding UTF8
}

#clean folder
#Remove-Item C:\temp\* -include *.xml,*.zip,*.txt
#Remove-Item C:\temp\arquivo\ 

$import = "C:\temp\importdb.csv"
$CSV = Get-Content "C:\temp\importdb.csv" | Select-Object -Skip 1

foreach ($i in $CSV){
	$linha = $i -Split ";"
	$Strike = $linha[1] -Split " "
	$StrikeF = $Strike[-1].Replace('"','')
	$Data = $linha[3].Replace('"','')
	$Datas = $Data -Split ":"
	$DataC = $Datas[0] + ":" + $Datas[1]
	$Codigo = $linha[0].Replace('"','')
	$Nome = $StrikeF.Replace(',','.')
	$Abertura = $linha[4].Replace('"','').Replace(',','.')
	$Minimo = $linha[5].Replace('"','').Replace(',','.')
	$Maximo = $linha[6].Replace('"','').Replace(',','.')
	$Medio = $linha[7].Replace('"','').Replace(',','.')
	$Ultimo = $linha[8].Replace('"','').Replace(',','.')
	$Oscilacao = $linha[9].Replace('"','').Replace(',','.')
	
	#Conexão DB e insert on table
	$MySQLAdminUserName = 'root'
	$MySQLAdminPassword = 'Varisco2012'
	$MySQLDatabase = 'bvmf'
	$MySQLHost = '127.0.0.1'
	[void][system.reflection.Assembly]::LoadWithPartialName("MySql.Data")
	$dbconnect = New-Object MySql.Data.MySqlClient.MySqlConnection
	$dbconnect.ConnectionString = "server=" + $MySQLHost + ";port=3306;uid=" + $MySQLAdminUserName + ";pwd=" + $MySQLAdminPassword + ";database="+$MySQLDatabase
	$dbconnect.Open()
	$sql = New-Object MySql.Data.MySqlClient.MySqlCommand
	$sql.Connection = $dbconnect
	
	#INSERT Janeiro
	if (($vencimento.Month -eq 1) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petra (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrb (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrm (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrn (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 1) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrb (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrc (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrn (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petro (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT fevereiro
	if (($vencimento.Month -eq 2) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrb (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrc (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrn (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petro (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 2) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrc (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrd (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petro (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrp (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Março
	if (($vencimento.Month -eq 3) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrc (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrd (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petro (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrp (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 3) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrd (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petre (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrp (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrq (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Abril
	if (($vencimento.Month -eq 4) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrd (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petre (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrp (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrq (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 4) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petre (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrf (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrq (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrr (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Maio
	if (($vencimento.Month -eq 5) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petre (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrf (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrq (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrr (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 5) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrf (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrg (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrr (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrs (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Junho
	if (($vencimento.Month -eq 6) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrf (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrg (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrr (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrs (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 6) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrg (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrh (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrs (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrt (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Julho
	if (($vencimento.Month -eq 7) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrg (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrh (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrs (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrt (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 7) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrh (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petri (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrt (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petru (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Agosto
	if (($vencimento.Month -eq 8) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrh (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petri (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrt (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petru (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 8) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petri (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrj (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petru (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrv (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Setembro
	if (($vencimento.Month -eq 9) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petri (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrj (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petru (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrv (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 9) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrj (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrk (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrv (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrw (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Outubro
	if (($vencimento.Month -eq 10) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrj (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrk (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrv (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrw (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 10) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrk (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrl (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrw (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrx (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Novembro
	if (($vencimento.Month -eq 11) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrk (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrl (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrw (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrx (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 11) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petrl (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petra (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrx (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrm (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
	
	#INSERT Dezembro
	if (($vencimento.Month -eq 12) -and ($tsegunda -eq 0)){
		$sql.CommandText="INSERT INTO petrl (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petra (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrx (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrm (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
		
	if (($vencimento.Month -eq 12) -and ($tsegunda -eq 1)){
		$sql.CommandText="INSERT INTO petra (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrb (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrm (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
		$sql.CommandText="INSERT INTO petrn (Codigo, Nome, Abertura, Minimo, Maximo, Medio, Ultimo, Oscilacao, Data) VALUES ('$Codigo','$Nome','$Abertura','$Minimo','$Maximo','$Medio','$Ultimo','$Oscilacao',STR_TO_DATE('$DataC','%d/%m/%Y %H:%i'))"
	}
			
	$dr = $sql.ExecuteNonQuery()
	$sql.Dispose()
	$dbconnect.Close()
}
#Remove-Item C:\temp\importdb.csv
