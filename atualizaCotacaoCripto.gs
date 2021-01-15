function geraCotacao(url){
  // Conecta na API do mercadobitcoin, faz o parse, pega a última cotação e formata em reais.
  var fetch = UrlFetchApp.fetch(url);
  var jsonFetch = JSON.parse(fetch);
  
  cotacao = jsonFetch.ticker.last;
  cotacao = (Math.round(cotacao * 100) / 100).toFixed(2).replace('.', ',').replace(/(\d)(?=(\d{3})+\,)/g, "$1.");
  cotacao = "R$ "+cotacao.toString();

  return cotacao;
}
  
function atualizaCripto() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Necessário 2 porque além de pular 1 linha que é a do cabeçalho, precisa de +1 por causa do Array 
  var indexRow=2;

  /* Explicação do getRange
  2: the starting row
  2: the starting col 
  1000/sheet.getLastRow(): the number of rows
  2: the number of cols 
  */
  var range  = sheet.getRange(2, 2, sheet.getLastRow(), 2);
  var values = range.getValues();

  for (var row in values) {
    for (var col in values[row]) {
      
    var cont = 1
    
    // Usado para alterar os tipos de criptomoedas sem fazer inúmeros IFs
	  while(cont <= 4){
      
		  switch (cont){
		    case 1: coin="BTC"; break;
		    case 2: coin="ETH"; break;
		    case 3: coin="LTC"; break;
		    case 4: coin="XRP"; break;
		  }

      // Aqui vai pegar a coluna 2, que contém os tickers
	    if ( values[row][col] == coin ){

		    var url = "https://www.mercadobitcoin.net/api/"+coin+"/ticker";

        // getRange 3 porque já vai buscar direto na coluna de Cotação 
        sheet.getRange(indexRow, 3).setValue(geraCotacao(url));
        
	    }
	    cont++;
	    }

      // Usado para debugar e apresentar os valores na tela 
      // Logger.log(values[row][col]);
    }
    // Não precisa de indexCol porque as colunas do excel são fixas, só precisa achar em qual linha está o valor
    indexRow++;
  } 
} 
