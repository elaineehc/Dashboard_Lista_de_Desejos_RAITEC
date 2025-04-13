function FormLista() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaLista = planilha.getSheetByName("Lista de Desejos");

  var ultimaLinha = guiaLista.getLastRow()-1;
  if(ultimaLinha==0){
    ultimaLinha=1;
  }

  var list = guiaLista.getRange(2,1,ultimaLinha,1).getValues();
  list.sort();

  var Form = HtmlService.createTemplateFromFile("Menu");
  Form.list = list.map(function(r){return r[0];});

  var MostrarForm = Form.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  MostrarForm.setTitle("Lista de Desejos").setHeight(400).setWidth(710);

  SpreadsheetApp.getUi().showModalDialog(MostrarForm, "Lista de Desejos"); 
}

function Chamar(Arquivo){
  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}

//---------------------------------------------------------------------------------

function AdicionarItem(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
    var Nome = Dados.Nome;
    var Item = Dados.Item;
    var Espec = Dados.Espec;
    var Motivo = Dados.Motivo;
    var Eixo = Dados.Eixo;
    var Status = Dados.Status;
    var Quantidade = Dados.Quantidade;
    var Valor = Dados.Valor;

    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var guiaLista = planilha.getSheetByName("Lista de Desejos");

    var ultimaLinha = guiaLista.getLastRow();
    var dadosItem = guiaLista.getRange(2,1,ultimaLinha, 1).getValues();

    // for(var linha=0; linha<dadosItem.length; linha++){
    //   if(dadosItem[linha][0]==Nome){
    //      dadosItem.length=0;
    //      return "VENDEDOR JÁ CADASTRADO!";
    //   }
    // }

    var linha = guiaLista.getLastRow();
    linha = linha + 1;
    var Data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");

    guiaLista.getRange(linha, 2).setValue(Nome);
    guiaLista.getRange(linha, 3).setValue(Data);
    guiaLista.getRange(linha, 4).setValue(Item);
    guiaLista.getRange(linha, 5).setValue(Espec);
    guiaLista.getRange(linha, 6).setValue(Eixo);
    guiaLista.getRange(linha, 7).setValue(Motivo);
    guiaLista.getRange(linha, 8).setValue(Quantidade);
    guiaLista.getRange(linha, 9).setValue(Valor);
    guiaLista.getRange(linha, 10).setValue(Status);

    dadosItem.length=0;
    return "REGISTRADO COM SUCESSO!";
  }
}

//-----------------------------------------------------------------------------

function buscarItens(termo){

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  const ultimaLinha = planilha.getLastRow();
  const ultimaColuna = planilha.getLastColumn();

  var dados = planilha.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
  var nome = planilha.getRange(2, 2, ultimaLinha - 1, 1).getValues();
  var data = planilha.getRange(2, 3, ultimaLinha - 1, 1).getValues();
  var item = planilha.getRange(2, 4, ultimaLinha - 1, 1).getValues();
  var espec = planilha.getRange(2, 5, ultimaLinha - 1, 1).getValues();
  var eixo = planilha.getRange(2, 6, ultimaLinha - 1, 1).getValues();
  var motivo = planilha.getRange(2, 7, ultimaLinha - 1, 1).getValues();
  var qtd = planilha.getRange(2, 8, ultimaLinha - 1, 1).getValues();
  var valor = planilha.getRange(2, 9, ultimaLinha - 1, 1).getValues();
  var status = planilha.getRange(2, 10, ultimaLinha - 1, 1).getValues();

  var resultados = [];
  
  termo = termo.toLowerCase();
  
  for(var i=0; i<dados.length; i++){
    var linha = dados[i];
    var linhaTexto = linha.join(" ").toLowerCase();

    if(linhaTexto.indexOf(termo)!=-1){
      resultados.push(nome[i]);
    }
  }
  //return "capuccino assassino";
  return resultados;
}













