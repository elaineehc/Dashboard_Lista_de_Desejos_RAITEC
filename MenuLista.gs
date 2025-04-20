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

//----------------------------FUNÇÕES ADICIONAR-----------------------------------------------------

function AdicionarItem(Dados){

  const user = LockService.getScriptLock();
  user.tryLock(10000);

  if(user.hasLock()){
    var ID = CalcularID();
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

    var linha = guiaLista.getLastRow();
    linha = linha + 1;
    var Data = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy");

    guiaLista.getRange(linha, 1).setValue(ID);
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
    return "Item adicionado com sucesso.";
  }
}

function CalcularID(){
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  var ultimaLinha = planilha.getLastRow();
  var ids = [];
  var teste = "";
  
  for(var linha=2; linha<=ultimaLinha; linha++){
    var id_atual = planilha.getRange(linha, 1).getValue();
    id_atual = id_atual.split("D");
    ids.push(parseInt(id_atual[1]));
  }

  ids.sort();

  for(var i=0; i<ids.length; i++){
    if(ids[i]!=i+1){
      var id_calculado = "LD"+(i+1<10?"0":"")+String(i+1);
      return id_calculado;
    }
  }
  var len = ids.length;
  return "LD"+(ids[len-1]+1<10?"0":"")+String(ids[len-1]+1);
}

//--------------------------FUNÇÕES EXCLUIR---------------------------------------------------

function ExcluirItem(id_procurado){

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  const ultimaLinha = planilha.getLastRow();

  var ids = planilha.getRange(2, 1, ultimaLinha - 1, 1).getValues();

  for(var i=0; i<ids.length; i++){
    var id_atual = ids[i][0];
    if(id_atual == id_procurado){
      planilha.deleteRow(i+2);
      return "Item excluído com sucesso.";
    }
    
  }
  return "Exclusão falhou.";
}

//------------------------FUNÇÕES PESQUISAR-----------------------------------------------------

function buscarItens(termo){

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  const ultimaLinha = planilha.getLastRow();
  const ultimaColuna = planilha.getLastColumn();

  var dados = planilha.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
  var id = planilha.getRange(2, 1, ultimaLinha - 1, 1).getValues();
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
      var texto=id[i]+' --  &#160 '+item[i]+' &#160 ('+espec[i]+')';
      resultados.push(texto);
    }
  }
  return resultados;
}

//--------------------------FUNÇÕES DETALHES---------------------------------------------------

function buscarDetalhes(id_procurado, num){

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  const ultimaLinha = planilha.getLastRow();
  const ultimaColuna = planilha.getLastColumn();

  var dados = planilha.getRange(2, 1, ultimaLinha - 1, ultimaColuna).getValues();
  var ids = planilha.getRange(2, 1, ultimaLinha - 1, 1).getValues();
  var nome = planilha.getRange(2, 2, ultimaLinha - 1, 1).getValues();
  var data = planilha.getRange(2, 3, ultimaLinha - 1, 1).getValues();
  var item = planilha.getRange(2, 4, ultimaLinha - 1, 1).getValues();
  var espec = planilha.getRange(2, 5, ultimaLinha - 1, 1).getValues();
  var eixo = planilha.getRange(2, 6, ultimaLinha - 1, 1).getValues();
  var motivo = planilha.getRange(2, 7, ultimaLinha - 1, 1).getValues();
  var qtd = planilha.getRange(2, 8, ultimaLinha - 1, 1).getValues();
  var valor = planilha.getRange(2, 9, ultimaLinha - 1, 1).getValues();
  var status = planilha.getRange(2, 10, ultimaLinha - 1, 1).getValues();
 
  for(var i=0; i<ids.length; i++){
    var id_atual = ids[i];

    if(id_atual == id_procurado){
      
      var dataFormatada = new Date(data[i]).toLocaleDateString('pt-BR');
      var dataFormatada2 = formatarData(dataFormatada);

      var texto1='<p>'+
          '<strong>ID:</strong> '+ids[i]+'<br>'+
          '<strong>Solicitante:</strong> '+nome[i]+'<br>'+
          '<strong>Data:</strong> '+dataFormatada2+'<br>'+
          '<strong>Item:</strong> '+item[i]+'<br>'+
          '<strong>Especificações:</strong> '+espec[i]+'<br>'+
          '<strong>Eixo:</strong> '+eixo[i]+'<br>'+
          '<strong>Motivo:</strong> '+motivo[i]+'<br>'+
          '<strong>Quantidade:</strong> '+qtd[i]+'<br>'+
          '<strong>Valor:</strong> R$ '+valor[i]+'<br>'+
          '<strong>Status:</strong> '+status[i]+
        '</p>';

      var texto2 = {ID: ids[i][0],
                    Nome: nome[i][0],
                    Item: item[i][0],
                    Espec: espec[i][0],
                    Eixo: eixo[i][0],
                    Motivo: motivo[i][0],
                    Qtd: qtd[i][0],
                    Valor: valor[i][0],
                    Status: status[i][0]};

      if(num==1) return texto1;
      if(num==2) return texto2;
    }
  }
}

function formatarData(dataTexto) {
  var meses = {
    January: '01',
    February: '02',
    March: '03',
    April: '04',
    May: '05',
    June: '06',
    July: '07',
    August: '08',
    September: '09',
    October: '10',
    November: '11',
    December: '12'
  };
  var partes = dataTexto.replace(',', '').split(' '); // ["March", "23", "2025"]
  var dia = partes[1];
  var mes = meses[partes[0]];
  var ano = partes[2];

  if(Number(dia)<10){
    dia = '0'+dia;
  }
  return dia + '/' + mes + '/' + ano;
}

//------------------------------FUNÇÕES EDITAR-----------------------------------------------

function EditarItem(Dados){

  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista de Desejos");
  const ultimaLinha = planilha.getLastRow();

  var ids = planilha.getRange(2, 1, ultimaLinha - 1, 1).getValues();

  for(var linha=0; linha<ids.length; linha++){
    if(ids[linha][0]==Dados.ID){
      planilha.getRange(linha, 2).setValue(Dados.Nome);
      planilha.getRange(linha, 4).setValue(Dados.Item);
      planilha.getRange(linha, 5).setValue(Dados.Espec);
      planilha.getRange(linha, 6).setValue(Dados.Eixo);
      planilha.getRange(linha, 7).setValue(Dados.Motivo);
      planilha.getRange(linha, 8).setValue(Dados.Quantidade);
      planilha.getRange(linha, 9).setValue(Dados.Valor);
      planilha.getRange(linha, 10).setValue(Dados.Status);
      return "Item editado com sucesso.";
    }
  }
  return "Edição falhou.";
}
