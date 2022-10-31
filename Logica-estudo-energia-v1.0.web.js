// @ts-nocheck
const ss = SpreadsheetApp.getActiveSpreadsheet();
const planLogin = ss.getSheetByName("Login");
const planDados = ss.getSheetByName("Base dados");
const planDadosCalculos = ss.getSheetByName("Dados e calculos");
const planPrincipal = ss.getSheetByName("Tela principal");
const planTelaDados = ss.getSheetByName("Tela dados");
const planImpressao = ss.getSheetByName("Impressao")
const planSobre = ss.getSheetByName("Sobre");

function seguranca(){

  senha = planLogin.getRange("C2").getValue();
  aux = planLogin.getRange("C3").getValue();
  aux1 = planLogin.getRange("C4").getValue();

  if(senha == 0){

    planLogin.showSheet();

    Utilities.sleep(1000);

    if(aux1 == 0){
      SpreadsheetApp.setActiveSheet(planLogin);
      planLogin.getRange("C3").setValue(0);
      planLogin.getRange("C4").setValue(1);
    }

    planPrincipal.hideSheet();
    planTelaDados.hideSheet();
  
  }else if(senha == 1){

    planPrincipal.showSheet();
    planTelaDados.showSheet();
    
    Utilities.sleep(1000);
    
    if(aux == 0){
      SpreadsheetApp.setActiveSheet(planPrincipal);
      planLogin.getRange("C3").setValue(1);
      planLogin.getRange("C4").setValue(0);
    }

    planLogin.hideSheet();
    
  } 
}

function telaLogin(){

  var template = HtmlService.createTemplateFromFile("userform");

  var html = template.evaluate();
  html.setTitle("SOFTWARE ESTUDO DE ENERGIA INCIDENTE").setHeight(350).setWidth (550);
  SpreadsheetApp.getUi().showModalDialog(html, "SOFTWARE ESTUDO DE ENERGIA INCIDENTE"); // TELA MÓVEL SEM SEGUNDO PLANO
}

function validacaoLogin(data){

  let validacao = [data.validacao];

  if(validacao == 1){

    planLogin.getRange("C2").setValue(1);
    seguranca();

  }
}

function continuar(){
  //planLogin.getRange("B2").setValue(0); //Tira dúvida;

  //PEGANDO AS INFORMAÇÕES DA TELA PRINCIPAL

  cl_tensao = planPrincipal.getRange("G19").getValue();
  freq = planPrincipal.getRange("G22").getValue();
  dist_barra = planPrincipal.getRange("G25").getValue();
  tip_amb = planPrincipal.getRange("G28").getValue();
  tag = planPrincipal.getRange("G31").getValue();
  tip_aterra = planPrincipal.getRange("M19").getValue();
  icc_sime = planPrincipal.getRange("M22").getValue();
  tip_equi = planPrincipal.getRange("M25").getValue();
  tip_disj_prot = planPrincipal.getRange("M28").getValue();
  tempo_atua = planDadosCalculos.getRange("D10").getValue();
  dist_traba = planDadosCalculos.getRange("D11").getValue();
  ener_inci_j = planDadosCalculos.getRange("D12").getValue();
  ener_inci_cal = planDadosCalculos.getRange("D13").getValue();
  limite = planDadosCalculos.getRange("D14").getValue();
  metodo = planDadosCalculos.getRange("D15").getValue();
  risco = planDadosCalculos.getRange("D16").getValue();
  cat_risco = planDadosCalculos.getRange("D17").getValue();
  atpv = planDadosCalculos.getRange("D18").getValue();
  epi = planDadosCalculos.getRange("D19").getValue();
  zero = planPrincipal.getRange("P19").getValue();

  // SE TODOS OS VALORES FOREM PREENCHIDOS VAI DAR CONTINUIDADE NA FUNÇÃO
 
  
  if (cl_tensao == zero || freq == zero || dist_barra == zero || tip_amb == zero || tag == zero || tip_aterra == zero || icc_sime == zero || tip_equi == zero || tip_disj_prot == zero  ) {

  }

  else{

  
  //INSERINDO DADOS NO CALCULO E DADOS

  planDadosCalculos.getRange("D1").setValue(tag);
  planDadosCalculos.getRange("D2").setValue(cl_tensao);
  planDadosCalculos.getRange("D3").setValue(freq);
  planDadosCalculos.getRange("D4").setValue(dist_barra);
  planDadosCalculos.getRange("D5").setValue(tip_amb);
  planDadosCalculos.getRange("D6").setValue(tip_aterra);
  planDadosCalculos.getRange("D7").setValue(icc_sime);
  planDadosCalculos.getRange("D8").setValue(tip_equi);
  planDadosCalculos.getRange("D9").setValue(tip_disj_prot);

  //INSERINDO NA TELA DE DADOS

  planTelaDados.getRange("G19").setValue(tempo_atua);
  planTelaDados.getRange("G22").setValue(dist_traba);
  planTelaDados.getRange("G25").setValue(ener_inci_j);
  planTelaDados.getRange("G28").setValue(ener_inci_cal);
  planTelaDados.getRange("G31").setValue(limite);
  planTelaDados.getRange("M19").setValue(metodo);
  planTelaDados.getRange("M22").setValue(risco);
  planTelaDados.getRange("M25").setValue(cat_risco);
  planTelaDados.getRange("M28").setValue(atpv);
  planTelaDados.getRange("J31").setValue(epi);

  // INSERINDO NA IMPRESSAO
  planImpressao.getRange("F13").setValue(tag);
  planImpressao.getRange("F14").setValue(cl_tensao);
  planImpressao.getRange("F15").setValue(ener_inci_cal);
  planImpressao.getRange("F16").setValue(risco);
  planImpressao.getRange("H13").setValue(limite); 
  planImpressao.getRange("H14").setValue(dist_traba);
  planImpressao.getRange("H15").setValue(atpv);
  planImpressao.getRange("H16").setValue(cat_risco);
  planImpressao.getRange("E18").setValue(epi);

  //OCULTAR TELA

  SpreadsheetApp.setActiveSheet(planTelaDados);
  
  }

}

//VOLTA NA TELA PRINCIPAL E OCULTA A TELA DE DADOS

function voltar(){
  
  SpreadsheetApp.setActiveSheet(planPrincipal);
 
}

function reiniciar(){

  //REINICIA O PROCESSO MANTENDO ALGUMAS INFORMAÇÕES
  SpreadsheetApp.setActiveSheet(planPrincipal);

  planPrincipal.getRange("G19").setValue("");
  planPrincipal.getRange("G22").setValue("");
  planPrincipal.getRange("G25").setValue("");
  planPrincipal.getRange("G28").setValue("Fechado");
  planPrincipal.getRange("G31").setValue("");
  planPrincipal.getRange("M19").setValue("Solidamente aterrado");
  planPrincipal.getRange("M22").setValue("");
  planPrincipal.getRange("M25").setValue("Painel de distribuição");
  planPrincipal.getRange("M28").setValue("");

}

function finalizar(){

  planLogin.getRange("C2").setValue(0);

  //apaga dados iniciais

  planPrincipal.getRange("G19").setValue("");
  planPrincipal.getRange("G22").setValue("");
  planPrincipal.getRange("G25").setValue("");
  planPrincipal.getRange("G28").setValue("Fechado");
  planPrincipal.getRange("G31").setValue("");
  planPrincipal.getRange("M19").setValue("Solidamente aterrado");
  planPrincipal.getRange("M22").setValue("");
  planPrincipal.getRange("M25").setValue("Painel de distribuição");
  planPrincipal.getRange("M28").setValue("");
  
  seguranca();  

}

function pdf(){
  planImpressao.showSheet();
  planTelaDados.hideSheet();
  planPrincipal.hideSheet();

  var nomePDF = planImpressao.getRange("F13").getValue(); //VARIAVEL RECEBENDO O NOME DO PDF

  //planLogin.getRange("B2").setValue(2);

  var folderIter = DriveApp.getFoldersByName("Estudos");//PEGA A PASTA DENTRO DO GOOGLE DRIVE
  var pdfFolder = folderIter.next();//ENTRA NA PASTA

  var spredsheet_id = ss.getId();//PEGA ID DA PLANILHA ATUAL
  var spredsheetFile = DriveApp.getFileById(spredsheet_id);//PEGA A PLANILHA DENTRO DO GOOGLE DRIVE PELO ID
  var blob = spredsheetFile.getAs (MimeType.PDF); // PEGA A PLANILHA DENTRO DO GOOGLE DRIVE COMO PDF
  pdfFolder.createFile(blob).setName(nomePDF);// SALVA A PLANILHA COMO PDF NO GOOGLE DRIVE 

  planTelaDados.showSheet();
  planPrincipal.showSheet();
  planImpressao.hideSheet();
  
}

function voltarsobre(){

  SpreadsheetApp.setActiveSheet(planPrincipal);
  planSobre.hideSheet();

}

function sobre(){

  planSobre.showSheet();
  SpreadsheetApp.setActiveSheet(planSobre);

}