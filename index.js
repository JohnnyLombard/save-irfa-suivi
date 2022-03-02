// info :
// -- Mettre déclancheur par fonction (selon heure, intervalle d'une heure entre chaque, TEST OK)
// -- Génération de lien de formulaire selon la date prévue TEST OK
// -- Envoyer mail après la création du lien TEST OK sortie et 3 mois et 6 mois
// -- Enlever obligation sur formulaire champ téléphone fixe
// -- Créer sheet reception données formulaire ( 1 par sortie m, m+3, m+6)

/***************************************************************/
/**                                                           **/
/**      Créé le lien du formulaire de suivi en sortie        **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
function createLinkFormOutLearn() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_Sortie"); //ID Google Sheets et nom de l'onglet
  var dataRange = sheet.getRange("A3:Q");
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    var emailAddress = rowData[7];
    var name = rowData[3];
    var firstname = rowData[4];
    var link = rowData[16];
    var status = rowData[0];
    var setRow = parseInt(i) + 3;
    var date01 = new Date();
    var out = rowData[13];

    if (date01.valueOf() >= out) {
      sheet
        .getRange(setRow, 17)
        .setFormula(
          '=IF(B3:B="";"";JOIN("";Bilan_Sortie!B43&Bilan_Sortie!B44&B3:B&Bilan_Sortie!B45&VLOOKUP(B3:B;BDD_Action!$A$2:$E;2;0)&Bilan_Sortie!B46&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);2))&Bilan_Sortie!B47&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);2))&Bilan_Sortie!B48&VLOOKUP(B3:B;BDD_Action!$A$2:$E;5;0)&Bilan_Sortie!B49&D3:D&Bilan_Sortie!B50&E3:E&Bilan_Sortie!B51&F3:F&Bilan_Sortie!B52&G3:G))'
        );
    }

    // //     // if (link == "")
    // //     //    break;
  } //}}
}

/***************************************************************/
/**                                                           **/
/**         Envoi un email au stagiaire à la sortie           **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
function sendMailOutLearn() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_Sortie"); //ID Google Sheets et nom
  var dataRange = sheet.getRange("A3:Q");
  var startcolumn = 2;
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var startcolumn = 3;
    var currentRow = data[i];
    var emailAddress = currentRow[7];
    var name = currentRow[3];
    var firstname = currentRow[4];
    var link = currentRow[16];
    var status = currentRow[0];
    var setRow = parseInt(i) + 3;

    if (currentRow[0] == "" && currentRow[16] != "") {
      // if (link == "")
      //    break;
      // if (status == "GForms envoyé")
      //    continue;
      // if (emailAddress == "")
      //    break;

      var link_url = '<a href="' + link + '">Lien vers le formulaire</a>';
      var message =
        "<p>Bonjour " +
        firstname +
        " " +
        name +
        ",</p>" +
        "<p>Conformément aux obligations du financeur de votre parcours de formation suivi au sein de Via Formation, nous vous envoyons un questionnaire afin de connaître votre situation actuelle à la sortie de votre formation </p>" +
        "<p> Veuillez remplir le questionnaire que vous trouverez en cliquant sur le lien suivant : " +
        link_url +
        "</p>";

      var subject = "Questionnaire de suivi de formation à la sortie";

      //Envoyer l'email
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message,
      });
      var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
      sheet.getRange(startcolumn + i, 1).setValue("GForms envoyé le : " + date);
    }
  }
}

/***************************************************************/
/**                                                           **/
/**      Créé le lien du formulaire de suivi à 3 mois         **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
function createLinkForm3Month() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_3mois"); //ID Google Sheets et nom de l'onglet
  var dataRange = sheet.getRange("A3:R");
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    var emailAddress = rowData[7];
    var name = rowData[3];
    var firstname = rowData[4];
    var link = rowData[16];
    var status = rowData[0];
    var setRow = parseInt(i) + 3;
    var date01 = new Date();
    var out = rowData[13];
    var plusTrois = rowData[17];

    // var formattedDebut = Utilities.formatDate(new Date(), "GMT +1", "dd/MM/yyyy");
    // var FormattedFin = Utilities.formatDate(new Date(plusTrois), "GMT + 1", "dd/MM/yyyy");

    if (date01 >= plusTrois && rowData[0] == "") {
      sheet
        .getRange(setRow, 17)
        .setFormula(
          '=IF(B3:B="";"";JOIN("";Bilan_3mois!B43&Bilan_Sortie!B44&B3:B&Bilan_Sortie!B45&VLOOKUP(B3:B;BDD_Action!$A$2:$E;2;0)&Bilan_Sortie!B46&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);2))&Bilan_Sortie!B47&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);2))&Bilan_Sortie!B48&VLOOKUP(B3:B;BDD_Action!$A$2:$E;5;0)&Bilan_Sortie!B49&D3:D&Bilan_Sortie!B50&E3:E&Bilan_Sortie!B51&F3:F&Bilan_Sortie!B52&G3:G))'
        );
    }
    //     if (link == "")
    //        break;
  }
}

// /***************************************************************/
// /**                                                           **/
// /**         Envoi un email au stagiaire à 3 mois              **/
// /**                                                           **/
// /**                 Code By Via formation                     **/
// /**                                                           **/
// /***************************************************************/
function sendMailOut3Month() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_3mois"); //ID Google Sheets et nom
  var dataRange = sheet.getRange("A3:Q");
  var startcolumn = 2;
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var startcolumn = 3;
    var currentRow = data[i];
    var emailAddress = currentRow[7];
    var name = currentRow[3];
    var firstname = currentRow[4];
    var link = currentRow[16];
    var status = currentRow[0];
    var setRow = parseInt(i) + 3;

    if (currentRow[0] == "" && currentRow[16] != "") {
      //     if (link == "")
      //        break;
      //     if (status == "GForms envoyé")
      //        continue;
      //     if (emailAddress == "")
      //        break;

      var link_url = '<a href="' + link + '">Lien vers le formulaire</a>';
      var message =
        "<p>Bonjour " +
        firstname +
        " " +
        name +
        ",</p>" +
        "<p>Conformément aux obligations du financeur de votre parcours de formation, nous vous envoyons un questionnaire afin de connaître votre situation actuelle à 3 mois après la fin de votre formation </p>" +
        "<p> Veuillez remplir le questionnaire que vous trouverez en cliquant sur le lien suivant : " +
        link_url +
        "</p>";

      var subject = "Questionnaire de suivi de formation à 3 mois";

      //Envoyer l'email
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message,
      });
      var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
      sheet.getRange(setRow, 1).setValue("GForms envoyé le : " + date);
    }
  }
}

/***************************************************************/
/**                                                           **/
/**      Créé le lien du formulaire de suivi à 6 mois         **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
function createLinkForm6Month() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_6mois"); //ID Google Sheets et nom de l'onglet
  var dataRange = sheet.getRange("A3:R");
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var rowData = data[i];
    var emailAddress = rowData[7];
    var name = rowData[3];
    var firstname = rowData[4];
    var link = rowData[16];
    var status = rowData[0];
    var setRow = parseInt(i) + 3;
    var date02 = new Date();
    var out = rowData[13];
    var plusSix = rowData[17];

    //var date01 = Utilities.formatDate(new Date(), "GMT +1", "dd/MM/yyyy");
    // var FormattedFin = Utilities.formatDate(new Date(plusSix), "GMT + 1", "dd/MM/yyyy");

    if (date02 >= plusSix && rowData[0] == "") {
      sheet
        .getRange(setRow, 17)
        .setFormula(
          '=IF(B3:B="";"";JOIN("";Bilan_6mois!B43&Bilan_Sortie!B44&B3:B&Bilan_Sortie!B45&VLOOKUP(B3:B;BDD_Action!$A$2:$E;2;0)&Bilan_Sortie!B46&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;3;0);2))&Bilan_Sortie!B47&CONCATENATE(RIGHT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);4);"-";RIGHT(LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);LENB(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0))-5);2);"-";LEFT(VLOOKUP(B3:B;BDD_Action!$A$2:$E;4;0);2))&Bilan_Sortie!B48&VLOOKUP(B3:B;BDD_Action!$A$2:$E;5;0)&Bilan_Sortie!B49&D3:D&Bilan_Sortie!B50&E3:E&Bilan_Sortie!B51&F3:F&Bilan_Sortie!B52&G3:G))'
        );
    }
    //     if (link == "")
    //        break;
  }
}

/***************************************************************/
/**                                                           **/
/**   Récupère le lien du Gforms du stagiaire à 6 mois        **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
// function ResponsesLinkForm_6month() {

//     //Récupére le lien du formulaire pour validation du RF
//   var form = FormApp.openById('1NC7LtaDuqL4Yb5hL0n72dUle0vV7C0V-9AHzGWlEUc8');
//   var sheet = SpreadsheetApp.openById("19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac").getSheetByName('BDD_FORM_6mois');
//   var data = sheet.getDataRange().getValues();
//   var urlCol = 32; // Numéro de colonne où l'URL va être indiquée; A = 1, B = 2 etc
//   var responses = form.getResponses();
//   var timestamps = [], urls = [], resultUrls = [];

//   for (var i = 0; i < responses.length; i++) {

//     timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
//     urls.push(responses[i].getEditResponseUrl());
//   }
//   for (var j = 1; j < data.length; j++) {

//     resultUrls.push([data[j][0]?urls[timestamps.indexOf(data[j][0].setMilliseconds(0))]:'']);
//  }
//   //le lien du formulaire à éditer
//    sheet.getRange(2, urlCol, resultUrls.length).setValues(resultUrls);
// }

/***************************************************************/
/**                                                           **/
/**         Envoi un email au stagiaire à 6 mois              **/
/**                                                           **/
/**                 Code By Via formation                     **/
/**                                                           **/
/***************************************************************/
function sendMailOut6Month() {
  var sheet = SpreadsheetApp.openById(
    "19icdiubW-8CB6y6AnJi2iB1U90aOembbqafsuJ0_Gac"
  ).getSheetByName("BDD_6mois"); //ID Google Sheets et nom
  var dataRange = sheet.getRange("A3:Q");
  var startcolumn = 2;
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; i++) {
    var startcolumn = 3;
    var currentRow = data[i];
    var emailAddress = currentRow[7];
    var name = currentRow[3];
    var firstname = currentRow[4];
    var link = currentRow[16];
    var status = currentRow[0];
    var setRow = parseInt(i) + 3;

    if (currentRow[0] == "" && currentRow[16] != "") {
      //     if (link == "")
      //        break;
      //     if (status == "GForms envoyé")
      //        continue;
      //     if (emailAddress == "")
      //        break;

      var link_url = '<a href="' + link + '">Lien vers le formulaire</a>';
      var message =
        "<p>Bonjour " +
        firstname +
        " " +
        name +
        ",</p>" +
        "<p>Conformément aux obligations du financeur de votre parcours de formation, nous vous envoyons un questionnaire afin de connaître votre situation actuelle à 3 mois après la fin de votre formation </p>" +
        "<p> Veuillez remplir le questionnaire que vous trouverez en cliquant sur le lien suivant : " +
        link_url +
        "</p>";

      var subject = "Questionnaire de suivi de formation à 6 mois";

      //Envoyer l'email
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message,
      });
      var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
      sheet.getRange(setRow, 1).setValue("GForms envoyé le : " + date);
    }
  }
}
