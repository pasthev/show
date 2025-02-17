// (Google Apps Script)
// This is the source code of the Show extension I published in 2016 in the Google Marketplace
//
// Show extension was used to display non-printable characters within a Google Docs document,
// it is now useless as Google finally implemented the function natively to Google Docs.
//
//
//
//
// Global variables
//
var body = DocumentApp.getActiveDocument().getBody();
var nbsp = '\u00A0';                                                   // Espace insécable, No Break Space
var nnbsp = '\u202F';                                                  // Espace courte insécable, Narrow No Break Space
// 
//
function pagebreaks_reveal_click(element) {
  pagebreaks_reveal();
  update_highlight(highlightOption());
}
//
function linebreaks_reveal_click(element) {
  linebreaks_reveal();
  update_highlight(highlightOption());
}
//
function unbreak_spaces_reveal_click() {
  unbreak_spaces_reveal();
  update_highlight(highlightOption());
}
//
function tabs_reveal_click() {
  tabs_reveal();
  update_highlight(highlightOption());
}
//
function spaces_reveal_click() {
  spaces_reveal();
  update_highlight(highlightOption());
}
//
function all_reveal() {
  unbreak_spaces_reveal();
  tabs_reveal();
  pagebreaks_reveal();
  linebreaks_reveal();
  update_highlight(highlightOption());
}
//
function all_hide() {
  unbreak_spaces_hide();
  tabs_hide();
  spaces_hide();
  pagebreaks_hide();
  linebreaks_hide();
}
//
function unbreak_spaces_reveal() {
  body.replaceText(nbsp + '{1}', '\u25CB');                            // Espace insécable -> ○ (nbsp + '' seul foire lamentablement)
  body.replaceText(nnbsp + '{1}', '\u25E6');                           // Espace fine insécable -> ◦ 
}
//
function unbreak_spaces_hide() {
  highlight_text('\u25CB', '#FFFFFF');                                 // Supprimer d'abord le surlignement
  highlight_text('\u25E6', '#FFFFFF');                                 // Supprimer d'abord le surlignement
  body.replaceText('\u25CB', nbsp);                                    // ○ -> Espace insécable
  body.replaceText('\u25E6', nnbsp);                                   // ◦ -> Espace fine insécable
}
//
function tabs_reveal() {
  body.replaceText('\u0009{1}', '\uFFEB\u0009');                       // Tabulation horizontale -> ￫
  body.replaceText('\u000B{1}', '\uFFEC\u000B');                       // Tabulation verticale  ->  ￬
}
//
function tabs_hide() {
  body.replaceText('\uFFEB', '');                                      // ￫ -> ''
  body.replaceText('\uFFEC', '');                                      // ￬ -> ''
}
//
function spaces_reveal() {
  body.replaceText(' {1}', '\u25AA');                                  // Espace -> ▪ (' ' seul foire lamentablement)
}
//
function spaces_hide() {
  highlight_text('\u25AA', '#FFFFFF');                                 // Supprimer d'abord le surlignement
  body.replaceText('\u25AA', ' ');                                     // ▪ -> Espace
}
//
function pagebreaks_reveal(element) {
  //
  if (!element) {                                                      // Si la fonction est appelée directement, et non en récursif
    element = DocumentApp.getActiveDocument().getBody();               //   Remplace l'argument vide par le document en cours
  }                                                                    //
  //
  if (element.getType() == DocumentApp.ElementType.PAGE_BREAK) {       // Si l'élément est un page break
    element.getParent().asText().appendText('\u21A1');                 //     ↡ Ajoute la marque de saut de page
  } else if (element.getNumChildren) {                                 // Si l'élément actuel a des objets enfants
    for (var i = element.getNumChildren() - 1; i >= 0; i--) {          //    Lis un à un tous les enfants
      pagebreaks_reveal(element.getChild(i));                          //    Appel récursif pour vérifier les objets enfants
    }                                                                  //
  }
}
//
function pagebreaks_hide() {
  body.replaceText('\u21A1', '');                                      // ↡ -> ''
}
//
function linebreaks_reveal(element) {
  //
  if (!element) {                                                      // Si la fonction est appelée directement, et non en récursif
    element = DocumentApp.getActiveDocument().getBody();               //   Remplace l'argument vide par le document en cours
  }                                                                    //
  //
  if (element.getType() == DocumentApp.ElementType.PARAGRAPH           // Si l'élément est un paragraphe
      && element.asParagraph().getText().replace(/\s/g, '') == '') {   //   et si ce paragraphe n'est pas blanc,
    element.asParagraph().appendText('\u2761');                        //     ❡ Ajoute la marque de paragraphe courbée pour la fin de section
    //
  } else if (element.getType() == DocumentApp.ElementType.TEXT) {      // Si l'élément est un texte
    var text = element.asText();                                       //   Lecture du texte
    var content = text.getText();                                      //   Lecture du contenu du texte
    var regexp = /\r/g;                                                //   Carriage return (\x0D) correspond à Ctrl-Entrée
    var pbsearch = /\\u000C/g;
    var resultsTable;                                                  //   Pour stocker résultats d'Exec
    var offset = 0;                                                    //   Compense le fait que \r = 0 espace
    while ((resultsTable = regexp.exec(content)) !== null) {           //   Tant que l'on trouve des \r
      text.insertText(regexp.lastIndex-1+offset, '\u23CE');            //     ⏎ Ajoute la marque de retour chariot
      offset++;                                                        //     Compense le décalage
    }                                                                  //    Fin de boucle
    content = text.getText();                                          //    Capture le texte de nouveau
    text.appendText('\u00B6');                                         //    ¶ Ajoute la marque de paragraphe à la fin du bloc texte
    //
  } else if (element.getNumChildren) {                                 // Si l'élément actuel a des objets enfants
    for (var i = element.getNumChildren() - 1; i >= 0; i--) {          //    Lis un à un tous les enfants
      var child = element.getChild(i);                                 //    Pour les passer ensuite en argument
      linebreaks_reveal(child);                                        //    Appel récursif pour vérifier les objets enfants
    }                                                                  //
  }                                                                    //
}                                                                      //
//
function linebreaks_hide() {
  body.replaceText('\u23CE', '');                                      // Supprime ⏎ u23CE Return symbol
  body.replaceText('\u00B6', '');                                      // Supprime ¶ u00B6 pilcrow sign
  body.replaceText('\u2761', '');                                      // Supprime ❡ u2761 curved stem paragraph
}
//
function symbols_highlight_click() {
//
  var documentProperties = PropertiesService.getDocumentProperties();
  var highlight = highlightOption();       // 
  highlight ^= 1;       // 
  try {
    documentProperties.setProperty('highlightSymbols', highlight);    // Only strings in properties service
  } catch (e) {
    DocumentApp.getUi().alert(e.message);
  }
  onOpen();
  symbols_highlight(highlight);
}
//
function update_highlight(option) {                                   // Appelle symbols_highlight() 
  if (option == true) {symbols_highlight(true)};                      // uniquement si l'option est activée
}
//
function symbols_highlight(option) {
  //
  var color = '#FFFFFF';
  if (option == true) {color = '#AAFFAA'};
  //
  highlight_text('\u23CE', color);                                    // Surligne ⏎
  highlight_text('\u00B6', color);                                    // Surligne ¶
  highlight_text('\u2761', color);                                    // Surligne ❡
  highlight_text('\u21A1', color);                                    // Surligne ↡
  highlight_text('\u25AA', color);                                    // Surligne ▪
  highlight_text('\uFFEB', color);                                    // Surligne ￫
  highlight_text('\uFFEC', color);                                    // Surligne ￬
  highlight_text('\u25CB', color);                                    // Surligne ○
  highlight_text('\u25E6', color);                                    // Surligne ◦
}
//
function highlight_text(target,background) {
  //
  var count = 0;
  var background = background;
  var searchResult = body.findText(target);
  //
  while (searchResult) {
    var thisElement = searchResult.getElement();
    var text = thisElement.asText();
    text.setBackgroundColor(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive(),background);
    count++;
    searchResult = body.findText(target, searchResult);
  }
  return count;
}
//
function highlightOption() {
    var documentProperties = PropertiesService.getDocumentProperties();
    var highlight = documentProperties.getProperty('highlightSymbols');
    if (highlight == null) {
      highlight = false
      documentProperties.setProperty('highlightSymbols', highlight);    // Only strings in properties service
    }
  return highlight
}
//
function show_about() {
  DocumentApp.getUi().alert('Show Extension' + '\n Pascal - 2016' + '\n pasgoude@gmail.com');
}
//
function onInstall(e) {
  onOpen(e);
}
//
function onOpen(e) {
  //

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // (things that work in all authorization modes belong to here).
    createInstallMenu();
  } else {
    // (things that don't work in AuthMode.NONE belong to here).
    createMenus();
  }
}
function createMenus() {
  UserLang=Session.getActiveUserLocale()
  //UserLang='fr'
  switch (UserLang) {
    case 'fr':
      var typo = 'Caractères masqués';
      var reveal = 'Révéler';
      var mask = 'Masquer';
      var revealAll = 'Tout révéler';
      var maskAll = 'Tout masquer';
      var all = 'Tous';
      var unbreak = 'Espaces insécables';
      var spaces = 'Espaces';
      var tabs = 'Tabulations';
      var pgbreaks = 'Sauts de page';
      var linefeed = 'Sauts de ligne';
      var shighlight = 'Surligner les symboles';
      var about = 'À propos';
      break;
    default:                                                            // var typo = LanguageApp.translate'Typography', 'en', UserLang;
      var typo = 'Hidden Characters';
      var reveal = 'Show';
      var mask = 'Hide';
      var revealAll = 'Show all';
      var maskAll = 'Hide all';
      var all = 'Tous';
      var unbreak = 'Non-breakable spaces';
      var spaces = 'Spaces';
      var tabs = 'Tabs';
      var pgbreaks = 'Page breaks';
      var linefeed = 'Line breaks';
      var shighlight = 'Highlight symbols';
      var about = 'About';
  }
  //
  if (highlightOption() == true) {
    shighlight = '✓ ' + shighlight                                      // checkmark \u2713
  }
  var ui = DocumentApp.getUi();                                         // If the script is published as an add-on, the caption parameter is ignored 
  ui.createMenu(typo)                                                   // and the menu is added as a sub-menu of the Add-ons menu, equivalent to createAddonMenu().
  .addItem(revealAll, 'all_reveal')
  .addItem(maskAll, 'all_hide')
  .addSubMenu(ui.createMenu(reveal)
              .addItem(unbreak, 'unbreak_spaces_reveal_click')
              .addItem(spaces, 'spaces_reveal_click')
              .addItem(tabs, 'tabs_reveal_click')
              .addItem(pgbreaks, 'pagebreaks_reveal_click')
              .addItem(linefeed, 'linebreaks_reveal_click')
              .addSeparator())
  .addSubMenu(ui.createMenu(mask)
              .addItem(unbreak, 'unbreak_spaces_hide')
              .addItem(spaces, 'spaces_hide')
              .addItem(tabs, 'tabs_hide')
              .addItem(pgbreaks, 'pagebreaks_hide')
              .addItem(linefeed, 'linebreaks_hide'))
  .addSeparator()
  .addItem(shighlight, 'symbols_highlight_click')
  .addSeparator()
  .addItem(about, 'show_about')
  .addToUi();
}
function createInstallMenu() {
  //
  var ui = DocumentApp.getUi();                                         // If the script is published as an add-on, the caption parameter is ignored 
  ui.createMenu('Show')                                                   // and the menu is added as a sub-menu of the Add-ons menu, equivalent to createAddonMenu().
  .addItem('Activate Show extension', 'activateShow')
  .addToUi();
}
//
function activateShow() {
  // var UserLang=Session.getActiveUserLocale();                       // would be nince but forbidden in authmode.NONE
  var UserLang='en';
  switch (UserLang) {
    case 'fr':
      var text = 'Show - Affichage des caractères masqués';
      var prompt = 'Activer l\'extension pour ce document ?';
    default:
      var text = 'Show - Hidden Characters';
      var prompt = 'Activate extension for this document?';
  }
  var ui = DocumentApp.getUi();
  var response = ui.alert('Show ¶', text + '\n' +
                          '\n Pascal - 2016' +
                          '\n pasgoude@gmail.com\n\n' + prompt,
                          ui.ButtonSet.YES_NO);
  //  DocumentApp.getUi().alert('Show ¶' +
  //                            '\n' + text +
  //                            '\n Pascal - 2016' +
  //                            '\n pasgoude@gmail.com');
 if (response == ui.Button.YES) {
   createMenus();
 }
}
