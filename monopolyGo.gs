function dispoCartes() {
  var app = SpreadsheetApp;
  var feuille = app.getActiveSpreadsheet().getActiveSheet();
  var nomCarte;
  var qteCarte;
  var rangeQte;
  var rechCarte;
  var carteDoree;

  var lastRow = feuille.getLastRow();

  for (var i=2; i<=lastRow; i++) {
    carteDoree = feuille.getRange(i, 4).getValue();
    qteCarte = feuille.getRange(i, 3).getValue();
    rechCarte = feuille.getRange(i, 5).getValue();
    nomCarte = feuille.getRange(i, 1);
    rangeQte = feuille.getRange(i, 3);

    if (carteDoree === 1) {
      nomCarte.setBackground('#c89c13');
    } else {
      if (rechCarte === 1) {
        nomCarte.setBackground('#d51010');
      } else {
        if (qteCarte >= 1) {
          nomCarte.setBackground('#12d90c');
          rangeQte.setBackground('#12d90c');
        } else {
          nomCarte.setBackground('#605f5f');
        }
      }
    }
  }
}

function onSelectionChange(e) {
  var app = SpreadsheetApp;
  var feuille = app.getActiveSpreadsheet().getActiveSheet();
  var editCell = e.range.getRow();
  var editCellCol = e.range.getColumn();

  if (editCell >= 2 && editCellCol <= 5) {
    carteDoree = feuille.getRange(editCell, 4).getValue();
    qteCarte = feuille.getRange(editCell, 3).getValue();
    rechCarte = feuille.getRange(editCell, 5).getValue();
    nomCarte = feuille.getRange(editCell, 1);
    rangeQte = feuille.getRange(editCell, 3);

    if (carteDoree === 1) {
      nomCarte.setBackground('#c89c13');
    } else {
      if (rechCarte === 1) {
        nomCarte.setBackground('#d51010');
      } else {
        if (qteCarte >= 1) {
          nomCarte.setBackground('#12d90c');
          rangeQte.setBackground('#12d90c');
        } else {
          nomCarte.setBackground('#605f5f');
          rangeQte.setBackground('#ffffff');
        }
      }
    }
  }
}
