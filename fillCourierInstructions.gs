/**
 * Skrypt do automatycznego wypełniania dokumentu instrukcji dla kuriera Google Docs danymi z arkusza Google Sheets.
 * Zastępuje linie po "TYLKO I WYŁĄCZNIE" imionami i nazwiskami z zaznaczonych komórek.
 */
function fillCourierInstructions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ranges = sheet.getActiveRangeList().getRanges(); // Pobiera listę zaznaczonych zakresów
  var values = [];

  // Przetwarzaj każdy zaznaczony zakres
  for (var i = 0; i < ranges.length; i++) {
    var rangeValues = ranges[i].getValues(); // Pobiera wartości z każdego zakresu
    values = values.concat(rangeValues); // Dodaje wartości do listy
  }

  // Uzupełnij wartości spacjami, jeśli zaznaczono mniej niż 4 komórki
  while (values.length < 4) {
    values.push([" "]); // Dodaj spację zamiast pustego tekstu
  }

  var doc = DocumentApp.openById(
    "YOUR_DOCUMENT_ID" // Wstaw tutaj ID swojego dokumentu, które znajdziesz w linku Google Docs, np. https://docs.google.com/document/d/YOUR_DOCUMENT_ID/edit
  );
  var body = doc.getBody();
  var paragraphs = body.getParagraphs();

  var namesIndex = 0;

  // Przetwarzaj akapity w dokumencie
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraphText = paragraphs[i].getText().trim();

    // Znajdź akapit z "TYLKO I WYŁĄCZNIE"
    if (paragraphText === "TYLKO I WYŁĄCZNIE") {
      if (namesIndex < values.length) {
        var entry = values[namesIndex][0].trim();

        // Rozdziel imię i nazwisko na podstawie spacji lub kropki
        var nameParts = entry.split(/\.|\s+/); // Dzielimy na kropkę lub spację
        var firstName = nameParts[0] || ""; // Pierwszy element to nazwisko
        var lastName = nameParts[1] || ""; // Drugi element to imię

        // Łączymy nazwisko i imię, jeśli oba istnieją
        var fullName = (firstName + " " + lastName).trim();

        // Pomiń pusty akapit po "TYLKO I WYŁĄCZNIE"
        var targetIndex = i + 1;
        if (
          targetIndex < paragraphs.length &&
          paragraphs[targetIndex].getText().trim() === ""
        ) {
          targetIndex++; // Skaczemy do następnego akapitu
        }

        // Jeśli znaleźliśmy akapit, zamieniamy treść
        if (targetIndex < paragraphs.length) {
          var targetParagraph = paragraphs[targetIndex];

          // Jeśli wartość jest pusta, wstaw spację
          if (fullName === "") {
            targetParagraph.setText(" "); // Wstaw spację zamiast pustego tekstu
          } else {
            targetParagraph.setText(fullName);
          }

          // Ustawienie pogrubienia i usunięcie podkreślenia
          targetParagraph.editAsText().setBold(true).setUnderline(false);
        }

        namesIndex++;
      }
    }
  }

  // Zapisuje dokument
  doc.saveAndClose();
}
