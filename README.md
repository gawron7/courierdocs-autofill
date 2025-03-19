## Courierdocs Autofill

Skrypt Google Apps Script do automatycznego wypełniania dokumentów Google Docs danymi z arkusza Google Sheets.

## Jak to działa?

Skrypt zastępuje linie po "TYLKO I WYŁĄCZNIE" imionami i nazwiskami z zaznaczonych komórek w arkuszu Google Sheets.

## Instalacja

Skopiuj zawartość skryptu fillCourierInstructions.gs.

1. Otwórz arkusz Google Sheets z danymi.
2. Przejdź do zakładki Rozszerzenia -> Apps Script.
3. Po otworzeniu Apps Script wklej gotowy kod i podmień ID dokumentu Google Docs w tej linijce.:
   var doc = DocumentApp.openById(
   "YOUR_DOCUMENT_ID" // Wstaw tutaj ID swojego dokumentu, które znajdziesz w linku Google Docs, np. https://docs.google.com/document/d/YOUR_DOCUMENT_ID/edit
   );
4. Zapisz.

https://github.com/user-attachments/assets/f4e7d2c1-dec6-4282-a6dd-8dfa55d320de

5. Przejdź z powrotem do swojego arkusza Google Sheets.
6. Przejdź do zakładki Rozszerzenia -> Makra -> Zaimportuj makro / Zarządzaj makrami.
7. Po pojawieniu się okienka, importuj skrypt fillCourierInstructions, kliknij przycisk Dodaj funkcję.
8. Gotowy przycisk skryptu powinien pojawić się w Rozszerzenia -> Makra -> fillCourierInstructions.

---

9. Aby zmienić nazwę lub dodać skrót klawiszowy do makra, przejdź do Rozszerzenia -> Makra -> Zarządzaj makrami.
10. Po pojawieniu się okienka zmień nazwę fillCourierInstructions na własną, np. Wklej instrukcje dla kuriera, lub dodaj skrót klawiszowy, np. Ctrl + Alt + Shift + "twój_klawisz", i kliknij Aktualizuj.

https://github.com/user-attachments/assets/eede47b0-f7bf-4b8b-adfe-1b5b65c74285

## Uwaga

Przy pierwszym uruchomieniu będziemy musieli autoryzować skrypt.

1. Przejdź do zakładki Rozszerzenia -> Makra -> fillCourierInstructions / Wklej instrukcje dla kuriera.
2. Przy pierwszym uruchomieniu pojawi się okienko z wymaganą autoryzacją do skryptu, kliknij OK.
3. Wybierz swoje konto Google.
4. Otworzy się nowa karta z ostrzeżeniem, że ta aplikacja nie została zweryfikowana przez Google.
5. Kliknij Zaawansowane -> Otwórz: Projekt (niebezpieczne).

https://github.com/user-attachments/assets/04a6198e-6d6b-480f-9e14-b6a8e8a8e452

## Jak używać?

1. Zaznacz od 1 do 4 kolumn w arkuszu, z których chcesz skopiować dane. (Shift + zaznaczenie lewym przyciskiem myszy kolumny od której do której chcesz zaznaczyć lub Ctrl + lewy przycisk myszy, aby wybrać poszczególne kolumny).
2. Jeśli chcesz wybrać tylko jedną kolumnę, zaznacz ją oraz kolumnę obok (nie da się zaznaczyć tylko jednej kolumny, skrypt nie odczyta danych).
3. Przejdź do zakładki Rozszerzenia -> Makra -> fillCourierInstructions / Wklej instrukcje dla kuriera.
4. Gotowe :)

https://github.com/user-attachments/assets/eb7c40a9-ccbf-497c-a8f2-5a90f423a7d4
