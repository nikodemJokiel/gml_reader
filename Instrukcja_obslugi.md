# Instrukcja obsługi: Konwerter GML (RCiWN) do Excela

Poniższa instrukcja krok po kroku wyjaśnia, jak przygotować środowisko i uruchomić program na całkowicie "czystym" komputerze. Przeznaczona jest dla osób nietechnicznych.

---

## CZĘŚĆ 1: Pierwsze uruchomienie (Konfiguracja komputera)

Te kroki wykonujesz **tylko raz** na danym komputerze.

### Krok 1: Instalacja języka Python
Program został napisany w języku Python. Aby komputer mógł go zrozumieć, musisz go najpierw zainstalować.
1. Wejdź na oficjalną stronę: [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Kliknij duży żółty przycisk **"Download Python"**.
3. Uruchom pobrany plik instalacyjny.
4. **BARDZO WAŻNE:** Na samym dole okienka instalatora, zanim klikniesz "Install Now", **ZAZNACZ POLE wyboru o nazwie `Add Python.exe to PATH`** (lub podobnie brzmiące). Jeśli tego nie zrobisz, komputer "nie będzie widział" Pythona.
5. Kliknij **Install Now** i poczekaj na zakończenie instalacji, a następnie zamknij okno.

### Krok 2: Instalacja dodatkowych bibliotek
Program potrzebuje kilku paczek danych (np. do obsługi plików Excel), które musimy pobrać.
1. Wejdź do folderu `gml_reader` (tego, w którym znajduje się plik `gml_reader.py`).
2. Kliknij **lewym przyciskiem myszy w pasek adresu** na samej górze okna folderu (tam, gdzie wyświetla się ścieżka np. `C:\Users\Nazwa\gml_reader`).
3. Skasuj to, co tam jest, wpisz słowo **`cmd`** i wciśnij **Enter**.
4. Otworzy się czarne okno konsoli (wiersz polecenia).
5. Wpisz w nim poniższą komendę i wciśnij Enter:
   ```bash
   pip install -r requirements.txt
   ```
6. Komputer pobierze potrzebne pliki. Poczekaj, aż proces się zakończy (gdy znów pojawi się na dole linijka ze ścieżką do folderu gotowa do wpisywania). Możesz teraz zamknąć to czarne okno.

---

## CZĘŚĆ 2: Codzienne korzystanie z programu

Kiedy masz już zainstalowanego Pythona oraz biblioteki, konwertowanie plików wygląda następująco:

1. **Przygotowanie plików:** 
   Przekopiuj pobrany plik GML (np. `moje_dane.gml`) do folderu `gml_reader`. W tym samym folderze powinien znajdować się plik `TERC.xml` (odpowiada za nazwy gmin) oraz opcjonalnie `Obreby.dbf` (odpowiada za nazwy obrębów).

2. **Otwieranie wiersza poleceń (cmd):**
   Wejdź do folderu `gml_reader`. Kliknij w górny pasek adresu, wpisz **`cmd`** i wciśnij **Enter**.

3. **Uruchomienie programu:**
   W czarnym oknie wpisz komendę, podając nazwę swojego pliku GML:
   ```bash
   python gml_reader.py moje_dane.gml
   ```
   Wciśnij **Enter**. Program zacznie działać i na ekranie pojawią się informacje o przebiegu (np. ile znalazł transakcji).

4. **Wynik:**
   Program utworzy plik Excel o **tej samej nazwie** co plik GML (np. `moje_dane.xlsx`).
   
   Plik Excel zawiera **3 arkusze**:
   - **Działki** — transakcje dotyczące wyłącznie działek (bez budynków i lokali)
   - **Budynki** — transakcje z budynkami (ale bez lokali)
   - **Lokale** — transakcje z lokalami (nawet jeśli w transakcji jest też budynek)

---

## CZĘŚĆ 3: Rozwiązywanie najczęstszych problemów

* **Błąd `PermissionError: [Errno 13] Permission denied: '....xlsx'`**
  **Przyczyna:** Program próbuje zapisać plik Excel, ale plik o tej nazwie jest obecnie **otwarty** w programie Microsoft Excel na Twoim komputerze.
  **Rozwiązanie:** Zamknij plik w programie Excel i ponownie uruchom program.

* **Błąd `python nie jest rozpoznawany jako polecenie wewnętrzne...`**
  **Przyczyna:** Podczas pierwszego uruchomienia zapomniano zaznaczyć "Add Python to PATH" (Krok 1).
  **Rozwiązanie:** Należy odinstalować program Python w systemie i zainstalować go ponownie, tym razem pamiętając o zaznaczeniu tego okienka.

* **Brak polskich nazw gmin / puste miejsca**
  **Przyczyna:** Brak pliku `TERC.xml` w folderze `gml_reader`.
  **Rozwiązanie:** Upewnij się, że plik `TERC.xml` leży dokładnie obok pliku `gml_reader.py`. Jeśli go nie będzie, program zadziała, ale w tabeli wypisze surowe kody TERYT.

* **Brak nazw obrębów / puste miejsca w kolumnie Obręb**
  **Przyczyna:** Brak pliku `Obreby.dbf` (lub `Obręby.dbf`) w folderze `gml_reader`.
  **Rozwiązanie:** Umieść plik `Obreby.dbf` w folderze ze skryptem. Jeśli pliku nie ma, kolumny z obrębem pozostaną puste.
