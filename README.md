# wycenniczek - Zarządzanie kosztorysem

**wycenniczek** to prosty, konsolowy program napisany w Pythonie, służący do zarządzania kosztorysami zapisanymi w plikach Excel (`.xlsx`). Umożliwia tworzenie, edytowanie, sortowanie i filtrowanie pozycji kosztorysowych, a także przeglądanie i zapisywanie danych z formatowaniem i automatycznymi kopiami zapasowymi. Program jest intuicyjny, obsługuje indexy górne (np. `m²`, `m³`) i umożliwia nawigację po folderach oraz wczytywanie plików z wiersza poleceń.

## Funkcjonalności

- **Wczytywanie i zapisywanie kosztorysów**: Obsługuje pliki `.xlsx` z predefiniowanymi kolumnami: Pozycja, Ilość, Jednostka, Cena jednostkowa (PLN), Koszt całkowity (PLN), Kategoria, Opis.
- **Dodawanie pozycji**: Umożliwia dodawanie nowych pozycji z wyborem jednostek (np. `szt`, `m²`, `godz`) i kategorii.
- **Edycja pozycji**: Intuicyjna edycja istniejących pozycji z obsługą strzałek (dzięki `prompt_toolkit`).
- **Usuwanie pozycji**: Usuwanie pozycji z potwierdzeniem.
- **Sortowanie**: Sortowanie kosztorysu po nazwie pozycji, kategorii lub koszcie (rosnąco/malejąco).
- **Filtrowanie**: Filtrowanie po kategorii lub zakresie kosztów.
- **Nawigacja po folderach**: Zmiana bieżącego katalogu i przeglądanie dostępnych plików `.xlsx`.
- **Obsługa wiersza poleceń**:
  - Wczytanie konkretnego pliku `.xlsx` (np. `Kosztorysy/projekt1.xlsx`).
  - Przejście do określonego katalogu i wyświetlenie listy plików `.xlsx`.
- **Kopie zapasowe**: Automatyczne tworzenie kopii zapasowej przy zapisie.
- **Ostrzeżenie o niezapisanych zmianach**: Pyta o potwierdzenie przed wyjściem, jeśli dane zostały zmodyfikowane.
- **Formatowanie Excela**: Zapisuje pliki z estetycznym formatowaniem (pogrubione nagłówki, obramowania, wyrównanie, format liczb).

## Wymagania

- **System operacyjny**: Linux, Windows lub macOS.
- **Python**: Wersja 3.8 lub nowsza.
- **Zależności Python**:
  - `numpy==1.26.4`
  - `pandas==1.5.3`
  - `openpyxl==3.1.2`
  - `prompt_toolkit==3.0.43`

## Przygotowanie środowiska

Poniżej znajdują się instrukcje krok po kroku, jak skonfigurować środowisko Python i uruchomić program.

### 1. Zainstaluj Pythona
Upewnij się, że masz zainstalowaną wersję Pythona 3.8 lub nowszą:
- **Linux/macOS**:
  ```bash
  python3 --version
  ```
  Jeśli Python nie jest zainstalowany, zainstaluj go, np.:
  - Ubuntu: `sudo apt update && sudo apt install python3 python3-pip`
  - macOS (z Homebrew): `brew install python3`
- **Windows**: Pobierz instalator ze strony [python.org](https://www.python.org/downloads/) i upewnij się, że opcja "Add Python to PATH" jest zaznaczona podczas instalacji.

### 2. Sklonuj lub pobierz program
Pobierz plik `wycenniczek.py` do wybranego katalogu, np. `/home/user` lub `C:\Users\YourUser`.

### 3. Utwórz i aktywuj wirtualne środowisko
Wirtualne środowisko pozwala izolować zależności programu.

- **Linux/macOS**:
  ```bash
  cd /home/user
  python3 -m venv venv
  source venv/bin/activate
  ```
- **Windows**:
  ```cmd
  cd C:\Users\YourUser
  python -m venv venv
  venv\Scripts\activate
  ```

Po aktywacji powinieneś zobaczyć `(venv)` w wierszu poleceń.

### 4. Zainstaluj zależności
W aktywnym środowisku wirtualnym zainstaluj wymagane biblioteki:
```bash
pip install numpy==1.26.4 pandas==1.5.3 openpyxl==3.1.2 prompt_toolkit==3.0.43
```

Sprawdź, czy zależności są poprawnie zainstalowane:
```bash
pip list
```

## Użycie programu

Program można uruchomić w wierszu poleceń, podając ścieżkę do pliku `.xlsx`, katalogu lub bez argumentów.

### Uruchomienie z wiersza poleceń
- **Wczytanie konkretnego pliku `.xlsx`**:
  ```bash
  python wycenniczek.py Kosztorysy/projekt1.xlsx
  ```
  lub z cudzysłowami:
  ```bash
  python wycenniczek.py "Kosztorysy/projekt1.xlsx"
  ```
  Program wczyta plik, wyświetli jego zawartość i przejdzie do menu głównego.

- **Przejście do katalogu**:
  ```bash
  python wycenniczek.py Kosztorysy
  ```
  lub:
  ```bash
  python wycenniczek.py "Kosztorysy"
  ```
  Program zmieni katalog na wskazany i wyświetli listę plików `.xlsx` do wyboru.

- **Uruchomienie w trybie interaktywnym**:
  ```bash
  python wycenniczek.py
  ```
  Program wyświetli pliki `.xlsx` w bieżącym katalogu i pozwoli wybrać plik lub utworzyć nowy kosztorys.

### Przykładowe użycie
1. Uruchom program z plikiem:
   ```bash
   python wycenniczek.py Kosztorysy/projekt1.xlsx
   ```
   Wyjście:
   ```
   === Witaj w programie wycenniczek! ===
   Bieżący folder: /home/user/Kosztorysy
   Kosztorys wczytany z pliku: /home/user/Kosztorysy/projekt1.xlsx

   === Aktualny kosztorys ===
   Nr Pozycja  Ilość Jednostka  Cena jednostkowa (PLN)  Koszt całkowity (PLN)  Kategoria        Opis
    1  Kamera      4       szt                    200.00                 800.00    Sprzęt    Kamera 4K
    2   Kable    100         m                      5.00                 500.00 Materiały   Kabel UTP
     Łączny koszt: 1300.00 PLN

   === wycenniczek - Zarządzanie kosztorysem ===
   ...
   ```

2. Uruchom program z katalogiem:
   ```bash
   python wycenniczek.py Kosztorysy
   ```
   Wyjście:
   ```
   === Witaj w programie wycenniczek! ===
   Bieżący folder: /home/user/Kosztorysy
   Wybierz plik kosztorysu lub utwórz nowy.

   Dostępne pliki Excel w folderze /home/user/Kosztorysy (posortowane według daty modyfikacji):
     1. projekt1.xlsx (zmodyfikowany: [data modyfikacji])

   Wpisz numer pliku, Enter dla nowego kosztorysu lub 'q' aby anulować:
   ```

3. Wybierz opcje w menu głównym (1-10) do edycji, sortowania, filtrowania itp.

## Struktura pliku Excel
Program oczekuje plików `.xlsx` z kolumnami:
- **Pozycja**: Nazwa pozycji (np. "Kamera").
- **Ilość**: Liczba (np. 4).
- **Jednostka**: Jednostka miary (np. `szt`, `m²`).
- **Cena jednostkowa (PLN)**: Cena za jednostkę (np. 200.00).
- **Koszt całkowity (PLN)**: Ilość × Cena jednostkowa.
- **Kategoria**: Kategoria pozycji (np. "Sprzęt").
- **Opis**: Dodatkowy opis (opcjonalny).

## Uwagi
- **Polskie znaki i jednostki**: Program obsługuje znaki takie jak `m²`, `m³`. W systemie Windows ustaw kodowanie konsoli na UTF-8:
  ```cmd
  chcp 65001
  ```
- **Ścieżki z spacjami**: Używaj backslashy (`\`) lub cudzysłowów (`"`) dla ścieżek z spacjami.
- **Kopie zapasowe**: Przy zapisie tworzony jest plik backup z sygnaturą czasową (np. `backup_20250815_183000_wycenniczek.xlsx`).
- **Sugerowane ulepszenia**:
  - Autouzupełnianie ścieżek w wierszu poleceń.
  - Skontaktuj się z twórcą, jeśli potrzebujesz dodatkowych funkcji!

## Rozwiązywanie problemów
- **Błąd brakujących zależności**: Upewnij się, że wszystkie wymagane biblioteki są zainstalowane (`pip list`).
- **Błąd wczytywania pliku**: Sprawdź, czy plik `.xlsx` istnieje i ma poprawną strukturę kolumn.
- **Problemy z kodowaniem w Windows**: Uruchom `chcp 65001` przed startem programu.
