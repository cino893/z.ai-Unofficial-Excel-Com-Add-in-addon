# Z.AI Excel Add-in

Dodatek do Microsoft Excel umożliwiający korzystanie z agenta AI platformy **z.ai** (Zhipu AI) bezpośrednio w arkuszu kalkulacyjnym.

## Możliwości

Agent AI może wykonywać następujące operacje na Twoim arkuszu:

| Skill | Opis |
|-------|------|
| `read_cell` | Odczyt wartości, formuły i typu z komórki |
| `write_cell` | Zapis wartości do komórki |
| `read_range` | Odczyt danych z zakresu komórek |
| `write_range` | Zapis tablicy danych od wskazanej komórki |
| `get_sheet_info` | Informacje o arkuszu (zakres, nagłówki, wymiary) |
| `get_workbook_info` | Informacje o skoroszycie (arkusze, nazwa, ścieżka) |
| `format_range` | Formatowanie (pogrubienie, kolory, ramki, wyrównanie, itp.) |
| `insert_formula` | Wstawianie formuł Excel |
| `sort_range` | Sortowanie danych |
| `add_sheet` | Dodawanie nowego arkusza |
| `delete_rows` | Usuwanie wierszy |
| `insert_rows` | Wstawianie wierszy |
| `create_chart` | Tworzenie wykresów (kolumnowy, liniowy, kołowy, itp.) |

## Wymagania

- Microsoft Excel 2016 lub nowszy (Windows)
- Klucz API z platformy [z.ai](https://open.z.ai/) (rejestracja darmowa)
- Włączony dostęp do modelu obiektów VBA (instrukcja poniżej)

## Instalacja

### Krok 1: Włącz dostęp do VBA

1. Otwórz Excel
2. **Plik** → **Opcje** → **Centrum zaufania** → **Ustawienia Centrum zaufania**
3. **Ustawienia makr** → zaznacz **Ufaj dostępowi do modelu obiektów projektu VBA**
4. Kliknij **OK**

### Krok 2: Zbuduj dodatek

Uruchom skrypt budujący (wymaga Excela na komputerze):

```
cscript build.vbs
```

Lub kliknij dwukrotnie plik `build.vbs`.

Skrypt automatycznie:
- Uruchomi Excel w tle
- Zaimportuje wszystkie moduły VBA
- Utworzy formularz czatu
- Zapisze plik `ZaiExcelAddin.xlam`

### Krok 3: Zainstaluj dodatek

1. Otwórz Excel
2. **Plik** → **Opcje** → **Dodatki**
3. Na dole: **Zarządzaj** → **Dodatki programu Excel** → **Przejdź**
4. Kliknij **Przeglądaj** i wskaż plik `ZaiExcelAddin.xlam`
5. Zaznacz **ZaiExcelAddin** i kliknij **OK**

### Alternatywnie: Instalacja ręczna

Jeśli skrypt `build.vbs` nie działa, możesz zaimportować moduły ręcznie:

1. Otwórz Excel → **Alt+F11** (edytor VBA)
2. **Plik** → **Importuj plik** (Ctrl+M)
3. Zaimportuj po kolei pliki: `modJSON.bas`, `modDebug.bas`, `modAuth.bas`, `modZaiAPI.bas`, `modExcelSkills.bas`, `modConversation.bas`, `modRibbon.bas`
4. Zapisz jako `.xlam` (**Plik** → **Zapisz jako** → typ: **Dodatek programu Excel (*.xlam)**)

## Użytkowanie

### Logowanie

1. Kliknij **Z.AI** → **Zaloguj (Klucz API)** w menu Excel
2. Wpisz swój klucz API z platformy z.ai
3. Klucz zostanie zweryfikowany i zapisany (w rejestrze Windows)

### Czat z agentem

1. Kliknij **Z.AI** → **Asystent AI (Chat)**
2. Wpisz polecenie po polsku, np.:
   - "Przeczytaj dane z komórek A1:D10"
   - "Dodaj formułę SUM do komórki E1 sumującą kolumnę D"
   - "Sformatuj wiersz 1 na pogrubiony z szarym tłem"
   - "Posortuj dane malejąco po kolumnie B"
   - "Stwórz wykres kołowy z danych A1:B5"
3. Agent automatycznie przeczyta Twój arkusz, wykona operacje i potwierdzi

### Szybkie polecenie

**Z.AI** → **Szybkie polecenie** — jednorazowe polecenie bez historii czatu.

## Debugowanie

- **Z.AI** → **Pokaż log debugowania** — otwiera plik logu w Notatniku
- **Z.AI** → **Wyczyść log** — czyści plik logu
- Logi zapisywane w: `%APPDATA%\ZaiExcelAddin\zai_debug_RRRR-MM-DD.log`
- Logowane są: żądania/odpowiedzi API, wywołania narzędzi, błędy

## Struktura projektu

```
dodatek-z-ai-opus/
├── modJSON.bas          # Parser/builder JSON dla VBA
├── modDebug.bas         # Logowanie debugowe
├── modAuth.bas          # Zarządzanie kluczem API
├── modZaiAPI.bas        # Komunikacja z API z.ai
├── modExcelSkills.bas   # 13 umiejętności (tools) do edycji Excela
├── modConversation.bas  # Pętla konwersacji z tool-calling
├── modRibbon.bas        # Menu w pasku Excel
├── frmZaiChat.frm       # Formularz czatu (backup)
├── build.vbs            # Skrypt budujący .xlam
└── README.md            # Ta dokumentacja
```

## Architektura

```
┌─────────────┐     HTTP/JSON      ┌──────────────────┐
│  z.ai API   │◄──────────────────►│   modZaiAPI.bas   │
│  (GLM-4+)   │                    └────────┬─────────┘
└─────────────┘                             │
                                   ┌────────▼─────────┐
                                   │ modConversation   │
                                   │ (tool-calling     │
                                   │  loop)            │
                                   └────────┬─────────┘
                                            │
                                   ┌────────▼─────────┐
                                   │ modExcelSkills    │──► ActiveWorkbook
                                   │ (13 narzędzi)     │
                                   └──────────────────┘
```

Agent z.ai otrzymuje definicje narzędzi (tools) w formacie OpenAI-compatible, a następnie autonomicznie decyduje które wywołać. Dodatek wykonuje te wywołania na aktywnym skoroszycie i zwraca wyniki agentowi.

## Rozwiązywanie problemów

| Problem | Rozwiązanie |
|---------|-------------|
| "Brak dostępu do VBA" | Włącz opcję w Centrum zaufania (Krok 1) |
| "HTTP 401" | Nieprawidłowy klucz API — sprawdź na open.z.ai |
| "Network error" | Sprawdź połączenie internetowe |
| Menu Z.AI nie pojawia się | Upewnij się, że dodatek jest załadowany (Opcje → Dodatki) |
| Agent nie widzi danych | Agent musi najpierw użyć `get_sheet_info` — opisz co chcesz zrobić |

## Licencja

Projekt open-source. Wykorzystuje API platformy z.ai — wymagane konto i klucz API.
