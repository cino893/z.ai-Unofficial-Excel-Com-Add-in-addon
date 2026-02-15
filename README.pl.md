# Z.AI Dodatek do Excela

[🇬🇧 English version](README.md)

Nieoficjalny dodatek do Microsoft Excel integrujący platformę **[Z.AI](https://z.ai)** (Zhipu AI) — czatuj z agentem AI, który czyta, pisze, formatuje, tworzy wykresy i automatyzuje arkusze kalkulacyjne.

> ⚠️ **Uwaga:** To jest nieoficjalny dodatek społecznościowy. Nie jest powiązany z Zhipu AI / Z.AI ani przez nich wspierany.

![Demo](show-reel.gif)

## Pobierz

- **[⬇ Pobierz najnowszy .xll](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases/latest/download/ZaiExcelAddin-AddIn64-packed.xll)**
- [Wszystkie wydania](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases)

## Instalacja

1. Pobierz plik `.xll` powyżej
2. Otwórz Excel → **Plik** → **Opcje** → **Dodatki**
3. Na dole: **Zarządzaj** → **Dodatki programu Excel** → **Przejdź…**
4. Kliknij **Przeglądaj** i wskaż pobrany plik `ZaiExcelAddin-AddIn64-packed.xll`
5. Zatwierdź — zakładka **Z.AI** pojawi się na wstążce

> 📖 Potrzebujesz instrukcji ze zrzutami ekranu? Zobacz [Jak dodać dodatek do Excela (Microsoft Support)](https://support.microsoft.com/pl-pl/office/dodawanie-lub-usuwanie-dodatk%C3%B3w-w-programie-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460).

### Wymagania

- Microsoft Excel 2016+ (Windows, zalecany 64-bit)
- [.NET 8.0 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)
- Klucz API Z.AI — [uzyskaj tutaj](https://z.ai/manage-apikey/apikey-list) (darmowy tier dostępny)

## Użytkowanie

1. **Login** — kliknij **Z.AI → Login**, wklej klucz API (strona [zarządzania kluczami](https://z.ai/manage-apikey/apikey-list) otworzy się automatycznie)
2. **Chat** — kliknij **💬 Chat** aby otworzyć panel boczny; poproś AI o pracę z arkuszem
3. **Model** — kliknij **Model** aby wybrać z 12 modeli (od darmowych ⚡ po premium 💎)
4. **Język** — kliknij **Language** aby zmienić język interfejsu (PL, EN, DE, FR, ES, UK, ZH, JA)
5. **Doładuj** — otwiera [stronę płatności](https://z.ai/manage-apikey/billing) do doładowania salda

### Przykładowe polecenia

- *"Przeczytaj dane z A1:D10 i podsumuj je"*
- *"Dodaj formułę SUM do E1"*
- *"Sformatuj nagłówki na pogrubione z zielonym tłem"*
- *"Stwórz wykres kołowy z A1:B5"*
- *"Posortuj po kolumnie C malejąco"*

## Funkcje

- **Panel boczny z czatem AI** — Custom Task Pane po prawej stronie Excela
- **Interfejs WPF** — dymki czatu, zielony motyw Excela, animacje, emoji
- **15 narzędzi Excel** — AI czyta/pisze komórki, formatuje, tworzy wykresy, sortuje
- **12 modeli** — pełen katalog z cenami (darmowe modele flash w zestawie)
- **8 języków** — automatyczne wykrywanie z ustawień Windows
- **Wykrywanie pętli** — AI nie powtarza tych samych operacji w nieskończoność
- **Dedykowana zakładka Ribbon** — logowanie, wybór modelu, język, saldo, logi, informacje

### Narzędzia AI

| Narzędzie | Opis |
|-----------|------|
| `read_cell` / `write_cell` | Odczyt/zapis komórki |
| `read_range` / `write_range` | Odczyt/zapis zakresu (tablice 2D) |
| `get_sheet_info` | Wymiary arkusza, nagłówki, zakres użyty |
| `get_workbook_info` | Arkusze w skoroszycie, ścieżka pliku |
| `format_range` | Czcionka, kolory, ramki, wyrównanie, scalanie |
| `insert_formula` | Wstawianie formuł Excel |
| `sort_range` | Sortowanie danych po kolumnie |
| `add_sheet` | Dodawanie arkusza |
| `delete_rows` / `insert_rows` | Usuwanie/wstawianie wierszy |
| `create_chart` | Tworzenie wykresów (kolumnowy, słupkowy, liniowy, kołowy, punktowy, obszarowy) |
| `delete_chart` / `list_charts` | Usuwanie lub lista wykresów |

## Budowanie ze źródeł

Wymagany .NET SDK 8.0+:

```powershell
cd src
dotnet build -c Release
```

Wynik: `src\bin\Release\net8.0-windows\publish\ZaiExcelAddin-AddIn64-packed.xll`

## Struktura projektu

```
dodatek-z-ai-opus/
├── src/                            # .NET 8 COM Add-in (ExcelDNA)
│   ├── ZaiExcelAddin.csproj        # Projekt C#
│   ├── AddIn.cs                    # Punkt wejścia (IExcelAddIn)
│   ├── RibbonController.cs         # Wstążka + Custom Task Pane
│   ├── Models/
│   │   └── ChatMessage.cs          # Model wiadomości czatu
│   ├── Services/
│   │   ├── AuthService.cs          # Klucz API (rejestr Windows)
│   │   ├── ConversationService.cs  # Pętla tool-calling + wykrywanie pętli
│   │   ├── DebugLogger.cs          # Logowanie do pliku
│   │   ├── ExcelSkillService.cs    # 15 narzędzi Excel
│   │   ├── I18nService.cs          # 8 języków
│   │   └── ZaiApiService.cs        # Klient HTTP Z.AI + katalog modeli
│   └── UI/
│       ├── ChatPanel.xaml/.cs      # Panel czatu WPF
│       ├── ChatPaneHost.cs         # Host WinForms dla CTP (COM-visible)
│       ├── WpfLoginDialog.xaml/.cs # Dialog logowania WPF
│       └── WpfSelectDialog.xaml/.cs# Dialog wyboru WPF
├── legacy/                         # v1.0 VBA (zdeprecjonowany)
├── show-reel.gif                   # Animacja demo
├── dodatek-z-ai-opus.sln          # Plik solution
└── README.md
```

## Architektura

```
┌──────────────┐    HTTP/JSON     ┌──────────────────┐
│   Z.AI API   │◄───────────────►│   ZaiApiService   │
│  (modele GLM)│                 └────────┬─────────┘
└──────────────┘                          │
                                ┌─────────▼─────────┐
                                │ ConversationService │  pętla tool-calling
                                │  (max 15 rund,      │  + wykrywanie duplikatów
                                │   detekcja pętli)   │
                                └─────────┬─────────┘
                                          │
                     ┌────────────────────┼────────────────────┐
                     │                    │                    │
              ┌──────▼──────┐    ┌───────▼───────┐    ┌──────▼───────┐
              │  ChatPanel   │    │ ExcelSkillSvc  │    │  I18nService  │
              │  (WPF CTP)   │    │  (15 narzędzi) │    │  (8 języków)  │
              └─────────────┘    └───────────────┘    └──────────────┘
```

## Stara wersja (v1.0 — VBA)

Oryginalna wersja VBA (`.xlam`) znajduje się w katalogu [`legacy/`](legacy/). Nie jest już rozwijana — została w pełni zastąpiona wersją .NET powyżej. Aby ją zbudować: `cscript legacy\build.vbs`.

## Licencja

Projekt open-source. Wykorzystuje [API Z.AI](https://z.ai) — wymagane konto i klucz API.
