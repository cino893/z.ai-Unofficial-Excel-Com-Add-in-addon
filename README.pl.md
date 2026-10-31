# Z.AI Dodatek do Excela

[![Licencja: MIT](https://img.shields.io/badge/Licencja-MIT-yellow.svg)](LICENSE)
[![.NET](https://img.shields.io/badge/.NET-8.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![Platforma](https://img.shields.io/badge/Platforma-Windows-0078D4?logo=windows)](https://www.microsoft.com/windows)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?logo=microsoftexcel)](https://www.microsoft.com/excel)
[![Wydanie](https://img.shields.io/github/v/release/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon)](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases/latest)
[![Zgłoszenia](https://img.shields.io/github/issues/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon)](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/issues)

[🇬🇧 English version](README.md)

**Darmowy, otwartoźródłowy dodatek do Excela zasilany przez Z.AI (Zhipu AI)** — Rozmawiaj z asystentem AI, który może czytać, pisać, formatować, tworzyć wykresy i automatyzować Twoje arkusze kalkulacyjne. Idealny do analizy danych, automatyzacji raportów i zwiększania produktywności w Excelu.

> ⚠️ **Uwaga:** To jest nieoficjalny dodatek społecznościowy. Nie jest powiązany z Zhipu AI / Z.AI ani przez nich wspierany.

![Demo](show-reel.gif)

## Pobierz

- **[⬇ Pobierz najnowszy .xll](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases/latest/download/ZaiExcelAddin-AddIn64-packed.xll)**
- [Wszystkie wydania](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/releases)

Wydania GitHub dołączają spakowane pliki `.xll` (64-bit oraz 32-bit, jeśli powstanie) przez Actions (`release.yml`). Aktualna wersja produktu to **5.0.0**.

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

✨ **Najważniejsze:**

- 🤖 **Automatyzacja AI** — 31 narzędzi Excel do czytania, pisania, formatowania, wykresów, tabel przestawnych i więcej
- 💬 **Czat w Panelu Bocznym** — Piękny interfejs WPF z animacją pisania i emoji
- 🆓 **Darmowy** — Licencja open-source MIT, działa z darmowymi modelami Z.AI
- 🌍 **Wielojęzyczny** — 8 języków (EN, PL, DE, FR, ES, UK, ZH, JA)
- 🚀 **Szybka Wydajność** — Zoptymalizowane z kontrolą aktualizacji ekranu i wykrywaniem pętli
- 🎨 **Natywny dla Excela** — Własna zakładka w Ribbon i panel zadań, wygląda jak wbudowana funkcja

### Możliwości

| Narzędzie | Opis |
|-----------|------|
| `read_cell` / `write_cell` | Odczyt/zapis komórki (`write_cell` zapisuje liczby jako wartości numeryczne) |
| `read_range` / `write_range` | Odczyt/zapis zakresu (`read_range` ogranicza do 500 wierszy) |
| `get_sheet_info` | Wymiary, nagłówki, próbka danych, flagi wykresów/pivotów/filtrów |
| `get_workbook_info` | Arkusze, ścieżka pliku, liczba cache'y przestawnych |
| `format_range` | Czcionka (w tym podkreślenie / nazwa), kolory, ramki, wyrównanie, scalanie |
| `insert_formula` | Wstawianie formuł Excel z opcjonalnym `fill_to` (AutoFill) |
| `sort_range` | Sortowanie danych po kolumnie |
| `add_sheet` | Dodawanie arkusza |
| `delete_rows` / `insert_rows` | Usuwanie/wstawianie wierszy |
| `delete_columns` / `insert_columns` | Usuwanie/wstawianie kolumn |
| `create_chart` | Tworzenie wykresów (kolumnowy, słupkowy, liniowy, kołowy, punktowy, obszarowy) |
| `delete_chart` / `list_charts` | Usuwanie lub lista wykresów |
| `create_pivot_table` | Tworzenie tabeli przestawnej z polami wierszy/kolumn/wartości |
| `move_pivot_table` | Przenoszenie tabeli przestawnej (Location, potem odtworzenie z PivotCache) |
| `modify_pivot_table` | Konfiguracja pól wierszy/kolumn/wartości/stron na istniejącej tabeli |
| `auto_filter` | Zastosowanie lub usunięcie AutoFiltra na zakresie |
| `find_replace` | Znajdź i zamień wartości w arkuszu |
| `conditional_format` | Skala kolorów, paski danych, zestawy ikon, formuła, wartość komórki i inne |
| `copy_range` | Kopiowanie zakresu do innej lokalizacji (wartości lub z formatowaniem) |
| `rename_sheet` / `delete_sheet` | Zmiana nazwy lub usuwanie arkusza |
| `freeze_panes` | Zablokowanie/odblokowanie okienek w danej komórce |
| `remove_duplicates` | Usuwanie zduplikowanych wierszy z zakresu |
| `set_validation` | Walidacja danych (lista, liczba, data, długość tekstu) |
| `list_pivot_tables` | Lista wszystkich tabel przestawnych w skoroszycie |
| `clear_range` | Czyszczenie zawartości, formatowania lub wszystkiego z zakresu |

## Budowanie ze źródeł

Wymagany .NET SDK 8.0+:

```powershell
cd src
dotnet build -c Release
```

Wynik (Excel-DNA 1.9.0 packed XLL): `src\bin\Release\net8.0-windows\publish\ZaiExcelAddin-AddIn64-packed.xll`  
CI używa `dotnet publish src/ZaiExcelAddin.csproj -c Release -o src/publish`.

## Struktura projektu

```
dodatek-z-ai-opus/
├── src/                            # .NET 8 COM Add-in (ExcelDNA) — zalecany
│   ├── ZaiExcelAddin.csproj        # Projekt C# (wersja 5.0.0)
│   ├── AddIn.cs                    # Punkt wejścia (IExcelAddIn)
│   ├── RibbonController.cs         # Wstążka + Custom Task Pane
│   ├── i18n/                       # 8 plików JSON języków
│   ├── Models/
│   │   └── ChatMessage.cs
│   ├── Services/
│   │   ├── AuthService.cs
│   │   ├── ConversationService.cs  # Pętla tool-calling (max 45 rund)
│   │   ├── DebugLogger.cs
│   │   ├── ExcelSkillService*.cs   # 31 narzędzi Excel
│   │   ├── I18nService.cs
│   │   └── ZaiApiService.cs
│   └── UI/
│       ├── ChatPanel.xaml/.cs
│       ├── ChatPaneHost.cs
│       ├── WpfLoginDialog.xaml/.cs
│       └── WpfSelectDialog.xaml/.cs
├── legacy/                         # Dodatek VBA z parytetem funkcji COM
│   ├── build.vbs / build.ps1 / build.bat
│   ├── modExcelSkills.bas          # JSON narzędzi + dispatch
│   ├── modExcelSkillsOps.bas       # Dane/arkusz/format/wykres/filtr
│   ├── modExcelSkillsPivot.bas     # Narzędzia tabel przestawnych
│   ├── modConversation.bas
│   ├── modZaiAPI.bas / modAuth.bas / modI18n.bas / modRibbon.bas
│   ├── frmZaiChat.frm
│   └── i18n/                       # Te same JSON co src/i18n
├── tests/                          # Testy offline (bez Excela)
├── .github/workflows/              # ci.yml + release.yml
├── CHANGELOG.md
├── show-reel.gif
├── dodatek-z-ai-opus.sln
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
                                │  (max 45 rund,      │  + wykrywanie duplikatów
                                │   detekcja pętli)   │
                                └─────────┬─────────┘
                                          │
                     ┌────────────────────┼────────────────────┐
                     │                    │                    │
              ┌──────▼──────┐    ┌───────▼───────┐    ┌──────▼───────┐
              │  ChatPanel   │    │ ExcelSkillSvc  │    │  I18nService  │
              │  (WPF CTP)   │    │  (31 narzędzi) │    │  (8 języków)  │
              └─────────────┘    └───────────────┘    └──────────────┘
```

## Wersja VBA (parytet funkcji)

Katalog [`legacy/`](legacy/) to **utrzymywany dodatek VBA** z tymi samymi 31 narzędziami Excel, pętlą 45 rund, 8 językami i katalogiem modeli co dodatek COM .NET. Zalecana instalacja to nadal spakowany `.xll`. Aby zbudować wersję VBA na komputerze z Excelem: `cscript legacy\build.vbs` (kopiuje też `legacy\i18n\` obok `.xlam`).

## Licencja

Ten projekt jest licencjonowany na warunkach [licencji MIT](LICENSE) — szczegóły w pliku LICENSE.

**Otwartoźródłowy i darmowy.** Używa [API Z.AI](https://z.ai) — wymagane jest konto i klucz API (dostępny darmowy tier).

## Współpraca

Wkład w projekt jest mile widziany! Przeczytaj nasze [Wytyczne dla Współpracowników](CONTRIBUTING.md) oraz [Kodeks Postępowania](CODE_OF_CONDUCT.md) przed wysłaniem pull requestów.

### Sposoby Współpracy

- 🐛 Zgłaszanie błędów i problemów
- 💡 Sugerowanie nowych funkcji lub ulepszeń
- 📝 Poprawa dokumentacji
- 🔧 Wysyłanie poprawek lub ulepszeń
- 🌍 Dodawanie tłumaczeń na nowe języki
- ⭐ Oznaczanie gwiazdką repozytorium aby pokazać wsparcie

## Bezpieczeństwo

W sprawach bezpieczeństwa, zobacz naszą [Politykę Bezpieczeństwa](SECURITY.md). Nie zgłaszaj podatności bezpieczeństwa przez publiczne zgłoszenia GitHub.

## Wsparcie

- 📖 [Dokumentacja](README.pl.md) — Instrukcja instalacji i użytkowania
- 💬 [Dyskusje GitHub](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/discussions) — Zadawaj pytania i dziel się pomysłami
- 🐛 [Zgłoszenia](https://github.com/cino893/z.ai-Unofficial-Excel-Com-Add-in-addon/issues) — Zgłaszaj błędy i proś o funkcje
- 🌐 [Platforma Z.AI](https://z.ai) — Oficjalna dokumentacja API

## Podziękowania

- Zbudowano z [ExcelDna](https://excel-dna.net/) dla infrastruktury dodatków COM
- Napędzane przez [Z.AI](https://z.ai) modele GLM od Zhipu AI
- Inspirowane potrzebą automatyzacji Excela z użyciem AI

## Słowa Kluczowe

`excel` `ai` `automatyzacja` `chatbot` `dodatek` `excel-addin` `dotnet` `csharp` `zhipu-ai` `glm` `arkusz-kalkulacyjny` `produktywnosc` `analiza-danych` `excel-automation` `asystent-ai` `darmowy` `open-source` `windows` `excel-tools` `tabela-przestawna` `wykresy` `formatowanie`
