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

- 🤖 **Automatyzacja AI** — 28 narzędzi Excel do czytania, pisania, formatowania, tworzenia wykresów i więcej
- 💬 **Czat w Panelu Bocznym** — Piękny interfejs WPF z animacją pisania i emoji
- 🆓 **Darmowy** — Licencja open-source MIT, działa z darmowymi modelami Z.AI
- 🌍 **Wielojęzyczny** — 8 języków (EN, PL, DE, FR, ES, UK, ZH, JA)
- 🚀 **Szybka Wydajność** — Zoptymalizowane z kontrolą aktualizacji ekranu i wykrywaniem pętli
- 🎨 **Natywny dla Excela** — Własna zakładka w Ribbon i panel zadań, wygląda jak wbudowana funkcja

### Możliwości

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
| `create_pivot_table` | Tworzenie tabeli przestawnej z polami wierszy/kolumn/wartości |
| `move_table` | Przenoszenie zakresu danych lub tabeli przestawnej na inny arkusz |
| `auto_filter` | Zastosowanie lub usunięcie AutoFiltra na zakresie |
| `find_replace` | Znajdź i zamień wartości w arkuszu |
| `conditional_format` | Formatowanie warunkowe (podświetlanie, skala kolorów, paski danych) |
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
│   │   ├── ExcelSkillService.cs    # 28 narzędzi Excel
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
              │  (WPF CTP)   │    │  (28 narzędzi) │    │  (8 języków)  │
              └─────────────┘    └───────────────┘    └──────────────┘
```

## Stara wersja (v1.0 — VBA)

Oryginalna wersja VBA (`.xlam`) znajduje się w katalogu [`legacy/`](legacy/). Nie jest już rozwijana — została w pełni zastąpiona wersją .NET powyżej. Aby ją zbudować: `cscript legacy\build.vbs`.

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
