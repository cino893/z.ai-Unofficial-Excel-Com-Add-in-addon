Attribute VB_Name = "modI18n"
'==============================================================================
' modI18n - Internationalization Module
' Detects Windows language, provides T() translation function
'==============================================================================
Option Explicit

Private m_lang As String
Private m_translations As Object
Private m_initialized As Boolean

' --- Initialize i18n ---
Public Sub InitI18n()
    If m_initialized Then Exit Sub
    
    Set m_translations = CreateObject("Scripting.Dictionary")
    
    ' Detect Windows language
    m_lang = DetectLanguage()
    
    ' Load translations
    LoadTranslations_PL
    LoadTranslations_EN
    
    m_initialized = True
End Sub

' --- Detect Windows language from Excel settings ---
Private Function DetectLanguage() As String
    On Error Resume Next
    Dim langCode As Long
    langCode = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    On Error GoTo 0
    
    Select Case langCode
        Case 1045 ' Polish
            DetectLanguage = "pl"
        Case 1033, 2057, 3081, 4105 ' English US/UK/AU/CA
            DetectLanguage = "en"
        Case Else
            ' Fallback: check country setting
            On Error Resume Next
            Dim country As Long
            country = Application.International(xlCountrySetting)
            On Error GoTo 0
            If country = 48 Then
                DetectLanguage = "pl"
            Else
                DetectLanguage = "en"
            End If
    End Select
End Function

' --- Get current language ---
Public Function GetLanguage() As String
    If Not m_initialized Then InitI18n
    GetLanguage = m_lang
End Function

' --- Set language manually ---
Public Sub SetLanguage(ByVal langCode As String)
    If Not m_initialized Then InitI18n
    m_lang = LCase(langCode)
    
    ' Save to registry
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite "HKCU\Software\ZaiExcelAddin\Language", m_lang, "REG_SZ"
    On Error GoTo 0
End Sub

' --- Translate ---
Public Function T(ByVal key As String) As String
    If Not m_initialized Then InitI18n
    
    Dim fullKey As String
    fullKey = m_lang & "." & key
    
    If m_translations.Exists(fullKey) Then
        T = m_translations(fullKey)
    Else
        ' Fallback to English
        fullKey = "en." & key
        If m_translations.Exists(fullKey) Then
            T = m_translations(fullKey)
        Else
            ' Return key as-is
            T = key
        End If
    End If
End Function

' --- Check if saved language preference exists ---
Private Function LoadSavedLanguage() As String
    On Error GoTo ErrHandler
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    LoadSavedLanguage = wsh.RegRead("HKCU\Software\ZaiExcelAddin\Language")
    Exit Function
ErrHandler:
    LoadSavedLanguage = ""
End Function

' ======================== POLISH TRANSLATIONS ========================
Private Sub LoadTranslations_PL()
    Dim L As String: L = "pl."
    
    ' Menu
    m_translations(L & "menu.chat") = "&Asystent AI (Chat)"
    m_translations(L & "menu.quick") = "&Szybkie polecenie"
    m_translations(L & "menu.login") = "&Zaloguj (Klucz API)"
    m_translations(L & "menu.logout") = "&Wyloguj"
    m_translations(L & "menu.model") = "&Wybierz model"
    m_translations(L & "menu.viewlog") = "&Pokaz log debugowania"
    m_translations(L & "menu.clearlog") = "&Wyczysc log"
    m_translations(L & "menu.about") = "&O dodatku Z.AI"
    m_translations(L & "menu.language") = "&Jezyk / Language"
    
    ' Chat form
    m_translations(L & "chat.title") = "Z.AI - Asystent Excel"
    m_translations(L & "chat.send") = "Wyslij"
    m_translations(L & "chat.new") = "Nowa rozmowa"
    m_translations(L & "chat.clear") = "Wyczysc"
    m_translations(L & "chat.ready") = "Gotowy"
    m_translations(L & "chat.processing") = "Przetwarzanie..."
    m_translations(L & "chat.ready_count") = "Gotowy ({0} wiadomosci)"
    m_translations(L & "chat.new_started") = "Nowa rozmowa rozpoczeta. Jak moge pomoc?"
    m_translations(L & "chat.welcome") = "Witaj! Jestem asystentem AI zintegrowanym z Excel." & vbCrLf & _
        "Moge pomoc Ci edytowac arkusz - po prostu opisz co chcesz zrobic." & vbCrLf & vbCrLf & _
        "Przyklady polecen:" & vbCrLf & _
        "  - Przeczytaj dane z kolumny A" & vbCrLf & _
        "  - Dodaj formule SUM do komorki B10" & vbCrLf & _
        "  - Sformatuj naglowki na pogrubione" & vbCrLf & _
        "  - Posortuj dane wedlug kolumny C malejaco" & vbCrLf & _
        "  - Stworz wykres kolumnowy z danych A1:B5"
    
    ' Auth
    m_translations(L & "auth.need_login") = "Musisz najpierw sie zalogowac (podac klucz API)."
    m_translations(L & "auth.want_login") = "Czy chcesz to zrobic teraz?"
    m_translations(L & "auth.prompt") = "Podaj klucz API z.ai:" & vbCrLf & vbCrLf & _
        "Klucz mozesz uzyskac na: https://open.z.ai/" & vbCrLf & "(Sekcja: API Keys)"
    m_translations(L & "auth.current_key") = "Aktualny klucz: "
    m_translations(L & "auth.login_title") = "Z.AI - Logowanie"
    m_translations(L & "auth.cancelled") = "Logowanie anulowane. Musisz podac klucz API aby korzystac z Z.AI."
    m_translations(L & "auth.validating") = "Z.AI: Weryfikacja klucza API..."
    m_translations(L & "auth.success") = "Zalogowano pomyslnie!" & vbCrLf & "Klucz API zostal zapisany."
    m_translations(L & "auth.failed") = "Nie udalo sie zweryfikowac klucza API." & vbCrLf & "Czy chcesz go mimo to zapisac?"
    m_translations(L & "auth.not_logged") = "Nie jestes zalogowany."
    m_translations(L & "auth.confirm_logout") = "Czy na pewno chcesz sie wylogowac?" & vbCrLf & "Klucz API zostanie usuniety."
    m_translations(L & "auth.logged_out") = "Wylogowano."
    
    ' Quick command
    m_translations(L & "quick.prompt") = "Wpisz polecenie dla asystenta AI:" & vbCrLf & vbCrLf & _
        "Przyklady:" & vbCrLf & _
        "  - Podsumuj dane w kolumnie A" & vbCrLf & _
        "  - Dodaj formule SUM do B10" & vbCrLf & _
        "  - Sformatuj naglowki na pogrubione" & vbCrLf & _
        "  - Stworz wykres z danych A1:B10"
    m_translations(L & "quick.title") = "Z.AI - Szybkie polecenie"
    m_translations(L & "quick.status") = "Z.AI: Przetwarzanie polecenia..."
    m_translations(L & "quick.result_title") = "Z.AI - Odpowiedz"
    
    ' Model selector
    m_translations(L & "model.prompt") = "Wybierz model z.ai:" & vbCrLf & vbCrLf & _
        "Dostepne modele:" & vbCrLf & _
        "  glm-4-plus  (domyslny, szybki)" & vbCrLf & _
        "  glm-4-long  (dlugi kontekst)" & vbCrLf & _
        "  glm-4       (standardowy)" & vbCrLf & _
        "  glm-3-turbo (najszybszy)"
    m_translations(L & "model.current") = "Aktualny: "
    m_translations(L & "model.title") = "Z.AI - Wybor modelu"
    m_translations(L & "model.changed") = "Model zmieniony na: "
    
    ' About
    m_translations(L & "about.text") = _
        "Z.AI Excel Add-in" & vbCrLf & _
        "Wersja: 1.0.0" & vbCrLf & vbCrLf & _
        "Asystent AI zintegrowany z Microsoft Excel." & vbCrLf & _
        "Wykorzystuje platforme z.ai (Zhipu AI) do" & vbCrLf & _
        "inteligentnej edycji arkuszy kalkulacyjnych." & vbCrLf & vbCrLf & _
        "Mozliwosci:" & vbCrLf & _
        "  - Czytanie i zapisywanie komorek" & vbCrLf & _
        "  - Formatowanie danych" & vbCrLf & _
        "  - Wstawianie formul" & vbCrLf & _
        "  - Sortowanie danych" & vbCrLf & _
        "  - Tworzenie wykresow" & vbCrLf & _
        "  - Zarzadzanie arkuszami" & vbCrLf & vbCrLf & _
        "Strona: https://z.ai" & vbCrLf & _
        "Dokumentacja API: https://docs.z.ai"
    m_translations(L & "about.title") = "Z.AI - O dodatku"
    
    ' Conversation
    m_translations(L & "conv.status_round") = "Z.AI: Przetwarzanie... (runda {0})"
    m_translations(L & "conv.status_exec") = "Z.AI: Wykonywanie {0} operacji na Excelu..."
    m_translations(L & "conv.max_rounds") = "[Uwaga]: Osiagnieto maksymalna liczbe rund ({0}). Ostatnia odpowiedz moze byc niekompletna."
    m_translations(L & "conv.loop_detected") = "[Uwaga]: Wykryto powtarzajaca sie operacje. Przerywam petle."
    m_translations(L & "conv.api_error") = "[Blad API]: "
    m_translations(L & "conv.no_response") = "Brak odpowiedzi z serwera"
    m_translations(L & "conv.error") = "[Blad]: "
    m_translations(L & "conv.no_assistant") = "Brak odpowiedzi asystenta"
    
    ' Debug
    m_translations(L & "debug.no_log") = "Brak pliku logu."
    m_translations(L & "debug.path") = "Sciezka: "
    
    ' Errors
    m_translations(L & "error.generic") = "Blad: "
    
    ' Language
    m_translations(L & "lang.changed") = "Jezyk zmieniony na: Polski"
    m_translations(L & "lang.title") = "Z.AI - Jezyk"
    
    ' System prompt
    m_translations(L & "system.prompt") = _
        "Jestes inteligentnym asystentem AI zintegrowanym z Microsoft Excel. " & _
        "Pomagasz uzytkownikowi edytowac i analizowac dane w arkuszu kalkulacyjnym. " & _
        "Masz dostep do narzedzi (tools) ktore pozwalaja Ci czytac i zapisywac komorki, " & _
        "formatowac dane, wstawiac formuly, sortowac, tworzyc wykresy i wiele wiecej." & vbLf & vbLf & _
        "ZASADY:" & vbLf & _
        "1. Zawsze NAJPIERW uzyj get_sheet_info lub get_workbook_info aby poznac kontekst danych." & vbLf & _
        "2. Przed modyfikacja danych, przeczytaj odpowiedni zakres aby zrozumiec strukture." & vbLf & _
        "3. Po wykonaniu zmian, potwierdzaj co zrobiles." & vbLf & _
        "4. Uzywaj polskich nazw w komunikatach do uzytkownika." & vbLf & _
        "5. Formuly Excel pisz w skladni angielskiej (SUM, AVERAGE, IF, VLOOKUP itp.)." & vbLf & _
        "6. Kolory podawaj jako RGB long: Red=255, Green=65280, Blue=16711680, Yellow=65535, " & _
        "LightGray=12632256, White=16777215, Orange=33023." & vbLf & _
        "7. Jezeli uzytkownik nie sprecyzuje arkusza, uzyj aktywnego arkusza." & vbLf & _
        "8. Badz zwiezly ale informatywny w odpowiedziach." & vbLf & _
        "9. Przed tworzeniem nowego wykresu, uzyj list_charts aby sprawdzic istniejace wykresy. " & _
        "Jezeli chcesz poprawic wykres, NAJPIERW usun stary (delete_chart) a potem stworz nowy." & vbLf & _
        "10. Nie powtarzaj tej samej operacji wiele razy - jezeli narzedzie zwrocilo sukces, przejdz dalej."
End Sub

' ======================== ENGLISH TRANSLATIONS ========================
Private Sub LoadTranslations_EN()
    Dim L As String: L = "en."
    
    ' Menu
    m_translations(L & "menu.chat") = "&AI Assistant (Chat)"
    m_translations(L & "menu.quick") = "&Quick Command"
    m_translations(L & "menu.login") = "&Login (API Key)"
    m_translations(L & "menu.logout") = "&Logout"
    m_translations(L & "menu.model") = "&Select Model"
    m_translations(L & "menu.viewlog") = "&View Debug Log"
    m_translations(L & "menu.clearlog") = "&Clear Log"
    m_translations(L & "menu.about") = "&About Z.AI"
    m_translations(L & "menu.language") = "&Language / Jezyk"
    
    ' Chat form
    m_translations(L & "chat.title") = "Z.AI - Excel Assistant"
    m_translations(L & "chat.send") = "Send"
    m_translations(L & "chat.new") = "New Chat"
    m_translations(L & "chat.clear") = "Clear"
    m_translations(L & "chat.ready") = "Ready"
    m_translations(L & "chat.processing") = "Processing..."
    m_translations(L & "chat.ready_count") = "Ready ({0} messages)"
    m_translations(L & "chat.new_started") = "New conversation started. How can I help?"
    m_translations(L & "chat.welcome") = "Hello! I'm an AI assistant integrated with Excel." & vbCrLf & _
        "I can help you edit your spreadsheet - just describe what you want to do." & vbCrLf & vbCrLf & _
        "Example commands:" & vbCrLf & _
        "  - Read data from column A" & vbCrLf & _
        "  - Add SUM formula to cell B10" & vbCrLf & _
        "  - Format headers as bold" & vbCrLf & _
        "  - Sort data by column C descending" & vbCrLf & _
        "  - Create a column chart from data A1:B5"
    
    ' Auth
    m_translations(L & "auth.need_login") = "You need to log in first (provide API key)."
    m_translations(L & "auth.want_login") = "Would you like to do it now?"
    m_translations(L & "auth.prompt") = "Enter z.ai API key:" & vbCrLf & vbCrLf & _
        "Get your key at: https://open.z.ai/" & vbCrLf & "(Section: API Keys)"
    m_translations(L & "auth.current_key") = "Current key: "
    m_translations(L & "auth.login_title") = "Z.AI - Login"
    m_translations(L & "auth.cancelled") = "Login cancelled. You need an API key to use Z.AI."
    m_translations(L & "auth.validating") = "Z.AI: Validating API key..."
    m_translations(L & "auth.success") = "Logged in successfully!" & vbCrLf & "API key has been saved."
    m_translations(L & "auth.failed") = "Could not validate API key." & vbCrLf & "Save it anyway?"
    m_translations(L & "auth.not_logged") = "You are not logged in."
    m_translations(L & "auth.confirm_logout") = "Are you sure you want to log out?" & vbCrLf & "API key will be removed."
    m_translations(L & "auth.logged_out") = "Logged out."
    
    ' Quick command
    m_translations(L & "quick.prompt") = "Enter a command for the AI assistant:" & vbCrLf & vbCrLf & _
        "Examples:" & vbCrLf & _
        "  - Summarize data in column A" & vbCrLf & _
        "  - Add SUM formula to B10" & vbCrLf & _
        "  - Format headers as bold" & vbCrLf & _
        "  - Create a chart from data A1:B10"
    m_translations(L & "quick.title") = "Z.AI - Quick Command"
    m_translations(L & "quick.status") = "Z.AI: Processing command..."
    m_translations(L & "quick.result_title") = "Z.AI - Response"
    
    ' Model selector
    m_translations(L & "model.prompt") = "Select z.ai model:" & vbCrLf & vbCrLf & _
        "Available models:" & vbCrLf & _
        "  glm-4-plus  (default, fast)" & vbCrLf & _
        "  glm-4-long  (long context)" & vbCrLf & _
        "  glm-4       (standard)" & vbCrLf & _
        "  glm-3-turbo (fastest)"
    m_translations(L & "model.current") = "Current: "
    m_translations(L & "model.title") = "Z.AI - Model Selection"
    m_translations(L & "model.changed") = "Model changed to: "
    
    ' About
    m_translations(L & "about.text") = _
        "Z.AI Excel Add-in" & vbCrLf & _
        "Version: 1.0.0" & vbCrLf & vbCrLf & _
        "AI assistant integrated with Microsoft Excel." & vbCrLf & _
        "Uses the z.ai platform (Zhipu AI) for" & vbCrLf & _
        "intelligent spreadsheet editing." & vbCrLf & vbCrLf & _
        "Capabilities:" & vbCrLf & _
        "  - Read and write cells" & vbCrLf & _
        "  - Format data" & vbCrLf & _
        "  - Insert formulas" & vbCrLf & _
        "  - Sort data" & vbCrLf & _
        "  - Create charts" & vbCrLf & _
        "  - Manage worksheets" & vbCrLf & vbCrLf & _
        "Website: https://z.ai" & vbCrLf & _
        "API Docs: https://docs.z.ai"
    m_translations(L & "about.title") = "Z.AI - About"
    
    ' Conversation
    m_translations(L & "conv.status_round") = "Z.AI: Processing... (round {0})"
    m_translations(L & "conv.status_exec") = "Z.AI: Executing {0} Excel operations..."
    m_translations(L & "conv.max_rounds") = "[Warning]: Max rounds reached ({0}). Response may be incomplete."
    m_translations(L & "conv.loop_detected") = "[Warning]: Repetitive operation detected. Breaking loop."
    m_translations(L & "conv.api_error") = "[API Error]: "
    m_translations(L & "conv.no_response") = "No response from server"
    m_translations(L & "conv.error") = "[Error]: "
    m_translations(L & "conv.no_assistant") = "No assistant response"
    
    ' Debug
    m_translations(L & "debug.no_log") = "No log file found."
    m_translations(L & "debug.path") = "Path: "
    
    ' Errors
    m_translations(L & "error.generic") = "Error: "
    
    ' Language
    m_translations(L & "lang.changed") = "Language changed to: English"
    m_translations(L & "lang.title") = "Z.AI - Language"
    
    ' System prompt
    m_translations(L & "system.prompt") = _
        "You are an intelligent AI assistant integrated with Microsoft Excel. " & _
        "You help the user edit and analyze data in the spreadsheet. " & _
        "You have access to tools that allow you to read and write cells, " & _
        "format data, insert formulas, sort, create charts and more." & vbLf & vbLf & _
        "RULES:" & vbLf & _
        "1. Always FIRST use get_sheet_info or get_workbook_info to understand the data context." & vbLf & _
        "2. Before modifying data, read the relevant range to understand the structure." & vbLf & _
        "3. After making changes, confirm what you did." & vbLf & _
        "4. Communicate with the user in English." & vbLf & _
        "5. Write Excel formulas in English syntax (SUM, AVERAGE, IF, VLOOKUP etc.)." & vbLf & _
        "6. Specify colors as RGB long: Red=255, Green=65280, Blue=16711680, Yellow=65535, " & _
        "LightGray=12632256, White=16777215, Orange=33023." & vbLf & _
        "7. If the user doesn't specify a sheet, use the active sheet." & vbLf & _
        "8. Be concise but informative in responses." & vbLf & _
        "9. Before creating a new chart, use list_charts to check existing charts. " & _
        "If you want to fix a chart, FIRST delete the old one (delete_chart) then create a new one." & vbLf & _
        "10. Do not repeat the same operation multiple times - if a tool returned success, move on."
End Sub

' --- Helper: Replace {0} placeholder ---
Public Function TFormat(ByVal key As String, ByVal arg0 As Variant) As String
    TFormat = Replace(T(key), "{0}", CStr(arg0))
End Function
