using System.Globalization;
using Microsoft.Win32;

namespace ZaiExcelAddin.Services;

public class I18nService
{
    private const string RegKeyPath = @"SOFTWARE\ZaiExcelAddin";
    private const string RegValueName = "Language";
    private const string DefaultLanguage = "en";

    public static Dictionary<string, string> SupportedLanguages { get; } = new()
    {
        { "en", "English" },
        { "pl", "Polski" },
        { "de", "Deutsch" },
        { "fr", "Français" },
        { "es", "Español" },
        { "uk", "Українська" },
        { "zh", "简体中文" },
        { "ja", "日本語" }
    };

    public string CurrentLanguage { get; private set; } = DefaultLanguage;

    private readonly Dictionary<string, Dictionary<string, string>> _translations = new();

    public I18nService()
    {
        InitTranslations();
        var saved = LoadLanguageFromRegistry();
        if (saved != null && SupportedLanguages.ContainsKey(saved))
        {
            CurrentLanguage = saved;
        }
        else
        {
            CurrentLanguage = DetectLanguage();
        }
        try { AddIn.Logger?.Info($"I18nService initialized, language: {CurrentLanguage}"); } catch { }
    }

    public string T(string key)
    {
        if (_translations.TryGetValue(CurrentLanguage, out var lang) && lang.TryGetValue(key, out var val))
            return val;
        if (_translations.TryGetValue(DefaultLanguage, out var en) && en.TryGetValue(key, out var fallback))
            return fallback;
        return key;
    }

    public string TFormat(string key, object arg0)
    {
        return T(key).Replace("{0}", arg0?.ToString() ?? "");
    }

    public void SetLanguage(string code)
    {
        if (!SupportedLanguages.ContainsKey(code)) return;
        CurrentLanguage = code;
        SaveLanguageToRegistry(code);
        try { AddIn.Logger?.Info($"Language changed to: {code}"); } catch { }
    }

    private string DetectLanguage()
    {
        var culture = CultureInfo.CurrentUICulture;
        var twoLetter = culture.TwoLetterISOLanguageName.ToLowerInvariant();
        if (SupportedLanguages.ContainsKey(twoLetter))
            return twoLetter;
        return DefaultLanguage;
    }

    private string? LoadLanguageFromRegistry()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegKeyPath);
            return key?.GetValue(RegValueName) as string;
        }
        catch
        {
            return null;
        }
    }

    private void SaveLanguageToRegistry(string code)
    {
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegKeyPath);
            key.SetValue(RegValueName, code);
        }
        catch (Exception ex)
        {
            try { AddIn.Logger?.Error($"Failed to save language to registry: {ex.Message}"); } catch { }
        }
    }

    private void InitTranslations()
    {
        InitEnglish();
        InitPolish();
        InitGerman();
        InitFrench();
        InitSpanish();
        InitUkrainian();
        InitChinese();
        InitJapanese();
        AddExtraKeys();
    }

    // ── English ──────────────────────────────────────────────────────────

    private void InitEnglish()
    {
        _translations["en"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Chat",
            ["menu.login"] = "Login",
            ["menu.logout"] = "Logout",
            ["menu.model"] = "Model",
            ["menu.viewlog"] = "View Log",
            ["menu.clearlog"] = "Clear Log",
            ["menu.about"] = "About",
            ["menu.language"] = "Language",

            // Chat
            ["chat.title"] = "Z.AI Chat",
            ["chat.send"] = "Send",
            ["chat.new"] = "New Chat",
            ["chat.clear"] = "Clear",
            ["chat.ready"] = "Ready",
            ["chat.processing"] = "Processing...",
            ["chat.ready_count"] = "Ready ({0} messages)",
            ["chat.new_started"] = "New conversation started.",
            ["chat.welcome"] = "Welcome to Z.AI Chat!\n\nI can help you work with Excel. Try asking:\n• \"Summarize the data in this sheet\"\n• \"Create a chart from columns A and B\"\n• \"Format the header row with bold and blue background\"\n• \"Calculate the average of column C\"",

            // Auth
            ["auth.prompt"] = "Enter your API key:",
            ["auth.current_key"] = "Current key: {0}",
            ["auth.login_title"] = "Z.AI Login",
            ["auth.cancelled"] = "Login cancelled.",
            ["auth.validating"] = "Validating API key...",
            ["auth.success"] = "Login successful!",
            ["auth.failed"] = "Login failed. Invalid API key.",
            ["auth.not_logged"] = "You are not logged in. Please log in first.",
            ["auth.confirm_logout"] = "Are you sure you want to log out?",
            ["auth.logged_out"] = "You have been logged out.",

            // Model
            ["model.prompt"] = "Select a model:\n1. glm-4-plus (recommended)\n2. glm-4-long (long context)\n3. glm-4 (standard)\n4. glm-3-turbo (fast)",
            ["model.current"] = "Current model: {0}",
            ["model.title"] = "Select Model",
            ["model.changed"] = "Model changed to: {0}",

            // About
            ["about.text"] = "Z.AI Excel Add-in\nVersion 2.0\n\nAI-powered assistant for Microsoft Excel.\nPowered by ZhipuAI GLM models.\n\n© 2024 Z.AI",
            ["about.title"] = "About Z.AI",

            // Conversation
            ["conv.status_round"] = "Round {0}",
            ["conv.status_exec"] = "Executing tool: {0}",
            ["conv.max_rounds"] = "Maximum rounds reached ({0}). Stopping.",
            ["conv.loop_detected"] = "Loop detected. Stopping to prevent infinite execution.",
            ["conv.api_error"] = "API error occurred. Please try again.",
            ["conv.no_response"] = "No response received from the API.",
            ["conv.error"] = "An error occurred during processing.",
            ["conv.no_assistant"] = "No assistant response in API reply.",

            // Debug
            ["debug.no_log"] = "No log entries.",

            // Language
            ["lang.changed"] = "Language changed. Some changes may require restarting the add-in.",
            ["lang.title"] = "Language",

            // System prompt
            ["system.prompt"] = "You are an AI assistant integrated into Microsoft Excel through the Z.AI add-in. You have access to tools that can read and modify Excel workbooks.\n\nRules you must follow:\n1. Always call get_sheet_info or get_workbook_info first to understand the current state of the workbook before taking any action.\n2. Always read data before modifying it. Never assume the contents of cells.\n3. After making changes, confirm what was done by reading back the affected cells.\n4. Write all formulas using English function names (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN, etc.).\n5. When setting colors, use RGB Long values: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. Default to the active sheet unless the user specifies otherwise.\n7. Before creating charts, call list_charts first. If a similar chart already exists, delete it before creating a new one.\n8. Do not repeat operations that have already been completed successfully.\n9. Communicate with the user in English.\n10. Be concise and helpful. Explain what you are doing step by step."
        };
    }

    // ── Polish ───────────────────────────────────────────────────────────

    private void InitPolish()
    {
        _translations["pl"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Czat",
            ["menu.login"] = "Zaloguj",
            ["menu.logout"] = "Wyloguj",
            ["menu.model"] = "Model",
            ["menu.viewlog"] = "Pokaż log",
            ["menu.clearlog"] = "Wyczyść log",
            ["menu.about"] = "O programie",
            ["menu.language"] = "Język",

            // Chat
            ["chat.title"] = "Z.AI Czat",
            ["chat.send"] = "Wyślij",
            ["chat.new"] = "Nowy czat",
            ["chat.clear"] = "Wyczyść",
            ["chat.ready"] = "Gotowy",
            ["chat.processing"] = "Przetwarzanie...",
            ["chat.ready_count"] = "Gotowy ({0} wiadomości)",
            ["chat.new_started"] = "Rozpoczęto nową rozmowę.",
            ["chat.welcome"] = "Witaj w Z.AI Chat!\n\nMog\u0119 pom\u00f3c Ci w pracy z Excelem. Spr\u00f3buj zapyta\u0107:\n\u2022 Podsumuj dane w tym arkuszu\n\u2022 Utw\u00f3rz wykres z kolumn A i B\n\u2022 Sformatuj wiersz nag\u0142\u00f3wka pogrubieniem\n\u2022 Oblicz \u015bredni\u0105 z kolumny C",

            // Auth
            ["auth.prompt"] = "Podaj klucz API:",
            ["auth.current_key"] = "Aktualny klucz: {0}",
            ["auth.login_title"] = "Z.AI Logowanie",
            ["auth.cancelled"] = "Logowanie anulowane.",
            ["auth.validating"] = "Weryfikacja klucza API...",
            ["auth.success"] = "Logowanie udane!",
            ["auth.failed"] = "Logowanie nieudane. Nieprawidłowy klucz API.",
            ["auth.not_logged"] = "Nie jesteś zalogowany. Najpierw się zaloguj.",
            ["auth.confirm_logout"] = "Czy na pewno chcesz się wylogować?",
            ["auth.logged_out"] = "Zostałeś wylogowany.",

            // Model
            ["model.prompt"] = "Wybierz model:\n1. glm-4-plus (zalecany)\n2. glm-4-long (długi kontekst)\n3. glm-4 (standardowy)\n4. glm-3-turbo (szybki)",
            ["model.current"] = "Aktualny model: {0}",
            ["model.title"] = "Wybór modelu",
            ["model.changed"] = "Model zmieniony na: {0}",

            // About
            ["about.text"] = "Z.AI Dodatek do Excela\nWersja 2.0\n\nAsystent AI dla Microsoft Excel.\nWykorzystuje modele ZhipuAI GLM.\n\n© 2024 Z.AI",
            ["about.title"] = "O Z.AI",

            // Conversation
            ["conv.status_round"] = "Runda {0}",
            ["conv.status_exec"] = "Wykonywanie narzędzia: {0}",
            ["conv.max_rounds"] = "Osiągnięto maksymalną liczbę rund ({0}). Zatrzymywanie.",
            ["conv.loop_detected"] = "Wykryto pętlę. Zatrzymywanie, aby zapobiec nieskończonemu wykonywaniu.",
            ["conv.api_error"] = "Wystąpił błąd API. Spróbuj ponownie.",
            ["conv.no_response"] = "Nie otrzymano odpowiedzi z API.",
            ["conv.error"] = "Wystąpił błąd podczas przetwarzania.",
            ["conv.no_assistant"] = "Brak odpowiedzi asystenta w odpowiedzi API.",

            // Debug
            ["debug.no_log"] = "Brak wpisów w logu.",

            // Language
            ["lang.changed"] = "Język został zmieniony. Niektóre zmiany mogą wymagać ponownego uruchomienia dodatku.",
            ["lang.title"] = "Język",

            // System prompt
            ["system.prompt"] = "Jesteś asystentem AI zintegrowanym z Microsoft Excel poprzez dodatek Z.AI. Masz dostęp do narzędzi, które mogą odczytywać i modyfikować skoroszyty Excela.\n\nZasady, których musisz przestrzegać:\n1. Zawsze najpierw wywołaj get_sheet_info lub get_workbook_info, aby poznać aktualny stan skoroszytu przed podjęciem jakiegokolwiek działania.\n2. Zawsze odczytaj dane przed ich modyfikacją. Nigdy nie zakładaj zawartości komórek.\n3. Po dokonaniu zmian potwierdź, co zostało zrobione, odczytując zmienione komórki.\n4. Zapisuj wszystkie formuły używając angielskich nazw funkcji (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN itp.).\n5. Przy ustawianiu kolorów używaj wartości RGB Long: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. Domyślnie pracuj na aktywnym arkuszu, chyba że użytkownik wskaże inaczej.\n7. Przed utworzeniem wykresów wywołaj list_charts. Jeśli podobny wykres już istnieje, usuń go przed utworzeniem nowego.\n8. Nie powtarzaj operacji, które zostały już pomyślnie wykonane.\n9. Komunikuj się z użytkownikiem po polsku.\n10. Bądź zwięzły i pomocny. Wyjaśniaj krok po kroku, co robisz."
        };
    }

    // ── German ───────────────────────────────────────────────────────────

    private void InitGerman()
    {
        _translations["de"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Chat",
            ["menu.login"] = "Anmelden",
            ["menu.logout"] = "Abmelden",
            ["menu.model"] = "Modell",
            ["menu.viewlog"] = "Log anzeigen",
            ["menu.clearlog"] = "Log löschen",
            ["menu.about"] = "Über",
            ["menu.language"] = "Sprache",

            // Chat
            ["chat.title"] = "Z.AI Chat",
            ["chat.send"] = "Senden",
            ["chat.new"] = "Neuer Chat",
            ["chat.clear"] = "Löschen",
            ["chat.ready"] = "Bereit",
            ["chat.processing"] = "Verarbeitung...",
            ["chat.ready_count"] = "Bereit ({0} Nachrichten)",
            ["chat.new_started"] = "Neue Unterhaltung gestartet.",
            ["chat.welcome"] = "Willkommen bei Z.AI Chat!\n\nIch kann Ihnen bei der Arbeit mit Excel helfen. Versuchen Sie:\n\u2022 Fasse die Daten in diesem Blatt zusammen\n\u2022 Erstelle ein Diagramm aus Spalten A und B\n\u2022 Formatiere die Kopfzeile fett mit blauem Hintergrund\n\u2022 Berechne den Durchschnitt der Spalte C",

            // Auth
            ["auth.prompt"] = "Geben Sie Ihren API-Schlüssel ein:",
            ["auth.current_key"] = "Aktueller Schlüssel: {0}",
            ["auth.login_title"] = "Z.AI Anmeldung",
            ["auth.cancelled"] = "Anmeldung abgebrochen.",
            ["auth.validating"] = "API-Schlüssel wird überprüft...",
            ["auth.success"] = "Anmeldung erfolgreich!",
            ["auth.failed"] = "Anmeldung fehlgeschlagen. Ungültiger API-Schlüssel.",
            ["auth.not_logged"] = "Sie sind nicht angemeldet. Bitte melden Sie sich zuerst an.",
            ["auth.confirm_logout"] = "Möchten Sie sich wirklich abmelden?",
            ["auth.logged_out"] = "Sie wurden abgemeldet.",

            // Model
            ["model.prompt"] = "Modell auswählen:\n1. glm-4-plus (empfohlen)\n2. glm-4-long (langer Kontext)\n3. glm-4 (Standard)\n4. glm-3-turbo (schnell)",
            ["model.current"] = "Aktuelles Modell: {0}",
            ["model.title"] = "Modell auswählen",
            ["model.changed"] = "Modell geändert zu: {0}",

            // About
            ["about.text"] = "Z.AI Excel-Add-in\nVersion 2.0\n\nKI-gestützter Assistent für Microsoft Excel.\nBasiert auf ZhipuAI GLM-Modellen.\n\n© 2024 Z.AI",
            ["about.title"] = "Über Z.AI",

            // Conversation
            ["conv.status_round"] = "Runde {0}",
            ["conv.status_exec"] = "Werkzeug wird ausgeführt: {0}",
            ["conv.max_rounds"] = "Maximale Rundenanzahl erreicht ({0}). Wird gestoppt.",
            ["conv.loop_detected"] = "Schleife erkannt. Wird gestoppt, um Endlosausführung zu verhindern.",
            ["conv.api_error"] = "API-Fehler aufgetreten. Bitte versuchen Sie es erneut.",
            ["conv.no_response"] = "Keine Antwort von der API erhalten.",
            ["conv.error"] = "Bei der Verarbeitung ist ein Fehler aufgetreten.",
            ["conv.no_assistant"] = "Keine Assistenten-Antwort in der API-Antwort.",

            // Debug
            ["debug.no_log"] = "Keine Log-Einträge.",

            // Language
            ["lang.changed"] = "Sprache geändert. Einige Änderungen erfordern möglicherweise einen Neustart des Add-ins.",
            ["lang.title"] = "Sprache",

            // System prompt
            ["system.prompt"] = "Du bist ein KI-Assistent, der über das Z.AI-Add-in in Microsoft Excel integriert ist. Du hast Zugriff auf Werkzeuge, die Excel-Arbeitsmappen lesen und bearbeiten können.\n\nRegeln, die du befolgen musst:\n1. Rufe immer zuerst get_sheet_info oder get_workbook_info auf, um den aktuellen Zustand der Arbeitsmappe zu verstehen, bevor du Maßnahmen ergreifst.\n2. Lies immer Daten, bevor du sie änderst. Nimm niemals den Inhalt von Zellen an.\n3. Bestätige nach Änderungen, was getan wurde, indem du die betroffenen Zellen zurückliest.\n4. Schreibe alle Formeln mit englischen Funktionsnamen (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN usw.).\n5. Verwende beim Setzen von Farben RGB-Long-Werte: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. Arbeite standardmäßig auf dem aktiven Blatt, sofern der Benutzer nichts anderes angibt.\n7. Rufe vor dem Erstellen von Diagrammen list_charts auf. Wenn ein ähnliches Diagramm bereits existiert, lösche es, bevor du ein neues erstellst.\n8. Wiederhole keine Operationen, die bereits erfolgreich abgeschlossen wurden.\n9. Kommuniziere mit dem Benutzer auf Deutsch.\n10. Sei prägnant und hilfreich. Erkläre Schritt für Schritt, was du tust."
        };
    }

    // ── French ───────────────────────────────────────────────────────────

    private void InitFrench()
    {
        _translations["fr"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Discussion",
            ["menu.login"] = "Connexion",
            ["menu.logout"] = "Déconnexion",
            ["menu.model"] = "Modèle",
            ["menu.viewlog"] = "Voir le journal",
            ["menu.clearlog"] = "Effacer le journal",
            ["menu.about"] = "À propos",
            ["menu.language"] = "Langue",

            // Chat
            ["chat.title"] = "Z.AI Discussion",
            ["chat.send"] = "Envoyer",
            ["chat.new"] = "Nouvelle discussion",
            ["chat.clear"] = "Effacer",
            ["chat.ready"] = "Prêt",
            ["chat.processing"] = "Traitement...",
            ["chat.ready_count"] = "Prêt ({0} messages)",
            ["chat.new_started"] = "Nouvelle conversation démarrée.",
            ["chat.welcome"] = "Bienvenue dans Z.AI Chat !\n\nJe peux vous aider à travailler avec Excel. Essayez de demander :\n• « Résume les données de cette feuille »\n• « Crée un graphique à partir des colonnes A et B »\n• « Formate la ligne d'en-tête en gras avec un fond bleu »\n• « Calcule la moyenne de la colonne C »",

            // Auth
            ["auth.prompt"] = "Entrez votre clé API :",
            ["auth.current_key"] = "Clé actuelle : {0}",
            ["auth.login_title"] = "Connexion Z.AI",
            ["auth.cancelled"] = "Connexion annulée.",
            ["auth.validating"] = "Validation de la clé API...",
            ["auth.success"] = "Connexion réussie !",
            ["auth.failed"] = "Échec de la connexion. Clé API invalide.",
            ["auth.not_logged"] = "Vous n'êtes pas connecté. Veuillez d'abord vous connecter.",
            ["auth.confirm_logout"] = "Êtes-vous sûr de vouloir vous déconnecter ?",
            ["auth.logged_out"] = "Vous avez été déconnecté.",

            // Model
            ["model.prompt"] = "Sélectionnez un modèle :\n1. glm-4-plus (recommandé)\n2. glm-4-long (contexte long)\n3. glm-4 (standard)\n4. glm-3-turbo (rapide)",
            ["model.current"] = "Modèle actuel : {0}",
            ["model.title"] = "Sélection du modèle",
            ["model.changed"] = "Modèle changé en : {0}",

            // About
            ["about.text"] = "Z.AI Complément Excel\nVersion 2.0\n\nAssistant IA pour Microsoft Excel.\nPropulsé par les modèles ZhipuAI GLM.\n\n© 2024 Z.AI",
            ["about.title"] = "À propos de Z.AI",

            // Conversation
            ["conv.status_round"] = "Tour {0}",
            ["conv.status_exec"] = "Exécution de l'outil : {0}",
            ["conv.max_rounds"] = "Nombre maximum de tours atteint ({0}). Arrêt en cours.",
            ["conv.loop_detected"] = "Boucle détectée. Arrêt pour éviter une exécution infinie.",
            ["conv.api_error"] = "Erreur API survenue. Veuillez réessayer.",
            ["conv.no_response"] = "Aucune réponse reçue de l'API.",
            ["conv.error"] = "Une erreur est survenue lors du traitement.",
            ["conv.no_assistant"] = "Pas de réponse de l'assistant dans la réponse API.",

            // Debug
            ["debug.no_log"] = "Aucune entrée dans le journal.",

            // Language
            ["lang.changed"] = "Langue modifiée. Certains changements peuvent nécessiter un redémarrage du complément.",
            ["lang.title"] = "Langue",

            // System prompt
            ["system.prompt"] = "Tu es un assistant IA intégré à Microsoft Excel via le complément Z.AI. Tu as accès à des outils qui peuvent lire et modifier les classeurs Excel.\n\nRègles à suivre :\n1. Appelle toujours get_sheet_info ou get_workbook_info en premier pour comprendre l'état actuel du classeur avant toute action.\n2. Lis toujours les données avant de les modifier. Ne suppose jamais le contenu des cellules.\n3. Après avoir effectué des modifications, confirme ce qui a été fait en relisant les cellules affectées.\n4. Écris toutes les formules avec les noms de fonctions en anglais (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN, etc.).\n5. Pour les couleurs, utilise les valeurs RGB Long : Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. Travaille par défaut sur la feuille active, sauf indication contraire de l'utilisateur.\n7. Avant de créer des graphiques, appelle list_charts. Si un graphique similaire existe déjà, supprime-le avant d'en créer un nouveau.\n8. Ne répète pas les opérations qui ont déjà été effectuées avec succès.\n9. Communique avec l'utilisateur en français.\n10. Sois concis et utile. Explique étape par étape ce que tu fais."
        };
    }

    // ── Spanish ──────────────────────────────────────────────────────────

    private void InitSpanish()
    {
        _translations["es"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Chat",
            ["menu.login"] = "Iniciar sesión",
            ["menu.logout"] = "Cerrar sesión",
            ["menu.model"] = "Modelo",
            ["menu.viewlog"] = "Ver registro",
            ["menu.clearlog"] = "Borrar registro",
            ["menu.about"] = "Acerca de",
            ["menu.language"] = "Idioma",

            // Chat
            ["chat.title"] = "Z.AI Chat",
            ["chat.send"] = "Enviar",
            ["chat.new"] = "Nuevo chat",
            ["chat.clear"] = "Borrar",
            ["chat.ready"] = "Listo",
            ["chat.processing"] = "Procesando...",
            ["chat.ready_count"] = "Listo ({0} mensajes)",
            ["chat.new_started"] = "Nueva conversación iniciada.",
            ["chat.welcome"] = "¡Bienvenido a Z.AI Chat!\n\nPuedo ayudarte a trabajar con Excel. Intenta preguntar:\n• \"Resume los datos de esta hoja\"\n• \"Crea un gráfico con las columnas A y B\"\n• \"Formatea la fila de encabezado en negrita con fondo azul\"\n• \"Calcula el promedio de la columna C\"",

            // Auth
            ["auth.prompt"] = "Introduce tu clave API:",
            ["auth.current_key"] = "Clave actual: {0}",
            ["auth.login_title"] = "Inicio de sesión Z.AI",
            ["auth.cancelled"] = "Inicio de sesión cancelado.",
            ["auth.validating"] = "Validando clave API...",
            ["auth.success"] = "¡Inicio de sesión exitoso!",
            ["auth.failed"] = "Error de inicio de sesión. Clave API no válida.",
            ["auth.not_logged"] = "No has iniciado sesión. Por favor, inicia sesión primero.",
            ["auth.confirm_logout"] = "¿Estás seguro de que quieres cerrar sesión?",
            ["auth.logged_out"] = "Has cerrado sesión.",

            // Model
            ["model.prompt"] = "Selecciona un modelo:\n1. glm-4-plus (recomendado)\n2. glm-4-long (contexto largo)\n3. glm-4 (estándar)\n4. glm-3-turbo (rápido)",
            ["model.current"] = "Modelo actual: {0}",
            ["model.title"] = "Seleccionar modelo",
            ["model.changed"] = "Modelo cambiado a: {0}",

            // About
            ["about.text"] = "Z.AI Complemento para Excel\nVersión 2.0\n\nAsistente de IA para Microsoft Excel.\nImpulsado por los modelos ZhipuAI GLM.\n\n© 2024 Z.AI",
            ["about.title"] = "Acerca de Z.AI",

            // Conversation
            ["conv.status_round"] = "Ronda {0}",
            ["conv.status_exec"] = "Ejecutando herramienta: {0}",
            ["conv.max_rounds"] = "Número máximo de rondas alcanzado ({0}). Deteniendo.",
            ["conv.loop_detected"] = "Bucle detectado. Deteniendo para evitar ejecución infinita.",
            ["conv.api_error"] = "Error de API. Por favor, inténtalo de nuevo.",
            ["conv.no_response"] = "No se recibió respuesta de la API.",
            ["conv.error"] = "Ocurrió un error durante el procesamiento.",
            ["conv.no_assistant"] = "Sin respuesta del asistente en la respuesta de la API.",

            // Debug
            ["debug.no_log"] = "No hay entradas en el registro.",

            // Language
            ["lang.changed"] = "Idioma cambiado. Algunos cambios pueden requerir reiniciar el complemento.",
            ["lang.title"] = "Idioma",

            // System prompt
            ["system.prompt"] = "Eres un asistente de IA integrado en Microsoft Excel a través del complemento Z.AI. Tienes acceso a herramientas que pueden leer y modificar libros de Excel.\n\nReglas que debes seguir:\n1. Siempre llama primero a get_sheet_info o get_workbook_info para entender el estado actual del libro antes de realizar cualquier acción.\n2. Siempre lee los datos antes de modificarlos. Nunca asumas el contenido de las celdas.\n3. Después de hacer cambios, confirma lo que se hizo releyendo las celdas afectadas.\n4. Escribe todas las fórmulas usando nombres de funciones en inglés (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN, etc.).\n5. Al establecer colores, usa valores RGB Long: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. Trabaja por defecto en la hoja activa, a menos que el usuario indique lo contrario.\n7. Antes de crear gráficos, llama a list_charts. Si ya existe un gráfico similar, elimínalo antes de crear uno nuevo.\n8. No repitas operaciones que ya se completaron con éxito.\n9. Comunícate con el usuario en español.\n10. Sé conciso y útil. Explica paso a paso lo que estás haciendo."
        };
    }

    // ── Ukrainian ────────────────────────────────────────────────────────

    private void InitUkrainian()
    {
        _translations["uk"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "Чат",
            ["menu.login"] = "Увійти",
            ["menu.logout"] = "Вийти",
            ["menu.model"] = "Модель",
            ["menu.viewlog"] = "Переглянути журнал",
            ["menu.clearlog"] = "Очистити журнал",
            ["menu.about"] = "Про програму",
            ["menu.language"] = "Мова",

            // Chat
            ["chat.title"] = "Z.AI Чат",
            ["chat.send"] = "Надіслати",
            ["chat.new"] = "Новий чат",
            ["chat.clear"] = "Очистити",
            ["chat.ready"] = "Готово",
            ["chat.processing"] = "Обробка...",
            ["chat.ready_count"] = "Готово ({0} повідомлень)",
            ["chat.new_started"] = "Розпочато нову розмову.",
            ["chat.welcome"] = "Ласкаво просимо до Z.AI Chat!\n\nЯ можу допомогти вам працювати з Excel. Спробуйте запитати:\n• «Підсумуй дані на цьому аркуші»\n• «Створи діаграму з колонок A та B»\n• «Відформатуй рядок заголовка жирним шрифтом і синім фоном»\n• «Обчисли середнє значення колонки C»",

            // Auth
            ["auth.prompt"] = "Введіть ваш API-ключ:",
            ["auth.current_key"] = "Поточний ключ: {0}",
            ["auth.login_title"] = "Вхід Z.AI",
            ["auth.cancelled"] = "Вхід скасовано.",
            ["auth.validating"] = "Перевірка API-ключа...",
            ["auth.success"] = "Вхід успішний!",
            ["auth.failed"] = "Помилка входу. Недійсний API-ключ.",
            ["auth.not_logged"] = "Ви не увійшли. Будь ласка, спочатку увійдіть.",
            ["auth.confirm_logout"] = "Ви впевнені, що хочете вийти?",
            ["auth.logged_out"] = "Ви вийшли з системи.",

            // Model
            ["model.prompt"] = "Оберіть модель:\n1. glm-4-plus (рекомендовано)\n2. glm-4-long (довгий контекст)\n3. glm-4 (стандартна)\n4. glm-3-turbo (швидка)",
            ["model.current"] = "Поточна модель: {0}",
            ["model.title"] = "Вибір моделі",
            ["model.changed"] = "Модель змінено на: {0}",

            // About
            ["about.text"] = "Z.AI Надбудова для Excel\nВерсія 2.0\n\nАсистент зі штучним інтелектом для Microsoft Excel.\nПрацює на моделях ZhipuAI GLM.\n\n© 2024 Z.AI",
            ["about.title"] = "Про Z.AI",

            // Conversation
            ["conv.status_round"] = "Раунд {0}",
            ["conv.status_exec"] = "Виконання інструменту: {0}",
            ["conv.max_rounds"] = "Досягнуто максимальну кількість раундів ({0}). Зупинка.",
            ["conv.loop_detected"] = "Виявлено цикл. Зупинка для запобігання нескінченному виконанню.",
            ["conv.api_error"] = "Помилка API. Будь ласка, спробуйте ще раз.",
            ["conv.no_response"] = "Відповідь від API не отримано.",
            ["conv.error"] = "Під час обробки сталася помилка.",
            ["conv.no_assistant"] = "Відповідь асистента відсутня у відповіді API.",

            // Debug
            ["debug.no_log"] = "Записів у журналі немає.",

            // Language
            ["lang.changed"] = "Мову змінено. Деякі зміни можуть потребувати перезапуску надбудови.",
            ["lang.title"] = "Мова",

            // System prompt
            ["system.prompt"] = "Ти — асистент зі штучним інтелектом, інтегрований у Microsoft Excel через надбудову Z.AI. Ти маєш доступ до інструментів, які можуть читати та змінювати книги Excel.\n\nПравила, яких ти повинен дотримуватися:\n1. Завжди спочатку викликай get_sheet_info або get_workbook_info, щоб зрозуміти поточний стан книги, перш ніж виконувати будь-які дії.\n2. Завжди читай дані перед їх зміною. Ніколи не припускай вміст комірок.\n3. Після внесення змін підтверджуй, що було зроблено, перечитуючи змінені комірки.\n4. Записуй усі формули з англійськими назвами функцій (SUM, AVERAGE, IF, VLOOKUP, COUNT, MAX, MIN тощо).\n5. При встановленні кольорів використовуй значення RGB Long: Red=255, Green=65280, Blue=16711680, Yellow=65535, White=16777215, Black=0.\n6. За замовчуванням працюй на активному аркуші, якщо користувач не вказав інше.\n7. Перед створенням діаграм викликай list_charts. Якщо подібна діаграма вже існує, видали її перед створенням нової.\n8. Не повторюй операції, які вже були успішно виконані.\n9. Спілкуйся з користувачем українською мовою.\n10. Будь лаконічним і корисним. Пояснюй крок за кроком, що ти робиш."
        };
    }

    // ── Chinese (Simplified) ─────────────────────────────────────────────

    private void InitChinese()
    {
        _translations["zh"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "聊天",
            ["menu.login"] = "登录",
            ["menu.logout"] = "退出登录",
            ["menu.model"] = "模型",
            ["menu.viewlog"] = "查看日志",
            ["menu.clearlog"] = "清除日志",
            ["menu.about"] = "关于",
            ["menu.language"] = "语言",

            // Chat
            ["chat.title"] = "Z.AI 聊天",
            ["chat.send"] = "发送",
            ["chat.new"] = "新聊天",
            ["chat.clear"] = "清除",
            ["chat.ready"] = "就绪",
            ["chat.processing"] = "处理中...",
            ["chat.ready_count"] = "就绪（{0} 条消息）",
            ["chat.new_started"] = "已开始新对话。",
            ["chat.welcome"] = "\u6b22\u8fce\u4f7f\u7528 Z.AI Chat\uff01\n\n\u6211\u53ef\u4ee5\u5e2e\u52a9\u60a8\u4f7f\u7528 Excel\u3002\u8bf7\u5c1d\u8bd5\uff1a\n\u2022 \u603b\u7ed3\u8fd9\u4e2a\u5de5\u4f5c\u8868\u4e2d\u7684\u6570\u636e\n\u2022 \u6839\u636e A \u5217\u548c B \u5217\u521b\u5efa\u56fe\u8868\n\u2022 \u5c06\u6807\u9898\u884c\u8bbe\u7f6e\u4e3a\u7c97\u4f53\u5e76\u6dfb\u52a0\u84dd\u8272\u80cc\u666f\n\u2022 \u8ba1\u7b97 C \u5217\u7684\u5e73\u5747\u503c",

            // Auth
            ["auth.prompt"] = "请输入您的 API 密钥：",
            ["auth.current_key"] = "当前密钥：{0}",
            ["auth.login_title"] = "Z.AI 登录",
            ["auth.cancelled"] = "登录已取消。",
            ["auth.validating"] = "正在验证 API 密钥...",
            ["auth.success"] = "登录成功！",
            ["auth.failed"] = "登录失败。API 密钥无效。",
            ["auth.not_logged"] = "您尚未登录。请先登录。",
            ["auth.confirm_logout"] = "确定要退出登录吗？",
            ["auth.logged_out"] = "您已退出登录。",

            // Model
            ["model.prompt"] = "选择模型：\n1. glm-4-plus（推荐）\n2. glm-4-long（长上下文）\n3. glm-4（标准）\n4. glm-3-turbo（快速）",
            ["model.current"] = "当前模型：{0}",
            ["model.title"] = "选择模型",
            ["model.changed"] = "模型已更改为：{0}",

            // About
            ["about.text"] = "Z.AI Excel 加载项\n版本 2.0\n\n适用于 Microsoft Excel 的 AI 助手。\n由智谱AI GLM 模型提供支持。\n\n© 2024 Z.AI",
            ["about.title"] = "关于 Z.AI",

            // Conversation
            ["conv.status_round"] = "第 {0} 轮",
            ["conv.status_exec"] = "正在执行工具：{0}",
            ["conv.max_rounds"] = "已达到最大轮次（{0}）。正在停止。",
            ["conv.loop_detected"] = "检测到循环。正在停止以防止无限执行。",
            ["conv.api_error"] = "发生 API 错误。请重试。",
            ["conv.no_response"] = "未收到 API 响应。",
            ["conv.error"] = "处理过程中发生错误。",
            ["conv.no_assistant"] = "API 响应中没有助手回复。",

            // Debug
            ["debug.no_log"] = "没有日志记录。",

            // Language
            ["lang.changed"] = "语言已更改。某些更改可能需要重新启动加载项。",
            ["lang.title"] = "语言",

            // System prompt
            ["system.prompt"] = "你是一个通过 Z.AI 加载项集成到 Microsoft Excel 中的 AI 助手。你可以使用工具来读取和修改 Excel 工作簿。\n\n你必须遵循的规则：\n1. 在执行任何操作之前，始终先调用 get_sheet_info 或 get_workbook_info 来了解工作簿的当前状态。\n2. 在修改数据之前始终先读取数据。永远不要假设单元格的内容。\n3. 做出更改后，通过回读受影响的单元格来确认已完成的操作。\n4. 使用英文函数名编写所有公式（SUM、AVERAGE、IF、VLOOKUP、COUNT、MAX、MIN 等）。\n5. 设置颜色时，使用 RGB Long 值：Red=255、Green=65280、Blue=16711680、Yellow=65535、White=16777215、Black=0。\n6. 除非用户另有指定，否则默认在活动工作表上操作。\n7. 创建图表之前，先调用 list_charts。如果已存在类似图表，请先删除再创建新图表。\n8. 不要重复已成功完成的操作。\n9. 使用简体中文与用户交流。\n10. 简洁且有帮助。逐步解释你正在做的事情。"
        };
    }

    // ── Japanese ─────────────────────────────────────────────────────────

    private void InitJapanese()
    {
        _translations["ja"] = new Dictionary<string, string>
        {
            // Menu
            ["menu.chat"] = "チャット",
            ["menu.login"] = "ログイン",
            ["menu.logout"] = "ログアウト",
            ["menu.model"] = "モデル",
            ["menu.viewlog"] = "ログ表示",
            ["menu.clearlog"] = "ログ消去",
            ["menu.about"] = "バージョン情報",
            ["menu.language"] = "言語",

            // Chat
            ["chat.title"] = "Z.AI チャット",
            ["chat.send"] = "送信",
            ["chat.new"] = "新規チャット",
            ["chat.clear"] = "クリア",
            ["chat.ready"] = "準備完了",
            ["chat.processing"] = "処理中...",
            ["chat.ready_count"] = "準備完了（{0} 件のメッセージ）",
            ["chat.new_started"] = "新しい会話を開始しました。",
            ["chat.welcome"] = "Z.AI Chat へようこそ！\n\nExcel での作業をお手伝いします。次のように質問してみてください：\n• 「このシートのデータを要約して」\n• 「A列とB列からグラフを作成して」\n• 「ヘッダー行を太字にして青い背景にして」\n• 「C列の平均値を計算して」",

            // Auth
            ["auth.prompt"] = "APIキーを入力してください：",
            ["auth.current_key"] = "現在のキー：{0}",
            ["auth.login_title"] = "Z.AI ログイン",
            ["auth.cancelled"] = "ログインがキャンセルされました。",
            ["auth.validating"] = "APIキーを検証中...",
            ["auth.success"] = "ログイン成功！",
            ["auth.failed"] = "ログイン失敗。APIキーが無効です。",
            ["auth.not_logged"] = "ログインしていません。先にログインしてください。",
            ["auth.confirm_logout"] = "ログアウトしてもよろしいですか？",
            ["auth.logged_out"] = "ログアウトしました。",

            // Model
            ["model.prompt"] = "モデルを選択：\n1. glm-4-plus（推奨）\n2. glm-4-long（長文コンテキスト）\n3. glm-4（標準）\n4. glm-3-turbo（高速）",
            ["model.current"] = "現在のモデル：{0}",
            ["model.title"] = "モデル選択",
            ["model.changed"] = "モデルを変更しました：{0}",

            // About
            ["about.text"] = "Z.AI Excel アドイン\nバージョン 2.0\n\nMicrosoft Excel 用 AI アシスタント。\nZhipuAI GLM モデルで動作。\n\n© 2024 Z.AI",
            ["about.title"] = "Z.AI について",

            // Conversation
            ["conv.status_round"] = "ラウンド {0}",
            ["conv.status_exec"] = "ツール実行中：{0}",
            ["conv.max_rounds"] = "最大ラウンド数に達しました（{0}）。停止します。",
            ["conv.loop_detected"] = "ループを検出しました。無限実行を防ぐため停止します。",
            ["conv.api_error"] = "APIエラーが発生しました。もう一度お試しください。",
            ["conv.no_response"] = "APIから応答がありませんでした。",
            ["conv.error"] = "処理中にエラーが発生しました。",
            ["conv.no_assistant"] = "APIレスポンスにアシスタントの応答がありません。",

            // Debug
            ["debug.no_log"] = "ログエントリがありません。",

            // Language
            ["lang.changed"] = "言語が変更されました。一部の変更はアドインの再起動が必要な場合があります。",
            ["lang.title"] = "言語",

            // System prompt
            ["system.prompt"] = "あなたは Z.AI アドインを通じて Microsoft Excel に統合された AI アシスタントです。Excel ブックを読み取りおよび変更できるツールにアクセスできます。\n\n従うべきルール：\n1. 操作を行う前に、必ず最初に get_sheet_info または get_workbook_info を呼び出して、ブックの現在の状態を把握してください。\n2. データを変更する前に必ず読み取ってください。セルの内容を推測してはいけません。\n3. 変更を行った後、影響を受けたセルを読み返して、何が行われたかを確認してください。\n4. すべての数式は英語の関数名で記述してください（SUM、AVERAGE、IF、VLOOKUP、COUNT、MAX、MIN など）。\n5. 色を設定する際は、RGB Long 値を使用してください：Red=255、Green=65280、Blue=16711680、Yellow=65535、White=16777215、Black=0。\n6. ユーザーが特に指定しない限り、アクティブシートをデフォルトとしてください。\n7. グラフを作成する前に list_charts を呼び出してください。類似のグラフが既に存在する場合は、新しいグラフを作成する前に削除してください。\n8. 既に正常に完了した操作を繰り返さないでください。\n9. ユーザーとは日本語でコミュニケーションしてください。\n10. 簡潔で役立つ応答をしてください。何をしているかをステップごとに説明してください。"
        };
    }

    // ── Extra keys (added to all languages) ──────────────────────────────

    private void AddExtraKeys()
    {
        var extras = new Dictionary<string, Dictionary<string, string>>
        {
            ["en"] = new()
            {
                ["ribbon.group_main"] = "Assistant",
                ["ribbon.group_info"] = "Status",
                ["ribbon.group_tools"] = "Tools",
                ["ribbon.logged_in"] = "Logged in",
                ["ribbon.not_logged"] = "Not logged in",
                ["menu.add_tokens"] = "Add Tokens",
                ["auth.show_key"] = "Show key",
                ["auth.open_site"] = "Get API Key on z.ai",
                ["select.cancel"] = "Cancel",
                ["error.balance_empty"] = "⚠️ Your API balance is empty.\nPlease add tokens at z.ai to continue using the AI assistant.",
                ["error.content_filter"] = "⚠️ Your message was blocked by the content filter.\nPlease rephrase and try again.",
                ["error.invalid_key"] = "⚠️ Invalid API key.\nPlease check your key and log in again.",
                ["error.rate_limit"] = "⚠️ Too many requests.\nPlease wait a moment and try again.",
                ["error.api_generic"] = "⚠️ An API error occurred.",
                ["error.ctp_hint"] = "Try restarting Excel.\nIf the error persists, the chat panel may not be supported in this Excel version.",
                ["about.text"] = "Z.AI Excel Add-in\nVersion 2.0\n\nAI-powered assistant for Microsoft Excel.\nPowered by ZhipuAI GLM models.\n\n⚠️ DISCLAIMER:\nThis is an UNOFFICIAL add-in.\nNot affiliated with, endorsed by, or\nassociated with z.ai / Zhipu AI in any way.\nAll rights belong to their respective owners.",
                ["model.prompt"] = "Select an AI model.\n💰 = paid, ⚡ = cheap/free:",
                ["lang.select_prompt"] = "Select interface language:",
            },
            ["pl"] = new()
            {
                ["ribbon.group_main"] = "Asystent",
                ["ribbon.group_info"] = "Status",
                ["ribbon.group_tools"] = "Narzędzia",
                ["ribbon.logged_in"] = "Zalogowano",
                ["ribbon.not_logged"] = "Nie zalogowano",
                ["menu.add_tokens"] = "Doładuj tokeny",
                ["auth.show_key"] = "Pokaż klucz",
                ["auth.open_site"] = "Pobierz klucz API na z.ai",
                ["select.cancel"] = "Anuluj",
                ["error.balance_empty"] = "⚠️ Twoje saldo API jest puste.\nDoładuj tokeny na z.ai, aby kontynuować korzystanie z asystenta AI.",
                ["error.content_filter"] = "⚠️ Twoja wiadomość została zablokowana przez filtr treści.\nPrzeformułuj i spróbuj ponownie.",
                ["error.invalid_key"] = "⚠️ Nieprawidłowy klucz API.\nSprawdź klucz i zaloguj się ponownie.",
                ["error.rate_limit"] = "⚠️ Zbyt wiele zapytań.\nPoczekaj chwilę i spróbuj ponownie.",
                ["error.api_generic"] = "⚠️ Wystąpił błąd API.",
                ["error.ctp_hint"] = "Spróbuj zrestartować Excela.\nJeśli błąd się powtarza, panel czatu może nie być obsługiwany w tej wersji Excela.",
                ["about.text"] = "Z.AI Dodatek do Excela\nWersja 2.0\n\nAsystent AI dla Microsoft Excel.\nWykorzystuje modele ZhipuAI GLM.\n\n⚠️ ZASTRZEŻENIE:\nTo jest NIEOFICJALNY dodatek.\nNie jest powiązany z, zatwierdzony przez,\nani stowarzyszony z z.ai / Zhipu AI.\nWszelkie prawa należą do ich właścicieli.",
                ["model.prompt"] = "Wybierz model AI.\n💰 = płatny, ⚡ = tani/darmowy:",
                ["lang.select_prompt"] = "Wybierz język interfejsu:",
            },
            ["de"] = new()
            {
                ["ribbon.group_main"] = "Assistent",
                ["ribbon.group_info"] = "Status",
                ["ribbon.group_tools"] = "Werkzeuge",
                ["ribbon.logged_in"] = "Eingeloggt",
                ["ribbon.not_logged"] = "Nicht eingeloggt",
                ["menu.add_tokens"] = "Tokens aufladen",
                ["auth.show_key"] = "Schlüssel anzeigen",
                ["auth.open_site"] = "API-Schlüssel auf z.ai holen",
                ["select.cancel"] = "Abbrechen",
                ["error.balance_empty"] = "⚠️ Ihr API-Guthaben ist leer.\nBitte laden Sie Tokens auf z.ai auf.",
                ["error.content_filter"] = "⚠️ Ihre Nachricht wurde vom Inhaltsfilter blockiert.",
                ["error.invalid_key"] = "⚠️ Ungültiger API-Schlüssel.",
                ["error.rate_limit"] = "⚠️ Zu viele Anfragen. Bitte warten.",
                ["error.api_generic"] = "⚠️ Ein API-Fehler ist aufgetreten.",
                ["error.ctp_hint"] = "Versuchen Sie Excel neu zu starten.",
                ["about.text"] = "Z.AI Excel Add-in\nVersion 2.0\n\nKI-Assistent für Microsoft Excel.\nBetrieben mit ZhipuAI GLM-Modellen.\n\n⚠️ HAFTUNGSAUSSCHLUSS:\nDies ist ein INOFFIZIELLES Add-in.\nNicht verbunden mit z.ai / Zhipu AI.",
                ["model.prompt"] = "KI-Modell wählen.\n💰 = kostenpflichtig, ⚡ = günstig/kostenlos:",
                ["lang.select_prompt"] = "Sprache wählen:",
            },
            ["fr"] = new()
            {
                ["ribbon.group_main"] = "Assistant",
                ["ribbon.group_info"] = "Statut",
                ["ribbon.group_tools"] = "Outils",
                ["ribbon.logged_in"] = "Connecté",
                ["ribbon.not_logged"] = "Non connecté",
                ["menu.add_tokens"] = "Recharger les tokens",
                ["auth.show_key"] = "Afficher la clé",
                ["auth.open_site"] = "Obtenir une clé API sur z.ai",
                ["select.cancel"] = "Annuler",
                ["error.balance_empty"] = "⚠️ Votre solde API est vide.\nRechargez vos tokens sur z.ai.",
                ["error.content_filter"] = "⚠️ Votre message a été bloqué par le filtre.",
                ["error.invalid_key"] = "⚠️ Clé API invalide.",
                ["error.rate_limit"] = "⚠️ Trop de requêtes. Veuillez patienter.",
                ["error.api_generic"] = "⚠️ Une erreur API s'est produite.",
                ["error.ctp_hint"] = "Essayez de redémarrer Excel.",
                ["about.text"] = "Z.AI Complément Excel\nVersion 2.0\n\nAssistant IA pour Microsoft Excel.\nAlimenté par les modèles ZhipuAI GLM.\n\n⚠️ AVERTISSEMENT:\nCeci est un complément NON OFFICIEL.\nNon affilié à z.ai / Zhipu AI.",
                ["model.prompt"] = "Sélectionnez un modèle IA.\n💰 = payant, ⚡ = économique/gratuit:",
                ["lang.select_prompt"] = "Sélectionnez la langue:",
            },
            ["es"] = new()
            {
                ["ribbon.group_main"] = "Asistente",
                ["ribbon.group_info"] = "Estado",
                ["ribbon.group_tools"] = "Herramientas",
                ["ribbon.logged_in"] = "Conectado",
                ["ribbon.not_logged"] = "No conectado",
                ["menu.add_tokens"] = "Recargar tokens",
                ["auth.show_key"] = "Mostrar clave",
                ["auth.open_site"] = "Obtener clave API en z.ai",
                ["select.cancel"] = "Cancelar",
                ["error.balance_empty"] = "⚠️ Su saldo API está vacío.\nRecargue tokens en z.ai.",
                ["error.content_filter"] = "⚠️ Su mensaje fue bloqueado por el filtro.",
                ["error.invalid_key"] = "⚠️ Clave API inválida.",
                ["error.rate_limit"] = "⚠️ Demasiadas solicitudes. Espere un momento.",
                ["error.api_generic"] = "⚠️ Ocurrió un error de API.",
                ["error.ctp_hint"] = "Intente reiniciar Excel.",
                ["about.text"] = "Z.AI Complemento de Excel\nVersión 2.0\n\nAsistente de IA para Microsoft Excel.\nImpulsado por modelos ZhipuAI GLM.\n\n⚠️ AVISO:\nEste es un complemento NO OFICIAL.\nNo afiliado a z.ai / Zhipu AI.",
                ["model.prompt"] = "Seleccione un modelo de IA.\n💰 = de pago, ⚡ = económico/gratis:",
                ["lang.select_prompt"] = "Seleccione el idioma:",
            },
            ["uk"] = new()
            {
                ["ribbon.group_main"] = "Асистент",
                ["ribbon.group_info"] = "Статус",
                ["ribbon.group_tools"] = "Інструменти",
                ["ribbon.logged_in"] = "Увійшли",
                ["ribbon.not_logged"] = "Не увійшли",
                ["menu.add_tokens"] = "Поповнити токени",
                ["auth.show_key"] = "Показати ключ",
                ["auth.open_site"] = "Отримати ключ API на z.ai",
                ["select.cancel"] = "Скасувати",
                ["error.balance_empty"] = "⚠️ Ваш баланс API порожній.\nПоповніть токени на z.ai.",
                ["error.content_filter"] = "⚠️ Ваше повідомлення заблоковано фільтром контенту.",
                ["error.invalid_key"] = "⚠️ Недійсний ключ API.",
                ["error.rate_limit"] = "⚠️ Забагато запитів. Зачекайте.",
                ["error.api_generic"] = "⚠️ Помилка API.",
                ["error.ctp_hint"] = "Спробуйте перезапустити Excel.",
                ["about.text"] = "Z.AI Додаток для Excel\nВерсія 2.0\n\nAI-асистент для Microsoft Excel.\nПрацює на моделях ZhipuAI GLM.\n\n⚠️ ЗАСТЕРЕЖЕННЯ:\nЦе НЕОФІЦІЙНИЙ додаток.\nНе пов'язаний з z.ai / Zhipu AI.",
                ["model.prompt"] = "Оберіть модель AI.\n\U0001f4b0 = платна, ⚡ = дешева/безкоштовна:",
                ["lang.select_prompt"] = "Оберіть мову інтерфейсу:",
            },
            ["zh"] = new()
            {
                ["ribbon.group_main"] = "助手",
                ["ribbon.group_info"] = "状态",
                ["ribbon.group_tools"] = "工具",
                ["ribbon.logged_in"] = "已登录",
                ["ribbon.not_logged"] = "未登录",
                ["menu.add_tokens"] = "充值令牌",
                ["auth.show_key"] = "显示密钥",
                ["auth.open_site"] = "在 z.ai 获取 API 密钥",
                ["select.cancel"] = "取消",
                ["error.balance_empty"] = "⚠️ 您的API余额已用完。\n请在 z.ai 充值以继续使用。",
                ["error.content_filter"] = "⚠️ 您的消息被内容过滤器拦截。\n请修改后重试。",
                ["error.invalid_key"] = "⚠️ API密钥无效。\n请检查密钥并重新登录。",
                ["error.rate_limit"] = "⚠️ 请求过多。\n请稍后再试。",
                ["error.api_generic"] = "⚠️ 发生API错误。",
                ["error.ctp_hint"] = "请尝试重启Excel。",
                ["about.text"] = "Z.AI Excel 插件\n版本 2.0\n\n适用于 Microsoft Excel 的 AI 助手。\n由 ZhipuAI GLM 模型驱动。\n\n⚠️ 免责声明：\n这是一个非官方插件。\n与 z.ai / 智谱AI 无关。",
                ["model.prompt"] = "选择AI模型。\n\U0001f4b0 = 付费，⚡ = 便宜/免费：",
                ["lang.select_prompt"] = "选择界面语言：",
            },
            ["ja"] = new()
            {
                ["ribbon.group_main"] = "アシスタント",
                ["ribbon.group_info"] = "ステータス",
                ["ribbon.group_tools"] = "ツール",
                ["ribbon.logged_in"] = "ログイン中",
                ["ribbon.not_logged"] = "未ログイン",
                ["menu.add_tokens"] = "トークンを追加",
                ["auth.show_key"] = "キーを表示",
                ["auth.open_site"] = "z.ai で API キーを取得",
                ["select.cancel"] = "キャンセル",
                ["error.balance_empty"] = "⚠️ APIの残高がなくなりました。\nz.ai でトークンを追加してください。",
                ["error.content_filter"] = "⚠️ メッセージがコンテンツフィルターでブロックされました。",
                ["error.invalid_key"] = "⚠️ APIキーが無効です。",
                ["error.rate_limit"] = "⚠️ リクエストが多すぎます。しばらくお待ちください。",
                ["error.api_generic"] = "⚠️ APIエラーが発生しました。",
                ["error.ctp_hint"] = "Excelを再起動してみてください。",
                ["about.text"] = "Z.AI Excel アドイン\nバージョン 2.0\n\nMicrosoft Excel 用 AI アシスタント。\nZhipuAI GLM モデルで動作。\n\n⚠️ 免責事項：\nこれは非公式アドインです。\nz.ai / Zhipu AI とは無関係です。",
                ["model.prompt"] = "AIモデルを選択。\n\U0001f4b0 = 有料、⚡ = 安い/無料：",
                ["lang.select_prompt"] = "インターフェース言語を選択：",
            }
        };

        foreach (var (lang, keys) in extras)
        {
            if (!_translations.ContainsKey(lang)) continue;
            foreach (var (key, val) in keys)
                _translations[lang][key] = val;
        }
    }
}
