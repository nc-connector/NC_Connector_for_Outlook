/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Globalization;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Einfache Lokalisierungs-Hilfe fuer sichtbare UI-Texte (de/en/fr).
     */
    internal static class Strings
    {
        private static readonly string[] LanguageCandidates = BuildLanguageCandidates();
        private static readonly Dictionary<string, Dictionary<string, string>> Translations = BuildTranslations();

        private static string[] BuildLanguageCandidates()
        {
            var list = new List<string>();

            TryAddCulture(list, () => CultureInfo.CurrentUICulture);
            TryAddCulture(list, () => CultureInfo.CurrentCulture);
            TryAddCulture(list, () => CultureInfo.InstalledUICulture);

            if (!list.Contains("en"))
            {
                list.Add("en");
            }

            return list.ToArray();
        }

        private static void TryAddCulture(List<string> list, Func<CultureInfo> provider)
        {
            try
            {
                CultureInfo culture = provider();
                if (culture == null)
                {
                    return;
                }

                string code = culture.TwoLetterISOLanguageName.ToLowerInvariant();
                if (!string.IsNullOrEmpty(code) && !list.Contains(code))
                {
                    list.Add(code);
                }
            }
            catch
            {
            }
        }

        private static Dictionary<string, Dictionary<string, string>> BuildTranslations()
        {
            var translations = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            var german = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            german["SettingsFormTitle"] = "Nextcloud Enterprise for Outlook - Einstellungen";
            german["TabGeneral"] = "Allgemein";
            german["TabIfb"] = "IFB";
            german["TabAdvanced"] = "Erweitert";
            german["TabInvitations"] = "Einladungen";
            german["TabDebug"] = "Debug";
            german["TabAbout"] = "Über";
            german["TabFileLink"] = "Filelink";
            german["LabelServerUrl"] = "Server-URL";
            german["LabelUsername"] = "Benutzername";
            german["LabelAppPassword"] = "App-Passwort";
            german["GroupAuthentication"] = "Authentifizierung";
            german["RadioManual"] = "Manuell (Benutzername + App-Passwort eingeben)";
            german["RadioLoginFlow"] = "Login mit Nextcloud (App-Passwort automatisch holen)";
            german["ButtonLoginFlow"] = "Login mit Nextcloud...";
            german["ButtonTestConnection"] = "Verbindung testen";
            german["GroupMuzzle"] = "Maulkorb fuer Outlook";
            german["CheckMuzzle"] = "Outlook-Einladungen unterdruecken (Versand erfolgt ueber Nextcloud)";
            german["LabelMuzzleHint"] = "Hinweis: Antworten auf Einladungen sowie normale E-Mails werden weiterhin versendet.";
            german["LabelMuzzleAccountsHint"] = "Waehlen Sie die Konten aus, deren Outlook-Einladungen unterdrueckt werden sollen.";
            german["CheckIfbEnabled"] = "IFB-Endpunkt aktivieren";
            german["LabelIfbDays"] = "Zeitraum (Tage):";
            german["LabelIfbCacheHours"] = "Adressbuch-Cache (Stunden):";
            german["ButtonSave"] = "Speichern";
            german["ButtonCancel"] = "Abbrechen";
            german["DebugCheckbox"] = "Debug-Logging in Datei schreiben";
            german["DebugPathPrefix"] = "Ablage: ";
            german["DebugOpenLog"] = "Logdatei oeffnen";
            german["DebugLogMissingMessage"] = "Es wurde noch keine Logdatei erzeugt.";
            german["DebugLogOpenErrorMessage"] = "Logdatei konnte nicht geoeffnet werden.";
            german["ButtonNext"] = "Weiter";
            german["ButtonBack"] = "Zur\u00fcck";
            german["TalkFormTitle"] = "Talk-Link erstellen";
            german["TalkTitleLabel"] = "Titel";
            german["TalkPasswordLabel"] = "Passwort (optional)";
            german["TalkLobbyCheck"] = "Lobby bis Startzeit aktiv";
            german["TalkSearchCheck"] = "In Nextcloud Suche anzeigen";
            german["TalkRoomGroup"] = "Raumtyp";
            german["TalkEventRadio"] = "Event-Konversation (Termin binden)";
            german["TalkStandardRadio"] = "Standard-Raum (eigenstaendig)";
            german["TalkVersionUnknown"] = "unbekannt";
            german["TalkEventHint"] = "Event-Konversationen benoetigen Nextcloud 31 oder neuer.\nErkannte Version: {0}";
            german["DialogOk"] = "OK";
            german["DialogCancel"] = "Abbrechen";
            german["TalkPasswordTooShort"] = "Das Passwort muss mindestens 5 Zeichen lang sein.";
            german["TalkDefaultTitle"] = "Besprechung";
            german["ErrorMissingCredentials"] = "Bitte hinterlege zuerst Server-URL, Benutzername und App-Passwort in den Einstellungen.";
            german["ErrorNoAppointment"] = "Es konnte kein Termin-Element ermittelt werden.";
            german["ConfirmReplaceRoom"] = "Fuer diesen Termin existiert bereits ein Talk-Raum. Soll er ersetzt werden?";
            german["ErrorCreateRoom"] = "Talk-Raum konnte nicht erstellt werden: {0}";
            german["ErrorCreateRoomUnexpected"] = "Unerwarteter Fehler beim Erstellen des Talk-Raums: {0}";
            german["InfoRoomCreated"] = "Der Talk-Raum \"{0}\" wurde erstellt.";
            german["PromptOpenSettings"] = "{0}\n\nEinstellungen jetzt oeffnen?";
            german["ErrorServerUnavailable"] = "Der Nextcloud Server ist derzeit nicht erreichbar. Bitte pruefe die Internetverbindung.";
            german["ErrorAuthenticationRejected"] = "Anmeldedaten werden nicht akzeptiert: {0}";
            german["ErrorConnectionFailed"] = "Verbindung zum Nextcloud Server fehlgeschlagen: {0}";
            german["ErrorUnknownAuthentication"] = "Unbekannter Fehler bei der Authentifizierung: {0}";
            german["WarningIfbStartFailed"] = "IFB konnte nicht gestartet werden: {0}";
            german["WarningRoomDeleteFailed"] = "Talk-Raum konnte nicht geloescht werden: {0}";
            german["WarningLobbyUpdateFailed"] = "Lobby-Zeit konnte nicht aktualisiert werden: {0}";
            german["WarningDescriptionUpdateFailed"] = "Raumbeschreibung konnte nicht aktualisiert werden: {0}";
            german["RibbonAppointmentGroupLabel"] = "Nextcloud Enterprise";
            german["RibbonTalkButtonLabel"] = "Talk-Link einfuegen";
            german["RibbonTalkButtonScreenTip"] = "Nextcloud Talk-Unterhaltung anlegen";
            german["RibbonTalkButtonSuperTip"] = "Erzeugt eine Nextcloud Talk Unterhaltung fuer diesen Termin.";
            german["RibbonExplorerTabLabel"] = "Nextcloud Enterprise";
            german["RibbonExplorerGroupLabel"] = "Konfiguration";
            german["RibbonSettingsButtonLabel"] = "Einstellungen";
            german["RibbonSettingsScreenTip"] = "Nextcloud Enterprise Einstellungen";
            german["RibbonSettingsSuperTip"] = "Server- und Zugangsdaten verwalten.";
            german["RibbonMailGroupLabel"] = "Nextcloud Enterprise";
            german["RibbonFileLinkButtonLabel"] = "Nextcloud Freigabe hinzuf\u00fcgen";
            german["RibbonFileLinkButtonScreenTip"] = "Freigabelink aus Nextcloud einf\u00fcgen";
            german["RibbonFileLinkButtonSuperTip"] = "Erstellt einen Nextcloud Freigabelink f\u00fcr ausgew\u00e4hlte Dateien und f\u00fcgt ihn der E-Mail hinzu.";
            german["ErrorNoMailItem"] = "Es ist kein E-Mail-Element aktiv.";
            german["StatusLoginFlowStarting"] = "Login-Flow wird gestartet...";
            german["StatusLoginFlowBrowser"] = "Browser geoeffnet. Bitte Login in Nextcloud bestaetigen...";
            german["StatusLoginFlowSuccess"] = "Anmeldung erfolgreich. App-Passwort uebernommen.";
            german["StatusLoginFlowFailure"] = "Login fehlgeschlagen: {0}";
            german["StatusMissingFields"] = "Bitte Server, Benutzername und App-Passwort angeben.";
            german["StatusServerUrlRequired"] = "Bitte zuerst die Server-URL angeben.";
            german["StatusInvalidServerUrl"] = "Bitte eine gueltige Server-URL angeben.";
            german["StatusTestRunning"] = "Verbindungstest laeuft...";
            german["StatusTestSuccessFormat"] = "Verbindungstest erfolgreich{0}";
            german["StatusTestSuccessVersionFormat"] = "Nextcloud {0}";
            german["StatusTestFailureUnknown"] = "Unbekannter Fehler";
            german["StatusTestFailure"] = "Verbindungstest fehlgeschlagen: {0}";
            german["DialogTitle"] = "Nextcloud Enterprise for Outlook";
            german["AboutVersionFormat"] = "Version: {0}";
            german["AboutLicenseLabel"] = "Lizenz:";
            german["AboutLicenseLink"] = "AGPL v3";
            german["LicenseFileMissingMessage"] = "Die Lizenzdatei konnte nicht gefunden werden:\r\n{0}";
            german["LicenseFileOpenErrorMessage"] = "Lizenzdatei konnte nicht geoeffnet werden: {0}";
            german["AboutCopyright"] = "Copyright 2025 Bastian Kleinschmidt";
            german["FileLinkBaseLabel"] = "Basisverzeichnis";
            german["FileLinkBaseHint"] = "Dateien werden unterhalb dieses Verzeichnisses abgelegt (z.B. \"90 Freigaben - extern\").";
            german["FileLinkDownloadLabel"] = "Downloadlink";
            german["FileLinkPasswordLabel"] = "Passwort";
            german["FileLinkExpireLabel"] = "Ablaufdatum";
            german["FileLinkPermissionsLabel"] = "Ihre Berechtigungen";
            german["FileLinkHtmlIntro"] = "Ich m\u00f6chte Dateien sicher und unter Wahrung Ihrer Privatsph\u00e4re mit Ihnen teilen. Klicken Sie auf den untenstehenden Link, um Ihre Dateien herunterzuladen.";
            german["FileLinkHtmlFooter"] = "{0} ist eine L\u00f6sung f\u00fcr sicheren E-Mail- und Datenaustausch.";
            german["FileLinkNoFilesConfirm"] = "Es wurden keine Dateien oder Ordner zur Freigabe hinzugef\u00fcgt.\r\nDer Empf\u00e4nger kann lediglich eigene Dateien hochladen.\r\nM\u00f6chten Sie trotzdem fortfahren?";
            translations["de"] = german;

            var french = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            french["SettingsFormTitle"] = "Nextcloud Enterprise for Outlook - Parametres";
            french["TabGeneral"] = "General";
            french["TabIfb"] = "IFB";
            french["TabAdvanced"] = "Avance";
            french["TabInvitations"] = "Invitations";
            french["TabDebug"] = "Debogage";
            french["TabAbout"] = "À propos";
            french["TabFileLink"] = "Lien fichier";
            french["LabelServerUrl"] = "URL du serveur";
            french["LabelUsername"] = "Nom d'utilisateur";
            french["LabelAppPassword"] = "Mot de passe d'application";
            french["GroupAuthentication"] = "Authentification";
            french["RadioManual"] = "Manuel (saisir utilisateur et mot de passe d'application)";
            french["RadioLoginFlow"] = "Connexion avec Nextcloud (recuperation automatique)";
            french["ButtonLoginFlow"] = "Connexion avec Nextcloud...";
            french["ButtonTestConnection"] = "Tester la connexion";
            french["GroupMuzzle"] = "Mode silence Outlook";
            french["CheckMuzzle"] = "Bloquer les invitations Outlook (envoyees via Nextcloud)";
            french["LabelMuzzleHint"] = "Remarque : les reponses et e-mails ordinaires sont toujours envoyes.";
            french["LabelMuzzleAccountsHint"] = "Selectionnez les comptes dont il faut bloquer les invitations Outlook.";
            french["CheckIfbEnabled"] = "Activer l'endpoint IFB";
            french["LabelIfbDays"] = "Periode (jours) :";
            french["LabelIfbCacheHours"] = "Cache du carnet (heures) :";
            french["ButtonSave"] = "Enregistrer";
            french["ButtonCancel"] = "Annuler";
            french["DebugCheckbox"] = "Ecrire un journal de debogage";
            french["DebugPathPrefix"] = "Emplacement : ";
            french["DebugOpenLog"] = "Ouvrir le journal";
            french["DebugLogMissingMessage"] = "Aucun journal cree pour le moment.";
            french["DebugLogOpenErrorMessage"] = "Impossible d ouvrir le fichier journal.";
            french["TalkFormTitle"] = "Creer un lien Talk";
            french["TalkTitleLabel"] = "Titre";
            french["TalkPasswordLabel"] = "Mot de passe (optionnel)";
            french["TalkLobbyCheck"] = "Salle d attente active jusqu au debut";
            french["TalkSearchCheck"] = "Afficher dans la recherche Nextcloud";
            french["TalkRoomGroup"] = "Type de salle";
            french["TalkEventRadio"] = "Conversation evenement (lier au rendez-vous)";
            french["TalkStandardRadio"] = "Salle standard (independante)";
            french["TalkVersionUnknown"] = "inconnu";
            french["TalkEventHint"] = "Les conversations evenement requierent Nextcloud 31 ou plus.\nVersion detectee : {0}";
            french["DialogOk"] = "OK";
            french["DialogCancel"] = "Annuler";
            french["TalkPasswordTooShort"] = "Le mot de passe doit contenir au moins 5 caracteres.";
            french["TalkDefaultTitle"] = "Reunion";
            french["ErrorMissingCredentials"] = "Merci de definir URL serveur, nom d utilisateur et mot de passe application dans les parametres.";
            french["ErrorNoAppointment"] = "Aucun element de rendez-vous trouve.";
            french["ConfirmReplaceRoom"] = "Une salle Talk existe deja pour ce rendez-vous. La remplacer ?";
            french["ErrorCreateRoom"] = "Creation de la salle Talk impossible : {0}";
            french["ErrorCreateRoomUnexpected"] = "Erreur inattendue lors de la creation de la salle Talk : {0}";
            french["InfoRoomCreated"] = "La salle Talk \"{0}\" a ete creee.";
            french["PromptOpenSettings"] = "{0}\n\nOuvrir les parametres maintenant ?";
            french["ErrorServerUnavailable"] = "Le serveur Nextcloud est indisponible. Verifier la connexion Internet.";
            french["ErrorAuthenticationRejected"] = "Identifiants refuses : {0}";
            french["ErrorConnectionFailed"] = "Connexion au serveur Nextcloud echouee : {0}";
            french["ErrorUnknownAuthentication"] = "Erreur inconnue lors de l authentification : {0}";
            french["WarningIfbStartFailed"] = "Impossible de demarrer l IFB : {0}";
            french["WarningRoomDeleteFailed"] = "Suppression de la salle Talk impossible : {0}";
            french["WarningLobbyUpdateFailed"] = "Mise a jour de la salle d attente impossible : {0}";
            french["WarningDescriptionUpdateFailed"] = "Mise a jour de la description impossible : {0}";
            french["RibbonAppointmentGroupLabel"] = "Nextcloud Enterprise";
            french["RibbonTalkButtonLabel"] = "Inserer lien Talk";
            french["RibbonTalkButtonScreenTip"] = "Creer une conversation Nextcloud Talk";
            french["RibbonTalkButtonSuperTip"] = "Cree une conversation Nextcloud Talk pour ce rendez-vous.";
            french["RibbonExplorerTabLabel"] = "Nextcloud Enterprise";
            french["RibbonExplorerGroupLabel"] = "Configuration";
            french["RibbonSettingsButtonLabel"] = "Parametres";
            french["RibbonSettingsScreenTip"] = "Parametres Nextcloud Enterprise";
            french["RibbonSettingsSuperTip"] = "Gerer serveur et identifiants.";
            french["RibbonMailGroupLabel"] = "Nextcloud Enterprise";
            french["RibbonFileLinkButtonLabel"] = "Ajouter un lien Nextcloud";
            french["RibbonFileLinkButtonScreenTip"] = "Inserer un lien de partage Nextcloud";
            french["RibbonFileLinkButtonSuperTip"] = "Cree un lien de partage Nextcloud pour les fichiers selectionnes et l'ajoute au message.";
            french["ErrorNoMailItem"] = "Aucun message actif.";
            french["StatusLoginFlowStarting"] = "Demarrage du flux de connexion...";
            french["StatusLoginFlowBrowser"] = "Navigateur ouvert. Merci de confirmer dans Nextcloud...";
            french["StatusLoginFlowSuccess"] = "Connexion reussie. Mot de passe d'application importe.";
            french["StatusLoginFlowFailure"] = "Echec de la connexion : {0}";
            french["StatusMissingFields"] = "Veuillez indiquer serveur, utilisateur et mot de passe d'application.";
            french["StatusServerUrlRequired"] = "Veuillez d'abord saisir l'URL du serveur.";
            french["StatusInvalidServerUrl"] = "Veuillez saisir une URL de serveur valide.";
            french["StatusTestRunning"] = "Test de connexion en cours...";
            french["StatusTestSuccessFormat"] = "Connexion reussie{0}";
            french["StatusTestSuccessVersionFormat"] = "Nextcloud {0}";
            french["StatusTestFailureUnknown"] = "Erreur inconnue";
            french["StatusTestFailure"] = "Echec du test de connexion : {0}";
            french["DialogTitle"] = "Nextcloud Enterprise for Outlook";
            french["AboutVersionFormat"] = "Version : {0}";
            french["AboutLicenseLabel"] = "Licence :";
            french["AboutLicenseLink"] = "AGPL v3";
            french["LicenseFileMissingMessage"] = "Fichier de licence introuvable :\r\n{0}";
            french["LicenseFileOpenErrorMessage"] = "Impossible d'ouvrir le fichier de licence : {0}";
            french["AboutCopyright"] = "Copyright 2025 Bastian Kleinschmidt";
            french["FileLinkBaseLabel"] = "Repertoire de base";
            french["FileLinkBaseHint"] = "Les fichiers sont stockes sous ce repertoire (ex. \"90 Freigaben - extern\").";
            french["FileLinkDownloadLabel"] = "Lien de t\u00e9l\u00e9chargement";
            french["FileLinkPasswordLabel"] = "Mot de passe";
            french["FileLinkExpireLabel"] = "Date d'expiration";
            french["FileLinkPermissionsLabel"] = "Vos autorisations";
            french["FileLinkHtmlIntro"] = "Je souhaite partager des fichiers de mani\u00e8re s\u00fbre et confidentielle. Cliquez sur le lien ci-dessous pour t\u00e9l\u00e9charger vos fichiers.";
            french["FileLinkHtmlFooter"] = "{0} est une solution pour des \u00e9changes d'e-mails et de fichiers s\u00e9curis\u00e9s.";
            french["FileLinkNoFilesConfirm"] = "Aucun fichier ni dossier n'a \u00e9t\u00e9 ajout\u00e9 \u00e0 ce partage.\r\nLe destinataire pourra uniquement envoyer ses propres fichiers.\r\nVoulez-vous continuer ?";
            translations["fr"] = french;

            return translations;
        }

        private static string Get(string key, string defaultValue)
        {
            try
            {
                foreach (string code in LanguageCandidates)
                {
                    Dictionary<string, string> dictionary;
                    if (Translations.TryGetValue(code, out dictionary))
                    {
                        string value;
                        if (dictionary.TryGetValue(key, out value))
                        {
                            return value;
                        }
                    }
                }
            }
            catch
            {
            }

            return defaultValue;
        }

        internal static string SettingsFormTitle { get { return Get("SettingsFormTitle", "Nextcloud Enterprise for Outlook - Settings"); } }
        internal static string TabGeneral { get { return Get("TabGeneral", "General"); } }
        internal static string TabIfb { get { return Get("TabIfb", "IFB"); } }
        internal static string TabAdvanced { get { return Get("TabAdvanced", "Advanced"); } }
        internal static string TabInvitations { get { return Get("TabInvitations", "Invitations"); } }
        internal static string TabDebug { get { return Get("TabDebug", "Debug"); } }
        internal static string TabAbout { get { return Get("TabAbout", "About"); } }
        internal static string TabFileLink { get { return Get("TabFileLink", "Filelink"); } }
        internal static string LabelServerUrl { get { return Get("LabelServerUrl", "Server URL"); } }
        internal static string LabelUsername { get { return Get("LabelUsername", "Username"); } }
        internal static string LabelAppPassword { get { return Get("LabelAppPassword", "App password"); } }
        internal static string GroupAuthentication { get { return Get("GroupAuthentication", "Authentication"); } }
        internal static string RadioManual { get { return Get("RadioManual", "Manual (enter username and app password)"); } }
        internal static string RadioLoginFlow { get { return Get("RadioLoginFlow", "Login with Nextcloud (fetch app password automatically)"); } }
        internal static string ButtonLoginFlow { get { return Get("ButtonLoginFlow", "Login with Nextcloud..."); } }
        internal static string ButtonTestConnection { get { return Get("ButtonTestConnection", "Test connection"); } }
        internal static string GroupMuzzle { get { return Get("GroupMuzzle", "Outlook muzzle"); } }
        internal static string CheckMuzzle { get { return Get("CheckMuzzle", "Suppress Outlook invitations (Nextcloud sends them)"); } }
        internal static string LabelMuzzleHint { get { return Get("LabelMuzzleHint", "Note: Responses to invitations and regular e-mails are still sent."); } }
        internal static string LabelMuzzleAccounts { get { return Get("LabelMuzzleAccounts", "Accounts"); } }
        internal static string LabelMuzzleAccountsHint { get { return Get("LabelMuzzleAccountsHint", "Select the accounts whose Outlook invitations should be suppressed."); } }
        internal static string LabelMuzzleNoAccounts { get { return Get("LabelMuzzleNoAccounts", "No Outlook mail accounts were detected."); } }
        internal static string LabelMuzzleUnknownAccount { get { return Get("LabelMuzzleUnknownAccount", "(Unknown account)"); } }
        internal static string CheckIfbEnabled { get { return Get("CheckIfbEnabled", "Enable IFB endpoint"); } }
        internal static string LabelIfbDays { get { return Get("LabelIfbDays", "Range (days):"); } }
        internal static string LabelIfbCacheHours { get { return Get("LabelIfbCacheHours", "Address book cache (hours):"); } }
        internal static string ButtonSave { get { return Get("ButtonSave", "Save"); } }
        internal static string ButtonCancel { get { return Get("ButtonCancel", "Cancel"); } }
        internal static string DebugCheckbox { get { return Get("DebugCheckbox", "Write debug log file"); } }
        internal static string DebugPathPrefix { get { return Get("DebugPathPrefix", "Location: "); } }
        internal static string DebugOpenLog { get { return Get("DebugOpenLog", "Open log file"); } }
        internal static string DebugLogMissingMessage { get { return Get("DebugLogMissingMessage", "No log file has been created yet."); } }
        internal static string DebugLogOpenErrorMessage { get { return Get("DebugLogOpenErrorMessage", "Could not open log file."); } }
        internal static string ButtonNext { get { return Get("ButtonNext", "Next"); } }
        internal static string ButtonBack { get { return Get("ButtonBack", "Back"); } }
        internal static string TalkFormTitle { get { return Get("TalkFormTitle", "Create Talk link"); } }
        internal static string TalkTitleLabel { get { return Get("TalkTitleLabel", "Title"); } }
        internal static string TalkPasswordLabel { get { return Get("TalkPasswordLabel", "Password (optional)"); } }
        internal static string TalkLobbyCheck { get { return Get("TalkLobbyCheck", "Keep lobby active until start time"); } }
        internal static string TalkSearchCheck { get { return Get("TalkSearchCheck", "Show in Nextcloud search"); } }
        internal static string TalkRoomGroup { get { return Get("TalkRoomGroup", "Room type"); } }
        internal static string TalkEventRadio { get { return Get("TalkEventRadio", "Event conversation (bind to appointment)"); } }
        internal static string TalkStandardRadio { get { return Get("TalkStandardRadio", "Standard room (standalone)"); } }
        internal static string TalkVersionUnknown { get { return Get("TalkVersionUnknown", "unknown"); } }
        internal static string TalkEventHint { get { return Get("TalkEventHint", "Event conversations require Nextcloud 31 or newer.\nDetected version: {0}"); } }
        internal static string DialogOk { get { return Get("DialogOk", "OK"); } }
        internal static string DialogCancel { get { return Get("DialogCancel", "Cancel"); } }
        internal static string TalkPasswordTooShort { get { return Get("TalkPasswordTooShort", "The password must be at least 5 characters long."); } }
        internal static string TalkDefaultTitle { get { return Get("TalkDefaultTitle", "Meeting"); } }
        internal static string StatusLoginFlowStarting { get { return Get("StatusLoginFlowStarting", "Starting login flow..."); } }
        internal static string StatusLoginFlowBrowser { get { return Get("StatusLoginFlowBrowser", "Browser opened. Please confirm the login in Nextcloud..."); } }
        internal static string StatusLoginFlowSuccess { get { return Get("StatusLoginFlowSuccess", "Login successful. App password copied."); } }
        internal static string StatusLoginFlowFailure { get { return Get("StatusLoginFlowFailure", "Login failed: {0}"); } }
        internal static string StatusMissingFields { get { return Get("StatusMissingFields", "Please provide server URL, username, and app password."); } }
        internal static string StatusServerUrlRequired { get { return Get("StatusServerUrlRequired", "Please specify the server URL first."); } }
        internal static string StatusInvalidServerUrl { get { return Get("StatusInvalidServerUrl", "Please enter a valid server URL."); } }
        internal static string StatusTestRunning { get { return Get("StatusTestRunning", "Testing connection..."); } }
        internal static string StatusTestSuccessFormat { get { return Get("StatusTestSuccessFormat", "Connection successful{0}"); } }
        internal static string StatusTestSuccessVersionFormat { get { return Get("StatusTestSuccessVersionFormat", "Nextcloud {0}"); } }
        internal static string StatusTestFailureUnknown { get { return Get("StatusTestFailureUnknown", "Unknown error"); } }
        internal static string StatusTestFailure { get { return Get("StatusTestFailure", "Connection failed: {0}"); } }
        internal static string DialogTitle { get { return Get("DialogTitle", "Nextcloud Enterprise for Outlook"); } }
        internal static string ErrorMissingCredentials { get { return Get("ErrorMissingCredentials", "Please configure server URL, username, and app password in the settings first."); } }
        internal static string ErrorNoAppointment { get { return Get("ErrorNoAppointment", "Could not determine the current appointment item."); } }
        internal static string ConfirmReplaceRoom { get { return Get("ConfirmReplaceRoom", "A Talk room already exists for this appointment. Replace it?"); } }
        internal static string ErrorCreateRoom { get { return Get("ErrorCreateRoom", "Talk room could not be created: {0}"); } }
        internal static string ErrorCreateRoomUnexpected { get { return Get("ErrorCreateRoomUnexpected", "Unexpected error while creating the Talk room: {0}"); } }
        internal static string InfoRoomCreated { get { return Get("InfoRoomCreated", "The Talk room \"{0}\" has been created."); } }
        internal static string PromptOpenSettings { get { return Get("PromptOpenSettings", "{0}\n\nOpen settings now?"); } }
        internal static string ErrorServerUnavailable { get { return Get("ErrorServerUnavailable", "The Nextcloud server is currently unreachable. Please check your Internet connection."); } }
        internal static string ErrorAuthenticationRejected { get { return Get("ErrorAuthenticationRejected", "Credentials were not accepted: {0}"); } }
        internal static string ErrorConnectionFailed { get { return Get("ErrorConnectionFailed", "Connection to the Nextcloud server failed: {0}"); } }
        internal static string ErrorUnknownAuthentication { get { return Get("ErrorUnknownAuthentication", "Unknown error during authentication: {0}"); } }
        internal static string WarningIfbStartFailed { get { return Get("WarningIfbStartFailed", "IFB could not be started: {0}"); } }
        internal static string WarningRoomDeleteFailed { get { return Get("WarningRoomDeleteFailed", "Talk room could not be deleted: {0}"); } }
        internal static string WarningLobbyUpdateFailed { get { return Get("WarningLobbyUpdateFailed", "Lobby time could not be updated: {0}"); } }
        internal static string WarningDescriptionUpdateFailed { get { return Get("WarningDescriptionUpdateFailed", "Room description could not be updated: {0}"); } }
        internal static string RibbonAppointmentGroupLabel { get { return Get("RibbonAppointmentGroupLabel", "Nextcloud Enterprise"); } }
        internal static string RibbonTalkButtonLabel { get { return Get("RibbonTalkButtonLabel", "Insert Talk link"); } }
        internal static string RibbonTalkButtonScreenTip { get { return Get("RibbonTalkButtonScreenTip", "Create a Nextcloud Talk conversation"); } }
        internal static string RibbonTalkButtonSuperTip { get { return Get("RibbonTalkButtonSuperTip", "Creates a Nextcloud Talk conversation for this appointment."); } }
        internal static string RibbonExplorerTabLabel { get { return Get("RibbonExplorerTabLabel", "Nextcloud Enterprise"); } }
        internal static string RibbonExplorerGroupLabel { get { return Get("RibbonExplorerGroupLabel", "Configuration"); } }
        internal static string RibbonSettingsButtonLabel { get { return Get("RibbonSettingsButtonLabel", "Settings"); } }
        internal static string RibbonSettingsScreenTip { get { return Get("RibbonSettingsScreenTip", "Nextcloud Enterprise settings"); } }
        internal static string RibbonSettingsSuperTip { get { return Get("RibbonSettingsSuperTip", "Manage server and credentials."); } }
        internal static string RibbonMailGroupLabel { get { return Get("RibbonMailGroupLabel", "Nextcloud Enterprise"); } }
        internal static string RibbonFileLinkButtonLabel { get { return Get("RibbonFileLinkButtonLabel", "Insert Nextcloud share"); } }
        internal static string RibbonFileLinkButtonScreenTip { get { return Get("RibbonFileLinkButtonScreenTip", "Insert Nextcloud share link"); } }
        internal static string RibbonFileLinkButtonSuperTip { get { return Get("RibbonFileLinkButtonSuperTip", "Creates a Nextcloud share link for selected files and inserts it into the message."); } }
        internal static string ErrorNoMailItem { get { return Get("ErrorNoMailItem", "No active mail item found."); } }
        internal static string AboutVersionFormat { get { return Get("AboutVersionFormat", "Version: {0}"); } }
        internal static string AboutLicenseLabel { get { return Get("AboutLicenseLabel", "License:"); } }
        internal static string AboutLicenseLink { get { return Get("AboutLicenseLink", "AGPL v3"); } }
        internal static string LicenseFileMissingMessage { get { return Get("LicenseFileMissingMessage", "The license file could not be found:\r\n{0}"); } }
        internal static string LicenseFileOpenErrorMessage { get { return Get("LicenseFileOpenErrorMessage", "The license file could not be opened: {0}"); } }
        internal static string AboutCopyright { get { return Get("AboutCopyright", "Copyright 2025 Bastian Kleinschmidt"); } }
        internal static string FileLinkBaseLabel { get { return Get("FileLinkBaseLabel", "Base directory"); } }
        internal static string FileLinkBaseHint { get { return Get("FileLinkBaseHint", "Files will be stored beneath this directory (e.g. \"90 Freigaben - extern\")."); } }
        internal static string FileLinkDownloadLabel { get { return Get("FileLinkDownloadLabel", "Download link"); } }
        internal static string FileLinkPasswordLabel { get { return Get("FileLinkPasswordLabel", "Password"); } }
        internal static string FileLinkExpireLabel { get { return Get("FileLinkExpireLabel", "Expiration date"); } }
        internal static string FileLinkPermissionsLabel { get { return Get("FileLinkPermissionsLabel", "Your permissions"); } }
        internal static string FileLinkHtmlIntro { get { return Get("FileLinkHtmlIntro", "I would like to share files securely and protect your privacy. Click the link below to download your files."); } }
        internal static string FileLinkHtmlFooter { get { return Get("FileLinkHtmlFooter", "{0} is a solution for secure email and file exchange."); } }
        internal static string FileLinkNoFilesConfirm { get { return Get("FileLinkNoFilesConfirm", "No files or folders were added to this share.\r\nRecipients can only upload their own files.\r\nDo you still want to continue?"); } }
    }
}

