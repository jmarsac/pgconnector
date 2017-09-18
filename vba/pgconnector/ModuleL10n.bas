Attribute VB_Name = "ModuleL10n"
'/*MAN______________________________________________________________________
'
'
' FICHIER: $RCSfile$ $Revision$
'
'  $Author$  $Date$
'
' DESCRIPTION
' xxxxxxxxxxxxx
'
'
' FONCTIONS
' VOIR EGALEMENT
'
' Copyright   Azimut    Tous droits réservés
' Azimut S.A.R.L 3 rue du travail 67000 Strasbourg
' Tél 03 88 08 50 67
' SIREN 390 458 164   http://www.azimut.fr          grc@azimut.fr
'
'___________________________________________________________________ENDMAN*/
Option Explicit

Public Const TXT_IMPORTEE_DANS_LE_MODELE_ACTIF = "1"
Public Const TXT_CLAUSE_WHERE = "2"
Public Const TXT_ACTION = "3"
Public Const TXT_AUCUNE_TABLE_N_EST_DISPONIBLE_OU_BIEN_VOUS_N_AVEZ_AUCUN_DROIT_DE_CONSULTATION_SUR_LES_TABLES_EXISTANTES = "4"
Public Const TXT_ATTACHER_COUCHE_POSTGIS_EN_REFERENCE = "5"
Public Const TXT_AUCUNE = "6"
Public Const TXT_AUCUNE_CONNEXION_DEFINIE_VEUILLEZ_EN_CREER_UNE = "7"
Public Const TXT_B_D = "8"
Public Const TXT_CELLULE = "9"
Public Const TXT_CHOIX_COUCHE_POSTGIS = "10"
Public Const TXT_CONNECTE_A_LA_BD = "11"
Public Const TXT_CONNECTER = "12"
Public Const TXT_CONNEXION = "13"
Public Const TXT_CONNEXION_A_LA_BD_REUSSIE = "14"
Public Const TXT_CONNEXION_INVALIDE = "15"
Public Const TXT_CONTACTEZ_L_ADMINISTRATEUR = "16"
Public Const TXT_COPIEZ_LE_DANS_USTN_SITE_SEED_OU_DEFINISSEZ_CORRECTEMENT_LA_VARIABLE_PGC_SEED_FULLNAME = "17"
Public Const TXT_COUCHE = "18"
Public Const TXT_COUCHE_SEULEMENT = "19"
Public Const TXT_COUCHES_POSTGIS = "20"
Public Const TXT_COULEUR = "21"
Public Const TXT_DANS_LE_MODELE_ACTIF = "22"
Public Const TXT_DECONNECTE_DE_LA_BD = "23"
Public Const TXT_DETACHER_LES_COUCHES_POSTGIS = "24"
Public Const TXT_ECHEC_CONNEXION_CONTACTEZ_L_ADMINISTRATEUR = "25"
Public Const TXT_ECHEC_DE_L_IMPORT_DE_LA_COUCHE_POSTGIS = "26"
Public Const TXT_ECHEC_DE_LA_CONNEXION = "27"
Public Const TXT_ECHEC_DE_LA_CONNEXION_A_LA_BD = "28"
Public Const TXT_ELEMENT = "29"
Public Const TXT_ENREGISTRER = "30"
Public Const TXT_ERREUR = "31"
Public Const TXT_EST_INTROUVABLE = "32"
Public Const TXT_EST_VALIDE = "33"
Public Const TXT_EXPORT = "34"
Public Const TXT_EXPORTER = "35"
Public Const TXT_FICHIER_ENTIER = "36"
Public Const TXT_FICHIER_PROTOTYPE_INTROUVABLE = "37"
Public Const TXT_FILTRE = "38"
Public Const TXT_IDENTIFICATION = "39"
Public Const TXT_IMPORTER_COUCHE_POSTGIS = "40"
Public Const TXT_IMPOSSIBLE_DE_SE_CONNECTER_A_LA_BASE_DE_DONNEES_CONTACTEZ_L_ADMINISTRATEUR_POUR_RESOUDRE_LE_PROBLEME = "41"
Public Const TXT_INDIQUER_UN_NOM_POUR_LA_CONNEXION = "42"
Public Const TXT_LE_FICHIER = "43"
Public Const TXT_LE_MODELE_EST_EN_LECTURE_SEULE = "44"
Public Const TXT_MOT_DE_PASSE = "45"
Public Const TXT_NE_PEUT_ETRE_OUVERT = "46"
Public Const TXT_NE_PEUT_ETRE_REINITIALISE = "47"
Public Const TXT_N_EST_PLUS_VALIDE = "48"
Public Const TXT_NIVEAU = "49"
Public Const TXT_NOM = "50"
Public Const TXT_OK = "51"
Public Const TXT_OUVERTURE_DU_FICHIER_IMPOSSIBLE = "52"
Public Const TXT_PORT = "53"
Public Const TXT_POUR_UTILISER_CETTE_COMMANDE_VOUS_DEVEZ_ETRE_CONNECTE_A_LA_BD = "54"
Public Const TXT_QUITTER = "55"
Public Const TXT_REINITIALISATION_DU_FICHIER_IMPOSSIBLE = "56"
Public Const TXT_REINITIALISER = "57"
Public Const TXT_SCHEMA = "58"
Public Const TXT_SERVEUR = "59"
Public Const TXT_SYNCHRONISATION_BD_SPATIALE_ACTIVEE = "60"
Public Const TXT_SYNCHRONISATION_BD_SPATIALE_DESACTIVEE = "61"
Public Const TXT_TABLE = "62"
Public Const TXT_TABLES = "63"
Public Const TXT_TESTER_LA_CONNEXION = "64"
Public Const TXT_TOUS = "65"
Public Const TXT_TOUTES_LES_COUCHES_POSTGIS_DISPARAITRONT = "66"
Public Const TXT_UTILISATEUR = "67"
Public Const TXT_UTILISER_CLOTURE = "68"
Public Const TXT_VOULEZ_VOUS_DETACHER_LE_FICHIER = "69"
Public Const TXT_VUES = "70"
Public Const TXT_LABELS = "71"
Public Const TXT_MODE_SSL = "72"
Public Const TXT_COUCHE_POSTGIS_INTROUVABLE = "73"
Public Const TXT_DANS_LA_TABLE_DESCRIPTION = "74"
Public Const TXT_ERREUR_CONNEXION_OBJET_A_LA_BD = "75"
Public Const TXT_SUPPRIMER_CONNEXION = "76"

Private l10nColl As Collection
Private duColl As Collection

Public Sub l10nInit()
    
    gPgcUserLang = PgcUserLangFrench
    If Application.ActiveWorkspace.IsConfigurationVariableDefined("PGC_USER_LANG") Then
        gPgcUserLang = Application.ActiveWorkspace.ConfigurationVariableValue("PGC_USER_LANG")
        If gPgcUserLang <> PgcUserLangFrench And gPgcUserLang <> PgcUserLangDutch Then
            Application.MessageCenter.AddMessage "Langue inconnue", gPgcUserLang + " est inconnu. Le français sera utilisé.", msdMessageCenterPriorityWarning, True
            gPgcUserLang = PgcUserLangFrench
        End If
    End If
    l10nLoad (gPgcUserLang)

End Sub
Public Sub l10nLoad(langId As String)

Set l10nColl = Nothing
Set l10nColl = New Collection

If langId = PgcUserLangDutch Then
    l10nColl.Add "ingevoerd in het actieve model", TXT_IMPORTEE_DANS_LE_MODELE_ACTIF
    l10nColl.Add " (WHERE clausule)", TXT_CLAUSE_WHERE
    l10nColl.Add "aktie", TXT_ACTION
    l10nColl.Add "Er is geen tabel beschikbaar of u heeft geen rechten op de bestaande tabellen", TXT_AUCUNE_TABLE_N_EST_DISPONIBLE_OU_BIEN_VOUS_N_AVEZ_AUCUN_DROIT_DE_CONSULTATION_SUR_LES_TABLES_EXISTANTES
    l10nColl.Add "Koppelen van een postgislaag als referentie", TXT_ATTACHER_COUCHE_POSTGIS_EN_REFERENCE
    l10nColl.Add "Geen enkele", TXT_AUCUNE
    l10nColl.Add "Er is geen verbinding gedefinieerd, gelieve een verbinding te maken", TXT_AUCUNE_CONNEXION_DEFINIE_VEUILLEZ_EN_CREER_UNE
    l10nColl.Add "DB", TXT_B_D
    l10nColl.Add "Cel", TXT_CELLULE
    l10nColl.Add "Keuze Postgis laag", TXT_CHOIX_COUCHE_POSTGIS
    l10nColl.Add "Gekoppeld met de database", TXT_CONNECTE_A_LA_BD
    l10nColl.Add "Verbinden", TXT_CONNECTER
    l10nColl.Add "Verbinding", TXT_CONNEXION
    l10nColl.Add "Verbinding met de database geslaagd", TXT_CONNEXION_A_LA_BD_REUSSIE
    l10nColl.Add "Ongeldige verbinding", TXT_CONNEXION_INVALIDE
    l10nColl.Add "Contacteer de beheerder", TXT_CONTACTEZ_L_ADMINISTRATEUR
    l10nColl.Add "Kopiëer  in de $(_ustn_site)seed of definieer de variable   PGC_SEED_FULLNAME", TXT_COPIEZ_LE_DANS_USTN_SITE_SEED_OU_DEFINISSEZ_CORRECTEMENT_LA_VARIABLE_PGC_SEED_FULLNAME
    l10nColl.Add "Laag", TXT_COUCHE
    l10nColl.Add "Enkel de laag", TXT_COUCHE_SEULEMENT
    l10nColl.Add "Postgis lagen", TXT_COUCHES_POSTGIS
    l10nColl.Add "kleur", TXT_COULEUR
    l10nColl.Add "in het aktieve model", TXT_DANS_LE_MODELE_ACTIF
    l10nColl.Add "Loskoppelen van de database", TXT_DECONNECTE_DE_LA_BD
    l10nColl.Add "Loskoppelen van de postgis lagen", TXT_DETACHER_LES_COUCHES_POSTGIS
    l10nColl.Add "Verbindingsfout, contacteer de beheerder", TXT_ECHEC_CONNEXION_CONTACTEZ_L_ADMINISTRATEUR
    l10nColl.Add "Fout bij importeren van de postgis laag", TXT_ECHEC_DE_L_IMPORT_DE_LA_COUCHE_POSTGIS
    l10nColl.Add "Verbindingsfout ", TXT_ECHEC_DE_LA_CONNEXION
    l10nColl.Add "Verbindingsfout bij de database", TXT_ECHEC_DE_LA_CONNEXION_A_LA_BD
    l10nColl.Add "element", TXT_ELEMENT
    l10nColl.Add "Bewaren", TXT_ENREGISTRER
    l10nColl.Add "Fout bij importeren van de postgis laag", TXT_ERREUR
    l10nColl.Add "onvindbaar", TXT_EST_INTROUVABLE
    l10nColl.Add "Geldig", TXT_EST_VALIDE
    l10nColl.Add "Export", TXT_EXPORT
    l10nColl.Add "Exporteren", TXT_EXPORTER
    l10nColl.Add "Volledige bestand", TXT_FICHIER_ENTIER
    l10nColl.Add "Seed file onvindbaar", TXT_FICHIER_PROTOTYPE_INTROUVABLE
    l10nColl.Add "filter", TXT_FILTRE
    l10nColl.Add "identificatie", TXT_IDENTIFICATION
    l10nColl.Add "Importeren van een Postgis laag", TXT_IMPORTER_COUCHE_POSTGIS
    l10nColl.Add "Onmogelijk te verbinden met de database, contacteer de beheerder om dit probleem op te lossen", TXT_IMPOSSIBLE_DE_SE_CONNECTER_A_LA_BASE_DE_DONNEES_CONTACTEZ_L_ADMINISTRATEUR_POUR_RESOUDRE_LE_PROBLEME
    l10nColl.Add "Geef de naam van de connectie", TXT_INDIQUER_UN_NOM_POUR_LA_CONNEXION
    l10nColl.Add "het bestand", TXT_LE_FICHIER
    l10nColl.Add "Alleen lezen model", TXT_LE_MODELE_EST_EN_LECTURE_SEULE
    l10nColl.Add "Paswoord", TXT_MOT_DE_PASSE
    l10nColl.Add "Kan niet openen", TXT_NE_PEUT_ETRE_OUVERT
    l10nColl.Add "Kan niet initialiseren", TXT_NE_PEUT_ETRE_REINITIALISE
    l10nColl.Add "Niet meer geldig", TXT_N_EST_PLUS_VALIDE
    l10nColl.Add "Laag", TXT_NIVEAU
    l10nColl.Add "Naam", TXT_NOM
    l10nColl.Add "OK", TXT_OK
    l10nColl.Add "Kan het bestand niet openen", TXT_OUVERTURE_DU_FICHIER_IMPOSSIBLE
    l10nColl.Add "Poort", TXT_PORT
    l10nColl.Add "Dit commando vereist een connectie met de database", TXT_POUR_UTILISER_CETTE_COMMANDE_VOUS_DEVEZ_ETRE_CONNECTE_A_LA_BD
    l10nColl.Add "Afsluiten", TXT_QUITTER
    l10nColl.Add "Herinitialisatie van het bestand niet mogelijk", TXT_REINITIALISATION_DU_FICHIER_IMPOSSIBLE
    l10nColl.Add "Herinitialiseren", TXT_REINITIALISER
    l10nColl.Add "Schema", TXT_SCHEMA
    l10nColl.Add "Server", TXT_SERVEUR
    l10nColl.Add "Synchronisatie met de Spatial database is aktief", TXT_SYNCHRONISATION_BD_SPATIALE_ACTIVEE
    l10nColl.Add "Synchronisatie met de Spatial database is niet aktief", TXT_SYNCHRONISATION_BD_SPATIALE_DESACTIVEE
    l10nColl.Add "Tabel", TXT_TABLE
    l10nColl.Add "Tabellen", TXT_TABLES
    l10nColl.Add "Testen van de connectie", TXT_TESTER_LA_CONNEXION
    l10nColl.Add "Alle", TXT_TOUS
    l10nColl.Add "Alle Postgis lagen zullen verdwijnen", TXT_TOUTES_LES_COUCHES_POSTGIS_DISPARAITRONT
    l10nColl.Add "Gebruiker", TXT_UTILISATEUR
    l10nColl.Add "Gebruik het Fence commando", TXT_UTILISER_CLOTURE
    l10nColl.Add "Wil je het bestand loskoppelen", TXT_VOULEZ_VOUS_DETACHER_LE_FICHIER
    l10nColl.Add "Vensters", TXT_VUES
    l10nColl.Add "Labels", TXT_LABELS
    l10nColl.Add "SSL mode", TXT_MODE_SSL
    l10nColl.Add "Postgis laag onvindbaar", TXT_COUCHE_POSTGIS_INTROUVABLE
    l10nColl.Add " in 'description' tabel", TXT_DANS_LA_TABLE_DESCRIPTION
    l10nColl.Add "Verbindingsfout ", TXT_ERREUR_CONNEXION_OBJET_A_LA_BD
    l10nColl.Add "Verwijderen verbinding", TXT_SUPPRIMER_CONNEXION
    
Else
    l10nColl.Add " importée dans le modèle actif", TXT_IMPORTEE_DANS_LE_MODELE_ACTIF
    l10nColl.Add "(clause WHERE)", TXT_CLAUSE_WHERE
    l10nColl.Add "action", TXT_ACTION
    l10nColl.Add "Aucune table n'est disponible ou bien vous n'avez aucun droit de consultation sur les tables existantes", TXT_AUCUNE_TABLE_N_EST_DISPONIBLE_OU_BIEN_VOUS_N_AVEZ_AUCUN_DROIT_DE_CONSULTATION_SUR_LES_TABLES_EXISTANTES
    l10nColl.Add "Attacher couche PostGIS en référence", TXT_ATTACHER_COUCHE_POSTGIS_EN_REFERENCE
    l10nColl.Add "Aucune", TXT_AUCUNE
    l10nColl.Add "Aucune connexion définie. Veuillez en créer une.", TXT_AUCUNE_CONNEXION_DEFINIE_VEUILLEZ_EN_CREER_UNE
    l10nColl.Add "B.D.", TXT_B_D
    l10nColl.Add "Cellule", TXT_CELLULE
    l10nColl.Add "Choix couche PostGIS", TXT_CHOIX_COUCHE_POSTGIS
    l10nColl.Add "Connecté à la BD", TXT_CONNECTE_A_LA_BD
    l10nColl.Add "Connecter", TXT_CONNECTER
    l10nColl.Add "Connexion", TXT_CONNEXION
    l10nColl.Add "Connexion à la BD réussie", TXT_CONNEXION_A_LA_BD_REUSSIE
    l10nColl.Add "Connexion invalide", TXT_CONNEXION_INVALIDE
    l10nColl.Add "Contactez l'administrateur.", TXT_CONTACTEZ_L_ADMINISTRATEUR
    l10nColl.Add "Copiez le dans $(_ustn_site)seed ou définissez correctement la variable PGC_SEED_FULLNAME", TXT_COPIEZ_LE_DANS_USTN_SITE_SEED_OU_DEFINISSEZ_CORRECTEMENT_LA_VARIABLE_PGC_SEED_FULLNAME
    l10nColl.Add "Couche", TXT_COUCHE
    l10nColl.Add "Couche seulement", TXT_COUCHE_SEULEMENT
    l10nColl.Add "Couches PostGIS", TXT_COUCHES_POSTGIS
    l10nColl.Add "Couleur", TXT_COULEUR
    l10nColl.Add " dans le modèl actif", TXT_DANS_LE_MODELE_ACTIF
    l10nColl.Add "Déconnecté de la BD", TXT_DECONNECTE_DE_LA_BD
    l10nColl.Add "Détacher les couches PotGIS", TXT_DETACHER_LES_COUCHES_POSTGIS
    l10nColl.Add "Echec connexion. Contactez l'administrateur", TXT_ECHEC_CONNEXION_CONTACTEZ_L_ADMINISTRATEUR
    l10nColl.Add "Echec de l'import de la couche PostGIS", TXT_ECHEC_DE_L_IMPORT_DE_LA_COUCHE_POSTGIS
    l10nColl.Add "Echec de la connexion", TXT_ECHEC_DE_LA_CONNEXION
    l10nColl.Add "Echec de la connexion à la BD", TXT_ECHEC_DE_LA_CONNEXION_A_LA_BD
    l10nColl.Add "Elément", TXT_ELEMENT
    l10nColl.Add "Enregistrer", TXT_ENREGISTRER
    l10nColl.Add "Erreur", TXT_ERREUR
    l10nColl.Add "est introuvable", TXT_EST_INTROUVABLE
    l10nColl.Add "est valide", TXT_EST_VALIDE
    l10nColl.Add "Export", TXT_EXPORT
    l10nColl.Add "Exporter", TXT_EXPORTER
    l10nColl.Add "Fichier entier", TXT_FICHIER_ENTIER
    l10nColl.Add "Fichier prototype introuvable", TXT_FICHIER_PROTOTYPE_INTROUVABLE
    l10nColl.Add "Filtre", TXT_FILTRE
    l10nColl.Add "Identification", TXT_IDENTIFICATION
    l10nColl.Add "Importer couche PostGIS", TXT_IMPORTER_COUCHE_POSTGIS
    l10nColl.Add "Impossible de se connecter à la base de données. Contactez l'administrateur pour résoudre le problème. ", TXT_IMPOSSIBLE_DE_SE_CONNECTER_A_LA_BASE_DE_DONNEES_CONTACTEZ_L_ADMINISTRATEUR_POUR_RESOUDRE_LE_PROBLEME
    l10nColl.Add "Indiquer un nom pour la connexion", TXT_INDIQUER_UN_NOM_POUR_LA_CONNEXION
    l10nColl.Add "Le fichier", TXT_LE_FICHIER
    l10nColl.Add "Le modèle est en lecture seule", TXT_LE_MODELE_EST_EN_LECTURE_SEULE
    l10nColl.Add "Mot de passe", TXT_MOT_DE_PASSE
    l10nColl.Add "ne peut être ouvert", TXT_NE_PEUT_ETRE_OUVERT
    l10nColl.Add "ne peut être réinitialisé", TXT_NE_PEUT_ETRE_REINITIALISE
    l10nColl.Add "n'est plus valide", TXT_N_EST_PLUS_VALIDE
    l10nColl.Add "Niveau", TXT_NIVEAU
    l10nColl.Add "Nom", TXT_NOM
    l10nColl.Add "OK", TXT_OK
    l10nColl.Add "Ouverture du fichier impossible", TXT_OUVERTURE_DU_FICHIER_IMPOSSIBLE
    l10nColl.Add "Port", TXT_PORT
    l10nColl.Add "Pour utiliser cette commande, vous devez être connecté à la BD.", TXT_POUR_UTILISER_CETTE_COMMANDE_VOUS_DEVEZ_ETRE_CONNECTE_A_LA_BD
    l10nColl.Add "Quitter", TXT_QUITTER
    l10nColl.Add "Réinitialisation du fichier impossible", TXT_REINITIALISATION_DU_FICHIER_IMPOSSIBLE
    l10nColl.Add "Réinitialiser", TXT_REINITIALISER
    l10nColl.Add "Schéma", TXT_SCHEMA
    l10nColl.Add "Serveur", TXT_SERVEUR
    l10nColl.Add "Synchronisation BD spatiale activée", TXT_SYNCHRONISATION_BD_SPATIALE_ACTIVEE
    l10nColl.Add "Synchronisation BD spatiale desactivée", TXT_SYNCHRONISATION_BD_SPATIALE_DESACTIVEE
    l10nColl.Add "Table", TXT_TABLE
    l10nColl.Add "Tables", TXT_TABLES
    l10nColl.Add "Tester la connexion", TXT_TESTER_LA_CONNEXION
    l10nColl.Add "Tous", TXT_TOUS
    l10nColl.Add "Toutes les couches PostGIS disparaitront", TXT_TOUTES_LES_COUCHES_POSTGIS_DISPARAITRONT
    l10nColl.Add "Utilisateur", TXT_UTILISATEUR
    l10nColl.Add "Utiliser clôture", TXT_UTILISER_CLOTURE
    l10nColl.Add "Voulez-vous détacher le fichier", TXT_VOULEZ_VOUS_DETACHER_LE_FICHIER
    l10nColl.Add "Vues", TXT_VUES
    l10nColl.Add "Etiquettes", TXT_LABELS
    l10nColl.Add "Mode SSL", TXT_MODE_SSL
    l10nColl.Add "Couche PostGIS introuvable", TXT_COUCHE_POSTGIS_INTROUVABLE
    l10nColl.Add " dans la table 'description'", TXT_DANS_LA_TABLE_DESCRIPTION
    l10nColl.Add "Erreur de connexion de l'objet à la B.D.", TXT_ERREUR_CONNEXION_OBJET_A_LA_BD
    l10nColl.Add "Supprimer connexion", TXT_SUPPRIMER

End If

End Sub
Public Function l10n(txtId As String) As String
    On Error Resume Next
    l10n = l10nColl(txtId)
    On Error GoTo 0
End Function


