Attribute VB_Name = "ModuleGlobal"
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

Public Const PgcVersion = "1.1.2"


'
' Configuration variables :
' PGC_SEED_FULLNAME: nom complet du prototype utilisé ($(_USTN_SITE)seed/seed-postgis.dgn par défaut)
'


Public Enum EnumCheckoutMode
    enuCheckoutModeUnknown = 0
    enuCheckoutModeImport = 1
    enuCheckoutModeAttach = 2
End Enum

Public Enum EnumPostgisDgnResetLevel
    enuPostgisDgnResetNone = 0
    enuPostgisDgnResetLayerOnly = 1
    enuPostgisDgnResetFullFile = 2
End Enum

Public Enum EnumGeomType
    enuGeomTypeUnknown = 0
    enuGeomTypePoint = 1
    enuGeomTypeArea = 2
    enuGeomTypeLine = 3
End Enum

Public Const PostgisDgnRefLogicalName = "_pgis_layer_"
Public Const PgDefaultSrid = 32170

Public Const PgcUserLangFrench = "fr"
Public Const PgcUserLangDutch = "du"

Public gPgcInitDone As Boolean
Public gPgcUserLang As String

Public gSqlWhere As String
Public gSqlWhereFence As String
Public gSqlWhereUser As String
Public gSqlQuery As String
Public gSqlQueryLabels As String
Public gSqlUpsert As String

Public gSchemasArray() As String
Public gSchemaName As String
Public gTableName As String
Public gEntitynum As Integer
Public gCheckoutMode As EnumCheckoutMode
Public gPostgisDgnResetLevel As EnumPostgisDgnResetLevel
Public gLevelName As String
Public gLevel As Level
Public gCellname As String
Public gCellOrientation As Double
Public gTextOrientation As Double
Public gTextSize As Double
Public gTablenames As Collection

Public gPgSrid As Long

Public gPgConnexion As azidblib.aziDbConnexion
Public gPgCnxList As String
Public gPgConnectionName As String
Public gPgHost As String
Public gPgPort As String
Public gPgDbname As String
Public gPgUsername As String
Public gPgPassword As String

Public gUseSharedCell As Boolean

' variable de suspension momentanée de la gestion des modifications d'élément (synchro dao->SIG et topologie)
Public P_fl_no_elem_events_handle As Boolean
Public P_launchElemChangeTrack As ClassLaunchIChangeHandler
Public P_elemChangeTrack As ClassIChangeTrack
' variable de vérification si la fonction de gestion des modifications d'élément s'est bien déroulée
Public P_fl_elem_events_function_succeeded As Boolean
Public P_last_deleted_elem As Element
Public P_last_deleted_elem_id As DLong
Public P_soft_deleted_and_created_elem_id As DLong
Public P_fl_deleted_and_created As Boolean
Public P_last_deleted_and_created_mslink As Long



