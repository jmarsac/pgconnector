Attribute VB_Name = "ModuleGlobal"
'/*MAN______________________________________________________________________
'
'
' FICHIER: $RCSfile: ModuleGlobal.bas,v $ $Revision: 1.8 $
'
'  $Author: jma $  $Date: 2016/03/15 13:38:43 $
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

Public Declare Sub mdlElmdscr_getProperties Lib "stdmdlbltin.dll" (ByRef level As Long, ByRef ggNum As Long, ByRef dgnClass As Long, _
                                   ByRef locked As Long, ByRef newElm As Long, ByRef modified As Long, ByRef viewIndepend As Long, ByRef solidHole As Long, ByVal pElementDescr As Long)

Public Const csteCouleurFibreOptique = 66
Public Const csteCouleurEclairage = 58
Public Const csteCouleurElectrique = 35
Public Const csteCouleurCyclo = 58

Public Const csteA4PORTRAIT = "A4 Portrait"
Public Const csteA3PORTRAIT = "A3 Portrait"

Public Const csteA4LigneX1 = 0.15464
Public Const csteA4LigneX2 = 0.16518
Public Const csteA4LigneX3 = 0.16787

Public Const csteA4LigneY1 = 0.28713
Public Const csteA4LigneY2 = 0.28447
Public Const csteA4LigneY3 = 0.28185
Public Const csteA4LigneY4 = 0.28453

Public Const csteA3LigneX1 = 0.24816
Public Const csteA3LigneX2 = 0.25525
Public Const csteA3LigneX3 = 0.25794

Public Const csteA3LigneY1 = 0.41013
Public Const csteA3LigneY2 = 0.40747
Public Const csteA3LigneY3 = 0.40485
Public Const csteA3LigneY4 = 0.40223

Public Const csteTextelegendeHauteur = 0.002
Public Const csteTextelegendeLargeur = 0.002

Public Const csteTOPOGRAPHIQUE = "TOPOGRAPHIQUE"
Public Const csteORTHOPHOTOGRAPHIE = "ORTHOPHOTOGRAPHIE"
Public Const csteCADASTRE = "PLAN CADASTRAL"

Public Const cstePrecisionOrtho = "Précision: 1 pixel = 10 cm"
Public Const cstePrecision200 = "Echelle de précision 1/200"
Public Const cstePrecision500 = "Echelle de précision 1/500"
Public Const cstePrecisionCadastre = "Précision cadastrale"

Public Const csteTypeDemandeAtu = "ATU"
Public Const csteTypeDemandeDict = "DICT"
Public Const csteTypeDemandeDt = "DT"
Public Const csteTypeDemandeDtDict = "DT/DICT"

Public gCleCommune As String ' trigramme commune "all", "arb",.. pour PMA, code INSEE (381, 125 pour les autres)

Public Enum EnumTypeReseau
    enuTypeReseauInconnu = 0
    enuTypeReseauFibreOptique = 1 ' vert
    enuTypeReseauEclairagePublic = 2 ' rouge
    enuTypeReseauElectrique = 3 ' rouge
    enuTypeReseauGaz = 4 ' jaune
    enuTypeReseauHydrocarbure = 5 'jaune
    enuTypeReseauChimique = 6 ' orange
    enuTypeReseauEauPotable = 7 'bleu
    enuTypeReseauAssainissement = 8 ' marron
    enuTypeReseauEauPluviale = 9 ' marron
    enuTypeReseauChauffageClim = 10 ' violet
    enuTypeReseauTelecom = 11 ' vert
    enuTypeReseauFeuSignalisation = 12 ' blanc (noir)
    enuTypeReseauMultiples = 13 ' rose
    enuTypeReseauCyclable = 14
End Enum

Public Const cteClasseReseauInconnue = "inconnue"
Public Const cteClasseReseauA = "A"
Public Const cteClasseReseauB = "B"
Public Const cteClasseReseauC = "C"

Public Type ReseauTypeEtClasse
    typeReseau As EnumTypeReseau
    classeReseau As String
    couleur As Integer
End Type

Public gLevelEmprise As level
Public gLevelPlanche As level
Public gFormatName As String
Public gEchelle As Long
Public gNumeroDemande As String
Public gTypeDemande As String
Public gRootFilename As String
Public gCommune As String
Public gAdresse As String
Public gDossierDemandes As String
Public gDossierReponses As String
Public gIdRepertoireDemandes As Integer
Public gZoomFactor As Double
Public gFillColor As Integer

Public Sub PmaInit()

Dim dummy As String

' dossier des demandes
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_DOSSIER_DEMANDES") Then
    gDossierDemandes = Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_DOSSIER_DEMANDES")
End If

' dossier des réponses
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_DOSSIER_REPONSES") Then
    gDossierReponses = Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_DOSSIER_REPONSES")
End If

Set gLevelEmprise = Nothing

' niveau emprise travaux
dummy = "dict-emprise-travaux"
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_NIVEAU_EMPRISE") Then
    dummy = Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_NIVEAU_EMPRISE")
End If
If dummy <> "" Then Set gLevelEmprise = Application.ActiveDesignFile.Levels.Find(dummy)
If gLevelEmprise Is Nothing Then
    Call Application.MessageCenter.AddMessage("Le niveau de l'emprise des travaux est indéterminé." _
    , "Vérifiez que la variable AZI_DTDICT_NIVEAU_EMPRISE est bien définie et que le niveau " _
    + IIf(dummy <> "", "'" + dummy + "'", "spécifié") _
    + " est bien disponible dans le fichier dessin." _
    , msdMessageCenterPriorityError, True)
    End
End If

Set gLevelPlanche = Nothing

' niveau planches
dummy = "dict-emprise-planche"
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_NIVEAU_PLANCHE") Then
    dummy = Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_NIVEAU_PLANCHE")
End If
If dummy <> "" Then Set gLevelPlanche = Application.ActiveDesignFile.Levels.Find(dummy)
If gLevelPlanche Is Nothing Then
    Call Application.MessageCenter.AddMessage("Le niveau de l'emprise des planches est indéterminé." _
    , "Vérifiez que la variable AZI_DTDICT_NIVEAU_PLANCHE est bien définie et que le niveau " _
    + IIf(dummy <> "", "'" + dummy + "'", "spécifié") _
    + " est bien disponible dans le fichier dessin." _
    , msdMessageCenterPriorityError, True)
    End
End If

' facteur de zoom
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_ZOOM") Then
    gZoomFactor = Val(Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_ZOOM"))
End If
If gZoomFactor <= 0.01 Or gZoomFactor > 1000 Then gZoomFactor = 5

' couleur de remplissage de l'emprise
If Application.ActiveWorkspace.IsConfigurationVariableDefined("AZI_DTDICT_FILL_COLOR") Then
    gFillColor = Val(Application.ActiveWorkspace.ConfigurationVariableValue("AZI_DTDICT_FILL_COLOR"))
End If
If gFillColor <= 0 Or gZoomFactor > 254 Then gFillColor = 4

' format par défaut
gFormatName = csteA4PORTRAIT
' échelle par défaut
gEchelle = 200

End Sub
Public Function dictCouleur(typeReseau As EnumTypeReseau) As Integer

dictCouleur = 255
Select Case typeReseau
    Case enuTypeReseauFibreOptique ' vert
        dictCouleur = RGB(0, 255, 0)
    Case enuTypeReseauTelecom  ' vert
        dictCouleur = RGB(0, 255, 0)
    Case enuTypeReseauEclairagePublic ' rouge
        dictCouleur = RGB(255, 0, 0)
    Case enuTypeReseauElectrique  ' rouge
        dictCouleur = RGB(255, 0, 0)
    Case enuTypeReseauGaz  ' jaune
        dictCouleur = RGB(255, 255, 0)
    Case enuTypeReseauHydrocarbure  'jaune
        dictCouleur = RGB(255, 255, 0)
    Case enuTypeReseauChimique = 6 ' orange
        dictCouleur = RGB(255, 128, 0)
    Case enuTypeReseauEauPotable  'bleu
        dictCouleur = RGB(128, 0, 255)
    Case enuTypeReseauAssainissement ' marron
        dictCouleur = RGB(128, 64, 0)
    Case enuTypeReseauEauPluviale ' marron
        dictCouleur = RGB(128, 64, 0)
    Case enuTypeReseauChauffageClim  ' violet
        dictCouleur = RGB(160, 0, 160)
    Case enuTypeReseauFeuSignalisation  ' blanc (noir)
        dictCouleur = 0
    Case enuTypeReseauMultiples  ' rose
        dictCouleur = RGB(255, 128, 0)
    Case enuTypeReseauCyclable
End Select
dictCouleur = Application.ActiveModelReference.InternalColorFromRGBColor(dictCouleur)

End Function
