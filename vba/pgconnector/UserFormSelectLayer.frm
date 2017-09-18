VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSelectLayer 
   Caption         =   "Choix couche PostGIS 1.1"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8535
   OleObjectBlob   =   "UserFormSelectLayer.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserFormSelectLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*MAN______________________________________________________________________
'
'
' FICHIER: $RCSfile$ $Revision$
'
'         $Date$
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
' SIRET 390 458 164   http://www.azimut.fr          grc@azimut.fr
'
'___________________________________________________________________ENDMAN*/
Option Explicit


Public Sub updateComboBoxTableWithSchema(schemaname As String, Optional last_table As String = "")
Dim oRecordset As New ADODB.Recordset
Dim sql As String, tablename As String
Dim desc As String
Dim use_table As Boolean
Dim i As Integer

Set oRecordset.ActiveConnection = gPgConnexion.CONNEXION
  ' Ouverture du recordset recensant les tables du (des) schéma(s) par défaut
  sql = "SELECT table_name,table_type FROM information_schema.tables " _
        + " WHERE lower(table_schema) IN ('" & schemaname & "') AND table_type = 'BASE TABLE'" _
        + " ORDER BY table_schema,table_name"
  With oRecordset
    .Open sql, , , adLockReadOnly, adCmdText
  End With


   ' On affiche les tables standards et les vues de l'utilisateur gPgUsername
   If oRecordset.State = adStateOpen Then
    Do While Not oRecordset.EOF
        If oRecordset("table_type") = "TABLE" Or oRecordset("table_type") = "BASE TABLE" Or _
           oRecordset("table_type") = "VIEW" Or _
           oRecordset("table_type") = "LINK" Or oRecordset("table_type") = "PASS-THROUGH" Then
           ' On n'affiche que les tables autorisées en consultation
           tablename = oRecordset("table_name")
           
           ' les noms de tables exploitées sont entièrement en majuscules
           If UCase(tablename) = tablename Then
                use_table = True
           End If
           
           If use_table = True And gPgConnexion.TableExists(tablename) Then
                desc = gPgConnexion.SqlQuery(buildSqlSelectLayerDescription(gSchemaName, tablename))
                Me.ComboBoxTable.AddItem buildLayerKey(tablename, desc)
                gTablenames.Add tablename, buildLayerKey(tablename, desc)
           'ElseIf oRecordset("table_type") = "VIEW" Then
           '      Me.ComboBoxTable.AddItem tablename
           End If
        End If
        oRecordset.MoveNext
      Loop
     End If

If oRecordset.State = adStateOpen Then oRecordset.Close
If Me.ComboBoxTable.ListCount Then
   Me.ComboBoxTable.ListIndex = 0
   ' On fixe la table dernièrement sélectionnée (si elle existe)
   For i = 0 To Me.ComboBoxTable.ListCount - 1
      If UCase$(Me.ComboBoxTable.List(i)) = UCase$(last_table) Then
         Me.ComboBoxTable.ListIndex = i
         Exit For
      End If
   Next
Else
   MsgBox l10n(TXT_AUCUNE_TABLE_N_EST_DISPONIBLE_OU_BIEN_VOUS_N_AVEZ_AUCUN_DROIT_DE_CONSULTATION_SUR_LES_TABLES_EXISTANTES) & Chr$(13) & _
          l10n(TXT_CONTACTEZ_L_ADMINISTRATEUR)
    End
          
End If

End Sub





Private Sub ComboBoxSchema_Change()
If Me.ComboBoxSchema.text <> "" Then
    gSchemaName = Me.ComboBoxSchema.text
    Me.ComboBoxTable.Clear
    If Not gTablenames Is Nothing Then
        Set gTablenames = Nothing
    End If
    Set gTablenames = New Collection
    ' update ComboBoxTable
    updateComboBoxTableWithSchema (gSchemaName)
End If
End Sub

Private Sub ComboBoxTable_Change()
Dim tablename As String
On Error GoTo fin
tablename = gTablenames(Me.ComboBoxTable.text)
If gTableName <> tablename Then
    gTableName = tablename
    gEntitynum = gPgConnexion.SqlQuery(buildSqlSelectLayerEntitynum(tablename))
    gSqlQuery = ""
    gSqlWhere = ""
    gSqlWhereUser = ""
    gSqlWhereFence = ""
End If
fin:
End Sub

Private Sub CommandButtonOk_Click()
  Dim tablename As String
  Dim schemaname As String
  Dim s_tmp As String
   
  Me.OptionButtonNone.TripleState = False
  Me.OptionButtonLayerOnly.TripleState = True
  Me.OptionButtonFullFile.TripleState = False
    
  ' Sauvegarde le nom de la dernière table sélectionnée et l'utilisation de MSCATALOG.
  SaveSetting "PgConnector", "UserFormSelectLayer", "LastTableName", gTableName
  SaveSetting "PgConnector", "UserFormSelectLayer", "LastSchemaName", gSchemaName
   ' Enregistrement de la position de la fenètre dans la base de registre
  SaveSetting "PgConnector", "UserFormSelectLayer", "top", Me.Top
  SaveSetting "PgConnector", "UserFormSelectLayer", "left", Me.Left
    
' SRID
gPgSrid = getLayerSrid(gSchemaName, gTableName)

' requete
gSqlWhere = ""

If gSqlWhereUser <> "" Then
    gSqlWhere = " WHERE " + gSqlWhereUser
End If

gSqlQueryLabels = ""
' utiliser clôture (mode bloc)
If Me.CheckBoxUseFence.value = True And Application.ActiveDesignFile.Fence.IsDefined = True Then
    gSqlQuery = buildSqlSelectLayerObjects(gSchemaName, gTableName, gPgSrid, gSqlWhereUser, True)
    If Me.CheckBoxLabels.value = True Then
        gSqlQueryLabels = buildSqlSelectLayerLabels(gSchemaName, gTableName, gPgSrid, True)
    End If
Else
    gSqlQuery = buildSqlSelectLayerObjects(gSchemaName, gTableName, gPgSrid, gSqlWhereUser, False)
    If Me.CheckBoxLabels.value = True Then
        gSqlQueryLabels = buildSqlSelectLayerLabels(gSchemaName, gTableName, gPgSrid, False)
    End If
End If

'gSqlQuery = gSqlQuery + " LIMIT 19"
If gCheckoutMode = enuCheckoutModeImport Then
    Call pgcImportPgLayer
Else
    Call pgcAttachPgLayer
End If

End Sub

Private Sub CommandButtonCancel_Click()
Me.Hide
Application.CommandState.StartDefaultCommand
End Sub



Private Sub OptionButtonFullFile_Click()
gPostgisDgnResetLevel = enuPostgisDgnResetFullFile
Me.OptionButtonNone.value = False
Me.OptionButtonLayerOnly.value = False
Me.OptionButtonFullFile.value = True
End Sub

Private Sub OptionButtonLayerOnly_Click()
gPostgisDgnResetLevel = enuPostgisDgnResetLayerOnly
Me.OptionButtonNone.value = False
Me.OptionButtonLayerOnly.value = True
Me.OptionButtonFullFile.value = False
End Sub

Private Sub OptionButtonNone_Click()
gPostgisDgnResetLevel = enuPostgisDgnResetNone
Me.OptionButtonNone.value = True
Me.OptionButtonLayerOnly.value = False
Me.OptionButtonFullFile.value = False
End Sub

Private Sub TextBoxWhere_Change()
gSqlWhereUser = Me.TextBoxWhere
End Sub

Private Sub UserForm_Activate()
If gCheckoutMode = enuCheckoutModeImport Then
    Me.Caption = l10n(TXT_IMPORTER_COUCHE_POSTGIS) + " " + PgcVersion
    Me.OptionButtonNone.Top = 12
    Me.OptionButtonLayerOnly.Top = 36
    Me.OptionButtonFullFile.Top = 128
Else
    Me.Caption = l10n(TXT_ATTACHER_COUCHE_POSTGIS_EN_REFERENCE) + " " + PgcVersion
    Me.OptionButtonNone.Top = 6
    Me.OptionButtonLayerOnly.Top = 24
    Me.OptionButtonFullFile.Top = 48
End If
End Sub

Private Sub UserForm_Initialize()

Dim oRecordset As ADODB.Recordset
Dim sql As String
Dim i As Integer
Dim tablename As String
Dim table_arr() As String
Dim j As Integer
Dim last_table As String
Dim last_schema As String
Dim use_table As Boolean
Dim lngLeft As Long
Dim lngTop As Long
Dim schemas_names As String
Dim schemas_arr() As String

'Récupère la dernière position de la feuille. Cette position est enregistrée dans une clé du registre
'Il faut penser à mettre la propriété startUpPosition en Manuel
lngTop = GetSetting("PgConnector", "UserFormSelectLayer", "Top", "200")
lngLeft = GetSetting("PgConnector", "UserFormSelectLayer", "Left", "300")
Me.Move lngLeft, lngTop

Me.ComboBoxTable.Clear

'Récupère le nom de la dernière table sélectionnée.
last_table = GetSetting("PgConnector", "UserFormSelectLayer", "LastTableName", "")
last_schema = GetSetting("PgConnector", "UserFormSelectLayer", "LastSchemaName", "Utilisateur")

' chargement liste des schémas
If Not gPgConnexion Is Nothing Then
    If gPgConnexion.EstConnecte = True Then
        GoTo suite
    End If
End If
MsgBox "Pour utiliser cette commande, vous devez être connecté à la BD.", vbExclamation
Me.Hide
Application.CommandState.StartDefaultCommand
End
suite:
    Set oRecordset = New Recordset
    For i = 0 To UBound(gSchemasArray)
        Me.ComboBoxSchema.AddItem Trim(gSchemasArray(i))
        If last_schema = Trim(gSchemasArray(i)) Then Me.ComboBoxSchema.ListIndex = i
    Next i
    If Me.ComboBoxSchema.text = "" Then
        Me.ComboBoxSchema.text = gSchemasArray(0)
    End If


If Me.ComboBoxTable.ListCount < 1 Then
   MsgBox l10n(TXT_AUCUNE_TABLE_N_EST_DISPONIBLE_OU_BIEN_VOUS_N_AVEZ_AUCUN_DROIT_DE_CONSULTATION_SUR_LES_TABLES_EXISTANTES) & Chr$(13) & _
          l10n(TXT_CONTACTEZ_L_ADMINISTRATEUR)
    End
End If

Me.OptionButtonNone.value = True

gPgSrid = PgDefaultSrid

Me.Label2.Caption = l10n(TXT_SCHEMA)
Me.Label1.Caption = l10n(TXT_TABLE)
Me.Frame1.Caption = l10n(TXT_FILTRE)
Me.Frame2.Caption = l10n(TXT_REINITIALISER)

Me.CheckBoxLabels.Caption = l10n(TXT_LABELS)
Me.CheckBoxUseFence.Caption = l10n(TXT_UTILISER_CLOTURE)

Me.OptionButtonFullFile.Caption = l10n(TXT_FICHIER_ENTIER)
Me.OptionButtonLayerOnly.Caption = l10n(TXT_COUCHE_SEULEMENT)
Me.OptionButtonNone.Caption = l10n(TXT_AUCUNE)

Me.CommandButtonOk.Caption = l10n(TXT_OK)
Me.CommandButtonCancel.Caption = l10n(TXT_QUITTER)

End Sub


