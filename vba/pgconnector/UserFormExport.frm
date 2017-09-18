VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormExport 
   Caption         =   "Export 1.1"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   OleObjectBlob   =   "UserFormExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormExport"
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


Private Sub ComboBoxLevel_Change()
If Me.ComboBoxLevel.ListCount > 0 Then
    If Me.ComboBoxLevel.text <> "" Then
        gLevelName = Me.ComboBoxLevel.text
        Set gLevel = Application.ActiveDesignFile.Levels.Find(gLevelName)
    End If
End If
End Sub

Private Sub ComboBoxSchema_Change()
If Me.ComboBoxSchema.text <> "" Then
    gSchemaName = Me.ComboBoxSchema.text
End If
End Sub

Private Sub CommandButtonCancel_Click()
Me.Hide
Application.CommandState.StartDefaultCommand
End Sub

Private Sub CommandButtonOk_Click()
pgcExportLevelToPg
End Sub

Private Sub UserForm_Activate()
Dim lv As Variant

If Me.ComboBoxLevel.ListCount Then
    Me.ComboBoxLevel.Clear
End If
For Each lv In Application.ActiveModelReference.Levels
    Me.ComboBoxLevel.AddItem lv.name
Next
End Sub

Private Sub UserForm_Initialize()
Dim last_schema As String
Dim i As Integer

'Récupère le nom de la dernière table sélectionnée et l'utilisation de MSCATALOG.
last_schema = GetSetting("PgConnector", "UserFormSelectLayer", "LastSchemaName", "public")

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
    For i = 0 To UBound(gSchemasArray)
        Me.ComboBoxSchema.AddItem Trim(gSchemasArray(i))
        If last_schema = Trim(gSchemasArray(i)) Then Me.ComboBoxSchema.ListIndex = i
    Next i
    If Me.ComboBoxSchema.text = "" Then
        Me.ComboBoxSchema.text = gSchemasArray(0)
    End If

Me.Label2.Caption = l10n(TXT_NIVEAU)
Me.Label3.Caption = l10n(TXT_SCHEMA)

Me.CommandButtonOk.Caption = l10n(TXT_OK)
Me.CommandButtonCancel.Caption = l10n(TXT_QUITTER)

Me.Caption = l10n(TXT_EXPORT) + " " + PgcVersion

End Sub
