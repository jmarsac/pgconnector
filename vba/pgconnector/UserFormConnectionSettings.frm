VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormConnectionSettings 
   Caption         =   "Connexion 1.1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12990
   OleObjectBlob   =   "UserFormConnectionSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormConnectionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Function LoadConnectionSettings(cnxName As String)
Dim s_tmp As String

gPgPassword = ""
Me.TextBoxPwd.text = ""
Me.TextBoxUsr.text = ""

If cnxName <> "" Then
    Me.TextBoxCnxName.text = cnxName
    s_tmp = GetSetting("PgConnector", cnxName, "Host", "")
    If s_tmp <> "" Then
        Me.TextBoxHost.text = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", cnxName, "Port", "")
    If s_tmp <> "" Then
        Me.TextBoxPort.text = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", cnxName, "Dbname", "")
    If s_tmp <> "" Then
        Me.TextBoxDbName.text = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", cnxName, "Username", "")
    If s_tmp <> "" Then
        Me.TextBoxUsr.text = s_tmp
        s_tmp = GetSetting("PgConnector", cnxName, "Password", "")
        If s_tmp <> "" Then
            gPgPassword = s_tmp
            Me.TextBoxPwd.text = String(Len(gPgPassword), "*")
        End If
    End If
Else
    Me.TextBoxCnxName = "IBGE"
    Me.TextBoxHost = "wfs.ibgebim.be"
    Me.TextBoxPort = "5432"
    Me.TextBoxDbName = "ibge"
End If


End Function
Private Function UpdateCnxList(cnxName As String, Optional delete_it As Boolean = False)

Dim cnxNames() As String
Dim found As Boolean
Dim i As Integer

If delete_it = True Then
    If gPgCnxList <> cnxName Then
        gPgCnxList = Replace(gPgCnxList, ";" + cnxName + ";", ";")
        gPgCnxList = Replace(gPgCnxList, ";" + cnxName, "")
        gPgCnxList = Replace(gPgCnxList, cnxName + ";", "")
    Else
        gPgCnxList = ""
    End If
Else
    If gPgCnxList <> "" Then
        cnxNames = Split(gPgCnxList, ";")
        For i = 0 To UBound(cnxNames)
            If cnxName = cnxNames(i) Then
                found = True
                Exit For
            End If
        Next i
        If found = False Then
            gPgCnxList = gPgCnxList + ";" + cnxName
        End If
    Else
        gPgCnxList = cnxName
    End If
End If
End Function
Private Function CheckConnection() As Boolean
Dim azicnx As azidblib.aziDbConnexion
Dim pgOdbcDriver As String

pgOdbcDriver = "PostgreSQL Unicode"
If Application.ActiveWorkspace.IsConfigurationVariableDefined("PGC_ODBC_DRIVER") = True Then
    pgOdbcDriver = Application.ActiveWorkspace.ConfigurationVariableValue("PGC_ODBC_DRIVER")
End If
Set azicnx = azidblib.CreerConnexion
If Me.TextBoxHost.text <> "" And Me.TextBoxPort.text <> "" And Me.TextBoxDbName.text <> "" And Me.TextBoxUsr.text <> "" And gPgPassword <> "" Then
    If checkOpenConnection(azicnx, Me.TextBoxHost.text, Me.TextBoxPort.text, Me.TextBoxDbName.text, Me.TextBoxUsr.text, gPgPassword) = True Then
        CheckConnection = True
    Else
        CheckConnection = False
        MsgBox "Echec de la connexion" + vbCrLf + Err.Description, vbCritical
    End If
End If
Set azicnx = Nothing
End Function

Private Sub CommandButtonCancel_Click()
Me.Hide
Application.CommandState.StartDefaultCommand
End Sub

Private Sub CommandButtonCheckConnection_Click()
If CheckConnection = True Then
    MsgBox l10n(TXT_CONNEXION_A_LA_BD_REUSSIE)
Else
    MsgBox l10n(TXT_ECHEC_DE_LA_CONNEXION_A_LA_BD) + vbCrLf + Err.Description
End If
End Sub

Private Sub CommandButtonDelConnection_Click()
Dim cnxNames() As String
Dim y_or_n As VbMsgBoxResult
Dim defaultCnxName As String
Dim cnxName As String

y_or_n = MsgBox(l10n(TXT_SUPPRIMER_CONNEXION) + " ?", vbYesNo, l10n(TXT_CONNEXION) + " " + PgcVersion)
If y_or_n = vbYes Then
    On Error Resume Next
    Call DeleteSetting("PgConnector", Me.TextBoxCnxName.text)
    On Error GoTo 0
    UpdateCnxList Me.TextBoxCnxName.text, True
    SaveSetting "PgConnector", "Connections", "ConnectionList", gPgCnxList
    
    cnxNames = Split(gPgCnxList, ";")
    If UBound(cnxNames) >= 0 Then
        SaveSetting "PgConnector", "Connections", "ConnectionName", cnxNames(0)
        SaveSetting "PgConnector", "Connections", "DefaultConnectionName", cnxNames(0)
        LoadConnectionSettings cnxNames(0)
    Else
        SaveSetting "PgConnector", "Connections", "ConnectionName", ""
        SaveSetting "PgConnector", "Connections", "DefaultConnectionName", ""
    End If
    
    gPgPassword = ""
    Me.TextBoxPwd.text = ""
    Me.TextBoxUsr.text = ""
End If

End Sub

Private Sub CommandButtonOk_Click()

If Me.TextBoxCnxName.text = "" Then
    MsgBox l10n(TXT_INDIQUER_UN_NOM_POUR_LA_CONNEXION)
    Exit Sub
End If
If CheckConnection = True Then
    SaveSetting "PgConnector", Me.TextBoxCnxName.text, "Host", Me.TextBoxHost.text
    SaveSetting "PgConnector", Me.TextBoxCnxName.text, "Port", Me.TextBoxPort.text
    SaveSetting "PgConnector", Me.TextBoxCnxName.text, "Dbname", Me.TextBoxDbName.text
    If Me.CheckBoxRecLogin.value = True Then
        SaveSetting "PgConnector", Me.TextBoxCnxName.text, "Username", Me.TextBoxUsr.text
        If Me.CheckBoxRecPwd.value = True Then
            SaveSetting "PgConnector", Me.TextBoxCnxName.text, "Password", gPgPassword
        End If
    End If
    SaveSetting "PgConnector", "Connections", "ConnectionName", Me.TextBoxCnxName.text
    UpdateCnxList Me.TextBoxCnxName.text
    SaveSetting "PgConnector", "Connections", "ConnectionList", gPgCnxList
    SaveSetting "PgConnector", "Connections", "DefaultConnectionName", Me.TextBoxCnxName.text
    Me.Hide
Else
    MsgBox l10n(TXT_CONNEXION_INVALIDE) + vbCrLf + Err.Description
End If

End Sub


Private Sub TextBoxCnxName_Change()
Dim pgCnxList As String
Dim cnxNames() As String
Dim i As Integer

gPgPassword = ""
Me.TextBoxPwd.text = ""
Me.TextBoxUsr.text = ""
If Me.TextBoxCnxName.text <> "" Then
    pgCnxList = GetSetting("PgConnector", "Connections", "ConnectionList")
    If pgCnxList <> "" Then
        cnxNames = Split(pgCnxList, ";")
        If UBound(cnxNames) >= 0 Then
            For i = 0 To UBound(cnxNames)
                If Me.TextBoxCnxName.text = cnxNames(i) Then
                    LoadConnectionSettings cnxNames(i)
                End If
            Next i
        End If
    End If
End If
End Sub

Private Sub TextBoxPwd_Change()
If Me.TextBoxPwd.text = "" Then
    gPgPassword = ""
Else
    If Right(Me.TextBoxPwd.text, 1) <> "*" Then
        gPgPassword = gPgPassword + Right(Me.TextBoxPwd.text, 1)
        Me.TextBoxPwd.text = String(Len(gPgPassword), "*")
    End If
End If
End Sub





Private Sub UserForm_Activate()

Dim s_tmp As String
Dim cnxName As String

Me.ComboBoxSslMode.Visible = False
Me.ComboBoxSslMode.Enabled = False
Me.Label6.Visible = False
Me.Label6.Enabled = False

gPgPassword = ""
Me.TextBoxPwd.text = ""
Me.TextBoxUsr.text = ""

gPgCnxList = GetSetting("PgConnector", "Connections", "ConnectionList", "")
s_tmp = GetSetting("PgConnector", "Connections", "DefaultConnectionName", "")
LoadConnectionSettings s_tmp

End Sub


Private Sub UserForm_Initialize()

Me.TextBoxCnxName = "Test"
Me.TextBoxHost = "wfs.ibgebim.be"
Me.TextBoxPort = "5432"
Me.TextBoxDbName = "ibge"

Me.Label2.Caption = l10n(TXT_NOM)
Me.Label3.Caption = l10n(TXT_SERVEUR)
Me.Label4.Caption = l10n(TXT_PORT)
Me.Label5.Caption = l10n(TXT_B_D)
Me.Label6.Caption = l10n(TXT_MODE_SSL)
Me.Label7.Caption = l10n(TXT_NOM)
Me.Label8.Caption = l10n(TXT_MOT_DE_PASSE)
Me.Frame1.Caption = l10n(TXT_IDENTIFICATION)
Me.CheckBoxRecLogin.Caption = l10n(TXT_ENREGISTRER)
Me.CheckBoxRecPwd.Caption = l10n(TXT_ENREGISTRER)
Me.CommandButtonCheckConnection.Caption = l10n(TXT_TESTER_LA_CONNEXION)

Me.CommandButtonOk.Caption = l10n(TXT_OK)
Me.CommandButtonCancel.Caption = l10n(TXT_QUITTER)
Me.CommandButtonDelConnection.Caption = l10n(TXT_SUPPRIMER_CONNEXION)

Me.Caption = l10n(TXT_CONNEXION) + " " + PgcVersion

End Sub
