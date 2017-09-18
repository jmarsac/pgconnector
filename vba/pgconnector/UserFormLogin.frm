VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLogin 
   Caption         =   "Connexion 1.1"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5700
   OleObjectBlob   =   "UserFormLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormLogin"
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

Private Sub ComboBoxConnection_Change()
gPgConnectionName = Me.ComboBoxConnection.text
If gPgConnectionName <> "" Then
    gPgHost = GetSetting("PgConnector", gPgConnectionName, "Host", "")
    gPgPort = GetSetting("PgConnector", gPgConnectionName, "Port", "")
    gPgDbname = GetSetting("PgConnector", gPgConnectionName, "Dbname", "")
    gPgUsername = GetSetting("PgConnector", gPgConnectionName, "Username", "")
    If gPgUsername <> "" Then
        Me.TextBoxUser.text = gPgUsername
        Me.CheckBoxRecLogin.value = True
        gPgPassword = GetSetting("PgConnector", gPgConnectionName, "Password", "")
        If gPgPassword <> "" Then
            Me.TextBoxPassword.text = String(Len(gPgPassword), "*")
            Me.CheckBoxRecPwd.value = True
        Else
            Me.CheckBoxRecPwd.value = False
        End If
    Else
        Me.CheckBoxRecLogin.value = False
    End If
End If
End Sub

Private Sub CommandButtonCancel_Click()
Me.Hide
Application.CommandState.StartDefaultCommand
End Sub

Private Sub CommandButtonOk_Click()
If gPgConnexion Is Nothing Then
    Set gPgConnexion = azidblib.CreerConnexion
End If
If checkOpenConnection(gPgConnexion, gPgHost, gPgPort, gPgDbname, gPgUsername, gPgPassword) = False Then
    Application.MessageCenter.AddMessage l10n(TXT_ECHEC_CONNEXION_CONTACTEZ_L_ADMINISTRATEUR), l10n(TXT_IMPOSSIBLE_DE_SE_CONNECTER_A_LA_BASE_DE_DONNEES_CONTACTEZ_L_ADMINISTRATEUR_POUR_RESOUDRE_LE_PROBLEME), msdMessageCenterPriorityError, True
    End
Else
    Application.MessageCenter.AddMessage l10n(TXT_CONNECTE_A_LA_BD) + " '" + gPgDbname + "'", , msdMessageCenterPriorityInfo, False
    SaveSetting "PgConnector", "Connections", "ConnectionName", Me.ComboBoxConnection.text
    SaveSetting "PgConnector", "Connections", "DefaultConnectionName", Me.ComboBoxConnection.text
    If Me.CheckBoxRecLogin.value = True Then
        SaveSetting "PgConnector", Me.ComboBoxConnection.text, "Username", Me.TextBoxUser.text
        If Me.CheckBoxRecPwd.value = True Then
            SaveSetting "PgConnector", Me.ComboBoxConnection.text, "Password", gPgPassword
        End If
    End If
    Me.Hide
End If

End Sub

Private Sub TextBoxPassword_Change()
If Me.TextBoxPassword.text = "" Then
    gPgPassword = ""
Else
    If Right(Me.TextBoxPassword.text, 1) <> "*" Then
        gPgPassword = gPgPassword + Right(Me.TextBoxPassword.text, 1)
        Me.TextBoxPassword.text = String(Len(gPgPassword), "*")
    End If
End If
End Sub

Private Sub TextBoxUser_Change()
If Me.TextBoxUser.text <> "" And Me.TextBoxUser.text <> gPgUsername Then
    gPgUsername = Me.TextBoxUser.text
End If
End Sub

Private Sub UserForm_Activate()
Dim s_tmp As String
Dim cnxNames() As String
Dim i As Integer

gPgCnxList = GetSetting("PgConnector", "Connections", "ConnectionList", "")
Me.ComboBoxConnection.Clear
If gPgCnxList <> "" Then
    cnxNames = Split(gPgCnxList, ";")
    For i = 0 To UBound(cnxNames)
        Me.ComboBoxConnection.AddItem cnxNames(i)
    Next i
Else
    MsgBox l10n(TXT_AUCUNE_CONNEXION_DEFINIE_VEUILLEZ_EN_CREER_UNE), vbCritical
    Me.Hide
End If
s_tmp = GetSetting("PgConnector", "Connections", "DefaultConnectionName", "")
If s_tmp <> "" Then
    Me.ComboBoxConnection.text = s_tmp
    gPgConnectionName = s_tmp
    s_tmp = GetSetting("PgConnector", gPgConnectionName, "Host", "")
    If s_tmp <> "" Then
        gPgHost = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", gPgConnectionName, "Port", "")
    If s_tmp <> "" Then
        gPgPort = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", gPgConnectionName, "Dbname", "")
    If s_tmp <> "" Then
        gPgDbname = s_tmp
    End If
    s_tmp = GetSetting("PgConnector", gPgConnectionName, "Username", "")
    If s_tmp <> "" Then
        Me.TextBoxUser.text = s_tmp
        gPgUsername = s_tmp
        s_tmp = GetSetting("PgConnector", gPgConnectionName, "Password", "")
        If s_tmp <> "" Then
            gPgPassword = s_tmp
            Me.TextBoxPassword.text = String(Len(gPgPassword), "*")
            s_tmp = ""
        End If
    End If
End If

End Sub

Private Sub UserForm_Initialize()

Me.Label7.Caption = l10n(TXT_UTILISATEUR)
Me.Label8.Caption = l10n(TXT_MOT_DE_PASSE)
Me.Label9.Caption = l10n(TXT_CONNEXION)
Me.CommandButtonOk.Caption = l10n(TXT_OK)
Me.CommandButtonCancel.Caption = l10n(TXT_QUITTER)

Me.Caption = l10n(TXT_CONNEXION) + " " + PgcVersion

End Sub
