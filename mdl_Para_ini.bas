Attribute VB_Name = "mdl_Para_ini"
Global File_para_ini As String
' Variable global du banc
Global Non_banc As String ' Non du banc

Type Structure_unite_adresse
    Host_Network As Integer
    Host_Station As Integer
    Remote_Network As Integer
    Remote_Station As Integer
End Type

'-----------------------------------------------------------------------------
' Ecriture fichier au format des fichiers .INI de Windows
' Write Windows .INI file format
' NomFichier    = File Name & Location
' Section       = Section
' Rubrique      = Parameter will be save
' chaine        = Value will be save
'-----------------------------------------------------------------------------
 Function EcrireFicIni(ByRef NomFichier As String, ByVal Section As String, ByVal Rubrique As String, ByVal chaine As String) As String
 
 Dim ret As Long

Début_EcrireFicIni:

    'Ecrire l'information demandée dans le fichier spécifié
    If chaine = "" Then chaine = "[NEANT]"
    'GetPrivateProfileString Section, Rubrique, Defaut, Resultat, Len(Resultat), NomFichier
    ret = WritePrivateProfileString(Section, Rubrique, chaine, NomFichier)

    
    If ret <> 1 And ret <> 0 Then
        'If MsgBox("Problème décriture dans le fichier" & Nomfich, vbExclamation + vbRetryCancel, "Erreur fichier INI !") = vbRetry Then
        If MsgBox("Problem writing file" & Nomfich, vbExclamation + vbRetryCancel, "ERROR INI File !") = vbRetry Then
            GoTo Début_EcrireFicIni
        End If
    End If
    
End Function

'-----------------------------------------------------------------------------
' Lecture d'un fichier au format des fichiers .INI de Windows
' Read file format WINDOW .ini
' NomFichier    = Location of .ini File
' Section       = Section
' Rubrique      = Parameter looking for
' Defaut        = Default value if not found
'-----------------------------------------------------------------------------
 Function LireFicIni(ByVal NomFichier As String, ByVal Section As String, _
    ByVal Rubrique As String, Optional ByVal Defaut = "") As String
'Lire l'information demandée dans le fichier spécifié
''Read the requested information in the file specified
Dim Resultat As String
Dim ret As Long
    Resultat = String(1024, 0)
    ret = GetPrivateProfileString(Section, Rubrique, Defaut, Resultat, Len(Resultat), NomFichier)
    Resultat = Left(Resultat, InStr(Resultat, Chr(0)) - 1)
    
    If Resultat = "[NEANT]" Then Resultat = ""
    
    LireFicIni = Resultat
End Function

Function find_path(Text As String) As String
Dim i As Integer
Dim m As Integer

m = 0
For i = 1 To Len(Text)
    If Mid(Text, i, 1) = "\" Then
        m = i
    End If
Next
If m <> 0 Then
    find_path = Left(Text, m)
Else
    find_path = ""
End If
End Function

Function find_space(Text As String) As Long
Dim i As Integer
Dim aze As Integer
Dim qsd As String


For i = 1 To Len(Text)
    aze = Asc(Mid(Text, i, 1))
    qsd = Mid(Text, i, 1)
    If Asc((Mid(Text, i, 1))) = 0 Then
    find_space = i
    Exit Function
    End If
Next
find_pace = 0
End Function

Function Get_Non_banc(ByVal NomFichier As String) As String
    Get_Non_banc = LireFicIni(NomFichier, "BANC", "Non")
End Function

Function Get_DB_Banc(ByVal NomFichier As String) As String
    Get_DB_Banc = LireFicIni(NomFichier, "Base_donnée", "Banc")
End Function

Function Get_DB_Banc_log(ByVal NomFichier As String) As String
    Get_DB_Banc_log = LireFicIni(NomFichier, "Base_donnée", "Banc_log")
End Function

Function Get_DB_Banc_Cycles(ByVal NomFichier As String) As String
    Get_DB_Banc_Cycles = LireFicIni(NomFichier, "Base_donnée", "Banc_cycles")
End Function
'======DEBUT MODIF BEEA 30/11/2007===========
Function Get_DB_TestClaquage(ByVal NomFichier As String) As String
    Get_DB_TestClaquage = LireFicIni(NomFichier, "Base_donnée", "TestClaquage")
End Function
'======FIN MODIF BEEA 30/11/2007=============

Function Set_Non_banc(ByVal NomFichier As String, Non As String)
    EcrireFicIni NomFichier, "BANC", "Non", Non
End Function

Function Get_IP_API(ByVal NomFichier As String) As String
    Get_IP_API = LireFicIni(NomFichier, "API", "Adresse_IP")
End Function

Function Set_IP_API(ByVal NomFichier As String, IP As String)
    EcrireFicIni NomFichier, "API", "Adresse_IP", IP
End Function

Function Get_Adresse_API(NomFichier As String) As Structure_unite_adresse
    Get_Adresse_API.Host_Network = CInt(LireFicIni(NomFichier, "API", "Host_Network", "0"))
    Get_Adresse_API.Host_Station = CInt(LireFicIni(NomFichier, "API", "Host_Station", "0"))
    Get_Adresse_API.Remote_Network = CInt(LireFicIni(NomFichier, "API", "Remote_Network", "0"))
    Get_Adresse_API.Remote_Station = CInt(LireFicIni(NomFichier, "API", "Remote_Station", "0"))
End Function

Function Set_Adresse_API(NomFichier As String, Adresse As Structure_unite_adresse)
    EcrireFicIni NomFichier, "API", "Host_Network", Adresse.Host_Network
    EcrireFicIni NomFichier, "API", "Host_Station", Adresse.Host_Station
    EcrireFicIni NomFichier, "API", "Remote_Network", Adresse.Remote_Network
    EcrireFicIni NomFichier, "API", "Remote_Station", Adresse.Remote_Station
End Function

Function get_soft(File As String) As String
Text$ = Space$(255)
t_len& = Len(Text$)
Dim num As Long
num = GetPrivateProfileString("menu", "soft", "", Text$, t_len&, File)
get_soft = Left(Text$, find_space(Text$) - 1)
End Function

Function set_soft(Text As String, File As String)
WritePrivateProfileString "menu", "soft", Text, File
End Function

Function get_soft_parameter(File As String) As String
Text$ = Space$(255)
t_len& = Len(Text$)
Dim num As Long
num = GetPrivateProfileString("menu", "soft start up argument", "", Text$, t_len&, File)
get_soft_parameter = Left(Text$, find_space(Text$) - 1)
End Function

Function set_soft_parameter(Text As String, File As String)
WritePrivateProfileString "menu", "soft start up argument", Text, File
End Function

Function get_log_path() As String
    Text$ = Space$(255)
    t_len& = Len(Text$)
    Dim num As Long
    num = GetPrivateProfileString("Application", "Log_path", "C:\", Text$, t_len&, File_para_ini)
    get_log_path = Left(Text$, find_space(Text$) - 1)
End Function

Function set_log_path(Text As String)
WritePrivateProfileString "Application", "Log_path", Text, File_para_ini
End Function

Function get_log_path_copy() As String
Text$ = Space$(255)
t_len& = Len(Text$)
Dim num As Long
num = GetPrivateProfileString("Application", "Log_path_copy", "C:\", Text$, t_len&, File_para_ini)
get_log_path_copy = Left(Text$, find_space(Text$) - 1)
End Function

Function set_log_path_copy(Text As String)
WritePrivateProfileString "Application", "Log_path_copy", Text, File_para_ini
End Function

