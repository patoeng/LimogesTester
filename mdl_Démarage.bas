Attribute VB_Name = "mdl_D�marage"
Option Explicit
Global Const CST_FILE_PARA_INI = "c:\banc_xs\config.ini"
Global Mode_ouverture_fenetre As Integer
Global Imprime_etiquette_ecran As Boolean
Global Mode_execution_debug As Boolean

Sub Main()

Dim fs 'object system
Dim f 'object fichier

Dim commandLine
Dim i As Integer
    
Dim Num_enregistrement_log As Boolean
    
    If App.PrevInstance = True Then
        'commandLine = GetCommandLine(3)
        'If UBound(commandLine) = 0 Then
            'MsgBox " L'application IHM banc XS est d�ja lanc�e ", vbInformation, "Lancement IHM banc XS"
            End
        'Else
            
        'End If
    Else
        commandLine = GetCommandLine()
        
        If UBound(commandLine) = 0 Then
            'MsgBox "Il faut au moins un argument pour lancer l'application IHM banc XS ", vbInformation, "Lancement IHM banc XS"
            MsgBox "Required an argument to start Program HMI Final Test Bench XS", vbInformation, "Startup HMI Bench XS"
            Print_parameter_startup_setint_help
            End
        End If
        'Valeur  par d�faut
        Mode_ouverture_fenetre = vbNormal
        File_para_ini = CST_FILE_PARA_INI
    
        'Valeur  par d�faut
        Imprime_etiquette_ecran = False
        Mode_execution_debug = False
        Mode_ouverture_fenetre = vbNormal
        
        Num_enregistrement_log = False
        
        File_para_ini = "" 'App.Path & "/config.ini"
        
        For i = 1 To UBound(commandLine)
            Select Case UCase(commandLine(i))
                Case "HELP"
                   Print_parameter_startup_setint_help
                   End
                Case "MAX_WINDOWS"
                    Mode_ouverture_fenetre = vbMaximized
                Case "LOG_DATA_BASE_REGISTER_NUMBER_VISIBLE"
                    Num_enregistrement_log = True
                Case "DEBUG"
                    Mode_ouverture_fenetre = vbNormal
                    Mode_execution_debug = True
                Case "SCREEN_PRINT"
                    Imprime_etiquette_ecran = True
                Case "PRINTER_PRINT"
                    Imprime_etiquette_ecran = False
                Case Else
                    If i = UBound(commandLine) Then
                         File_para_ini = commandLine(i)
                    Else
                        Print_parameter_startup_setint_help
                    End If
            End Select
        Next i
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Not fs.FileExists(File_para_ini) Then
            'MsgBox " Le fichier de configuration" & vbCrLf & File_para_ini & vbCrLf & "n'exite pas", vbCritical + vbMsgBoxHelpButton, "Lancement IHM banc XS"
            MsgBox "Configuration file " & vbCrLf & File_para_ini & vbCrLf & "not exist", vbCritical + vbMsgBoxHelpButton, "Startup HMI Bench XS"
            End
        End If
        
        frmSplash.Show
        DoEvents
        
        'on va chercher les param�tres de r�seau unitel way de l'IHM et de API
        Unitel_Adresse = Get_Adresse_API(File_para_ini)
        
        Non_banc = Get_Non_banc(File_para_ini)
        DoEvents
        DataEnvironment1.Data_Base.ConnectionString = Get_DB_Banc(File_para_ini)
        DataEnvironment1.Data_Log.ConnectionString = Get_DB_Banc_log(File_para_ini)
        DataEnvironment1.Data_cycle.ConnectionString = Get_DB_Banc_Cycles(File_para_ini)
        '======DEBUT MODIF BEEA 30/11/2007===========
        DataEnvironment1.DataCtrlRefClaquage.ConnectionString = Get_DB_TestClaquage(File_para_ini)
        '======FIN MODIF BEEA 30/11/2007=============
        
        Utwmait.Winsock1.RemoteHost = Get_IP_API(File_para_ini)
        DoEvents
        
        If DataEnvironment1.rsTable_Log.State = adStateClosed Then
            DataEnvironment1.rsTable_Log.Open
        End If
        DoEvents
        If DataEnvironment1.rsTable_Log_Banc.State = adStateClosed Then
            DataEnvironment1.rsTable_Log_Banc.Open
        End If
        DoEvents
        If DataEnvironment1.rsTable_Message.State = adStateClosed Then
            DataEnvironment1.rsTable_Message.Open
        End If
        DoEvents
        
        FRM_startup.WindowState = Mode_ouverture_fenetre
        FRM_controle.WindowState = Mode_ouverture_fenetre
        DoEvents
        
        FRM_controle.Text1.Visible = Num_enregistrement_log
        FRM_controle.Text2.Visible = Num_enregistrement_log
        DoEvents
        
        Unload FRM_controle
        Unload FRM_startup
        DoEvents
        
        'Charge la forme Uni-Telway pour activer la communication
        Socket_valid = True 'Socket valid au d�marage de l'aplication par d�faut
        Load Utwmait
        DoEvents
        
        If Mode_execution_debug Then
            Utwmait.Show
        End If
         
        FRM_startup.Show
        DoEvents
    End If
    
End Sub

'Print a window help for starting seting parameter application
Sub Print_parameter_startup_setint_help()

    frm_help_seting_parameter_application.Show vbModal

End Sub

'
' Lie la chaine pas�e l'ors du lancement du programme
'' Link the argument string pass during startup program
Function GetCommandLine(Optional MaxArgs)
' D�clare les variables.
Dim C, CmdLine, CmdLnLen, InArg, i, NumArgs
   ' V�rifie si MaxArgs a �t� sp�cifi�.
   If IsMissing(MaxArgs) Then MaxArgs = 10
   ' D�finit un tableau au format appropri�.
   ReDim ArgArray(MaxArgs)
   NumArgs = 0: InArg = False
   ' R�cup�re les arguments de ligne de commande.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   ' Analyse de la ligne de commande caract�re par caract�re.
   For i = 1 To CmdLnLen
      C = Mid(CmdLine, i, 1)
      ' Analyse de caract�res d'espacement ou de tabulations.
      If (C <> " " And C <> vbTab) Then
      'If (C <> vbTab) Then
         ' Ni espace ni tabulation.
         ' V�rifie une �ventuelle pr�sence dans l'argument.
         If Not InArg Then
         ' Le nouvel argument commence.
         ' V�rifie si les arguments ne sont pas trop nombreux.
            If NumArgs = MaxArgs Then Exit For
               NumArgs = NumArgs + 1
               InArg = True
            End If
         ' Concat�ne un caract�re � l'argument courant.
         ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
         ' Recherche un espace ou une tabulation.
         ' L'indicateur InArg prend la valeur False.
         InArg = False
      End If
   Next i
   ' Redimensionne le tableau pour qu'il puisse
   ' juste contenir les arguments.
   ReDim Preserve ArgArray(NumArgs)
   ' Renvoie le tableau dans le nom de fonction.
   GetCommandLine = ArgArray()
End Function
