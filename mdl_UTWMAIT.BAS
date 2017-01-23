Attribute VB_Name = "mdl_UTWMAIT1"
Option Explicit
DefInt A-Z
'Global variables definition

Global com_en_cours, Lecture_en_cours, Ecriture_en_cours

' UNITELWAY <<<<<--------------------------
Global Const NB_ADRESSES = 1   ' Number of polled addresses

Global adresse_scrutee%    '

Global TX_requete$         ' Frame to be sent, ready-to-go (except double DLEs)
Global TX_requete_ok%      ' Frame to be sent is validated
Global TX_requete_len%     ' Frame length

Global tampon$             ' Buffer to stock
Global RX_requete$         ' Frame received
Global RX_requete_ok%      ' Received frame is validated
Global RX_requete_len%     ' Received frame length
Global RX_en_cours%        ' =TRUE if we are receiveing a frame


Global etape_polling%
Global nb_timer_no_RX%     ' Number of timer click without receiving anything

Global code_erreur_utw%    ' 0 (no error) to ...

Global chaine_polling$

'<<<<< FIN UTW ------------------------------------------------

