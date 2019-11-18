# CalendarScripts

Scripts pour extraire de l'information de facturation des calendriers Google pour les plateformes.

Scripts originaux de Simon Labrecque.

## Message de Simon

Afin de déterminer le nombre d’heures utilise par les labos vous devez aller au :

https://google-calendar-hours.com/

Selectionner le calendar du microscope qui vous intéresse, sélectionner show/hide details  et exporter le calendar en cvs.

Vous devez ensuite construire ou avoir déjà construit un fichier excel qui contient le nom du chef du labo sur la 1ere ligne et tous les membres de son labo dans la colonne en dessous. Voir le fichier : Members_LabList.xlsx.

*Le script ne reconnait pas les sections de string, il faut donc avoir les noms complets. Je crois que ce serait une belle implémentation à faire.

Il faut ensuite utiliser Google_Calendars_HourInUSe.py et mettre le directory du fichier exporté par goole-calendar-hours.com ainsi que le nom du fichier. Par exemple : ligne 26 et 27 du présent script:

directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\PDK-TIRF"

file = 'PDK_TIRF.xlsx'

Ceci va donc générer un fichier dont les users seront affiliés à un laboratoire. Le fichier sauvegardé aura le m^me nom que l’original mais avec un ‘_user.xlsx’ à la fin. Il suffit ensuite d’utiliser Google_Calendars_HourInUSe.py pour identifier le temps d’utilisation par labo. On doit encore entrer manuellement le directory et le nom du fichier i.e. :

directory = "Y:\Simon\Platforme_Imagerie\GoogleCalendar_time\PDK-TIRF"

file = 'PDK_TIRF_user.xlsx'

 On peut changer les dates à évaluer le temps d’utilisation aux ligne 61 et 65. Les mois en changeant les valeurs dans range(1,13) et l’année an changeant '.2018'.

En output un graph qui va déterminer l’utilisation pas mois de chaque labo avec la valeure numérique dans la legende. Voilà bonne planification.

Simon

 

 