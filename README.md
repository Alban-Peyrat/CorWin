# CorWin - Contrôle WinIBW

Corwin est un outil visant à contrôler des données du Sudoc en passant par WinIBW. Corwin impliquant un plus grand nombre d'actions de la part de l'utilisateur que [Constance](https://github.com/Alban-Peyrat/ConStance), il n'est pour l'instant utilisé que pour des contrôles impossible pour Constance.

**Évitez d'avoir d'autres fichiers Excel ouverts pendant l'analyse (dans le cas où une erreur de programmation pourrait faire intéragir Corwin avec des fichiers non prévus).**

_Version du 14/10/2021._

## Initialisation

### À partir d'une liste de PPN d'Alma

Exportez d'Alma une liste de Titres physiques, renommez-la `export_Alma_CorWin.xlsx` et placez-la dans le même dossier que `CorWin.xlsm` (CorWin dans le reste de la documentation).

Allumez ConStance, choisissez la feuille `Introduction` et remettez à zéro les données. Importez ensuite les données d'Alma. Sélectionnez en `H2` le contrôle à effectuer (en appuyant sur Alt + flèche du bas, une liste déroulante s'affichera).

### À partir d'une liste de PPN déjà établie

Allumez CorWin, choisissez la feuille `Introduction` et remettez à zéro les données. Collez votre liste de PPN dans la colonne associée. Sélectionnez en `H2` le contrôle à effectuer (en appuyant sur Alt + flèche du bas, une liste déroulante s'affichera).

Notes : ConStance prend en compte les 9 derniers caractères de la cellule, si votre liste se présente sous la forme `PPN 123456789` ou `(PPN)123456789` ce n'est pas la peine de la retoucher, ni de rajouter des 0 en début de PPN, elle les ajoutera automatiquement.

## Export de WinIBW

Pour exporter les données de WinIBW, commencez par vous y connecter, puis copier la liste de PPN depuis Corwin via le bouton dédié dans `Introduction`. __Il est obligatoire de passer par ce bouton car ce dernier génère au début de la liste de PPN l'emplacement de Corwin, sans lequel WinIBW ne saura pas où écrire les données.__ (L'emplacement de Corwin sera remplacé à la fin de l'analyse de Corwin par `Ø`.) Une fois la liste dans le presse-papier, lancez le script WinIBW du traitement que vous souhaitez, soit en l'appelant directement, soit en passant par le lanceur de CorWin.

Pendant le traitement, laissez WinIBW travailler. Une fois le traitement terminé, WinIBW affichera un pop-up vous invitant à lancer l'analyse depuis Corwin.

## Lancement de l'analyse dans Corwin

Une fois le traitement dans WinIBW terminé, vous pouvez lancer l'analyse de Corwin via le bouton dédié dans la feuille `Introduction`.

## Les analyses

### CW1 : vérification du format de l'UA103

A
