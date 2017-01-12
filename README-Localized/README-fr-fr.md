# <a name="college-credits-tracker-task-pane-add-in-sample-for-excel-2016"></a>Exemple de complément de volet Office - Suivi des crédits universitaires pour Excel 2016

_S’applique à : Excel 2016_

Ce complément de volet Office montre comment créer un outil de suivi des crédits universitaires à l’aide des API JavaScript dans Excel 2016. Il a deux versions : éditeur de texte et Visual Studio.

![Exemple d’outil de suivi des crédits universitaires](../Images/CollegeCreditsTracker_tracker.PNG)

## <a name="try-it-out"></a>Essayez !
### <a name="text-editor-version"></a>Version d’éditeur de texte

Pour déployer et tester votre complément, le plus simple consiste à copier les fichiers sur un partage réseau.

1.  Déployez les fichiers dans la version de l’éditeur de texte sur un serveur de votre choix.
2.  Modifiez les éléments \<SourceLocation\> et \<Url\> du fichier manifeste afin qu’il pointe vers l’emplacement hébergé de l’étape 1 (à savoir, https://localhost).
2.  Créez un dossier sur un partage réseau (par exemple, \\\MyShare\CollegeCreditsTracker) et copiez tous les fichiers dans le dossier d’éditeur de texte.
3.  Copiez le fichier manifeste (CollegeCreditsTrackerManifest.xml) dans le partage réseau (par exemple, \\\MyShare\MyManifests).
4.  Ajoutez l’emplacement de partage qui contient le fichier manifeste sous forme de catalogue d’applications approuvées dans Excel.

    a.  Lancez Excel et ouvrez une feuille de calcul vide.

    b.  Choisissez l’onglet **Fichier**, puis choisissez **Options**.

    c.  Choisissez l’onglet **Centre de gestion de la confidentialité**, puis choisissez **Paramètres du Centre de gestion de la confidentialité**.

    d.  Choisissez **Catalogues d’applications approuvées**.

    e.  Dans la zone **URL du catalogue**, entrez le chemin d’accès du partage réseau que vous avez créé à l’étape 1, puis choisissez **Ajouter un catalogue**.

   Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**. Un message s’affiche pour vous informer que vos paramètres seront appliqués la prochaine fois que vous démarrerez Office.

5.  Testez et exécutez le complément.

    a.  Dans l’onglet **Insertion** d’Excel 2016, choisissez **Mes compléments**.

    b.  Dans la boîte de dialogue **Compléments Office**, choisissez **Dossier partagé**.

    c.  Choisissez **Exemple de suivi des crédits universitaires**>**Insertion**. Le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante.

   ![Exemple d’outil de suivi des crédits universitaires](../Images/CollegeCreditsTracker_taskpane.PNG)

    d.  Cliquez sur le bouton **Créer un planificateur de crédits universitaires**. Cela permet de créer le suivi des crédits universitaires dans la feuille de calcul active, comme indiqué dans ce diagramme.

  ![Exemple d’outil de suivi des crédits universitaires](../Images/CollegeCreditsTracker_tracker.PNG)

    e.  Ajoutez des cours à l’aide de l’onglet **Ajouter un cours** et observez la façon dont les données et le graphique changent de manière dynamique.

### <a name="visual-studio-version"></a>Version de Visual Studio
1.  Copiez le projet dans un dossier local et ouvrez le fichier Excel-Add-in-JS-CollegeCreditsTracker.sln dans Visual Studio.
2.  Appuyez sur F5 pour créer et déployer l’exemple de complément. Excel démarre et le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante.

  ![Exemple d’outil de suivi des crédits universitaires](../Images/CollegeCreditsTracker_taskpane.PNG)

3.  Cliquez sur le bouton **Créer un planificateur de crédits universitaires**. Cela permet de créer le suivi des crédits universitaires dans la feuille de calcul active, comme indiqué dans ce diagramme.

  ![Exemple d’outil de suivi des crédits universitaires](../Images/CollegeCreditsTracker_tracker.PNG)

4. Ajoutez des cours à l’aide de l’onglet **Ajouter un cours** et observez la façon dont les données et le graphique changent de manière dynamique.


### <a name="learn-more"></a>En savoir plus

Les API JavaScript pour Excel peuvent vous offrir beaucoup pour l’élaboration de vos compléments. Voici quelques-unes des ressources disponibles :

1.  [Présentation de la programmation JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemples de code pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Référence de l’API JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Création de votre premier complément Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)