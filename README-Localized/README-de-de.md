# Aufgabenbereich-Add-In-Beispiel für Studiencreditsverfolgung für Excel 2016

_Gilt für: Excel 2016_

Dieses Aufgabenbereich-Add-In zeigt, wie mithilfe der JavaScript-APIs in Excel 2016 eine Studiencreditsverfolgung erstellt werden kann. Es ist in zwei Versionen verfügbar: Text-Editor und Visual Studio.

![Studiencredits-Verfolgungsbeispiel](../images/CollegeCreditsTracker_tracker.PNG)

## Probieren Sie es aus
### Text-Editor-Version

Am einfachsten können Sie Ihr Add-In bereitstellen und testen, indem Sie die Dateien in eine Netzwerkfreigabe kopieren.

1.  Erstellen Sie einen Ordner in einer Netzwerkfreigabe (z. B. \\\MyShare\CollegeCreditsTracker), und kopieren Sie alle Dateien im Ordner „Text-Editor”. 
2.  Bearbeiten Sie das <SourceLocation>-Element der Manifestdatei, damit es auf den Freigabepfad aus Schrittä 1 zeigt. 
3.  Kopieren Sie das Manifest (CollegeCreditsTrackerManifest.xml) in eine Netzwerkfreigabe (z. B. \\\MyShare\MyManifests).
4.  Fügen Sie den Freigabepfad, unter dem das Manifest enthalten ist, als vertrauenswürdigen App-Katalog in Excel hinzu.

    a. Starten Sie Excel, und öffnen Sie ein leeres Arbeitsblatt.  
    
    b. Klicken Sie auf die Registerkarte **Datei**, und klicken Sie dann auf **Optionen**.
    
    c. Wählen Sie **Trust Center** aus, und klicken Sie dann auf die Schaltfläche **Einstellungen für das Trust Center**.
    
  d. Wählen Sie **Vertrauenswürdige App-Kataloge** aus.
    
  e. Geben Sie im Feld **Katalog-URL** den Pfad zu der in Schritt 1 erstellten Netzwerkfreigabe ein, und klicken Sie auf **Katalog hinzufügen**.
    
   f. Aktivieren Sie das Kontrollkästchen **Im Menü anzeigen**, und wählen Sie dann **OK**. Eine Meldung wird angezeigt, dass Ihre Einstellungen angewendet werden, wenn Office das nächste Mal gestartet wird. 
        
5.  Testen und führen Sie das Add-In aus. 

  a. Klicken Sie auf der Registerkarte **Einfügen** in Excel 2016 auf **Meine-Add-Ins**. 
    
  b. Wählen Sie im Dialogfenster **Office-Add-Ins** die Option **Freigegebener Ordner** aus.
    
  c. Wählen Sie **Studiencredits-Verfolgungsbeispiel**>**Einfügen**. Das Add-In wird in einem Aufgabenbereich geöffnet, wie in diesem Diagramm dargestellt. 
        
   ![Studiencredits-Verfolgungsbeispiel](../images/CollegeCreditsTracker_taskpane.PNG) 

  d. Klicken Sie auf die Schaltfläche **Studiencreditsplaner erstellen**. Dabei wird die Studiencreditsverfolgung im aktiven Arbeitsblatt erstellt, wie in der folgenden Abbildung dargestellt. 
    
  ![Studiencredits-Verfolgungsbeispiel](../images/CollegeCreditsTracker_tracker.PNG)

  e. Fügen Sie einige Kurse mithilfe der Registerkarte **Kurs hinzufügen**, und sehen Sie, wie sich Daten und das Diagramm dynamisch ändern.
    
### Visual Studio-Version
1.  Kopieren Sie das Projekt in einen lokalen Ordner, und öffnen Sie die Datei „Excel-Add-in-JS-CollegeCreditsTracker.sln” in Visual Studio.
2.  Drücken Sie F5, um das Beispiel-Add-In zu erstellen und bereitzustellen. Excel wird gestartet und das Add-In wird in einem Aufgabenbereich rechts neben einem leeren Arbeitsblatt geöffnet, wie in der folgenden Abbildung dargestellt. 
        
  ![Studiencredits-Verfolgungsbeispiel](../images/CollegeCreditsTracker_taskpane.PNG) 

3.  Klicken Sie auf die Schaltfläche **Studiencreditsplaner erstellen**. Dabei wird die Studiencreditsverfolgung im aktiven Arbeitsblatt erstellt, wie in der folgenden Abbildung dargestellt. 
    
  ![Studiencredits-Verfolgungsbeispiel](../images/CollegeCreditsTracker_tracker.PNG) 
  
4. Fügen Sie einige Kurse mithilfe der Registerkarte **Kurs hinzufügen**, und sehen Sie, wie sich Daten und das Diagramm dynamisch ändern.


### Weitere Informationen

Die Excel-JavaScript-APIs haben viel mehr bei der Entwicklung von Add-Ins zu bieten. Im Folgenden werden nur einige der verfügbaren Ressourcen aufgeführt. 

1.  [Programmierungsübersicht für Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Codeausschnitt-Explorer für Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Codebeispiele zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [JavaScript-API-Referenz zu Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Erstellen Ihres ersten Excel-Add-Ins](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
