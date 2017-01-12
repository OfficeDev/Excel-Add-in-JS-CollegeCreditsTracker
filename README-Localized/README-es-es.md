# <a name="college-credits-tracker-task-pane-add-in-sample-for-excel-2016"></a>Ejemplo del complemento del panel de tareas del rastreador de créditos universitarios para Excel 2016

_Se aplica a: Excel 2016_

Este complemento del panel de tareas muestra cómo crear un rastreador de créditos universitarios mediante las API de JavaScript de Excel 2016. Hay dos tipos: editor de texto y Visual Studio.

![Ejemplo de rastreador de créditos universitarios](../Images/CollegeCreditsTracker_tracker.PNG)

## <a name="try-it-out"></a>Pruébelo
### <a name="text-editor-version"></a>Versión del editor de texto

La forma más sencilla de implementar y probar el complemento consiste en copiar los archivos en un recurso compartido de red.

1.  Implemente los archivos de la versión del Editor de texto en el servidor que prefiera.
2.  Edite los elementos \<SourceLocation\> y \<Urls\> del archivo de manifiesto para que apunte a la ubicación hospedada del paso 1 (p. ej., https://hostlocal).
2.  Cree una carpeta en un recurso compartido de red (por ejemplo, \\\MiRecursoCompartido\RastreadorCréditosUniversitarios) y copie todos los archivos en la carpeta del Editor de texto.
3.  Copie el manifiesto (CollegeCreditsTrackerManifest.xml) en el recurso compartido de red (por ejemplo, \\\MiRecursoCompartido\\MisManifiestos).
4.  Agregue la ubicación del recurso compartido que contiene el manifiesto como un catálogo de aplicaciones de confianza en Excel.

    a.  Inicie Excel y abra una hoja de cálculo en blanco.

    b.  Elija la pestaña **Archivo** y, a continuación, **Opciones**.

    c.  Elija **Centro de confianza** y, a continuación, el botón **Configuración del Centro de confianza**.

    d.  Elija **Catálogos de aplicaciones de confianza**.

    e.  En el cuadro **URL del catálogo**, escriba la ruta de acceso al recurso compartido de red que creó en el paso 1 y, a continuación, elija **Agregar catálogo**.

   f. Active la casilla **Mostrar en el menú** y elija **Aceptar**. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Office.

5.  Pruebe y ejecute el complemento.

    a.  En la pestaña **Insertar** de Excel 2016, elija **Mis complementos**.

    b.  En el cuadro de diálogo **Complementos de Office**, elija **Carpeta compartida**.

    c.  Elija **Ejemplo de rastreador de créditos universitarios**>**Insertar**. El complemento se abrirá en un panel de tareas a la derecha de la hoja de cálculo actual, como se muestra en la ilustración siguiente.

   ![Ejemplo de rastreador de créditos universitarios](../Images/CollegeCreditsTracker_taskpane.PNG)

    d.  Haga clic en el botón **Crear planificador de créditos universitarios**. Esto creará el rastreador de créditos universitarios en la hoja activa, tal y como se muestra en este diagrama.

  ![Ejemplo de rastreador de créditos universitarios](../Images/CollegeCreditsTracker_tracker.PNG)

    e.  Agregue algunos cursos mediante la pestaña **Agregar un curso** y vea cómo cambian de forma dinámica los datos y el gráfico.

### <a name="visual-studio-version"></a>Versión de Visual Studio
1.  Copie el proyecto en una carpeta local y abra Excel-Add-in-JS-CollegeCreditsTracker.sln en Visual Studio.
2.  Pulse F5 para crear e implementar el complemento de ejemplo. Excel se inicia y se abre el complemento en un panel de tareas a la derecha de una hoja de cálculo en blanco, como se muestra en la siguiente ilustración.

  ![Ejemplo de rastreador de créditos universitarios](../Images/CollegeCreditsTracker_taskpane.PNG)

3.  Haga clic en el botón **Create College Credits Planner** (Crear planificador de créditos universitarios). Esto crea el rastreador de créditos universitarios en la hoja activa tal y como se muestra en este diagrama.

  ![Ejemplo de rastreador de créditos universitarios](../Images/CollegeCreditsTracker_tracker.PNG)

4. Agregue algunos cursos mediante la pestaña **Add a course** (Agregar un curso) y vea cómo cambian dinámicamente los datos y el gráfico.


### <a name="learn-more"></a>Obtener más información

Las API de JavaScript de Excel tienen mucho que ofrecer para el desarrollo de complementos. A continuación se muestran algunos de los recursos disponibles.

1.  [Introducción a la programación de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de fragmentos de código para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Ejemplos de código de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referencia de la API de JavaScript de complementos de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Compilar el primer complemento de Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)