# <a name="college-credits-tracker-task-pane-add-in-sample-for-excel-2016"></a>Exemplo do Suplemento do Painel de Tarefas de Controlador de Créditos Acadêmicos para o Excel 2016

_Aplica-se a: Excel 2016_

Este suplemento do painel de tarefas mostra como criar um controlador de créditos acadêmicos usando as APIs JavaScript no Excel 2016. Há dois tipos: o editor de texto e o Visual Studio.

![Exemplo de Controlador de créditos acadêmicos](../images/CollegeCreditsTracker_tracker.PNG)

## <a name="try-it-out"></a>Experimente
### <a name="text-editor-version"></a>Versão do editor de texto

A maneira mais fácil de implantar e testar o suplemento é copiar os arquivos para um compartilhamento de rede.

1.  Implante os arquivos na versão do Editor de Texto para um servidor de sua escolha.
2.  Edite os elementos \<SourceLocation\> e \<URL\> do arquivo de manifesto para que ele aponte para o local hospedado da etapa 1. (isto é, https://localhost)
2.  Crie uma pasta em um compartilhamento de rede (por exemplo, \\\MyShare\CollegeCreditsTracker) e copie todos os arquivos na pasta do Editor de Texto.
3.  Copie o manifesto (CollegeCreditsTrackerManifest.xml) em um compartilhamento de rede (por exemplo, \\\MyShare\MyManifests).
4.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a.  Inicie o Excel e abra uma planilha em branco.

    b.  Escolha a guia **Arquivo** e escolha **Opções**.

    c.  Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

    d.  Escolha **Catálogos de Aplicativos Confiáveis**.

    e.  Na caixa **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 1 e escolha **Adicionar Catálogo**.

   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office.

5.  Teste e execute o suplemento.

    a.  Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**.

    b.  Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.

    c.  Escolha **Exemplo de Controlador de Créditos Acadêmicos**>**Inserir**. O suplemento abre em um painel de tarefas à direita da planilha atual, conforme mostrado na figura a seguir.

   ![Exemplo de Controlador de créditos acadêmicos](../images/CollegeCreditsTracker_taskpane.PNG)

    d.  Clique no botão **Criar Planejador de Créditos Acadêmicos**. Essa opção cria o controlador de créditos acadêmicos na planilha ativa, conforme mostrado neste diagrama.

  ![Exemplo de Controlador de créditos acadêmicos](../images/CollegeCreditsTracker_tracker.PNG)

    e.  Adicione alguns cursos usando a guia **Adicionar um curso** e veja como os dados e o gráfico mudam dinamicamente.

### <a name="visual-studio-version"></a>Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o Excel-Add-in-JS-CollegeCreditsTracker.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na figura a seguir.

  ![Exemplo de Controlador de créditos acadêmicos](../images/CollegeCreditsTracker_taskpane.PNG)

3.  Clique no botão **Criar Planejador de Créditos Acadêmicos**. Essa opção cria o controlador de créditos acadêmicos na planilha ativa, conforme mostrado neste diagrama.

  ![Exemplo de Controlador de créditos acadêmicos](../images/CollegeCreditsTracker_tracker.PNG)

4. Adicione alguns cursos usando a guia **Adicionar um curso** e veja como os dados e o gráfico mudam dinamicamente.


### <a name="learn-more"></a>Saiba mais

As APIs JavaScript para Excel têm muito mais a oferecer à medida que você desenvolve suplementos. Confira a seguir alguns dos recursos disponíveis.

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de Trechos para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Referência da API JavaScript de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Criar seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)