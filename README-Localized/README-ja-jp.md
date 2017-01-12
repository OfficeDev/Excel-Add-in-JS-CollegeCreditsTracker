# <a name="college-credits-tracker-task-pane-add-in-sample-for-excel-2016"></a>Excel 2016 用の大学単位追跡ツール作業ウィンドウ アドインのサンプル

_適用対象:Excel 2016_

この作業ウィンドウ アドインには、Excel 2016 で JavaScript API を使用して、大学単位追跡ツールを作成する方法が示されます。テキスト エディターと Visual Studio のいずれかを選択できます。

![大学単位追跡ツールのサンプル](../Images/CollegeCreditsTracker_tracker.PNG)

## <a name="try-it-out"></a>お試しください。
### <a name="text-editor-version"></a>テキスト エディターのバージョン

アドインを展開してテストする最も簡単な方法は、ファイルをネットワーク共有にコピーすることです。

1.  テキスト エディター バージョンでファイルを任意のサーバーに配置します。
2.  マニフェスト ファイルの \<SourceLocation\> 要素と \<URL\> 要素を編集して、手順 1 のホストの場所 (たとえば https://localhost) をポイントするようにします。
2.  ネットワーク共有 (たとえば \\\MyShare\CollegeCreditsTracker) にフォルダーを作成し、テキスト エディター フォルダー内のすべてのファイルをコピーします。
3.  マニフェスト (CollegeCreditsTrackerManifest.xml) をネットワーク共有 (たとえば \\\MyShare\MyManifests) にコピーします。
4.  マニフェストを格納する共有の場所を、Excel で信頼されるアプリ カタログとして追加します。

    a.Excel を起動し、空のスプレッドシートを開きます。

    b.**[ファイル]** タブを選び、**[オプション]** を選びます。

    c.**[セキュリティ センター]** を選択し、**[セキュリティ センターの設定]** ボタンを選択します。

    d.**[信頼されているアプリ カタログ]** を選択します。

    e.**[カタログの URL]** ボックスに、手順 1 で作成したネットワーク共有のパスを入力し、**[カタログの追加]** を選択します。

   f. **[メニューに表示する]** チェック ボックスをオンにしてから **[OK]** を選択します。これらの設定は Office を次回起動したときに適用されることを示すメッセージが表示されます。

5.  アドインをテストし、実行します。

    a.Excel 2016 の **[挿入]** タブで、**[個人用アドイン]** を選択します。

    b.**[Office アドイン]** ダイアログ ボックスで、**[共有フォルダー]** を選択します。

    c.**[大学の単位の追跡ツールのサンプル]**>**[挿入]** の順に選択します。次の図に示すように、現在のワークシートの右側の作業ウィンドウでアドインが開きます。

   ![大学単位追跡ツールのサンプル](../Images/CollegeCreditsTracker_taskpane.PNG)

    d.**[大学の単位のプランナーの作成]** ボタンをクリックします。この操作により、次の図に示されているように、作業中のシートに大学の単位の追跡ツールが作成されます。

  ![大学単位追跡ツールのサンプル](../Images/CollegeCreditsTracker_tracker.PNG)

    e.**[コースの追加]** タブを使用してコースを追加し、データとグラフがどのように動的に変化するかを確認します。

### <a name="visual-studio-version"></a>Visual Studio のバージョン
1.  プロジェクトをローカル フォルダーにコピーし、Visual Studio で Excel-Add-in-JS-CollegeCreditsTracker.sln を開きます。
2.  F5 キーを押して、サンプル アドインをビルドおよび展開します。Excel が起動し、次の図に示すように、空白のワークシートの右側の作業ウィンドウでアドインが開きます。

  ![大学単位追跡ツールのサンプル](../Images/CollegeCreditsTracker_taskpane.PNG)

3.  **[大学単位プランナーの作成]** ボタンをクリックします。この操作により、次の図に示されているように、作業中のシートに大学単位追跡ツールが作成されます。

  ![大学単位追跡ツールのサンプル](../Images/CollegeCreditsTracker_tracker.PNG)

4. **[コースの追加**] タブを使用してコースを追加し、データとグラフがどのように動的に変化するかを確認します。


### <a name="learn-more"></a>詳細を見る

アドインを開発する際に Excel JavaScript API でできることは他にも数多くあります。以下は、利用可能なリソースのほんの一例です。

1.  [Excel アドインのプログラミングの概要](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Excel のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel アドインのコード サンプル](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md)
4.  [Excel アドインの JavaScript API リファレンス](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [最初の Excel アドインをビルドする](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)