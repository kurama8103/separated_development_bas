# separated_development_bas
エクセルとVBAコードを分離し、管理しやすくするツール。

Excel VBAのコードはxlsmファイルの形式ではバイナリ形式で認識されるため直接Githubで管理できない。  
xlsmからコードを分離しbasファイルなどのモジュールにエクスポートすることで、Githubで管理できるようにする。  

モジュールのImport: NewBookWithModule.vbsを実行するとlibフォルダのすべてのbas, frm, clsファイルを読み込んだ新しいxlsmファイルを作成する。  
モジュールのExport: mdl_IO.basのVB_ExportModuleを実行すると指定のフォルダにモジュールをエクスポートする。

</br></br>

(補足)libフォルダの中身  
・mdl_init.bas  
上記Importを行うコード。NewBookWithModule.vbsがNewBookWithModule.vbsにおいて呼び出すモジュールであり、libフォルダを読み込んだ新しいブックが作られる。

・mdl_IO.bas  
ベースとなるプログラム。上記Exportに対応するモジュールのほか、VBComponentの操作、参照設定、よく使う関数等が記述されている。

・mdl_ODBC.bas  
ワークシート埋め込みのテーブルクエリ等に関するモジュール。パラメーターが多いのでデバッグ用が主。

・mdl_worksheet.bas  
ワークシートの図形や数値計算関数など、mdl_IO.basより重要度の低い関数モジュール。

・mdl_test.bas  
上記モジュールのテスト用。  


MEMO  
'基本的にはFunctionを使用、ワークシートから呼び出すコードはSub。  
'Subで呼び出す関数の引数は、利便性よりActiveBook/ActiveSheetとしている。  

'現状ADODB関連のみ事前バインディング  
'事前バインディングはコーディングが楽、実行速度も速い:New Scripting.FileSystemObject  
'実行時バインディングは参照設定が不要で配布が楽：CreateObject("Scripting.FileSystemObject")
