# ATMinVBA
  ExcelVBAを用いて簡易的なATMの動きを再現しました。

#開発内容
   
   EXcelVBAのユーザーフォームを用いて簡易的なATMの動きを再現し
  登録済みユーザーのExcelSheetを用いて通帳の記帳を行う

#工数
  
  40時間

#開発スケジュール

  1月10日         設計
  
  1月11日-1月14日 プログラミング
  
  1月21日         〃
  
  1月24-1月25日　 プログラミング・テスト
  
  1月26日         仕様書作成

#開発環境

ExcelVBA/Windows10

#変数一覧・定義

  RegistrationForm(新規登録画面)
  
    DB_ws       Sheet(DB）を格納するWoeksheet型変数   
    Cust_ws     新規登録で挿入した新規シートを格納するWorksheet型変数
    Flag        Sheet(DB)に既存の名前とIDと組み合わせがないか判断するためのBoolean型変数
    i           名前とIDの組み合わせの存在を確認するFor文でセルの行数のカウントに用いるLong型変数
    Lastdata    Sheet(DB)の最終行+1の行を表すLong型変数
    Enter_Id    RegistrationFormのId_txtオブジェクトのTextの文字列型数字を読み取るStrig型変数
    
  LoginForm(ログイン画面)
    
    Enter_Id    LoginFormのLogId_txtオブジェクトのTextの文字列型数字を読み取るStrig型変数
    i           名前とIDの組み合わせの存在を確認するFor文でセルの行数のカウントに用いるLong型変数
    DB_ws       Sheet(DB）を格納するWoeksheet型変数
    
  EnterMoneyForm(金額入力画面)
   
     Chg_txt    Changeイベントで用いる、EnterMoneyForm.E_AmountMoneyを格納するオブジェクト型変数
     Zero_txt   Key0プロシージャで用いる、EnterMoneyForm.E_AmountMoneyを格納するオブジェクト型変数
     E_txt      KeyClickプロシージャで用いる、EnterMoneyForm.E_AmountMoneyを格納するオブジェクト型変数
     ※ChangeイベントはEnterMoneyForm.E_AmountMoneyに変更を加えた際に実行され他のプロシージャと同時に実       行される為同じ変数名を宣言するとエラーが起こります。なので、中身は同じですが変数名を変えてエラーを       回避しています。
     
  PartnerForm(振込先情報入力画面)
    
      DB_ws     Sheet(DB）を格納するWoeksheet型変数
      i         名前とIDの組み合わせの存在を確認するFor文でセルの行数のカウントに用いるLong型変数
      Enter_Id   PartnerFormのPId_txtオブジェクトのTextの文字列型数字を読み取るStrig型変数

  ConfirmationForm(確認画面)
    
      DB_ws     Sheet(DB)を格納するWoeksheet型変数
      Name      LoginForm.LogName_txtを格納するString型変数
      ID        Loginorm.LogId_txt.Textを格納するStirin型変数
      Cust_ws_name  Sheet(DB)のログインユーザー行の５列目に入力してあるシート名を格納するString型変数
      Cust_ws   Cust_ws_nameに格納してあるシート名のSheetを格納するWorksheet型変数
      CustRow   NameとIDが存在するSheet(DB)の行を格納するLong型変数
      AppRow    Cust_wsの最終行+１の行を格納するLong型変数
      P_ID      PartnerForm.PId_txt.Textを格納するString型変数
      P_Name    PartnerForm.PName_txtを格納するString型変数
      P_Row     P_NameとP_IDが存在するSheet(DB)の行を格納するLong型変数
      P_ws_name Sheet(DB)のP_Row行５列目に入力してあるシート名を格納するString型変数
      P_ws      P_ws_nameに格納してあるシート名のSheetを格納するWorksheet型変数
      P_AppRow  P_wsの最終行の+1の行を格納するLong型変数
      












