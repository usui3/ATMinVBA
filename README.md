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
ExcelVBA

#変数一覧・定義

  RegistrationForm（新規登録画面）
  
    DB_ws       Sheet(DB）を格納するWoeksheet型変数   
    Cust_ws     新規登録で挿入した新規シートを格納するWorksheet型変数
    Flag        Sheet(DB)に既存の名前とIDと組み合わせがないか判断するためのBoolean型変数
    i           名前とIDの組み合わせの存在を確認するFor文でセルの行数のカウントに用いるLong型変数
    Lastdata    Sheet(DB)の最終行+1の行を表すLong型変数
    
    


















