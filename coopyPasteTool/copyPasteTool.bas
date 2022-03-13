Const SETTING_SHEEET = "Settings"  '設定用のシート
Const COPY_SHEEET = "copy"  '別ファイルへコピーするための値がセットされたシート(ここに入力された値がコピーされます。)
Const CELL_PATH = "B2"  'コピー先のフォルダパス
Const CELL_EXTENSION = "B3" 'ファイル名(【*.xlsx】とすることでフォルダ内のエクセルファイルが対象になります。)
Const CELL_S_AREA = "B4"  'copyシートの検索範囲
Const CELL_S_CELL = "B5"  '別ファイルの検索値が入力されているセル
Const CELL_PASTE = "B6"   '別ファイルのコピー先のセル

Sub bookOpen()

  Dim settingSheeet As Worksheet
  Dim copySheeet As Worksheet

  Dim strPath As String
  Dim strExtension As String
  Dim strFile As String
  Dim strSArea As String
  Dim strCopy As String
  Dim strSCell As String
  Dim strSearch As String
  Dim strPaste As String

  'ワークシートオブジェクトの設定
  With ThisWorkbook
    Set settingSheeet = .Worksheets(SETTING_SHEEET)
    Set copySheeet = .Worksheets(COPY_SHEEET)
  End With

  'Settingsシートから設定値を取得
  With settingSheeet
    strPath = .Range(CELL_PATH).Value
    strExtension = .Range(CELL_EXTENSION).Value
    strSArea = .Range(CELL_S_AREA).Value
    strSCell = .Range(CELL_S_CELL).Value
    strPaste = .Range(CELL_PASTE).Value
  End With

  strFile = Dir(strPath & strExtension)

  'ファルダ内の別シートへコピペ
  Do While strFile <> ""
    'コピー先のファイルを開く
    Workbooks.Open strPath & strFile
    '別ファイルの検索値を取得
    strSearch = Workbooks(strFile).Worksheets(1).Range(strSCell).Value
    
    'copyシートから検索値に合致する値の隣のセルをコピーする値として取得
    Set Rng = copySheeet.Range(strSArea).Find(what:=strSearch, LookAt:=xlWhole)
    strCopy = Rng.Next.Value

    '別シートのコピー先セルに値を貼り付け
    Workbooks(strFile).Worksheets(1).Range(strPaste).Value = strCopy
    
    'ファイルを保存して閉じる
    Workbooks(strFile).Save
    Workbooks(strFile).Close
    strFile = Dir
  Loop

End Sub
