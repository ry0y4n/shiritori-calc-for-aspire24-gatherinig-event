# 2次元タプルの各要素（Cellオブジェクト）から値を取得して2次元配列を返す
def get_value_list(t_2d):
  return([[cell.value for cell in row] for row in t_2d])

# 始まりの行番号・列番号で範囲を指定して2次元配列（リストのリスト）として取得する
def get_list_2d(sheet, start_row, end_row, start_col, end_col):
  return get_value_list(sheet.iter_rows(min_row=start_row,
                                        max_row=end_row,
                                        min_col=start_col,
                                        max_col=end_col))