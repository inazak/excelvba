Option Explicit

''' 印刷設定
Sub Step_PrintSetting()

  Dim Sheet As Worksheet
  
  Set Sheet = Worksheets("Sheet1")
  Sheet.Select
    
  'ページ設定
  Sheet.PageSetup.Orientation = xlLandscape '横向き
  Sheet.PageSetup.Zoom = False '倍率設定なし
  Sheet.PageSetup.FitToPagesWide = 1 '横1枚
  Sheet.PageSetup.FitToPagesTall = 1 '縦1枚
  Sheet.PageSetup.LeftHeader = "&A" '左ヘッダにシート名
  Sheet.PageSetup.RightHeader = "&D" '右ヘッダに日付
  Sheet.PageSetup.TopMargin = Application.InchesToPoints(1.69291338582677)
  Sheet.PageSetup.HeaderMargin = Application.InchesToPoints(1.14173228346457)

End Sub

