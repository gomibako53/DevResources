Attribute VB_Name = "ColorPaletModule"
Option Explicit

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChooseColor As ChooseColor) As Long

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type



Private Const CC_RGBINT = &H1 ' 色のデフォルト値を設定
Private Const CC_LFULLOPEN = &H2 ' 色の作成を行う部分を表示
Private Const CC_PREVENTFULLOPEN = &H4 ' 色の作成ボタンを無効にする
Private Const CC_SHOWHELP = &H8 ' ヘルプボタンを表示

' 機能  ：  色選択ダイアログを表示し、選択された色のRGB値を返す
' 引数  ：  lngDefColor デフォルト表示する色
' 返値  ：  成功時 RGB値、キャンセル時 -1、エラー時 -2 (ゼロは黒なので注意)
Public Function GetColorDlg(lngDefColor As Long) As Long
    Dim udtChooseColor As ChooseColor
    Dim lngRet As Long

    With udtChooseColor
        ' ダイアログの設定
        .lStructSize = Len(udtChooseColor)
        .IpCustColors = String$(64, Chr$(0))
        '.flags CC_RGBINT + CC_LFULLOPEN
        .flags = 0
        .rgbResult = lngDefColor
        ' ダイアログの表示
        lngRet = ChooseColor(udtChooseColor)
        ' ダイアログからの返り値をチェック
        If lngRet <> 0 Then
            If .rgbResult > RGB(255, 255, 255) Then
                ' エラー
                GetColorDlg = -2
            Else
                ' 正常終了、RGB値を返り値にセット
                GetColorDlg = .rgbResult
            End If
        Else
            ' キャンセルが押されたとき
            GetColorDlg = -1
        End If
    End With
End Function

