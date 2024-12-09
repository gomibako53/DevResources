Attribute VB_Name = "ColorPaletModule"
Option Explicit

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChooseColor As ChooseColor) As Long

Private Const CC_RGBINIT = &H1 ' 色のデフォルト値を設定
Private Const CC_FULLOPEN = &H2 ' 色の作成を行う部分を表示(右側の部分)
Private Const CC_PREVENTFULLOPEN = &H4 ' 色の作成ボタンを無効にする
Private Const CC_SHOWHELP = &H8 ' ヘルプボタンを表示
Private Const CC_ANYCOLOR = &H100 ' 利用可能な基本色をすべて表示

' 機能  ：  色の設定ダイアログを表示し、そこで選択された色のRGBの値を返す
' 引数  ：  lngDefColor デフォルト表示する色
' 返値  ：  成功時 RGB値、キャンセル時 -1、エラー時 -2 (ゼロは黒なので注意)
Public Function GetColorDlg(lngDefColor As Long) As Long
    Dim udtChooseColor As ChooseColor
    Dim lngRet As LongPtr
    Static CustomColors(16) As Long
    
    ' Some predefined color, there are 16 slots available for predefined colors
    CustomColors(0) = RGB(255, 255, 255)    ' White
    CustomColors(1) = RGB(0, 0, 0)  ' Black
    CustomColors(2) = RGB(255, 0, 0)    ' Red
    'CustomColors(3) = RGB(0, 255, 0)   ' Green
    CustomColors(3) = RGB(0, 176, 80)   ' Green(default)
    CustomColors(4) = RGB(0, 0, 255)  ' Blue
    CustomColors(8) = RGB(255, 255, 204)    ' Light Yellow
    CustomColors(9) = RGB(255, 204, 255)  ' Light Pink
    CustomColors(10) = RGB(204, 255, 255)  ' Light Blue
    CustomColors(11) = RGB(204, 255, 204)  ' Light Green
    CustomColors(12) = RGB(191, 191, 191) ' Light Gray

    With udtChooseColor
        .lStructSize = LenB(udtChooseColor)
        .flags = CC_RGBINIT Or CC_ANYCOLOR
        If IsNull(lngDefColor) = False And IsMissing(lngDefColor) = False Then
            .rgbResult = lngDefColor  'Set the initial color of the dialog
        End If
        .lpCustColors = VarPtr(CustomColors(0))
        
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

