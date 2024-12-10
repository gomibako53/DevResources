Attribute VB_Name = "ExportModule"
Option Explicit

' --------------------------------------------------------------------------
' 以下の参照設定が必要です。
' 設定は、[ツール]→[参照設定]で。
' "Microsoft Visual Basic for Applications Extensibility *.*"
' --------------------------------------------------------------------------
' 以下処理で VBA モジュールのエクスポート処理がありますが、エクセルの設定を変更しないとエラーになります。
' もしエラーになったら以下設定を見直してください。
'   1. Excelを開き、[ファイル] タブをクリックします。
'   2. [オプション] をクリックします。
'   3. [トラストセンター] をクリックし、[トラストセンターの設定] をクリックします。
'   4. [マクロの設定] をクリックし、[VBAプロジェクトオブジェクトモデルへのアクセスを信頼する]のチェックボックスをオンにします。
'   5. [OK] をクリックして、ダイアログボックスを閉じます。
' --------------------------------------------------------------------------

Sub VBAモジュールを一括Export()
    Dim module      As VBComponent      ' モジュール
    Dim moduleList  As VBComponents     ' VBAプロジェクトの全モジュール
    Dim extension                       ' モジュールの拡張子
    Dim sPath                           ' 処理対象ブックのパス
    Dim sFilePath                       ' エクスポートファイルパス
    Dim TargetBook                      ' 処理対象ブックオブジェクト

    Set TargetBook = ActiveWorkbook ' 表示しているブックを対象とする

    sPath = TargetBook.Path

    ' 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents

    ' VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        ' クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        ' フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            ' .frxも一緒にエクスポートされる
            extension = "frm"
        ' 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        ' その他
        Else
            ' エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If

        ' 新しいフォルダのパスを作成 (ブック名から拡張子を取り除いた名称のフォルダ)
        Dim newFolderPath As String: newFolderPath = sPath & "\" & "VBA_Export_" & Split(TargetBook.Name, ".")(0)
        ' フォルダが存在しない場合は作成
        If Dir(newFolderPath, vbDirectory) = "" Then
            MkDir newFolderPath
        End If

        ' エクスポート実施
        sFilePath = newFolderPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)

        ' 出力先確認用ログ出力
        Debug.Print sFilePath

CONTINUE:
    Next
End Sub

