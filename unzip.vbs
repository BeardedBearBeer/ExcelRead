'ZIPファイル名を取得
ZIP_SOURCE = WScript.Arguments(0)

If SearchDelimiter(ZIP_SOURCE) = False Then
    'カレントフォルダのパスとファイル名を結合
    ZIP_SOURCE = CreateObject("Scripting.FileSystemObject").getParentFolderName(WScript.ScriptFullName) & "\" & ZIP_SOURCE
End if

'オブジェクト生成
Set objShell = CreateObject("shell.application")
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set ZipFile = objShell.NameSpace(ZIP_SOURCE).items
Set objFolder = objShell.NameSpace (objWshShell.CurrentDirectory)

'進行状況ダイアログボックス非表示 + ダイアログボックスは[すべてはい]
objFolder.CopyHere ZipFile, &H14

'オブジェクトの破棄
Set objShell = Nothing
Set objWshShell = Nothing
Set ZipFile = Nothing
Set objFolder =Nothing


'=== 関数 ============================================================================================== 
'引数で指定した文字列にパスの区切り文字がないかチェックする関数
Function SearchDelimiter(strSerchObect)
    Dim obj
    Set obj = new RegExp    '正規表現を扱えるクラスRegExpのインスタンスを生成

    obj.Pattern = "\\"
    obj.Global = True       '引数で受けた文字列全体を検索するように設定

    SearchDelimiter = obj.Test(strSerchObect)
    Set obj = Nothing
End Function
