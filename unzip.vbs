Option Explicit
dim objShell, objWshShell, objFolder, ZipFile, i
if WScript.Arguments.Count < 1 or WScript.Arguments.Count > 2 then
    WScript.Echo "Usage: CScript.exe UnZip.VBS ZIPFile [objFolder]"
    WScript.Quit
end if
Set objShell = CreateObject("shell.application")
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set ZipFile = objShell.NameSpace (WScript.Arguments(0)).items
if WScript.Arguments.Count = 2 then
    Set objFolder = objShell.NameSpace (WScript.Arguments(1))          '�w�肳�ꂽ�𓀐�t�H���_
else
    Set objFolder = objShell.NameSpace (objWshShell.CurrentDirectory)  '�ȗ����̓J�����g�f�B���N�g����
end if
objFolder.CopyHere ZipFile, &H14            '�i�s�󋵃_�C�A���O�{�b�N�X��\�� + �_�C�A���O�{�b�N�X��[���ׂĂ͂�]
