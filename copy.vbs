Const FOF_CREATEPROGRESSDLG = &H0&

dim QuellDatei, ZielOrdner

QuellDatei="h:\ki_apd\gvf.fic"
ZielOrdner= "h:\"

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.NameSpace(ZielOrdner) 

objFolder.CopyHere Quelldatei, FOF_CREATEPROGRESSDLG