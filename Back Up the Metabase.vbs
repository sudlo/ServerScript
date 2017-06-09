Const MD_BACKUP_HIGHEST_VERSION = &HFFFFFFFF
Const MD_BACKUP_OVERWRITE = 1

strComputer = "LocalHost"
Set objComputer = GetObject("IIS://" & strComputer & "")
objComputer.Backup "ScriptedBackup", MD_BACKUP_HIGHEST_VERSION,_
     MD_BACKUP_OVERWRITE
