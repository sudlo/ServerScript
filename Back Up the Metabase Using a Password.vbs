Const MD_BACKUP_HIGHEST_VERSION = &HFFFFFFFE 
Const MD_BACKUP_OVERWRITE = 1

strComputer = "LocalHost"
Set objComputer = GetObject("IIS://" & strComputer & "")
objComputer.BackupWithPassword "ScriptedBackup", _
    MD_BACKUP_HIGHEST_VERSION, MD_BACKUP_OVERWRITE, "ie456@k"
