Set colDiskQuotas = CreateObject("Microsoft.DiskQuota.1")

colDiskQuotas.Initialize "C:\", True
Set objUser = colDiskQuotas.AddUser("kenmyer")
Set objUser = colDiskQuotas.FindUser("kenmyer")
objUser.QuotaLimit = 50000000
