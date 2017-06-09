Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objApplication = CreateObject("HNetCfg.FwAuthorizedApplication")
objApplication.Name = "Free Cell"
objApplication.IPVersion = 2
objApplication.ProcessImageFileName = "c:\windows\system32\freecell.exe"
objApplication.RemoteAddresses = "*"
objApplication.Scope = 0
objApplication.Enabled = True

Set colApplications = objPolicy.AuthorizedApplications
colApplications.Add(objApplication)
