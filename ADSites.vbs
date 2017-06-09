On Error Resume Next

	Set objSiteFSO = CreateObject("Scripting.FileSystemObject")
	set ObjSiteFile = objSiteFSO.createTextFile("c:\ADscripts\ADLogs\ADSites.txt", true)

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.Filter = Array("site")
 
For Each objSite In objSitesContainer
     objSiteFile.WriteLine objSite.CN
    strSiteName = objSite.Name
    strServerPath = "LDAP://cn=Servers," & strSiteName & ",cn=Sites," & _
        strConfigurationNC
    Set colServers = GetObject(strServerPath)
 
    For Each objServer In colServers
	objSiteFile.WriteLine vbTab & objServer.CN
    Next
   objSiteFile.WriteLine vbCRLF
Next

WScript.Echo "Script... Done....."

objSiteFile.Close

