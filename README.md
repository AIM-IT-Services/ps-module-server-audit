Generate an HTML file/folder audit report with TreeView, Dark Mode toggle,
and click-to-open file details modal. Compatible with Windows PowerShell 5.1
and servers without System.Web. 

```powershell
. 'C:\ProgramData\NinjaRMMAgent\CusotmPS'; New-FileDirectoryAuditReport -DrivePath `<insert file directory being audited>` -LastYears <insect years> -ClientName <insect client name> -TreeView -Verbose -OpenInBrowser <optional>
```
