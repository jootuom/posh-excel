# Excel

Excel module for PowerShell for easy document conversion.

Get-Help works for every cmdlet.

## ConvertFrom-Excel
Convert Excel documents into other formats

(.xlsx, .xls, .html, .txt, .csv, .pdf, .xps)
## ConvertTo-Excel
Convert text files into Excel documents
## Format-Excel
Displays a PowerShell object in Excel

Essentially this just runs Export-CSV on the piped object and then opens that in Excel.

It's a really cool way to view data from PowerShell in Excel.

Example: `Get-ADUser -Properties mail,mobilePhone | Format-Excel`
## Export-Excel
This allows you to export PowerShell data directly into the Excel format

This just runs Export-CSV then ConvertTo-Excel to convert it into Excel format.

