Function ConvertFrom-Excel {
	<#
	.SYNOPSIS
	Convert a file from the Excel format to something else
	.DESCRIPTION
	This cmdlet converts an Excel file (.xlsx,.xls) to another format:
	
	Spreadsheets: .xlsx, .xls, .ods
	Text: .html, .txt, .csv
	Other: .pdf, .xps
	.EXAMPLE
	ConvertFrom-Excel ".\ssheet.xlsx"
	Creates a CSV file with the same name.
	.EXAMPLE
	ConvertFrom-Excel ".\ssheet.xlsx" -Publish
	Creates a PDF file with the same name.
	.EXAMPLE
	ConvertFrom-Excel ".\ssheet.xlsx" -DestinationDirectory "C:\temp" -SheetNames "Sheet1"
	Creates a PDF file into C:\temp from the sheet named "Sheet1"
	.PARAMETER ExcelFile
	The file to convert from.
	.PARAMETER DestinationDirectory
	The destination directory.
	
	If the format doesn't support multiple sheets,
	the file is split into multiple files
	with the following naming convention: ..\File_SheetName.extension
	.PARAMETER SheetNames
	Only convert the selected sheets instead of everything.
	Only works for formats that don't support multiple sheets.
	
	default: convert all
	.PARAMETER Convert
	Convert to one of the formats listed in Format.
	This is the default option.
	.PARAMETER Publish
	"Publish" to one of the formats listed in PubFormat.
	.PARAMETER Format
	The format to convert to:
	
	Spreadsheets: .xlsx, .xls, .ods
	Text: .html, .txt, .csv
	
	default: .csv
	.PARAMETER PubFormat
	The format to publish to:
	
	.pdf, .xps
	
	default: .pdf
	.PARAMETER PubQuality
	Publication quality: Normal / Web.
	
	default: Normal
	.PARAMETER KeepProperties
	Keep document properties?
	
	default: no
	.PARAMETER IgnorePrintAreas
	Ignore print areas defined in document.
	
	default: yes
	.PARAMETER PrintRange
	Publish only the defined range.
	NOT USED.
	.PARAMETER OpenAfterPublish
	Open file after publishing.
	NOT USED.
	#>
	[CmdletBinding(
		DefaultParameterSetName="Convert",
		SupportsShouldProcess=$true,
		ConfirmImpact="Low"
	)]
	Param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
		$ExcelFile,
		
		[Parameter(Position=1)]
		[string] $DestinationDirectory = "",
		
		[Parameter()]
		[string] $SheetNames = "",
		
		[Parameter(ParameterSetName="Convert")]
		[switch] $Convert,
		
		[Parameter(ParameterSetName="Publish")]
		[switch] $Publish,
		
		[Parameter(ParameterSetName="Convert")]
		[ValidateSet(".xlsx",".xls",".ods",".html",".txt",".csv")]
		[string] $Format = ".csv",
		
		[Parameter(ParameterSetName="Publish")]
		[ValidateSet(".pdf",".xps")]
		[string] $PubFormat = ".pdf",
		
		[Parameter(ParameterSetName="Publish")]
		[ValidateSet("Normal","Web")]
		[string] $PubQuality = "Normal",
		
		[Parameter(ParameterSetName="Publish")]
		[bool] $KeepProperties = $false,
		
		[Parameter(ParameterSetName="Publish")]
		[bool] $IgnorePrintAreas = $true
		
		## NOT USED
		#[Parameter(ParameterSetName="Publish")]
		#[ValidatePattern("[0-9]+[-][0-9]+")]
		#[string] $PrintRange = "",
		#
		## NOT USED
		#[Parameter(ParameterSetName="Publish")]
		#[switch] $OpenAfterPublish
	)
	begin {
		$formats = @{
			".xlsx" = 51;
			".xls"  = 56;
			".ods"  = 60;
			".html" = 44;
			".txt"  = 42;
			".csv"  = 6;
		}
		
		$pubformats = @{
			".pdf" = 0;
			".xps" = 1;
		}
		
		$pubqualities = @{
			"Normal" = 0;
			"Web"    = 1;
		}
		
		$sheets = $SheetNames.split(",")
		
		# NOT USED
		#if ($PrintRange) {
		#	$printmin, $printmax = $PrintRange.split("-")
		#	
		#	if (!($printmin -le $printmax) -or $printmin -lt 1 -or $printmax -lt 1) {
		#		Throw "Invalid PrintRange!"
		#	}
		#}
		
		switch ($PSCmdlet.ParameterSetName) {
			"Convert" { $saveformat = $Format    }
			"Publish" { $saveformat = $PubFormat }
		}
	}
	process {
		if (!(Test-Path $ExcelFile)) {
			return "Excel file doesn't exist!"
		}
		
		$xlspath = (Resolve-Path $ExcelFile).Path
		
		# Destination is s/.xlsx/.extension by default.
		if (!$DestinationDirectory) {
			$dstpath = $xlspath -Replace "[.]\w+$", $saveformat
		} else {
			if (!(Test-Path $DestinationDirectory)) {
				return "Destination directory doesn't exist!"
			}
		
			$dstfn = (Split-Path -Leaf $xlspath) -Replace "[.]\w+$", $saveformat
			
			$dstpath = Join-Path (Resolve-Path $DestinationDirectory).Path $dstfn
		}
		
		$dstpath_noext = $dstpath -Replace "[.]\w+$", ""
		
		# See: http://support.microsoft.com/kb/320369
		[threading.thread]::CurrentThread.CurrentCulture = "en-US"

		$Excel = New-Object -ComObject Excel.Application
		
		$Excel.DisplayAlerts = $false

		$wb = $Excel.Workbooks.Open($xlspath)
		
		# Saveformat supports multiple sheets.
		if (@(".xlsx", ".xls", ".ods") -contains $saveformat) {
			if ($PSCmdlet.ShouldProcess($dstpath)) {
				$wb.SaveAs(
					$dstpath,
					$formats.get_item($saveformat)
				)
			}
		}
		# We have to save each sheet separately.
		else {
			ForEach ($sheet in $wb.Worksheets) {
				if ($SheetNames -and ($sheets -notcontains $sheet.Name)) {
					continue
				}
				
				$sheetfn = [string]::join("", $dstpath_noext, "_", $sheet.Name, $saveformat)
				
				$sheet.Activate()
				
				switch ($PSCmdlet.ParameterSetName) {
					"Convert" {
						if ($PSCmdlet.ShouldProcess($sheetfn)) {
							# http://msdn.microsoft.com/en-us/library/ff841185(v=office.14).aspx
							$wb.SaveAs(
								$sheetfn,
								$formats.get_item($saveformat)
							)
						}
					}
					"Publish" {
						if ($PSCmdlet.ShouldProcess($sheetfn)) {
							# http://msdn.microsoft.com/en-us/library/ff198122(v=office.14).aspx
							$sheet.ExportAsFixedFormat(
								$pubformats.get_item($saveformat),
								$sheetfn,
								$pubqualities.get_item($PubQuality),
								$KeepProperties,
								$IgnorePrintAreas
								# Have to specify these all if we want to use
								# OpenAfterPublish, so these are disabled
								#$printmin,
								#$printmax,
								#$OpenAfterPublish
							)
						}
					}
				}
			}
		}

		$Excel.Quit()
	}
	end {
		
	}
}

Function ConvertTo-Excel {
	<#
	.SYNOPSIS
	Convert a file to the Excel format
	.DESCRIPTION
	This cmdlet can convert a (CSV) file to the Excel format.
	.EXAMPLE
	ConvertTo-Excel ".\values.csv"
	Creates a file, "values.xlsx" into the current location.
	.EXAMPLE
	ConvertTo-Excel ".\values.csv" -SaveFormat ".ods"
	Creates an OpenDocument Spreadsheet into the current location.
	.EXAMPLE
	ConvertTo-Excel ".\data.txt" -Delimited -TextQualifier "None" -Separator "\t"
	Convert a tab-separated text file into an Excel file.
	.PARAMETER TextFile
	The file you want to convert.
	.PARAMETER DestinationDirectory
	The destination directory.
	.PARAMETER StartRow
	Which row to start conversion from.
	
	default: 1
	.PARAMETER Delimited
	Text is delimited.
	.PARAMETER FixedWidth
	Text is split into fixed width columns.
	.PARAMETER TextQualifier
	Qualifier for text.
	
	Double quote : "DoubleQuote"
	Single quote : "SingleQuote"
	None         : "None"
	
	default: DoubleQuote
	.PARAMETER MergeDelimiters
	Merge multiple delimiters into one.
	.PARAMETER Separator
	The text delimiter.
	
	Tab       : "\t"
	Semicolon : ";"
	Comma     : ","
	Space     : " "
	Other     : Any character
	
	default: ","
	.PARAMETER SaveFormat
	The format you want to convert to.
	
	default: .xlsx
	#>
	[CmdletBinding(
		DefaultParameterSetName="Delimited",
		SupportsShouldProcess=$true,
		ConfirmImpact="Low"
	)]
	Param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
		$TextFile,
		
		[Parameter(Position=1)]
		[string] $DestinationDirectory = "",
		
		[Parameter()]
		[int] $StartRow = 1,
		
		[Parameter(ParameterSetName="Delimited")]
		[switch] $Delimited,
		
		[Parameter(ParameterSetName="FixedWidth")]
		[switch] $FixedWidth,
		
		[Parameter()]
		[ValidateSet("DoubleQuote","SingleQuote","None")]
		[string] $TextQualifier = "DoubleQuote",
		
		[Parameter()]
		[bool] $MergeDelimiters = $false,
		
		[Parameter(ParameterSetName="Delimited")]
		[string] $Separator = ",",
		
		[Parameter()]
		[ValidateSet(".xlsx",".xls",".html",".ods")]
		[string] $SaveFormat = ".xlsx"
	)
	begin {
		# http://msdn.microsoft.com/en-us/library/ff838376(v=office.14).aspx
		$TextQ = @{
			"DoubleQuote" = 1;
			"SingleQuote" = 2;
			"None"        = -4142
		}
		
		# http://msdn.microsoft.com/en-us/library/ff198017(v=office.14).aspx
		$SaveF = @{
			".xlsx" = 51;
			".xls"  = 56;
			".html" = 44;
			".ods"  = 60
		}
		
		$SepTab       = $false
		$SepSemicolon = $false
		$SepComma     = $false
		$SepSpace     = $false
		$SepOther     = $false
		$OtherChar    = ""
		
		switch ($Separator) {
			"\t"    { $SepTab       = $true }
			";"     { $SepSemicolon = $true }
			","     { $SepComma     = $true }
			" "     { $SepSpace     = $true }
			default { $SepOther     = $true; $OtherChar = $Separator }
		}
	}
	process {
		if (!(Test-Path $TextFile)) {
			return "Text file doesn't exist!"
		}
		
		$txtpath = (Resolve-Path $TextFile).Path
		
		# Destination is s/.csv/.extension by default.
		if (!$DestinationDirectory) {
			$xlspath = $txtpath -Replace "[.]\w+$", $SaveFormat
		} else {
			if (!(Test-Path $DestinationDirectory)) {
				return "Destination directory doesn't exist!"
			}
		
			$xlsfn = (Split-Path -Leaf $txtpath) -Replace "[.]\w+$", $SaveFormat
			
			$xlspath = Join-Path (Resolve-Path $DestionationDirectory).Path $xlsfn
		}
		
		# See: http://support.microsoft.com/kb/320369
		[threading.thread]::CurrentThread.CurrentCulture = "en-US"

		$Excel = New-Object -ComObject Excel.Application
		
		switch ($PSCmdlet.ParameterSetName) {
			"Delimited"  {
				# http://msdn.microsoft.com/en-us/library/ff837097(v=office.14).aspx
				$Excel.Workbooks.OpenText(
					$txtpath,
					2,
					$StartRow,
					1,
					$TextQ.Get_Item($TextQualifier),
					$MergeDelimiters,
					$SepTab,
					$SepSemicolon,
					$SepComma,
					$SepSpace,
					$SepOther,
					$OtherChar
				)
			}
			"FixedWidth" {
				$Excel.Workbooks.OpenText(
					$txtpath,
					2,
					$StartRow,
					2,
					$TextQ.Get_Item($TextQualifier),
					$MergeDelimiters
				)
			}
		}
		
		$wb = $Excel.ActiveWorkbook
		
		if ($PSCmdlet.ShouldProcess($xlspath)) {
			try   { $wb.SaveAs($xlspath, $SaveF.Get_Item($SaveFormat)) }
			catch { "Unable to save file.`n`n"; $_ }
		}

		$Excel.Quit()
	}
	end {
		
	}
}

Function Format-Excel {
	<#
	.SYNOPSIS
	Displays the input object in Excel
	.DESCRIPTION
	This cmdlet displays the input object in Excel.
	i.e. it is opened in Excel.
	
	The object is exported into a temporary CSV file first.
	.EXAMPLE
	Get-ChildItem | Format-Excel
	Displays the current directory's contents in Excel.
	.PARAMETER Object
	The input object. Can come from a pipeline.
	#>
	[CmdletBinding(
		SupportsShouldProcess=$false,
		ConfirmImpact="None"
	)]
	Param(
		[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
		[object] $Object
	)
	begin {
		$obj = @()
	}
	process {
		$obj += $Object
	}
	end {
		$tf = [IO.Path]::GetTempFileName()

		$obj | Export-CSV -Delimiter "," -Path $tf `
			-Encoding "Unicode" -NoTypeInformation

		# See: http://support.microsoft.com/kb/320369
		[threading.thread]::CurrentThread.CurrentCulture = "en-US"

		$Excel = New-Object -ComObject Excel.Application

		# http://msdn.microsoft.com/en-us/library/ff837097(v=office.14).aspx
		$Excel.Workbooks.OpenText(
			$tf,
			2,
			1,
			1,
			1,
			$false,
			$false,
			$false,
			$true,
			$false,
			$false,
			""
		)

		$Excel.Visible = $true
	}
}

Function Export-Excel {
	<#
	.SYNOPSIS
	Export the input object into an Excel file
	.DESCRIPTION
	This cmdlet exports the input object into an
	Excel file.
	
	The object is exported into a temporary CSV file first.
	.EXAMPLE
	Get-ChildItem | Export-Excel ".\dir.xlsx"
	.PARAMETER FileName
	Name for the exported file.
	Must end in .xlsx.
	.PARAMETER Object
	The object to export. Can come from a pipeline.
	#>
	[CmdletBinding(
		SupportsShouldProcess=$false,
		ConfirmImpact="Low"
	)]
	Param(
		[Parameter(Position=0,Mandatory=$true)]
		[ValidateNotNull()]
		[string] $FileName = "",
	
		[Parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true)]
		[object] $Object
	)
	begin {
		$obj = @()
	}
	process {
		$obj += $Object
	}
	end {
		$tf = [IO.Path]::GetTempFileName()

		$obj | Export-CSV -Delimiter "," -Path $tf `
			-Encoding "Unicode" -NoTypeInformation
		
		ConvertTo-Excel -TextFile $tf
		
		$path = $tf -Replace "[.]\w+$", ".xlsx"
		
		try { Copy-Item $path $FileName }
		catch {
			"Unable to copy file to '$FileName'"
			"You can access it from '$path'"
			$_
		}
	}
}
