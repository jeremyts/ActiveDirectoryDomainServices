#requires -Version 3.0
#requires -Module DHCPServer
#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft 2012+ DHCP server using Microsoft Word or formatted text.
.DESCRIPTION
	Creates a complete inventory of a Microsoft 2012+ DHCP server using Microsoft Word and PowerShell or just formatted text.
	Creates a Text file, Word document or PDF named after the DHCP server.
	
	Requires PowerShell V3.0 or later.
	Requires the DHCPServer module.
	Can be run on a DHCP server or on a Windows 8.x computer with RSAT installed.
	
	For Windows Server 2003, 2008 and 2008 R2, use the following to export and import the DHCP data:
		Export from the 2003, 2008 or 2008 R2 server:
			netsh dhcp server export C:\DHCPExport.txt all
			
			Copy the C:\DHCPExport.txt file to the 2012+ server.
			
		Import on the 2012+ server:
			netsh dhcp server import c:\DHCPExport.txt all
			
		The script can now be run on the 2012+ DHCP server to document the older DHCP information.

	For Windows Server 2008 R2, the 2012+ DHCP Server PowerShell cmdlets can be used for the export and import.
		From the 2012+ DHCP server:
			Export-DhcpServer –ComputerName 2008R2Server.domain.tld -Leases -File C:\DHCPExport.xml -verbose
			
			Import-DhcpServer –ComputerName 2012Server.domain.tld -Leases –File C:\DHCPExport.xml -BackupPath C:\dhcp\backup\ -Verbose
			
			Note: The c:\dhcp\backup path must exist before the Import-DhcpServer cmdlet is run.
	
	Word and PDF Documents include a Cover Page, Table of Contents and Footer.
	
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(Default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Sideline.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	For Word 2007, the Microsoft add-in for saving as a PDF muct be installed.
	For Word 2007, please see http://www.microsoft.com/en-us/download/details.aspx?id=9943
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
	This parameter is reserved for a future update and no output is created at this time.
.PARAMETER ComputerName
	DHCP server to run the script against.
	This parameter is required.
.PARAMETER IncludeLeases
	Include DHCP lease information.
	Default is to not included lease information.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -ComputerName DHCPServer01
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -ComputerName localhost
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will resolve localhost to $env:computername, for example DHCPServer01.
	Script will be run remotely against DHCP server DHCPServer01 and not localhost.
	Output file name will use the server name DHCPServer01 and not localhost.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -PDF -ComputerName DHCPServer02
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -Text -ComputerName DHCPServer02
	
	Will use all Default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -MSWord -ComputerName DHCPServer02
	
	Will use all Default values and save the document as a Word DOCX file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -HTML -ComputerName DHCPServer02
	
	Will use all Default values and save the output as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory.ps1 -ComputerName DHCPServer03 -IncludeLeases
	
	Will use all Default values and add additional information for each server about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer03.
	Output will contain DHCP lease information.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ComputerName DHCPServer01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V41.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster" -ComputerName DHCPServer02 -IncludeLeases

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
	
	Script will be run remotely against DHCP server DHCPServer02.
	Output will contain DHCP lease information.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word, PDF or formatted text document.
.NOTES
	NAME: DHCP_Inventory.ps1
	VERSION: 1.0
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith)
	LASTEDIT: April 13, 2014
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "WordOrPDF") ]

Param(
	[parameter(ParameterSetName="WordOrPDF",
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordOrPDF",
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordOrPDF",
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 3, 
	Mandatory=$false )
	] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",
	Position = 4, 
	Mandatory=$false )
	] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 4, 
	Mandatory=$false )
	] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="HTML",
	Position = 4, 
	Mandatory=$false )
	] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="WordOrPDF",
	Position = 5, 
	Mandatory=$True )
	] 
	[parameter(ParameterSetName="Text",
	Position = 5, 
	Mandatory=$True )
	] 
	[parameter(ParameterSetName="HTML",
	Position = 5, 
	Mandatory=$True )
	] 
	[string]$ComputerName="", 

	[parameter(ParameterSetName="WordOrPDF",
	Position = 6, 
	Mandatory=$false )
	] 
	[parameter(ParameterSetName="Text",
	Position = 6, 
	Mandatory=$false )
	] 
	[parameter(ParameterSetName="HTML",
	Position = 6, 
	Mandatory=$false )
	] 
	[Switch]$IncludeLeases=$False
	)

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
If($PDF -eq $Null)
{
	$PDF = $False
}
If($Text -eq $Null)
{
	$Text = $False
}
If($HTML -eq $Null)
{
	$HTML = $False
}
If($IncludeLeases -eq $Null)
{
	$IncludeLeases = $False
}
If($MSWord -eq $Null)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($MSWord -eq $Null)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($PDF -eq $Null)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Text -eq $Null)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($HTML -eq $Null)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is " $MSWord
		Write-Verbose "$(Get-Date): PDF is " $PDF
		Write-Verbose "$(Get-Date): Text is " $Text
		Write-Verbose "$(Get-Date): HTML is " $HTML
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

Set-StrictMode -Version 2

[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

If($MSWord -or $PDF)
{
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdSaveFormatPDF = 17
	[int]$wdCaptionPositionAbove  = 0
	[int]$wdCaptionTable  = -2
	[bool]$wdExcludeLabel = $True
	[string]$wdTitleAutoText = ""

	$hash = @{}

	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish

	Switch ($PSCulture.Substring(0,3))
	{
		'ca-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Taula automática 2';
				}
			}

		'da-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automatisk tabel 2';
				}
			}

		'de-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automatische Tabelle 2';
				}
			}

		'en-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents'  = 'Automatic Table 2';
				}
			}

		'es-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Tabla automática 2';
				}
			}

		'fi-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automaattinen taulukko 2';
				}
			}

		'fr-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Sommaire Automatique 2';
				}
			}

		'nb-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automatisk tabell 2';
				}
			}

		'nl-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
				}
			}

		'pt-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Sumário Automático 2';
				}
			}

		'sv-'	{
				$hash.($($PSCulture)) = @{
					'Word_TableOfContents' = 'Automatisk innehållsförteckning2';
				}
			}

		Default	{$hash.('en-US') = @{
					'Word_TableOfContents'  = 'Automatic Table 2';
				}
			}
	}

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	$myHash = $hash.$PSCulture

	If($myHash -eq $Null)
	{
		$myHash = $hash.('en-US')
	}

	$myHash.Word_NoSpacing = $wdStyleNoSpacing
	$myHash.Word_Heading1 = $wdStyleheading1
	$myHash.Word_Heading2 = $wdStyleheading2
	$myHash.Word_Heading3 = $wdStyleheading3
	$myHash.Word_Heading4 = $wdStyleheading4
	$myHash.Word_TableGrid = $wdTableGrid

	Function ValidateCoverPage
	{
		Param([int]$xWordVersion, [string]$xCP)
		
		$xArray = ""
		
		Switch ($PSCulture.Substring(0,3))
		{
			'ca-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
						"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
						"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
						"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
						"Contrast", "Cubicles", "Diplomàtic", "Exposició",
						"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
						"Perspectiva", "Piles", "Quadrícula", "Sobri",
						"Transcendir", "Trencaclosques")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabet", "Anual", "Conservador", "Contrast",
						"Cubicles", "Diplomàtic", "En mosaic", "Exposició", "Línia lateral",
						"Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
					}
				}

			'da-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
						"Retro", "Semafor", "Visningsmaster", "Integral",
						"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
						"Udsnit (mørk)", "Ion (mørk)", "Austin")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
						"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
						"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
						"Nålestribet", "Årlig", "Avispapir", "Tradionel")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabet", "Årlig", "BevægElse", "Eksponering",
						"Enkel", "Firkanter", "Fliser", "Gåde", "Kontrast",
						"Mod", "Nålestribet", "Overskrid", "Sidelinje", "Stakke",
						"Tradionel")
					}
				}

			'de-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
						"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
						"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
						"Randlinie", "Austin", "Integral", "Facette")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
						"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
						"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
						"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt",
						"Jährlich", "Kacheln", "Kontrast", "Kubistisch", "Modern",
						"Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel",
						"Traditionell")
					}
				}

			'en-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
						"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
						"Sideline", "Stacks", "Tiles", "Transcend")
					}
				}

			'es-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
						"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
						"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
						"Ion (claro)", "Integral", "Con bandas")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
						"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
						"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
						"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabeto", "Anual", "Austero", "Conservador",
						"Contraste", "Cubículos", "Exposición", "Línea lateral",
						"Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle",
						"Rayas", "Sobrepasar")
					}
				}

			'fi-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
						"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
						"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
						"Kiehkura", "Liike", "Ruudukko", "Sivussa")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
						"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
						"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
						"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Aakkoset", "Alttius", "Kontrasti", "Kuvakkeet ja tiedot",
						"Liike" , "Liituraita" , "Mod" , "Palapeli", "Perinteinen", "Pinot",
						"Sivussa", "Työpisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
					}
				}

			'fr-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
						"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
						"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
						"Guide", "Whisp", "Lignes latérales", "Quadrillage")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
						"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
						"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
						"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annuel", "Austère", "Blocs empilés", "Blocs superposés",
						"Classique", "Contraste", "Exposition", "Guide", "Ligne latérale", "Moderne",
						"Mosaïques", "Mots croisés", "Rayures fines", "Transcendant")
					}
				}

			'nb-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
						"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
						"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
						"ViewMaster")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
						"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
						"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
						"Smale striper", "Stabler", "Transcenderende")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabet", "Årlig", "Avlukker", "BevegElse", "Engasjement",
						"Enkel", "Fliser", "Konservativ", "Kontrast", "Mod", "Puslespill",
						"Sidelinje", "Smale striper", "Stabler", "Transcenderende")
					}
				}

			'nl-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
						"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
						"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
						"Terugblik", "Terzijde", "ViewMaster")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
						"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
						"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
						"Puzzel", "Raster", "Stapels",
						"Tegels", "Terzijde")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging",
						"Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks", "Krijtstreep",
						"Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
					}
				}

			'pt-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
						"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana",
						"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
						"Retrospectiva", "Semáforo")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
						"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
						"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
						"Quebra-cabeça", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Baias", "Conservador",
						"Contraste", "Exposição", "Ladrilhos", "Linha Lateral", "Listras", "Mod",
						"Pilhas", "Quebra-cabeça", "Transcendente")
					}
				}

			'sv-'	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
						"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
						"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
						"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
						"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
						"Övergående")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alfabetmönster", "Årligt", "Enkelt", "Exponering", "Konservativt",
						"Kontrast", "Kritstreck", "Kuber", "Övergående", "Plattor", "Pussel", "RörElse",
						"Sidlinje", "Sobert", "Staplat")
					}
				}

			Default	{
						If($xWordVersion -eq $wdWord2013)
						{
							$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
							"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
							"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
						}
						ElseIf($xWordVersion -eq $wdWord2010)
						{
							$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
							"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
							"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
						}
						ElseIf($xWordVersion -eq $wdWord2007)
						{
							$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
							"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
							"Sideline", "Stacks", "Tiles", "Transcend")
						}
					}
		}
		
		If($xArray -contains $xCP)
		{
			Return $True
		}
		Else
		{
			Return $False
		}
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function CheckWord2007SaveAsPDFInstalled
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
	{
		Write-Host "Word 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
		Write-Host "The add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
		Write-Host "Install the SaveAs PDF add-in and rerun the script."
		Return $False
	}
	Return $True
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014

{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )

	While( $tabs -gt 0 ) { $global:output += "`t"; $tabs--; }

	If( $nonewline )
	{
		$global:output += $name + $value
	}
	Else
	{
		$global:output += $name + $value + $newline
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$null,
	[int]$fontSize=0,
	[bool]$italics=$false,
	[bool]$boldface=$false,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function AbortScript
{
	$Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

Function GetIPv4ScopeData_WordPDF
{
	#put each scope on a new page
	$selection.InsertNewPage()
	Write-Verbose "$(Get-Date): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"
	WriteWordLine $xStartLevel 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
	WriteWordLine ($xStartLevel + 1) 0 "Address Pool"
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	[int]$Rows = 5
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	$Table.Cell(1,1).Range.Text = "Start IP Address"
	$Table.Cell(1,2).Range.Text = $IPv4Scope.StartRange
	$Table.Cell(2,1).Range.Text = "End IP Address"
	$Table.Cell(2,2).Range.Text = $IPv4Scope.EndRange
	$Table.Cell(3,1).Range.Text = "Subnet Mask"
	$Table.Cell(3,2).Range.Text = $IPv4Scope.SubnetMask
	$Table.Cell(4,1).Range.Text = "Lease duration"
	If($IPv4Scope.LeaseDuration -eq "00:00:00")
	{
		$Table.Cell(4,2).Range.Text = "Unlimited"
	}
	Else
	{
		$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
			$IPv4Scope.LeaseDuration.Days, `
			$IPv4Scope.LeaseDuration.Hours, `
			$IPv4Scope.LeaseDuration.Minutes)

		$Table.Cell(4,2).Range.Text = $Str
	}
	$Table.Cell(5,1).Range.Text = "Description"
	$Table.Cell(5,2).Range.Text = $IPv4Scope.Description

	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		
		WriteWordLine ($xStartLevel + 1) 0 "Address Leases"
		$Leases = Get-DHCPServerV4Lease -ComputerName $DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
		If($? -and $Leases -ne $Null)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Leases -is [array])
			{
				[int]$Rows = ($Leases.Count * 11) - 1
				#subtract the very last row used for spacing
			}
			Else
			{
				[int]$Rows = 10
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			[int]$xRow = 0
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
				If($Lease.LeaseExpiryTime -ne $Null)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Days, `
						$Lease.LeaseExpiryTime.Hours, `
						$Lease.LeaseExpiryTime.Minutes)
				}
				Else
				{
					$LeaseStr = ""
				}

				If($Lease.ProbationEnds -ne $Null)
				{
					$ProbationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.ProbationEnds.Days, `
						$Lease.ProbationEnds.Hours, `
						$Lease.ProbationEnds.Minutes)
				}
				Else
				{
					$ProbationStr = ""
				}

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Client IP address"
				$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Lease.HostName
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
				If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
				{
					If($Lease.AddressState -eq "ActiveReservation")
					{
						$Table.Cell($xRow,2).Range.Text = "Reservation (active)"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = "Reservation (inactive)"
					}
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $LeaseStr
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Type"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientType
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Unique ID"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientID
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Lease.Description

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Network Access Protection"
				$Table.Cell($xRow,2).Range.Text = $Lease.NapStatus

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Probation Expiration"
				
				If([string]::IsNullOrEmpty($Lease.ProbationEnds))
				{
					$Table.Cell($xRow,2).Range.Text = "N/A"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ProbationStr
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Filter"
				
				$Filters | % { $index = $null }{ if( $_.MacAddress -eq $Lease.ClientID ) { $index = $_ } }
				If($Index -ne $Null)
				{
					$Table.Cell($xRow,2).Range.Text = $Index.List
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy"
				
				If([string]::IsNullOrEmpty($Lease.PolicyName))
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $Lease.PolicyName
				}
				
				#skip a row for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
		}
		Else
		{
			WriteWordLine 0 0 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	WriteWordLine ($xStartLevel + 1) 0 "Reservations"
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
	If($? -and $Reservations -ne $Null)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If([string]::IsNullOrEmpty($Reservation.Description))
			{
				[int]$Rows = 4
			}
			Else
			{
				[int]$Rows = 5
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			$Table.Cell(1,1).Range.Text = "Reservation name"
			$Table.Cell(1,2).Range.Text = $Reservation.Name
			$Table.Cell(2,1).Range.Text = "IP address"
			$Table.Cell(2,2).Range.Text = $Reservation.IPAddress
			$Table.Cell(3,1).Range.Text = "MAC address"
			$Table.Cell(3,2).Range.Text = $Reservation.ClientId
			$Table.Cell(4,1).Range.Text = "Supported types"
			$Table.Cell(4,2).Range.Text = $Reservation.Type
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$Table.Cell(5,1).Range.Text = "Description"
				$Table.Cell(5,2).Range.Text = $Reservation.Description
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null

			Write-Verbose "$(Get-Date):	`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $DNSSettings -ne $Null)
			{
				GetDNSSettings $DNSSettings "A"
			}
			Else
			{
				WriteWordLine 0 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
			WriteWordLine 0 0 ""
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting scope options"
	WriteWordLine ($xStartLevel + 1) 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $ScopeOptions -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopeOptions -is [array])
		{
			[int]$Rows = (($ScopeOptions.Count * 5) - 5) - 1
			#subtract option 51
			#subtract the very last row used for spacing
		}
		Else
		{
			[int]$Rows = 4
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($ScopeOption in $ScopeOptions)
		{
			If($ScopeOption.OptionId -ne 51)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Option Name"
				$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Vendor"
				If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
				{
					$Table.Cell($xRow,2).Range.Text = "Standard" 
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ScopeOption.VendorClass 
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Value"
				$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.Value)" 
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy Name"
				
				If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ScopeOption.PolicyName
				}
			
				#for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting policies"
	WriteWordLine ($xStartLevel + 1) 0 "Policies"
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $ScopePolicies -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopePolicies -is [array])
		{
			[int]$Rows = $ScopePolicies.Count * 6
		}
		Else
		{
			[int]$Rows = 6
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing policy name $($Policy.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Policy Name"
			$Table.Cell($xRow,2).Range.Text = $Policy.Name
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Description"
			$Table.Cell($xRow,2).Range.Text = $Policy.Description

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Processing Order"
			$Table.Cell($xRow,2).Range.Text = $Policy.ProcessingOrder

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Level"
			$Table.Cell($xRow,2).Range.Text = "Scope"

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Address Range"
			
			$IPRange = Get-DHCPServerV4PolicyIPRange -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0

			If($? -and $IPRange -ne $Null)
			{
				$Table.Cell($xRow,2).Range.Text = "$($IPRange.StartRange) - $($IPRange.EndRange)"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "<None>"
			}

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State"
			If($Policy.Enabled)
			{
				$Table.Cell($xRow,2).Range.Text = "Enabled"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "Disabled"
			}
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope policies"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting DNS"
	WriteWordLine ($xStartLevel + 1) 0 "DNS"
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		WriteWordLine 0 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date):	`t`tGetting Failover"
	WriteWordLine ($xStartLevel + 1) 0 "Failover"
	$Failovers = Get-DHCPServerV4Failover -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Failovers -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Failovers -is [array])
		{
			[int]$Rows = ($Failovers.Count * 10) - 1
			#subtract the very last row used for spacing
		}
		Else
		{
			[int]$Rows = 9
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Relationship name"
			$Table.Cell($xRow,2).Range.Text = $Failover.Name
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Partner Server"
			$Table.Cell($xRow,2).Range.Text = $Failover.PartnerServer
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Mode"
			$Table.Cell($xRow,2).Range.Text = $Failover.Mode
					
			If($Failover.MaxClientLeadTime -ne $Null)
			{
				$MaxLeadStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.MaxClientLeadTime.Days, `
					$Failover.MaxClientLeadTime.Hours, `
					$Failover.MaxClientLeadTime.Minutes)
			}
			Else
			{
				$MaxLeadStr = ""
			}

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Max Client Lead Time"
			$Table.Cell($xRow,2).Range.Text = $MaxLeadStr
					
			If($Failover.StateSwitchInterval -ne $Null)
			{
				$SwitchStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.StateSwitchInterval.Days, `
					$Failover.StateSwitchInterval.Hours, `
					$Failover.StateSwitchInterval.Minutes)
			}
			Else
			{
				$SwitchStr = "Disabled"
			}

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State Switchover Interval"
			$Table.Cell($xRow,2).Range.Text = $SwitchStr
					
			Switch($Failover.State)
			{
				"NoState" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "No State"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "No State"
				}
				"Normal" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
				}
				"Init" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Initializing"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Initializing"
				}
				"CommunicationInterrupted" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
				}
				"PartnerDown" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Down"
				}
				"PotentialConflict" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
				}
				"Startup" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Starting Up"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Starting Up"
				}
				"ResolutionInterrupted" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
				}
				"ConflictDone" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Conflict Done"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Conflict Done"
				}
				"Recover" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover"
				}
				"RecoverWait" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Wait"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Wait"
				}
				"RecoverDone" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Done"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Done"
				}
				Default 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
				}
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Local server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.LoadBalancePercent)%"
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Partner Server"
				$tmp = (100 - $($Failover.LoadBalancePercent))
				$Table.Cell($xRow,2).Range.Text = "$($tmp)%"
				$tmp = $Null
			}
			Else
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Role of this server"
				$Table.Cell($xRow,2).Range.Text = $Failover.ServerRole
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Addresses reserved for standby server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			$xRow++
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Failovers = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	WriteWordLine ($xStartLevel + 1) 0 "Statistics"

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Statistics -ne $Null)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv4ScopeData_HTML
{

}

Function GetIPv4ScopeData_Text
{
	Write-Verbose "$(Get-Date): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"
	Line 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
	Line 1 "Address Pool:"
	Line 2 "Start IP Address`t: " $IPv4Scope.StartRange
	Line 2 "End IP Address`t`t: " $IPv4Scope.EndRange
	Line 2 "Subnet Mask`t`t: " $IPv4Scope.SubnetMask
	Line 2 "Lease duration`t`t: " -NoNewLine
	If($IPv4Scope.LeaseDuration -eq "00:00:00")
	{
		Line 0 "Unlimited"
	}
	Else
	{
		$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
			$IPv4Scope.LeaseDuration.Days, `
			$IPv4Scope.LeaseDuration.Hours, `
			$IPv4Scope.LeaseDuration.Minutes)

		Line 0 $Str
	}
	If(![string]::IsNullOrEmpty($IPv4Scope.Description))
	{
		Line 2 "Description`t`t: " $IPv4Scope.Description
	}

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		
		Line 1 "Address Leases:"
		$Leases = Get-DHCPServerV4Lease -ComputerName $DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
		If($? -and $Leases -ne $Null)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
				If($Lease.LeaseExpiryTime -ne $Null)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Days, `
						$Lease.LeaseExpiryTime.Hours, `
						$Lease.LeaseExpiryTime.Minutes)
				}
				Else
				{
					$LeaseStr = ""
				}

				If($Lease.ProbationEnds -ne $Null)
				{
					$ProbationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.ProbationEnds.Days, `
						$Lease.ProbationEnds.Hours, `
						$Lease.ProbationEnds.Minutes)
				}
				Else
				{
					$ProbationStr = ""
				}

				Line 2 "Name: " $Lease.HostName
				Line 2 "Client IP address`t`t: " $Lease.IPAddress
				Line 2 "Lease Expiration`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
				{
					If($Lease.AddressState -eq "ActiveReservation")
					{
						Line 0 "Reservation (active)"
					}
					Else
					{
						Line 0 "Reservation (inactive)"
					}
				}
				Else
				{
					Line 0 $LeaseStr
				}
				Line 2 "Type`t`t`t`t: " $Lease.ClientType
				Line 2 "Unique ID`t`t`t: " $Lease.ClientID
				If(![string]::IsNullOrEmpty($Lease.Description))
				{
					Line 2 "Description`t`t`t: " $Lease.Description
				}
				Line 2 "Network Access Protection`t: " $Lease.NapStatus
				Line 2 "Probation Expiration`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($Lease.ProbationEnds))
				{
					Line 0 "N/A"
				}
				Else
				{
					Line 0 $ProbationStr
				}
				Line 2 "Filter`t`t`t`t: " -NoNewLine
				
				$Filters | % { $index = $null }{ if( $_.MacAddress -eq $Lease.ClientID ) { $index = $_ } }
				If($Index -ne $Null)
				{
					Line 0 $Index.List
				}
				Else
				{
					Line 0 "<None>"
				}
				Line 2 "Policy`t`t`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($Lease.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $Lease.PolicyName
				}
				
				#skip a row for spacing
				Line 0 ""
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
		}
		Else
		{
			Line 2 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	Line 1 "Reservations:"
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
	If($? -and $Reservations -ne $Null)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			Line 2 "Reservation name`t: " $Reservation.Name
			Line 2 "IP address`t`t: " $Reservation.IPAddress
			Line 2 "MAC address`t`t: " $Reservation.ClientId
			Line 2 "Supported types`t`t: " $Reservation.Type
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				Line 2 "Description`t`t: " $Reservation.Description
			}

			Write-Verbose "$(Get-Date):	`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $DNSSettings -ne $Null)
			{
				GetDNSSettings $DNSSettings "A"
			}
			Else
			{
				Line 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting scope options"
	Line 1 "Scope Options:"
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $ScopeOptions -ne $Null)
	{
		ForEach($ScopeOption in $ScopeOptions)
		{
			If($ScopeOption.OptionId -ne 51)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
				Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
				Line 2 "Vendor`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
				{
					Line 0 "Standard" 
				}
				Else
				{
					Line 0 $ScopeOption.VendorClass 
				}
				Line 2 "Value`t`t: $($ScopeOption.Value)" 
				Line 2 "Policy Name`t: " -NoNewLine
				
				If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $ScopeOption.PolicyName
				}
			
				#for spacing
				Line 0 ""
			}
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting policies"
	Line 1 "Policies:"
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $ScopePolicies -ne $Null)
	{
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing policy name $($Policy.Name)"
			Line 2 "Policy Name`t`t: " $Policy.Name
			If(![string]::IsNullOrEmpty($Policy.Description))
			{
				Line 2 "Description`t`t: " $Policy.Description
			}
			Line 2 "Processing Order`t: " $Policy.ProcessingOrder
			Line 2 "Level`t`t`t: Scope"
			Line 2 "Address Range`t: " -NoNewLine
			
			$IPRange = Get-DHCPServerV4PolicyIPRange -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0

			If($? -and $IPRange -ne $Null)
			{
				Line 0 "$($IPRange.StartRange) - $($IPRange.EndRange)"
			}
			Else
			{
				Line 0 "<None>"
			}
			Line 2 "State`t`t`t: " -NoNewLine
			If($Policy.Enabled)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope policies"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting DNS"
	Line 1 "DNS:"
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		Line 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date):	`t`tGetting Failover"
	Line 1 "Failover:"
	$Failovers = Get-DHCPServerV4Failover -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Failovers -ne $Null)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			Line 2 "Relationship name: " $Failover.Name
			Line 2 "Partner Server`t`t`t: " $Failover.PartnerServer
			Line 2 "Mode`t`t`t`t: " $Failover.Mode
					
			If($Failover.MaxClientLeadTime -ne $Null)
			{
				$MaxLeadStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.MaxClientLeadTime.Days, `
					$Failover.MaxClientLeadTime.Hours, `
					$Failover.MaxClientLeadTime.Minutes)
			}
			Else
			{
				$MaxLeadStr = ""
			}

			Line 2 "Max Client Lead Time`t`t: " $MaxLeadStr
					
			If($Failover.StateSwitchInterval -ne $Null)
			{
				$SwitchStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.StateSwitchInterval.Days, `
					$Failover.StateSwitchInterval.Hours, `
					$Failover.StateSwitchInterval.Minutes)
			}
			Else
			{
				$SwitchStr = "Disabled"
			}

			Line 2 "State Switchover Interval`t: " $SwitchStr
					
			Switch($Failover.State)
			{
				"NoState" 
				{
					Line 2 "State of this Server`t`t: No State"
					Line 2 "State of Partner Server`t`t: No State"
				}
				"Normal" 
				{
					Line 2 "State of this Server`t`t: Normal"
					Line 2 "State of Partner Server`t`t: Normal"
				}
				"Init" 
				{
					Line 2 "State of this Server`t`t: Initializing"
					Line 2 "State of Partner Server`t`t: Initializing"
				}
				"CommunicationInterrupted" 
				{
					Line 2 "State of this Server`t`t: Communication Interrupted"
					Line 2 "State of Partner Server`t`t: Communication Interrupted"
				}
				"PartnerDown" 
				{
					Line 2 "State of this Server`t`t: Normal"
					Line 2 "State of Partner Server`t`t: Down"
				}
				"PotentialConflict" 
				{
					Line 2 "State of this Server`t`t: Potential Conflict"
					Line 2 "State of Partner Server`t`t: Potential Conflict"
				}
				"Startup" 
				{
					Line 2 "State of this Server`t`t: Starting Up"
					Line 2 "State of Partner Server`t`t: Starting Up"
				}
				"ResolutionInterrupted" 
				{
					Line 2 "State of this Server`t`t: Resolution Interrupted"
					Line 2 "State of Partner Server`t`t: Resolution Interrupted"
				}
				"ConflictDone" 
				{
					Line 2 "State of this Server`t`t: Conflict Done"
					Line 2 "State of Partner Server`t`t: Conflict Done"
				}
				"Recover" 
				{
					Line 2 "State of this Server`t`t: Recover"
					Line 2 "State of Partner Server`t`t: Recover"
				}
				"RecoverWait" 
				{
					Line 2 "State of this Server`t`t: Recover Wait"
					Line 2 "State of Partner Server`t`t: Recover Wait"
				}
				"RecoverDone" 
				{
					Line 2 "State of this Server`t`t: Recover Done"
					Line 2 "State of Partner Server`t`t: Recover Done"
				}
				Default 
				{
					Line 2 "State of this Server`t`t: Unable to determine Server: state"
					Line 2 "State of Partner Server`t`t: Unable to determine Server: state"
				}
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				Line 2 "Local server`t`t`t: $($Failover.LoadBalancePercent)%"
				Line 2 "Partner Server`t`t`t: " -NoNewLine
				$tmp = (100 - $($Failover.LoadBalancePercent))
				Line 0 "$($tmp)%"
				$tmp = $Null
			}
			Else
			{
				Line 2 "Role of this server`t`t: " $Failover.ServerRole
				Line 2 "Addresses reserved for standby server: $($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			Line 0 ""
		}
	}
	Else
	{
		Line 2 "<None>"
	}
	$Failovers = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	Line 1 "Statistics:"

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Statistics -ne $Null)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		Line "Error retrieving scope statistics"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv4ScopeData
{
	Param([object]$IPv4Scope, [int]$xStartLevel)
	
	If($MSWord -or $PDF)
	{
		GetIPv4ScopeData_WordPDF
	}
	ElseIf($Text)
	{
		GetIPv4ScopeData_Text
	}
	ElseIf($HTML)
	{
		GetIPv4ScopeData_HTML
	}
}

Function GetIPv6ScopeData_WordPDF
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	WriteWordLine 3 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
	WriteWordLine 4 0 "General"
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		[int]$Rows = 6
	}
	Else
	{
		[int]$Rows = 5
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	$Table.Cell(1,1).Range.Text = "Prefix"
	$Table.Cell(1,2).Range.Text = $IPv6Scope.Prefix
	$Table.Cell(2,1).Range.Text = "Preference"
	$Table.Cell(2,2).Range.Text = $IPv6Scope.Preference
	$Table.Cell(3,1).Range.Text = "Available Range"
	$Table.Cell(3,2).Range.Text = ""
	$Table.Cell(4,1).Range.Text = "`tStart"
	$Table.Cell(4,2).Range.Text = "$($IPv6Scope.Prefix)0:0:0:1"
	$Table.Cell(5,1).Range.Text = "`tEnd"
	$Table.Cell(5,2).Range.Text = "$($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		$Table.Cell(6,1).Range.Text = "Description"
		$Table.Cell(6,2).Range.Text = $IPv6Scope.Description
	}
	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null

	Write-Verbose "$(Get-Date): `t`tGetting scope DNS settings"
	WriteWordLine 4 0 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting scope lease settings"
	WriteWordLine 4 0 "Lease"
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	[int]$Rows = 2
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	$Table.Cell(1,1).Range.Text = "Preferred life time"
	$Table.Cell(1,2).Range.Text = $PrefStr
	$Table.Cell(2,1).Range.Text = "Valid life time"
	$Table.Cell(2,2).Range.Text = $ValidStr

	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	
	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		WriteWordLine 4 0 "Address Leases"
		$Leases = Get-DHCPServerV6Lease -ComputerName $DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Leases -ne $Null)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Leases -is [array])
			{
				[int]$Rows = $Leases.Count * 8
			}
			Else
			{
				[int]$Rows = 7
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			[int]$xRow = 0
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
				$xRow++
				If($Lease.LeaseExpiryTime -ne $Null)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Days, `
						$Lease.LeaseExpiryTime.Hours, `
						$Lease.LeaseExpiryTime.Minutes)
				}
				Else
				{
					$LeaseStr = ""
				}

				$Table.Cell($xRow,1).Range.Text = "Client IPv6 address"
				$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Lease.HostName
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
				$Table.Cell($xRow,2).Range.Text = $LeaseStr
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "IAID"
				$Table.Cell($xRow,2).Range.Text = $Lease.Iaid
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Type"
				$Table.Cell($xRow,2).Range.Text = $Lease.AddressType
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Unique ID"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientDuid
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Lease.Description
				
				#skip a row for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
		}
		Else
		{
			WriteWordLine 0 1 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
	WriteWordLine 4 0 "Exclusions"
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Exclusions -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Exclusions -is [array])
		{
			[int]$Rows = $Exclusions.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Start IP Address"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "End IP Address"
		[int]$xRow = 1
		ForEach($Exclusion in $Exclusions)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Exclusion.StartRange
			$Table.Cell($xRow,2).Range.Text = $Exclusion.EndRange 
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	WriteWordLine 4 0 "Reservations"
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
	If($? -and $Reservations -ne $Null)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If([string]::IsNullOrEmpty($Reservation.Description))
			{
				[int]$Rows = 4
			}
			Else
			{
				[int]$Rows = 5
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			$Table.Cell(1,1).Range.Text = "Reservation name"
			$Table.Cell(1,2).Range.Text = $Reservation.Name
			$Table.Cell(2,1).Range.Text = "IPv6 address"
			$Table.Cell(2,2).Range.Text = $Reservation.IPAddress
			$Table.Cell(3,1).Range.Text = "DUID"
			$Table.Cell(3,2).Range.Text = $Reservation.ClientDuid
			$Table.Cell(4,1).Range.Text = "IAID"
			$Table.Cell(4,2).Range.Text = $Reservation.Iaid
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$Table.Cell(5,1).Range.Text = "Description"
				$Table.Cell(5,2).Range.Text = $Reservation.Description
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null

			Write-Verbose "$(Get-Date):	`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $DNSSettings -ne $Null)
			{
				GetDNSSettings $DNSSettings "AAAA"
			}
			Else
			{
				WriteWordLine 0 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Reservation = $Null

	Write-Verbose "$(Get-Date):	Getting IPv6 scope options"
	WriteWordLine 4 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $ScopeOptions -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopeOptions -is [array])
		{
			[int]$Rows = $ScopeOptions.Count * 4
		}
		Else
		{
			[int]$Rows = 3
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Option Name"
			$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Vendor"
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				$Table.Cell($xRow,2).Range.Text = "Standard" 
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = $ScopeOption.VendorClass 
			}
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Value"
			$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.Value)" 
			
			#for spacing
			$xRow++
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 scope options"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	WriteWordLine 4 0 "Statistics"

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Statistics -ne $Null)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv6ScopeData_HTML
{
}

Function GetIPv6ScopeData_Text
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	Line 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
	Line 1 "General"
	Line 2 "Prefix`t`t: " $IPv6Scope.Prefix
	Line 2 "Preference`t: " $IPv6Scope.Preference
	Line 2 "Available Range`t: "
	Line 3 "Start`t: $($IPv6Scope.Prefix)0:0:0:1"
	Line 3 "End`t: $($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		Line 2 "Description`t: " $IPv6Scope.Description
	}

	Write-Verbose "$(Get-Date): `t`tGetting scope DNS settings"
	Line 1 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		Line 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting scope lease settings"
	Line 1 "Lease"
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
	Line 2 "Preferred life time`t: " $PrefStr
	Line 2 "Valid life time`t`t: " $ValidStr

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		Line 1 "Address Leases:"
		$Leases = Get-DHCPServerV6Lease -ComputerName $DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Leases -ne $Null)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
				If($Lease.LeaseExpiryTime -ne $Null)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Days, `
						$Lease.LeaseExpiryTime.Hours, `
						$Lease.LeaseExpiryTime.Minutes)
				}
				Else
				{
					$LeaseStr = ""
				}

				Line 2 "Client IPv6 address: " $Lease.IPAddress
				Line 2 "Name`t`t`t: " $Lease.HostName
				Line 2 "Lease Expiration`t: " $LeaseStr
				Line 2 "IAID`t`t`t: " $Lease.Iaid
				Line 2 "Type`t`t`t: " $Lease.AddressType
				Line 2 "Unique ID`t`t: " $Lease.ClientDuid
				If(![string]::IsNullOrEmpty($Lease.Description))
				{
					Line 2 "Description`t`t: " $Lease.Description
				}
				
				#skip a row for spacing
				Line 0 ""
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
		}
		Else
		{
			Line 2 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
	Line 1 "Exclusions:"
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Exclusions -ne $Null)
	{
		ForEach($Exclusion in $Exclusions)
		{
			Line 2 "Start IP Address`t: " $Exclusion.StartRange
			Line 2 "End IP Address`t`t: " $Exclusion.EndRange 
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	Line 1 "Reservations:"
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
	If($? -and $Reservations -ne $Null)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			Line 2 "Reservation name: " $Reservation.Name
			Line 2 "IPv6 address: " $Reservation.IPAddress
			Line 2 "DUID`t`t: " $Reservation.ClientDuid
			Line 2 "IAID`t`t: " $Reservation.Iaid
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				Line 2 "Description`t: " $Reservation.Description
			}

			Write-Verbose "$(Get-Date):	`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $DNSSettings -ne $Null)
			{
				GetDNSSettings $DNSSettings "AAAA"
			}
			Else
			{
				Line 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Reservation = $Null

	Write-Verbose "$(Get-Date):	Getting IPv6 scope options"
	Line 1 "Scope Options:"
	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $ScopeOptions -ne $Null)
	{
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
			Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
			Line 2 "Vendor`t`t: " -NoNewLine
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				Line 0 "Standard" 
			}
			Else
			{
				Line 0 $ScopeOption.VendorClass 
			}
			Line 2 "Value`t`t: $($ScopeOption.Value)" 
			
			#for spacing
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving IPv6 scope options"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	Line 1 "Statistics:"

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Statistics -ne $Null)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope statistics"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Statistics = $Null
}

Function GetIPV6ScopeData
{
	Param([object]$IPv6Scope)

	If($MSWord -or $PDF)
	{
		GetIPv6ScopeData_WordPDF
	}
	ElseIf($Text)
	{
		GetIPv6ScopeData_Text
	}
	ElseIf($HTML)
	{
		GetIPv6ScopeData_HTML
	}
}

Function GetDNSSettings
{
	Param([object]$DNSSettings, [string]$As)
	
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 4
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xxRow = 1
		$Table.Cell($xxRow,1).Range.Text = "Enable DNS dynamic updates"
		If($DNSSettings.DynamicUpdates -eq "Never")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Disabled"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
			$xxRow++
			$Table.Cell($xxRow,1).Range.Text =  "Dynamically update DNS $($As) and PTR records only if requested by the DHCP clients"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "Always")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
			$xxRow++
			$Table.Cell($xxRow,1).Range.Text =  "Always dynamically update DNS $($As) and PTR records"
		}
		$xxRow++
		$Table.Cell($xxRow,1).Range.Text = "Discard $($As) and PTR records when lease is deleted"
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$Table.Cell($xxRow,2).Range.Text = "Enabled"
		}
		Else
		{
			$Table.Cell($xxRow,2).Range.Text = "Disabled"
		}
		$xxRow++
		$Table.Cell($xxRow,1).Range.Text = "Name Protection"
		If($DNSSettings.NameProtection)
		{
			$Table.Cell($xxRow,2).Range.Text = "Enabled"
		}
		Else
		{
			$Table.Cell($xxRow,2).Range.Text = "Disabled"
		}

		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
		If($DNSSettings.DynamicUpdates -eq "Never")
		{
			Line 0 "Disabled"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
		{
			Line 0 "Enabled"
			Line 2 "Dynamically update DNS $($As) & PTR records only if requested by the DHCP clients"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "Always")
		{
			Line 0 "Enabled"
			Line 2 "Always dynamically update DNS $($As) & PTR records"
		}
		Line 2 "Discard $($As) & PTR records when lease deleted`t: " -NoNewLine
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			Line 0 "Enabled"
		}
		Else
		{
			Line 0 "Disabled"
		}
		Line 2 "Name Protection`t`t`t`t`t: " -NoNewLine
		If($DNSSettings.NameProtection)
		{
			Line 0 "Enabled"
		}
		Else
		{
			Line 0 "Disabled"
		}
	}
	ElseIf($HTML)
	{
	}
}

Function GetShortStatistics
{
	Param([object]$Statistics)
	
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 4
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Description"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Details"

		$Table.Cell(2,1).Range.Text = "Total Addresses"
		[decimal]$TotalAddresses = "{0:N0}" -f ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$Table.Cell(2,2).Range.Text = $TotalAddresses
		$Table.Cell(3,1).Range.Text = "In Use"
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		$Table.Cell(3,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent))%"
		$Table.Cell(4,1).Range.Text = "Available"
		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$Table.Cell(4,2).Range.Text = "$($Statistics.AddressesFree) ($($AvailablePercent))%"

		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		Line 2 "Description" -NoNewLine
		Line 1 "Details"
		Line 2 "Total Addresses`t" -NoNewLine
		[decimal]$TotalAddresses = ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$tmp = "{0:N0}" -f $TotalAddresses
		Line 0 $tmp
		Line 2 "In Use`t`t" -NoNewLine
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		$tmp = "{0:N0}" -f $Statistics.AddressesInUse
		Line 0 "$($tmp) ($($InUsePercent))%"
		Line 2 "Available`t" -NoNewLine
		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$tmp = "{0:N0}" -f $Statistics.AddressesFree
		Line 0 "$($tmp) ($($AvailablePercent))%"
		Line 0 ""
	}
	ElseIf($HTML)
	{
	}
}

Function InsertBlankLine
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
	}
}

#Script begins

$script:startTime = Get-Date
$global:output = ""

#get server name

#first test to make sure the dhcp server is reachable
Write-Verbose "$(Get-Date): Testing to see if $($ComputerName) is online and reachable"
If(Test-Connection -ComputerName $ComputerName -quiet -EA 0)
{
	Write-Verbose "$(Get-Date): DHCP Server $($ComputerName) is online"
}
Else
{
	Write-Verbose "$(Get-Date): DHCP Server $($ComputerName) is offline"
	Write-Error "DHCP Server $($ComputerName) is offline.`nScript cannot continue."
	Exit
}

#if dhcp server name is localhost, get actual server name
If($ComputerName -eq "localhost")
{
	$DHCPServerName = $env:ComputerName
	Write-Verbose "$(Get-Date): DHCP Server has been renamed from localhost to $($DHCPServerName)"
}
Else
{
	$DHCPServerName = $ComputerName
}

If($MSWord -or $PDF)
{
	CheckWordPreReq

	Write-Verbose "$(Get-Date): Setting up Word"

	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
	$Word = New-Object -comobject "Word.Application" -EA 0

	If(!$? -or $Word -eq $Null)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		Write-Error "The Word object could not be created.  You may need to repair your Word installation.  Script cannot continue."
		Exit
	}

	[int]$WordVersion = [int]$Word.Version
	If($WordVersion -eq $wdWord2013)
	{
		$WordProduct = "Word 2013"
	}
	ElseIf($WordVersion -eq $wdWord2010)
	{
		$WordProduct = "Word 2010"
	}
	ElseIf($WordVersion -eq $wdWord2007)
	{
		$WordProduct = "Word 2007"
	}
	Else
	{
		Write-Error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
		AbortScript
	}

	Write-Verbose "$(Get-Date): Running Microsoft $WordProduct"

	If($PDF -and $WordVersion -eq $wdWord2007)
	{
		Write-Verbose "$(Get-Date): Verify the Word 2007 Save As PDF add-in is installed"
		If(CheckWord2007SaveAsPDFInstalled)
		{
			Write-Verbose "$(Get-Date): The Word 2007 Save As PDF add-in is installed"
		}
		Else
		{
			AbortScript
		}
	}

	Write-Verbose "$(Get-Date): Validate company name"
	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		$CompanyName = ValidateCompanyName
		If([String]::IsNullOrEmpty($CompanyName))
		{
			Write-Warning "Company Name cannot be blank."
			Write-Warning "Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Error "Script cannot continue.  See messages above."
			AbortScript
		}
	}

	Write-Verbose "$(Get-Date): Check Default Cover Page for language specific version"
	[bool]$CPChanged = $False
	Switch ($PSCulture.Substring(0,3))
	{
		'ca-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Línia lateral"
					$CPChanged = $True
				}
			}

		'da-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidelinje"
					$CPChanged = $True
				}
			}

		'de-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Randlinie"
					$CPChanged = $True
				}
			}

		'es-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Línea lateral"
					$CPChanged = $True
				}
			}

		'fi-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sivussa"
					$CPChanged = $True
				}
			}

		'fr-'	{
				If($CoverPage -eq "Sideline")
				{
					If($WordVersion -eq $wdWord2013)
					{
						$CoverPage = "Lignes latérales"
						$CPChanged = $True
					}
					Else
					{
						$CoverPage = "Ligne latérale"
						$CPChanged = $True
					}
				}
			}

		'nb-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidelinje"
					$CPChanged = $True
				}
			}

		'nl-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Terzijde"
					$CPChanged = $True
				}
			}

		'pt-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Linha Lateral"
					$CPChanged = $True
				}
			}

		'sv-'	{
				If($CoverPage -eq "Sideline")
				{
					$CoverPage = "Sidlinje"
					$CPChanged = $True
				}
			}
	}

	If($CPChanged)
	{
		Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
	}

	Write-Verbose "$(Get-Date): Validate cover page"
	[bool]$ValidCP = ValidateCoverPage $WordVersion $CoverPage
	If(!$ValidCP)
	{
		Write-Error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
		AbortScript
	}

	[string]$Title = "DHCP Inventory Report for Server $($DHCPServerName)"

	[string]$filename1 = "$($pwd.path)\$($DHCPServerName)_DHCP_Inventory.docx"

	If($PDF)
	{
		[string]$filename2 = "$($pwd.path)\$($DHCPServerName)_DHCP_Inventory.pdf"
	}
}
ElseIf($Text)
{
	[string]$filename1 = "$($pwd.path)\$($DHCPServerName)_DHCP_Inventory.txt"
}
ElseIf($HTML)
{
	[string]$filename1 = "$($pwd.path)\$($DHCPServerName)_DHCP_Inventory.html"
}

If($MSWord -or $PDF)
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Company Name  : $CompanyName"
	Write-Verbose "$(Get-Date): Cover Page    : $CoverPage"
	Write-Verbose "$(Get-Date): User Name     : $UserName"
	Write-Verbose "$(Get-Date): Save As PDF   : $PDF"
	Write-Verbose "$(Get-Date): Server Name   : $DHCPServerName"
	Write-Verbose "$(Get-Date): Include Leases: $IncludeLeases"
	Write-Verbose "$(Get-Date): Title         : $Title"
	Write-Verbose "$(Get-Date): Filename1     : $filename1"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2     : $filename2"
	}
	Write-Verbose "$(Get-Date): Word version  : $WordProduct"
	Write-Verbose "$(Get-Date): Word language : $($Word.Language)"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Save As Text File"
	Write-Verbose "$(Get-Date): Filename1     : $filename1"
	Write-Verbose "$(Get-Date): Server Name   : $DHCPServerName"
	Write-Verbose "$(Get-Date): Include Leases: $IncludeLeases"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): Save As HTML File"
	Write-Verbose "$(Get-Date): Filename1     : $filename1"
	Write-Verbose "$(Get-Date): Server Name   : $DHCPServerName"
	Write-Verbose "$(Get-Date): Include Leases: $IncludeLeases"
}
Write-Verbose "$(Get-Date): OS Detected   : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture   : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture     : $PSCulture"
Write-Verbose "$(Get-Date): PoSH version  : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start  : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Server Properties and Configuration"
Write-Verbose "$(Get-Date): "

If($MSWord -or $PDF)
{
	$Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	#down to $configlog = $False is from Jeff Hicks
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$word.Templates.LoadBuildingBlocks()
	If($WordVersion -eq $wdWord2007)
	{
		$BuildingBlocks = $word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
	}
	Else
	{
		$BuildingBlocks = $word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
	}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	If($BuildingBlocks -ne $Null)
	{
		$BuildingBlocksExist = $True

		Try 
			{$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)}

		Catch
			{$part = $Null}

		If($part -ne $Null)
		{
			$CoverPagesExist = $True
		}
	}

	If(!$CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Doc = $Word.Documents.Add()
	If($Doc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "An empty Word document could not be created.  Script cannot continue."
		AbortScript
	}

	$Selection = $Word.Selection
	If($Selection -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		Write-Error "An unknown error happened selecting the entire Word document for default formatting options.  Script cannot continue."
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$CurrentGrammarOption = $Word.Options.CheckGrammarAsYouType
	$CurrentSpellingOption = $Word.Options.CheckSpellingAsYouType
	$Word.Options.CheckGrammarAsYouType = $False
	$Word.Options.CheckSpellingAsYouType = $False

	If($CoverPagesExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($selection.Range,$True) | out-null
		$selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
		If($toc -eq $Null)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($selection.Range,$True) | out-null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $doc.Sections.Last.Footers
	ForEach($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	#return focus to main document
	Write-Verbose "$(Get-Date): Return focus to main document"
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	Write-Verbose "$(Get-Date): Move to the end of the current document"
	Write-Verbose "$(Get-Date):"
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	#end of Jeff Hicks 
}

Write-Verbose "$(Get-Date): Getting DHCP server information"
If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "DHCP Server Information"
	WriteWordLine 0 0 "Server name: "$DHCPServerName
}
ElseIf($Text)
{
	Line 0 "DHCP Server Information"
	Line 1 "Server name`t: "$DHCPServerName
}
ElseIf($HTML)
{
}

$DHCPDB = Get-DHCPServerDatabase -ComputerName $DHCPServerName -EA 0

If($? -and $DHCPDB -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Database path: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
		WriteWordLine 0 0 "Backup path: " $DHCPDB.BackupPath
	}
	ElseIf($Text)
	{
		Line 1 "Database path`t: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
		Line 1 "Backup path`t: " $DHCPDB.BackupPath
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving DHCP Server Database information"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving DHCP Server Database information"
	}
	ElseIf($HTML)
	{
	}
}

InsertBlankLine

$DHCPDB = $Null

[bool]$GotServerSettings = $False
$ServerSettings = Get-DHCPServerSetting -ComputerName $DHCPServerName -EA 0

If($? -and $ServerSettings -ne $Null)
{
	$GotServerSettings = $True
	#some properties of $ServerSettings will be needed later
	If($ServerSettings.IsAuthorized)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "DHCP server is authorized"
		}
		ElseIf($Text)
		{
			Line 1 "DHCP server is authorized"
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "DHCP server is not authorized"
		}
		ElseIf($Text)
		{
			Line 1 "DHCP server is not authorized"
		}
		ElseIf($HTML)
		{
		}
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving DHCP Server setting information"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving DHCP Server setting information"
	}
	ElseIf($HTML)
	{
	}
}

InsertBlankLine

Write-Verbose "$(Get-Date): `tGetting IPv4 bindings"
$IPv4Bindings = Get-DHCPServerV4Binding -ComputerName $DHCPServerName -EA 0 | Sort-Object IPAddress

If($? -and $IPv4Bindings -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Connections and server bindings"
	}
	ElseIf($Text)
	{
		Line 0 "Connections and server bindings"
	}
	ElseIf($HTML)
	{
	}
	
	ForEach($IPv4Binding in $IPv4Bindings)
	{
		If($IPv4Binding.BindingState)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "Enabled " -NoNewLine
			}
			ElseIf($Text)
			{
				Line 1 "Enabled " -NoNewLine
			}
			ElseIf($HTML)
			{
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "Disabled " -NoNewLine
			}
			ElseIf($Text)
			{
				Line 1 "Disabled " -NoNewLine
			}
			ElseIf($HTML)
			{
			}
		}
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
		}
		ElseIf($Text)
		{
			Line 0 "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
		}
		ElseIf($HTML)
		{
		}
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 server bindings"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 server bindings"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 server bindings"
	}
	ElseIf($Text)
	{
		Line 1 "There were no IPv4 server bindings"
	}
	ElseIf($HTML)
	{
	}
}
$IPv4Bindings = $Null

InsertBlankLine

Write-Verbose "$(Get-Date): `tGetting IPv6 bindings"
$IPv6Bindings = Get-DHCPServerV6Binding -ComputerName $DHCPServerName -EA 0 | Sort-Object IPAddress

If($? -and $IPv6Bindings -ne $Null)
{
	WriteWordLine 0 0 "Connections and server bindings:"
	ForEach($IPv6Binding in $IPv6Bindings)
	{
		If($IPv6Binding.BindingState)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "Enabled " -NoNewLine
			}
			ElseIf($Text)
			{
				Line 1 "Enabled " -NoNewLine
			}
			ElseIf($HTML)
			{
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "Disabled " -NoNewLine
			}
			ElseIf($Text)
			{
				Line 1 "Disabled " -NoNewLine
			}
			ElseIf($HTML)
			{
			}
		}
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
		}
		ElseIf($Text)
		{
			Line 0 "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
		}
		ElseIf($HTML)
		{
		}
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 server bindings"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv6 server bindings"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv6 server bindings"
	}
	ElseIf($Text)
	{
		Line 1 "There were no IPv6 server bindings"
	}
	ElseIf($HTML)
	{
	}
}
$IPv6Bindings = $Null

If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 2 0 "IPv4"
	WriteWordLine 3 0 "Properties"
}
ElseIf($Text)
{
	Line 0 ""
	Line 0 "IPv4"
	Line 0 "Properties"
}
ElseIf($HTML)
{
}

Write-Verbose "$(Get-Date): Getting IPv4 properties"
Write-Verbose "$(Get-Date): `tGetting IPv4 general settings"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "General"
}
ElseIf($Text)
{
	Line 1 "General"
}
ElseIf($HTML)
{
}

[bool]$GotAuditSettings = $False
$AuditSettings = Get-DHCPServerAuditLog -ComputerName $DHCPServerName -EA 0

If($? -and $AuditSettings -ne $Null)
{
	$GotAuditSettings = $True
	If($AuditSettings.Enable)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "DHCP audit logging is enabled"
		}
		ElseIf($Text)
		{
			Line 2 "DHCP audit logging is enabled"
		}
		ElseIf($HTML)
		{
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "DHCP audit logging is disabled"
		}
		ElseIf($Text)
		{
			Line 2 "DHCP audit logging is disabled"
		}
		ElseIf($HTML)
		{
		}
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving audit log settings"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving audit log settings"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "There were no audit log settings"
	}
	ElseIf($Text)
	{
		Line 0 "There were no audit log settings"
	}
	ElseIf($HTML)
	{
	}
}

#"HKLM:\SYSTEM\CurrentControlSet\Services\DHCPServer\Parameters" "BootFileTable"

#Define the variable to hold the BOOTP Table
$BOOTPKey="SYSTEM\CurrentControlSet\Services\DHCPServer\Parameters" 

#Create an instance of the Registry Object and open the HKLM base key
$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$DHCPServerName) 

#Drill down into the BOOTP key using the OpenSubKey Method
$regkey1=$reg.OpenSubKey($BOOTPKey) 

#Retrieve an array of string that contain all the subkey names
If($regkey1 -ne $Null)
{
	$BOOTPTable = $regkey1.GetValue("BootFileTable") 
}
Else
{
    $BOOTPTable = $Null
}

If($BOOTPTable -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Show the BOOTP table folder is enabled"
	}
	ElseIf($Text)
	{
		Line 2 "Show the BOOTP table folder is enabled"
	}
	ElseIf($HTML)
	{
	}
}

#DNS settings
Write-Verbose "$(Get-Date): `tGetting IPv4 DNS settings"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "DNS"
}
ElseIf($Text)
{
	Line 1 "DNS"
}
ElseIf($HTML)
{
}

$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $DHCPServerName -EA 0
If($? -and $DNSSettings -ne $Null)
{
	GetDNSSettings $DNSSettings "A"
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 DNS Settings for DHCP server $DHCPServerName"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 DNS Settings for DHCP server $DHCPServerName"
	}
	ElseIf($HTML)
	{
	}
}
$DNSSettings = $Null

#now back to some server settings
Write-Verbose "$(Get-Date): `tGetting IPv4 NAP settings"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "Network Access Protection"
}
ElseIf($Text)
{
	Line 1 "Network Access Protection"
}
ElseIf($HTML)
{
}

If($GotServerSettings)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Network Access Protection is " -NoNewLine
		If($ServerSettings.NapEnabled)
		{
			WriteWordLine 0 0 "Enabled on all scopes"
		}
		Else
		{
			WriteWordLine 0 0 "Disabled on all scopes"
		}
		WriteWordLine 0 1 "DHCP server behavior when NPS is unreachable: " -NoNewLine
		Switch($ServerSettings.NpsUnreachableAction)
		{
			"Full"		{WriteWordLine 0 0 "Full Access"}
			"Restricted"	{WriteWordLine 0 0 "Restricted Access"}
			"NoAccess"		{WriteWordLine 0 0 "Drop Client Packet"}
			Default		{WriteWordLine 0 0 "Unable to determine NPS unreachable action: $($ServerSettings.NpsUnreachableAction)"}
		}
	}
	ElseIf($Text)
	{
		Line 2 "Network Access Protection is " -NoNewLine
		If($ServerSettings.NapEnabled)
		{
			Line 0 "Enabled on all scopes"
		}
		Else
		{
			Line 0 "Disabled on all scopes"
		}
		Line 2 "DHCP server behavior when NPS is unreachable: " -NoNewLine
		Switch($ServerSettings.NpsUnreachableAction)
		{
			"Full"		{Line 0 "Full Access"}
			"Restricted"	{Line 0 "Restricted Access"}
			"NoAccess"		{Line 0 "Drop Client Packet"}
			Default		{Line 0 "Unable to determine NPS unreachable action: $($ServerSettings.NpsUnreachableAction)"}
		}
	}
	ElseIf($HTML)
	{
	}
}

#filters
Write-Verbose "$(Get-Date): `tGetting IPv4 filters"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "Filters"
}
ElseIf($Text)
{
	Line 1 "Filters"
}
ElseIf($HTML)
{
}

$MACFilters = Get-DHCPServerV4FilterList -ComputerName $DHCPServerName -EA 0

If($? -and $MACFilters -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Enable Allow list: " -NoNewLine
		If($MACFilters.Allow)
		{
			WriteWordLine 0 0 "Enabled"
		}
		Else
		{
			WriteWordLine 0 0 "Disabled"
		}
		WriteWordLine 0 1 "Enable Deny list: " -NoNewLine
		If($MACFilters.Deny)
		{
			WriteWordLine 0 0 "Enabled"
		}
		Else
		{
			WriteWordLine 0 0 "Disabled"
		}
	}
	ElseIf($Text)
	{
		Line 2 "Enable Allow list`t: " -NoNewLine
		If($MACFilters.Allow)
		{
			Line "Enabled"
		}
		Else
		{
			Line 0 "Disabled"
		}
		Line 2 "Enable Deny list`t: " -NoNewLine
		If($MACFilters.Deny)
		{
			Line 0 "Enabled"
		}
		Else
		{
			Line 0 "Disabled"
		}
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving MAC filters for DHCP server $DHCPServerName"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving MAC filters for DHCP server $DHCPServerName"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no MAC filters for DHCP server $DHCPServerName"
	}
	ElseIf($Text)
	{
		Line 2 "There were no MAC filters for DHCP server $DHCPServerName"
	}
	ElseIf($HTML)
	{
	}
}
$MACFilters = $Null

#failover
Write-Verbose "$(Get-Date): `tGetting IPv4 Failover"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "Failover"
}
ElseIf($Text)
{
	Line 1 "Failover"
}
ElseIf($HTML)
{
}

$Failovers = Get-DHCPServerV4Failover -ComputerName $DHCPServerName -EA 0

If($? -and $Failovers -ne $Null)
{
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Failovers -is [array])
		{
			[int]$Rows = ($Failovers.Count * 8) - 1
			#subtract the very last row used for spacing
		}
		Else
		{
			[int]$Rows = 7
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Relationship name"
			$Table.Cell($xRow,2).Range.Text = $Failover.Name
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State of the server"
			Switch($Failover.State)
			{
				"NoState" 
				{
					$Table.Cell($xRow,2).Range.Text = "No State"
				}
				"Normal" 
				{
					$Table.Cell($xRow,2).Range.Text = "Normal"
				}
				"Init" 
				{
					$Table.Cell($xRow,2).Range.Text = "Iitializing"
				}
				"CommunicationInterrupted" 
				{
					$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
				}
				"PartnerDown" 
				{
					$Table.Cell($xRow,2).Range.Text = "Normal"
				}
				"PotentialConflict" 
				{
					$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
				}
				"Startup" 
				{
					$Table.Cell($xRow,2).Range.Text = "Starting Up"
				}
				"ResolutionInterrupted" 
				{
					$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
				}
				"ConflictDone" 
				{
					$Table.Cell($xRow,2).Range.Text = "Conflict Done"
				}
				"Recover" 
				{
					$Table.Cell($xRow,2).Range.Text = "Recover"
				}
				"RecoverWait" 
				{
					$Table.Cell($xRow,2).Range.Text = "Recover Wait"
				}
				"RecoverDone" 
				{
					$Table.Cell($xRow,2).Range.Text = "Recover Done"
				}
				Default 
				{
					$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
				}
			}
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Partner Server"
			$Table.Cell($xRow,2).Range.Text = $Failover.PartnerServer
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Mode"
			$Table.Cell($xRow,2).Range.Text = $Failover.Mode
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Message Authentication"
			If($Failover.EnableAuth)
			{
				$Table.Cell($xRow,2).Range.Text = "Enabled"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "Disabled"
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Local server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.LoadBalancePercent)%"
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Partner Server"
				$tmp = (100 - $($Failover.LoadBalancePercent))
				$Table.Cell($xRow,2).Range.Text = "$($tmp)%"
				$tmp = $Null
			}
			Else
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Role of this server"
				$Table.Cell($xRow,2).Range.Text = $Failover.ServerRole
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Addresses reserved for standby server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			$xRow++
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			Line 2 "Relationship name: " $Failover.Name
					
			Line 2 "State of the server`t: " -NoNewLine
			Switch($Failover.State)
			{
				"NoState" 
				{
					Line 0 "No State"
				}
				"Normal" 
				{
					Line 0 "Normal"
				}
				"Init" 
				{
					Line 0 "Iitializing"
				}
				"CommunicationInterrupted" 
				{
					Line 0 "Communication Interrupted"
				}
				"PartnerDown" 
				{
					Line 0 "Normal"
				}
				"PotentialConflict" 
				{
					Line 0 "Potential Conflict"
				}
				"Startup" 
				{
					Line 0 "Starting Up"
				}
				"ResolutionInterrupted" 
				{
					Line 0 "Resolution Interrupted"
				}
				"ConflictDone" 
				{
					Line 0 "Conflict Done"
				}
				"Recover" 
				{
					Line 0 "Recover"
				}
				"RecoverWait" 
				{
					Line 0 "Recover Wait"
				}
				"RecoverDone" 
				{
					Line 0 "Recover Done"
				}
				Default 
				{
					Line 0 "Unable to determine server state"
				}
			}
					
			Line 2 "Partner Server`t`t: " $Failover.PartnerServer
			Line 2 "Mode`t`t`t: " $Failover.Mode
			Line 2 "Message Authentication`t: " -NoNewLine
			If($Failover.EnableAuth)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				$tmp = (100 - $($Failover.LoadBalancePercent))
				Line 2 "Local server`t`t: $($Failover.LoadBalancePercent)%"
				Line 2 "Partner Server`t`t: $($tmp)%"
				$tmp = $Null
			}
			Else
			{
				Line 2 "Role of this server`t: " $Failover.ServerRole
				Line 2 "Addresses reserved for standby server: $($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There was no Failover configured for DHCP server $DHCPServerName"
	}
	ElseIf($Text)
	{
		Line 2 "There was no Failover configured for DHCP server $DHCPServerName"
	}
	ElseIf($HTML)
	{
	}
}
$Failovers = $Null

#Advanced
Write-Verbose "$(Get-Date): `tGetting IPv4 advanced settings"
If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "Advanced"
}
ElseIf($Text)
{
	Line 1 "Advanced"
}
ElseIf($HTML)
{
}

If($GotServerSettings)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Conflict detection attempts: " $ServerSettings.ConflictDetectionAttempts
	}
	ElseIf($Text)
	{
		Line 2 "Conflict detection attempts`t: " $ServerSettings.ConflictDetectionAttempts
	}
	ElseIf($HTML)
	{
	}
}

If($GotAuditSettings)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "Audit log file path: " $AuditSettings.Path
	}
	ElseIf($Text)
	{
		Line 2 "Audit log file path`t`t: " $AuditSettings.Path
	}
	ElseIf($HTML)
	{
	}
}

If($MSWord -or $PDF)
{
	WriteWordLine 4 0 "Statistics"
}
ElseIf($Text)
{
	Line 1 "Statistics"
}
ElseIf($HTML)
{
}

$Statistics = Get-DHCPServerV4Statistics -ComputerName $DHCPServerName -EA 0

If($? -and $Statistics -ne $Null)
{
	$UpTime = $(Get-Date) - $Statistics.ServerStartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
		$UpTime.Days, `
		$UpTime.Hours, `
		$UpTime.Minutes, `
		$UpTime.Seconds)
	[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
	[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable
	
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 16
		Write-Verbose "$(Get-Date): `tAdd IPv4 statistics table to doc"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Description"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Details"

		$Table.Cell(2,1).Range.Text = "Start Time"
		$Table.Cell(2,2).Range.Text = $Statistics.ServerStartTime
		$Table.Cell(3,1).Range.Text = "Up Time"
		$Table.Cell(3,2).Range.Text = $Str
		$Table.Cell(4,1).Range.Text = "Discovers"
		$Table.Cell(4,2).Range.Text = $Statistics.Discovers
		$Table.Cell(5,1).Range.Text = "Offers"
		$Table.Cell(5,2).Range.Text = $Statistics.Offers
		$Table.Cell(6,1).Range.Text = "Delayed Offers"
		$Table.Cell(6,2).Range.Text = $Statistics.DelayedOffers
		$Table.Cell(7,1).Range.Text = "Requests"
		$Table.Cell(7,2).Range.Text = $Statistics.Requests
		$Table.Cell(8,1).Range.Text = "Acks"
		$Table.Cell(8,2).Range.Text = $Statistics.Acks
		$Table.Cell(9,1).Range.Text = "Nacks"
		$Table.Cell(9,2).Range.Text = $Statistics.Naks
		$Table.Cell(10,1).Range.Text = "Declines"
		$Table.Cell(10,2).Range.Text = $Statistics.Declines
		$Table.Cell(11,1).Range.Text = "Releases"
		$Table.Cell(11,2).Range.Text = $Statistics.Releases
		$Table.Cell(12,1).Range.Text = "Total Scopes"
		$Table.Cell(12,2).Range.Text = $Statistics.TotalScopes
		$Table.Cell(13,1).Range.Text = "Scopes with delay configured"
		$Table.Cell(13,2).Range.Text = $Statistics.ScopesWithDelayConfigured
		$Table.Cell(14,1).Range.Text = "Total Addresses"
		$Table.Cell(14,2).Range.Text = "{0:N0}" -f $Statistics.TotalAddresses
		$Table.Cell(15,1).Range.Text = "In Use"
		$Table.Cell(15,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent))%"
		$Table.Cell(16,1).Range.Text = "Available"
		$Table.Cell(16,2).Range.Text = "{0:N0}" -f "$($Statistics.AddressesAvailable) ($($AvailablePercent))%"

		$table.AutoFitBehavior(1)

		#return focus back to document
		Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		Write-Verbose "$(Get-Date): Finished IPv4 statistics table"
		Write-Verbose "$(Get-Date): "
	}
	ElseIf($Text)
	{
		Line 2 "Description" -NoNewLine
		Line 3 "Details"

		Line 2 "Start Time:" -NoNewLine
		Line 3 $Statistics.ServerStartTime
		Line 2 "Up Time:" -NoNewLine
		Line 3 $Str
		Line 2 "Discovers:" -NoNewLine
		Line 3 $Statistics.Discovers
		Line 2 "Offers:" -NoNewLine
		Line 4 $Statistics.Offers
		Line 2 "Delayed Offers:" -NoNewLine
		Line 3 $Statistics.DelayedOffers
		Line 2 "Requests:" -NoNewLine
		Line 3 $Statistics.Requests
		Line 2 "Acks:" -NoNewLine
		Line 4 $Statistics.Acks
		Line 2 "Nacks:" -NoNewLine
		Line 4 $Statistics.Naks
		Line 2 "Declines:" -NoNewLine
		Line 3 $Statistics.Declines
		Line 2 "Releases:" -NoNewLine
		Line 3 $Statistics.Releases
		Line 2 "Total Scopes:" -NoNewLine
		Line 3 $Statistics.TotalScopes
		Line 2 "Scopes w/delay configured:" -NoNewLine
		Line 1 $Statistics.ScopesWithDelayConfigured
		Line 2 "Total Addresses:" -NoNewLine
		$tmp = "{0:N0}" -f $Statistics.TotalAddresses
		Line 2 $tmp
		Line 2 "In Use:" -NoNewLine
		Line 4 "$($Statistics.AddressesInUse) ($($InUsePercent))%"
		Line 2 "Available:" -NoNewLine
		$tmp = "{0:N0}" -f $Statistics.AddressesAvailable
		Line 3 "$($tmp) ($($AvailablePercent))%"
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 statistics"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 statistics"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "There were no IPv4 statistics"
	}
	ElseIf($Text)
	{
		Line 0 "There were no IPv4 statistics"
	}
	ElseIf($HTML)
	{
	}
}
$Statistics = $Null

Write-Verbose "$(Get-Date):	Build array of Allow/Deny filters"
$Filters = Get-DHCPServerV4Filter -ComputerName $DHCPServerName -EA 0

Write-Verbose "$(Get-Date): Getting IPv4 Superscopes"
$IPv4Superscopes = Get-DHCPServerV4Superscope -ComputerName $DHCPServerName -EA 0

If($? -and $IPv4Superscopes -ne $Null)
{
	ForEach($IPv4Superscope in $IPv4Superscopes)
	{
		If(![string]::IsNullOrEmpty($IPv4Superscope.SuperscopeName))
		{
			If($MSWord -or $PDF)
			{
				#put each superscope on a new page
				$selection.InsertNewPage()
				Write-Verbose "$(Get-Date): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
				WriteWordLine 3 0 "Superscope [$($IPv4Superscope.SuperscopeName)]"

				#get superscope statistics first
				$Statistics = Get-DHCPServerV4SuperscopeStatistics -ComputerName $DHCPServerName -Name $IPv4Superscope.SuperscopeName -EA 0

				If($? -and $Statistics -ne $Null)
				{
					GetShortStatistics $Statistics
				}
				ElseIf(!$?)
				{
					WriteWordLine 0 0 "Error retrieving superscope statistics"
				}
				Else
				{
					WriteWordLine 0 1 "There were no statistics for the superscope"
				}
				$Statistics = $Null
			
				$xScopeIds = $IPv4Superscope.ScopeId
				[int]$StartLevel = 4
				ForEach($xScopeId in $xScopeIds)
				{
					Write-Verbose "$(Get-Date): Processing scope id $($xScopeId) for Superscope $($IPv4Superscope.SuperscopeName)"
					$IPv4Scope = Get-DHCPServerV4Scope -ComputerName $DHCPServerName -ScopeId $xScopeId -EA 0
					
					If($? -and $IPv4Scope -ne $Null)
					{
						GetIPv4ScopeData $IPv4Scope $StartLevel
					}
					Else
					{
						WriteWordLine 0 0 "Error retrieving Superscope data for scope $($xScopeId)"
					}
				}
			}
			ElseIf($Text)
			{
				Write-Verbose "$(Get-Date): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
				Line 0 ""
				Line 0 "Superscope [$($IPv4Superscope.SuperscopeName)]"

				#get superscope statistics first
				$Statistics = Get-DHCPServerV4SuperscopeStatistics -ComputerName $DHCPServerName -Name $IPv4Superscope.SuperscopeName -EA 0

				If($? -and $Statistics -ne $Null)
				{
					Line 1 "Statistics:"
					GetShortStatistics $Statistics
				}
				ElseIf(!$?)
				{
					Line 0 "Error retrieving superscope statistics"
				}
				Else
				{
					Line 2 "There were no statistics for the superscope"
				}
				$Statistics = $Null
			
				$xScopeIds = $IPv4Superscope.ScopeId
				[int]$StartLevel = 4
				ForEach($xScopeId in $xScopeIds)
				{
					Write-Verbose "$(Get-Date): Processing scope id $($xScopeId) for Superscope $($IPv4Superscope.SuperscopeName)"
					$IPv4Scope = Get-DHCPServerV4Scope -ComputerName $DHCPServerName -ScopeId $xScopeId -EA 0
					
					If($? -and $IPv4Scope -ne $Null)
					{
						GetIPv4ScopeData $IPv4Scope $StartLevel
					}
					Else
					{
						Line 0 "Error retrieving Superscope data for scope $($xScopeId)"
					}
				}
			}
			ElseIf($HTML)
			{
			}
		}
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 Superscopes"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 Superscopes"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 Superscopes"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 Superscopes"
	}
	ElseIf($HTML)
	{
	}
}
$IPv4Superscopes = $Null

Write-Verbose "$(Get-Date): Getting IPv4 scopes"
$IPv4Scopes = Get-DHCPServerV4Scope -ComputerName $DHCPServerName -EA 0

If($? -and $IPv4Scopes -ne $Null)
{
	[int]$StartLevel = 3
	ForEach($IPv4Scope in $IPv4Scopes)
	{
		GetIPv4ScopeData $IPv4Scope $StartLevel
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 scopes"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 scopes"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 scopes"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 scopes"
	}
	ElseIf($HTML)
	{
	}
}
$IPv4Scopes = $Null
$Filters = $Null

Write-Verbose "$(Get-Date): Getting IPv4 Multicast scopes"
$IPv4MulticastScopes = Get-DHCPServerV4MulticastScope -ComputerName $DHCPServerName -EA 0

If($? -and $IPv4MulticastScopes -ne $Null)
{
	ForEach($IPv4MulticastScope in $IPv4MulticastScopes)
	{
		If($IPv4MulticastScope.LeaseDuration -ne $Null)
		{
			$DurationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
				$IPv4MulticastScope.LeaseDuration.Days, `
				$IPv4MulticastScope.LeaseDuration.Hours, `
				$IPv4MulticastScope.LeaseDuration.Minutes)
		}
		Else
		{
			$DurationStr = "Unlimited"
		}
		
		If($MSWord -or $PDF)
		{
			#put each scope on a new page
			$selection.InsertNewPage()
			Write-Verbose "$(Get-Date): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
			WriteWordLine 3 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
			WriteWordLine 4 0 "General"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = 6
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			$Table.Cell(1,1).Range.Text = "Name"
			$Table.Cell(1,2).Range.Text = $IPv4MulticastScope.Name
			$Table.Cell(2,1).Range.Text = "Start IP address"
			$Table.Cell(2,2).Range.Text = $IPv4MulticastScope.StartRange
			$Table.Cell(3,1).Range.Text = "End IP address"
			$Table.Cell(3,2).Range.Text = $IPv4MulticastScope.EndRange
			$Table.Cell(4,1).Range.Text = "Time to live"
			$Table.Cell(4,2).Range.Text = $IPv4MulticastScope.Ttl
			$Table.Cell(5,1).Range.Text = "Lease duration"
			$Table.Cell(5,2).Range.Text = $DurationStr
			$Table.Cell(6,1).Range.Text = "Description"
			$Table.Cell(6,2).Range.Text = $IPv4MulticastScope.Description
			
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null

			WriteWordLine 4 0 "Lifetime"
			WriteWordLine 0 1 "Multicast scope lifetime: " -NoNewLine
			If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
			{
				WriteWordLine 0 0 "Infinite"
			}
			Else
			{
				WriteWordLine 0 0 "Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
			}
			
			Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
			WriteWordLine 4 0 "Exclusions"
			$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
			If($? -and $Exclusions -ne $Null)
			{
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 2
				If($Exclusions -is [array])
				{
					[int]$Rows = $Exclusions.Count + 1
				}
				Else
				{
					[int]$Rows = 2
				}
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = $myHash.Word_TableGrid
				$table.Borders.InsideLineStyle = 0
				$table.Borders.OutsideLineStyle = 0
				$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Start IP Address"
				$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,2).Range.Font.Bold = $True
				$Table.Cell(1,2).Range.Text = "End IP Address"
				[int]$xRow = 1
				ForEach($Exclusion in $Exclusions)
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = $Exclusion.StartRange
					$Table.Cell($xRow,2).Range.Text = $Exclusion.EndRange 
				}
				$Table.Rows.SetLeftIndent(36,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
			}
			ElseIf(!$?)
			{
				WriteWordLine 0 0 "Error retrieving exclusions for multicast scope"
			}
			Else
			{
				WriteWordLine 0 1 "<None>"
			}
			
			#leases
			If($IncludeLeases)
			{
				Write-Verbose "$(Get-Date):	`t`tGetting leases"
				
				WriteWordLine 4 0 "Address Leases"
				$Leases = Get-DHCPServerV4MulticastLease -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
				If($? -and $Leases -ne $Null)
				{
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					If($Leases -is [array])
					{
						[int]$Rows = ($Leases.Count * 7) - 1
						#subtract the very last row used for spacing
					}
					Else
					{
						[int]$Rows = 6
					}
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$table.Style = $myHash.Word_TableGrid
					$table.Borders.InsideLineStyle = 0
					$table.Borders.OutsideLineStyle = 0
					[int]$xRow = 0
					ForEach($Lease in $Leases)
					{
						Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
						If($Lease.LeaseExpiryTime -ne $Null)
						{
							$LeaseEndStr = [string]::format("{0} days, {1} hours, {2} minutes", `
								$Lease.LeaseExpiryTime.Days, `
								$Lease.LeaseExpiryTime.Hours, `
								$Lease.LeaseExpiryTime.Minutes)
						}
						Else
						{
							$LeaseEndStr = ""
						}

						If($Lease.LeaseExpiryTime -ne $Null)
						{
							$LeaseStartStr = [string]::format("{0} days, {1} hours, {2} minutes", `
								$Lease.LeaseStartTime.Days, `
								$Lease.LeaseStartTime.Hours, `
								$Lease.LeaseStartTime.Minutes)
						}
						Else
						{
							$LeaseStartStr = ""
						}

						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Client IP address"
						$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress
						
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Range.Text = $Lease.HostName
						
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
						If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
						{
							$Table.Cell($xRow,2).Range.Text = "Unlimited"
						}
						Else
						{
							$Table.Cell($xRow,2).Range.Text = $LeaseEndStr
						}
						
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Lease Start"
						If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
						{
							$Table.Cell($xRow,2).Range.Text = "Unlimited"
						}
						Else
						{
							$Table.Cell($xRow,2).Range.Text = $LeaseStartStr
						}
						
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "Address State"
						$Table.Cell($xRow,2).Range.Text = $Lease.AddressState
						
						$xRow++
						$Table.Cell($xRow,1).Range.Text = "MAC address"
						$Table.Cell($xRow,2).Range.Text = $Lease.ClientID
						
						#skip a row for spacing
						$xRow++
					}
					$Table.Rows.SetLeftIndent(36,1)
					$table.AutoFitBehavior(1)

					#return focus back to document
					$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$selection.EndKey($wdStory,$wdMove) | Out-Null
				}
				ElseIf(!$?)
				{
					WriteWordLine 0 0 "Error retrieving leases for scope"
				}
				Else
				{
					WriteWordLine 0 1 "<None>"
				}
				$Leases = $Null
			}
			
			Write-Verbose "$(Get-Date):	`t`tGetting Multicast Scope statistics"
			WriteWordLine 4 0 "Statistics"

			$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

			If($? -and $Statistics -ne $Null)
			{
				GetShortStatistics $Statistics
			}
			ElseIf(!$?)
			{
				WriteWordLine 0 0 "Error retrieving multicast scope statistics"
			}
			Else
			{
				WriteWordLine 0 1 "<None>"
			}
			$Statistics = $Null
		}
		ElseIf($Text)
		{
			Write-Verbose "$(Get-Date): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
			Line 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
			Line 1 "General:"
			Line 2 "Name`t`t`t: " $IPv4MulticastScope.Name
			Line 2 "Start IP address`t: " $IPv4MulticastScope.StartRange
			Line 2 "End IP address`t`t: " $IPv4MulticastScope.EndRange
			Line 2 "Time to live`t`t: " $IPv4MulticastScope.Ttl
			Line 2 "Lease duration`t`t: " $DurationStr
			If(![string]::IsNullOrEmpty($IPv4MulticastScope.Description))
			{
				Line 2 "Description`t`t: " $IPv4MulticastScope.Description
			}
			
			Line 1 "Lifetime:"
			Line 2 "Multicast scope lifetime: " -NoNewLine
			If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
			{
				Line 0 "Infinite"
			}
			Else
			{
				Line 0 "Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
			}
			
			Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
			Line 1 "Exclusions:"
			$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
			If($? -and $Exclusions -ne $Null)
			{
				Line 2 "Start IP Address`tEnd IP Address"
				ForEach($Exclusion in $Exclusions)
				{
					Line 2 $Exclusion.StartRange -NoNewLine
					Line 2 $Exclusion.EndRange 
				}
			}
			ElseIf(!$?)
			{
				Line 0 "Error retrieving exclusions for multicast scope"
			}
			Else
			{
				Line 2 "<None>"
			}
			
			#leases
			If($IncludeLeases)
			{
				Write-Verbose "$(Get-Date):	`t`tGetting leases"
				
				Line 1 "Address Leases:"
				$Leases = Get-DHCPServerV4MulticastLease -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
				If($? -and $Leases -ne $Null)
				{
					ForEach($Lease in $Leases)
					{
						Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
						If($Lease.LeaseExpiryTime -ne $Null)
						{
							$LeaseEndStr = [string]::format("{0} days, {1} hours, {2} minutes", `
								$Lease.LeaseExpiryTime.Days, `
								$Lease.LeaseExpiryTime.Hours, `
								$Lease.LeaseExpiryTime.Minutes)
						}
						Else
						{
							$LeaseEndStr = ""
						}

						If($Lease.LeaseExpiryTime -ne $Null)
						{
							$LeaseStartStr = [string]::format("{0} days, {1} hours, {2} minutes", `
								$Lease.LeaseStartTime.Days, `
								$Lease.LeaseStartTime.Hours, `
								$Lease.LeaseStartTime.Minutes)
						}
						Else
						{
							$LeaseStartStr = ""
						}

						Line 2 "Client IP address`t: " $Lease.IPAddress
						Line 2 "Name`t`t`t: " $Lease.HostName
						Line 2 "Lease Expiration`t: " -NoNewLine
						If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
						{
							Line 0 "Unlimited"
						}
						Else
						{
							Line 0 $LeaseEndStr
						}
						
						Line 2 "Lease Start`t`t: " -NoNewLine
						If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
						{
							Line 0 "Unlimited"
						}
						Else
						{
							Line 0 $LeaseStartStr
						}
						
						Line 2 "Address State`t`t: " $Lease.AddressState
						Line 2 "MAC address`t: " $Lease.ClientID
						
						#skip a row for spacing
						Line 0 ""
					}
				}
				ElseIf(!$?)
				{
					Line 0 "Error retrieving leases for scope"
				}
				Else
				{
					Line 2 "<None>"
				}
				$Leases = $Null
			}
			
			Write-Verbose "$(Get-Date):	`t`tGetting Multicast Scope statistics"
			Line 1 "Statistics:"

			$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

			If($? -and $Statistics -ne $Null)
			{
				GetShortStatistics $Statistics
			}
			ElseIf(!$?)
			{
				Line 0 "Error retrieving multicast scope statistics"
			}
			Else
			{
				Line 2 "<None>"
			}
			$Statistics = $Null
		}
		ElseIf($HTML)
		{
		}
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 Multicast scopes"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 Multicast scopes"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 Multicast scopes"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 Multicast scopes"
	}
	ElseIf($HTML)
	{
	}
}
$IPv4MulticastScopes = $Null

#bootp table
If($BOOTPTable -ne $Null)
{
	Write-Verbose "$(Get-Date):	IPv4 BOOTP Table"
	
	If($MSWord -or $PDF)
	{

		$selection.InsertNewPage()
		WriteWordLine 3 0 "BOOTP Table"
		
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 3
		If($BOOTPTable -is [array])
		{
			[int]$Rows = $BOOTPTable.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Boot Image"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "File Name"
		$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,3).Range.Font.Bold = $True
		$Table.Cell(1,3).Range.Text = "File Server"
		[int]$xRow = 1
		ForEach($Item in $BOOTPTable)
		{
			$xRow++
			$ItemParts = $Item.Split(",")
			$Table.Cell($xRow,1).Range.Text = $ItemParts[0]
			$Table.Cell($xRow,2).Range.Text = $ItemParts[1] 
			$Table.Cell($xRow,3).Range.Text = $ItemParts[2] 
		}
		
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		Line 1 "BOOTP Table"
		
		ForEach($Item in $BOOTPTable)
		{
			$ItemParts = $Item.Split(",")
			Line 2 "Boot Image`t: " $ItemParts[0]
			Line 2 "File Name`t: " $ItemParts[1]
			Line 2 "FIle Server`t: " $ItemParts[2] 
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
	}
}

#Server Options
Write-Verbose "$(Get-Date):	Getting IPv4 server options"

If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 3 0 "Server Options"
}
ElseIf($Text)
{
	Line 1 "Server Options"
}
ElseIf($HTML)
{
}

$ServerOptions = Get-DHCPServerV4OptionValue -All -ComputerName $DHCPServerName -EA 0 | Sort-Object OptionId

If($? -and $ServerOptions -ne $Null)
{
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ServerOptions -is [array])
		{
			[int]$Rows = ($ServerOptions.Count * 5) - 5 
			#subtract option 81
		}
		Else
		{
			[int]$Rows = 4
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($ServerOption in $ServerOptions)
		{
			If($ServerOption.OptionId -ne 81)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ServerOption.Name)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Option Name"
				$Table.Cell($xRow,2).Range.Text = "$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Vendor"
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					$Table.Cell($xRow,2).Range.Text = "Standard"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ServerOption.VendorClass
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Value"
				$Table.Cell($xRow,2).Range.Text = $ServerOption.Value

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy Name"
				If([string]::IsNullOrEmpty($ServerOption.PolicyName))
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ServerOption.PolicyName
				}
				#for spacing
				$xRow++
			}
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		ForEach($ServerOption in $ServerOptions)
		{
			If($ServerOption.OptionId -ne 81)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ServerOption.Name)"
				Line 2 "Option Name`t: $($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"
				Line 2 "Vendor`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					Line 0 "Standard"
				}
				Else
				{
					Line 0 $ServerOption.VendorClass
				}
				
				Line 2 "Value`t`t: " $ServerOption.Value
				Line 2 "Policy Name`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ServerOption.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $ServerOption.PolicyName
				}
				#for spacing
				Line 0 ""
			}
		}
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 server options"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 server options"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 server options"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 server options"
	}
	ElseIf($HTML)
	{
	}
}
$ServerOptions = $Null

#Policies
Write-Verbose "$(Get-Date):	Getting IPv4 policies"
If($MSWord -or $PDF)
{
	WriteWordLine 3 0 "Policies"
}
ElseIf($Text)
{
	Line 1 "Policies"
}
ElseIf($HTML)
{
}

$Policies = Get-DHCPServerV4Policy -ComputerName $DHCPServerName -EA 0 | Sort-Object ProcessingOrder

If($? -and $Policies -ne $Null)
{
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Policies -is [array])
		{
			[int]$Rows = $Policies.Count * 6
		}
		Else
		{
			[int]$Rows = 5
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($Policy in $Policies)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Policy Name"
			$Table.Cell($xRow,2).Range.Text = $Policy.Name
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Description"
			$Table.Cell($xRow,2).Range.Text = $Policy.Description

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Processing Order"
			$Table.Cell($xRow,2).Range.Text = $Policy.ProcessingOrder

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Level"
			$Table.Cell($xRow,2).Range.Text = "Server"

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State"
			If($Policy.Enabled)
			{
				$Table.Cell($xRow,2).Range.Text = "Enabled"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "Disabled"
			}
			#for spacing
			$xRow++
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}
	}
	ElseIf($Text)
	{
		ForEach($Policy in $Policies)
		{
			Line 2 "Policy Name`t`t: " $Policy.Name
			If(![string]::IsNullOrEmpty($Policy.Description))
			{
				Line 2 "Description`t`t: " $Policy.Description
			}
			Line 2 "Processing Order`t: " $Policy.ProcessingOrder
			Line 2 "Level`t`t`t: Server"
			Line 2 "State`t`t`t: " -NoNewLine
			If($Policy.Enabled)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
			#for spacing
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 policies"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 policies"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 policies"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 policies"
	}
	ElseIf($HTML)
	{
	}
}
$Policies = $Null

#Filters
Write-Verbose "$(Get-Date):	Getting IPv4 filters"
If($MSWord -or $PDF)
{
	WriteWordLine 3 0 "Filters"
}
ElseIf($Text)
{
	Line 1 "Filters"
}
ElseIf($HTML)
{
}

Write-Verbose "$(Get-Date):	`tAllow filters"
$AllowFilters = Get-DHCPServerV4Filter -List Allow -ComputerName $DHCPServerName -EA 0 | Sort-Object MacAddress

If($? -and $AllowFilters -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Allow"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($AllowFilters -is [array])
		{
			[int]$Rows = $AllowFilters.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "MAC Address"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Description"
		[int]$xRow = 1
		ForEach($AllowFilter in $AllowFilters)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $AllowFilter.MacAddress
			$Table.Cell($xRow,2).Range.Text = $AllowFilter.Description
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		Line 2 "Allow"
		ForEach($AllowFilter in $AllowFilters)
		{
			Line 3 "MAC Address`t: " $AllowFilter.MacAddress
			Line 3 "Description`t: " $AllowFilter.Description
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 allow filters"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 allow filters"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 allow filters"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 allow filters"
	}
	ElseIf($HTML)
	{
	}
}
$AllowFilters = $Null

Write-Verbose "$(Get-Date):	`tDeny filters"
$DenyFilters = Get-DHCPServerV4Filter -List Deny -ComputerName $DHCPServerName -EA 0 | Sort-Object MacAddress
If($? -and $DenyFilters -ne $Null)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Deny"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($DenyFilters -is [array])
		{
			[int]$Rows = $DenyFilters.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "MAC Address"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Description"
		[int]$xRow = 1

		ForEach($DenyFilter in $DenyFilters)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $DenyFilter.MacAddress
			$Table.Cell($xRow,2).Range.Text = $DenyFilter.Description
		}
		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf($Text)
	{
		Line 2 "Deny"
		ForEach($DenyFilter in $DenyFilters)
		{
			Line 3 "MAC Address`t: " $DenyFilter.MacAddress
			Line 3 "Description`t: " $DenyFilter.Description
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
	}
}
ElseIf(!$?)
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 "Error retrieving IPv4 deny filters"
	}
	ElseIf($Text)
	{
		Line 0 "Error retrieving IPv4 deny filters"
	}
	ElseIf($HTML)
	{
	}
}
Else
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 "There were no IPv4 deny filters"
	}
	ElseIf($Text)
	{
		Line 2 "There were no IPv4 deny filters"
	}
	ElseIf($HTML)
	{
	}
}
$DenyFilters = $Null

#IPv6

If($MSWord -or $PDF)
{
	$selection.InsertNewPage()
	WriteWordLine 2 0 "IPv6"
	WriteWordLine 3 0 "Properties"

	Write-Verbose "$(Get-Date): Getting IPv6 properties"
	Write-Verbose "$(Get-Date): `tGetting IPv6 general settings"
	WriteWordLine 4 0 "General"

	If($GotAuditSettings)
	{
		If($AuditSettings.Enable)
		{
			WriteWordLine 0 1 "DHCP audit logging is enabled"
		}
		Else
		{
			WriteWordLine 0 1 "DHCP audit logging is disabled"
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving audit log settings"
	}
	Else
	{
		WriteWordLine 0 1 "There were no audit log settings"
	}

	#DNS settings
	Write-Verbose "$(Get-Date): `tGetting IPv6 DNS settings"
	WriteWordLine 4 0 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for DHCP server $DHCPServerName"
	}
	$DNSSettings = $Null

	#Advanced
	Write-Verbose "$(Get-Date): `tGetting IPv6 advanced settings"
	WriteWordLine 4 0 "Advanced"
	If($GotAuditSettings)
	{
		WriteWordLine 0 1 "Audit log file path " $AuditSettings.Path
	}
	$AuditSettings = $Null

	WriteWordLine 4 0 "Statistics"
	$Statistics = Get-DHCPServerV6Statistics -ComputerName $DHCPServerName -EA 0

	If($? -and $Statistics -ne $Null)
	{
		$UpTime = $(Get-Date) - $Statistics.ServerStartTime
		$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
			$UpTime.Days, `
			$UpTime.Hours, `
			$UpTime.Minutes, `
			$UpTime.Seconds)

		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 16
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Description"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Details"

		$Table.Cell(2,1).Range.Text = "Start Time"
		$Table.Cell(2,2).Range.Text = $Statistics.ServerStartTime
		$Table.Cell(3,1).Range.Text = "Up Time"
		$Table.Cell(3,2).Range.Text = $Str
		$Table.Cell(4,1).Range.Text = "Solicits"
		$Table.Cell(4,2).Range.Text = $Statistics.Solicits
		$Table.Cell(5,1).Range.Text = "Advertises"
		$Table.Cell(5,2).Range.Text = $Statistics.Advertises
		$Table.Cell(6,1).Range.Text = "Requests"
		$Table.Cell(6,2).Range.Text = $Statistics.Requests
		$Table.Cell(7,1).Range.Text = "Replies"
		$Table.Cell(7,2).Range.Text = $Statistics.Replies
		$Table.Cell(8,1).Range.Text = "Renews"
		$Table.Cell(8,2).Range.Text = $Statistics.Renews
		$Table.Cell(9,1).Range.Text = "Rebinds"
		$Table.Cell(9,2).Range.Text = $Statistics.Rebinds
		$Table.Cell(10,1).Range.Text = "Confirms"
		$Table.Cell(10,2).Range.Text = $Statistics.Confirms
		$Table.Cell(11,1).Range.Text = "Declines"
		$Table.Cell(11,2).Range.Text = $Statistics.Declines
		$Table.Cell(12,1).Range.Text = "Releases"
		$Table.Cell(12,2).Range.Text = $Statistics.Releases
		$Table.Cell(13,1).Range.Text = "Total Scopes"
		$Table.Cell(13,2).Range.Text = $Statistics.TotalScopes
		$Table.Cell(14,1).Range.Text = "Total Addresses"
		$tmp = "{0:N0}" -f $Statistics.TotalAddresses
		$Table.Cell(14,2).Range.Text = $tmp
		$Table.Cell(15,1).Range.Text = "In Use"
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		$Table.Cell(15,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent)%)"
		$Table.Cell(16,1).Range.Text = "Available"
		[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable
		$tmp = "{0:N0}" -f $Statistics.AddressesAvailable
		$Table.Cell(16,2).Range.Text = "$($tmp) ($($AvailablePercent)%)"

		$Table.Rows.SetLeftIndent(36,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 statistics"
	}
	Else
	{
		WriteWordLine 0 0 "There were no IPv6 statistics"
	}
	$Statistics = $Null

	Write-Verbose "$(Get-Date):	Getting IPv6 scopes"
	$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $DHCPServerName -EA 0

	If($? -and $IPv6Scopes -ne $Null)
	{
		$selection.InsertNewPage()
		ForEach($IPv6Scope in $IPv6Scopes)
		{
			GetIPv6ScopeData $IPv6Scope
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 scopes"
	}
	Else
	{
		WriteWordLine 0 1 "There were no IPv6 scopes"
	}
	$IPv6Scopes = $Null

	Write-Verbose "$(Get-Date):	Getting IPv6 server options"
	$selection.InsertNewPage()
	WriteWordLine 3 0 "Server Options"

	$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $DHCPServerName -EA 0 | Sort-Object OptionId

	If($? -and $ServerOptions -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ServerOptions -is [array])
		{
			[int]$Rows = $ServerOptions.Count * 4
		}
		Else
		{
			[int]$Rows = 3
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		[int]$xRow = 0
		ForEach($ServerOption in $ServerOptions)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Option Name"
			$Table.Cell($xRow,2).Range.Text = "$($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Vendor"
			If([string]::IsNullOrEmpty($ServerOption.VendorClass))
			{
				$Table.Cell($xRow,2).Range.Text =  "Standard"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = $ServerOption.VendorClass
			}
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Value"
			$Table.Cell($xRow,2).Range.Text = $ServerOption.Value
			
			#for spacing
			$xRow++
		}
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 server options"
	}
	Else
	{
		WriteWordLine 0 1 "There were no IPv6 server options"
	}
	$ServerOptions = $Null
}
ElseIf($Text)
{
	Line 0 "IPv6"
	Line 0 "Properties"

	Write-Verbose "$(Get-Date): Getting IPv6 properties"
	Write-Verbose "$(Get-Date): `tGetting IPv6 general settings"
	Line 1 "General"

	If($GotAuditSettings)
	{
		If($AuditSettings.Enable)
		{
			Line 2 "DHCP audit logging is enabled"
		}
		Else
		{
			Line 2 "DHCP audit logging is disabled"
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving audit log settings"
	}
	Else
	{
		Line 2 "There were no audit log settings"
	}

	#DNS settings
	Write-Verbose "$(Get-Date): `tGetting IPv6 DNS settings"
	Line 1 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $DHCPServerName -EA 0
	If($? -and $DNSSettings -ne $Null)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		Line 0 "Error retrieving IPv6 DNS Settings for DHCP server $DHCPServerName"
	}
	$DNSSettings = $Null

	#Advanced
	Write-Verbose "$(Get-Date): `tGetting IPv6 advanced settings"
	Line 1 "Advanced"
	If($GotAuditSettings)
	{
		Line 2 "Audit log file path " $AuditSettings.Path
	}
	$AuditSettings = $Null

	Line 1 "Statistics"
	$Statistics = Get-DHCPServerV6Statistics -ComputerName $DHCPServerName -EA 0

	If($? -and $Statistics -ne $Null)
	{
		$UpTime = $(Get-Date) - $Statistics.ServerStartTime
		$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
			$UpTime.Days, `
			$UpTime.Hours, `
			$UpTime.Minutes, `
			$UpTime.Seconds)
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable

		Line 2 "Description" -NoNewLine
		Line 2 "Details"

		Line 2 "Start Time: " -NoNewLine
		Line 2 $Statistics.ServerStartTime
		Line 2 "Up Time: " -NoNewLine
		Line 2 $Str
		Line 2 "Solicits: " -NoNewLine
		Line 2 $Statistics.Solicits
		Line 2 "Advertises: " -NoNewLine
		Line 2 $Statistics.Advertises
		Line 2 "Requests: " -NoNewLine
		Line 2 $Statistics.Requests
		Line 2 "Replies: " -NoNewLine
		Line 2 $Statistics.Replies
		Line 2 "Renews: " -NoNewLine
		Line 2 $Statistics.Renews
		Line 2 "Rebinds: " -NoNewLine
		Line 2 $Statistics.Rebinds
		Line 2 "Confirms: " -NoNewLine
		Line 2 $Statistics.Confirms
		Line 2 "Declines: " -NoNewLine
		Line 2 $Statistics.Declines
		Line 2 "Releases: " -NoNewLine
		Line 2 $Statistics.Releases
		Line 2 "Total Scopes: " -NoNewLine
		Line 2 $Statistics.TotalScopes
		Line 2 "Total Addresses: " -NoNewLine
		$tmp = "{0:N0}" -f $Statistics.TotalAddresses
		Line 1 $tmp
		Line 2 "In Use: " -NoNewLine
		Line 2 "$($Statistics.AddressesInUse) ($($InUsePercent)%)"
		Line 2 "Available: " -NoNewLine
		$tmp = "{0:N0}" -f $Statistics.AddressesAvailable 
		Line 2 "$($tmp) ($($AvailablePercent)%)"
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving IPv6 statistics"
	}
	Else
	{
		Line 0 "There were no IPv6 statistics"
	}
	
	$Statistics = $Null

	Write-Verbose "$(Get-Date):	Getting IPv6 scopes"
	$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $DHCPServerName -EA 0

	If($? -and $IPv6Scopes -ne $Null)
	{
		ForEach($IPv6Scope in $IPv6Scopes)
		{
			GetIPv6ScopeData $IPv6Scope
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving IPv6 scopes"
	}
	Else
	{
		Line 1 "There were no IPv6 scopes"
	}
	$IPv6Scopes = $Null

	Write-Verbose "$(Get-Date): Getting IPv6 server options"
	Line 0 "Server Options"

	$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $DHCPServerName -EA 0 | Sort-Object OptionId

	If($? -and $ServerOptions -ne $Null)
	{
		ForEach($ServerOption in $ServerOptions)
		{
			Line 1 "Option Name`t: $($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"
			Line 1 "Vendor`t`t: " -NoNewLine
			If([string]::IsNullOrEmpty($ServerOption.VendorClass))
			{
				Line 0 "Standard"
			}
			Else
			{
				Line 0 $ServerOption.VendorClass
			}
			Line 1 "Value`t`t: " $ServerOption.Value
			
			#for spacing
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving IPv6 server options"
	}
	Else
	{
		Line 2 "There were no IPv6 server options"
	}
	$ServerOptions = $Null
}
ElseIf($HTML)
{
}

If($MSWord -or $PDF)
{
	Write-Verbose "$(Get-Date): Finishing up Word document"
	#end of document processing

	#Update document properties
	If($CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Set Cover Page Properties"
		_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
		_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
		_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "DHCP Inventory"
		_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

		#Get the Coverpage XML part
		$cp = $doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

		#get the abstract XML part
		$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
		#set the text
		[string]$abstract = "DHCP Inventory for $CompanyName"
		$ab.Text = $abstract

		$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
		#set the text
		[string]$abstract = (Get-Date -Format d).ToString()
		$ab.Text = $abstract

		Write-Verbose "$(Get-Date): Update the Table of Contents"
		#update the Table of Contents
		$doc.TablesOfContents.item(1).Update()
		$cp = $Null
		$ab = $Null
		$abstract = $Null
	}

	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Word.Options.CheckGrammarAsYouType = $CurrentGrammarOption
	$Word.Options.CheckSpellingAsYouType = $CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($WordVersion -eq $wdWord2007)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
		If($RunningOS.Contains("Server 2008 R2"))
		{
			$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
			$doc.SaveAs($filename1, $SaveFormat)
			If($PDF)
			{
				Write-Verbose "$(Get-Date): Now saving as PDF"
				$SaveFormat = $wdSaveFormatPDF
				$doc.SaveAs($filename2, $SaveFormat)
			}
		}
		Else
		{
			#works for Server 2008 and Windows 7
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
			$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
			If($PDF)
			{
				Write-Verbose "$(Get-Date): Now saving as PDF"
				$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
				$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
			}
		}
	}
	Else
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$doc.Close()
	$Word.Quit()
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Deleting $($filename1) since only $($filename2) is needed"
		Remove-Item $filename1 -EA 0
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()

	If($PDF)
	{
		Write-Verbose "$(Get-Date): $($filename2) is ready for use"
	}
	Else
	{
		Write-Verbose "$(Get-Date): $($filename1) is ready for use"
	}
}

Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($Text)
{
	write-output $global:output | Out-File $Filename1
}
ElseIf($HTML)
{
}

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
      $runtime.Days, `
      $runtime.Hours, `
      $runtime.Minutes, `
      $runtime.Seconds,
      $runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null