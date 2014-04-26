#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft Active Directory Forest using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Microsoft Active Directory Forest using Microsoft Word and PowerShell.
	Creates a Word document named after the Active Directory Forest.
	Document includes a Cover Page, Table of Contents and Footer.
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
	(default cover pages in Word en-US)
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
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	Will be used on Domain Controllers only.
	This parameters requires the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware infomation
	This parameter is disabled by default.
.PARAMETER ADForest
	Specifies an Active Directory forest object by providing one of the following attribute values. 
	The identifier in parentheses is the LDAP display name for the attribute.

	Fully qualified domain name
		Example: corp.contoso.com
	GUID (objectGUID)
		Example: 599c3d2e-f72d-4d20-8a88-030d99495f20
	DNS host name
		Example: dnsServer.corp.contoso.com
	NetBIOS name
		Example: corp
		
	This parameter is required.
.PARAMETER ComputerName
	Specifies which domain controller to use to run the script against.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -ADForest company.tld
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -PDF -ADForest corp.carlwebster.com
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory.ps1 -hardware
	
	Will use all default values and add additional information for each domain controller about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster" -ComputerName ADDC01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		Domain Controller named ADDC01 for the ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
		The computer running the script for the ComputerName.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.NOTES
	NAME: ADDS_Inventory.ps1
	VERSION: 0.52
	AUTHOR: Carl Webster
	LASTEDIT: April 23, 2014
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param(
	[parameter(
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(
	Position = 3, 
	Mandatory=$false )
	] 
	[Switch]$PDF=$False,

	[parameter(
	Position = 4, 
	Mandatory=$false )
	] 
	[Switch]$Hardware=$False,

	[parameter(
	Position = 5, 
	Mandatory=$True )
	] 
	[string]$ADForest="", 

	[parameter(
	Position = 6, 
	Mandatory=$false )
	] 
	[string]$ComputerName="")
	

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
If($PDF -eq $Null)
{
	$PDF = $False
}
If($Hardware -eq $Null)
{
	$Hardware = $False
}


#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Created on April 10, 2014

Set-StrictMode -Version 2

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
[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

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

Switch ($PSUICulture.Substring(0,3))
{
	'ca-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Taula automática 2';
			}
		}

	'da-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabel 2';
			}
		}

	'de-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische Tabelle 2';
			}
		}

	'en-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}

	'es-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Tabla automática 2';
			}
		}

	'fi-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automaattinen taulukko 2';
			}
		}

	'fr-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Sommaire Automatique 2';
			}
		}

	'nb-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabell 2';
			}
		}

	'nl-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
			}
		}

	'pt-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Sumário Automático 2';
			}
		}

	'sv-'	{
			$hash.($($PSUICulture)) = @{
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

$myHash = $hash.$PSUICulture

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
	
	Switch ($PSUICulture.Substring(0,3))
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana", "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)", "Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador", "Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari", "Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast", "Cubicles", "Diplomàtic", "En mosaic",
					"Exposició", "Línia lateral", "Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran", "Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)", "Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter", "Overskrid", "Alfabet", "Kontrast", "Stakke",
					"Fliser", "Gåde", "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel", "Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "BevægElse", "Eksponering", "Enkel", "Firkanter", "Fliser", "Gåde", "Kontrast",
					"Mod", "Nålestribet", "Overskrid", "Sidelinje", "Stakke", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)", "Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung", "Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend", "Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt", "Jährlich", "Kacheln", "Kontrast",
					"Kubistisch", "Modern", "Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel", "Traditionell")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion",
					"Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid",
					"Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion",
					"Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin", "Slice (luz)", "Faceta", "Semáforo",
					"Retrospectiva", "Cuadrícula", "Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador", "Contraste", "Cuadrícula",
					"Cubículos", "Exposición", "Línea lateral", "Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador", "Contraste", "Cubículos", "Exposición",
					"Línea lateral", "Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)", "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin", "Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti", "Laatikot", "Liike", "Liituraita", "Mod",
					"Osittain peitossa", "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko", "Ruudut", "Sanomalehtipaperi",
					"Sivussa", "Vuotuinen", "Ylitys")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti" ,"Kuvakkeet ja tiedot" ,"Liike" ,"Liituraita" ,"Mod" ,"Palapeli",
					"Perinteinen", "Pinot", "Sivussa", "Työpisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster","Secteur (foncé)","Sémaphore","Rétrospective","Ion (foncé)","Ion (clair)","Intégrale",
					"Filigrane","Facette","Secteur (clair)","À bandes", "Austin", "Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective", "Contraste", "Emplacements de bureau",
					"Moderne","Blocs empilés", "Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage", "Exposition",
					"Alphabet", "Mots croisés", "Papier journal", "Austin","Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Blocs empilés", "Blocs superposés", "Classique", "Contraste",
					"Exposition","Guide", "Ligne latérale", "Moderne", "Mosaïques", "Mots croisés", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran", "Integral", "Ion (lys)", "Ion (mørk)",
					"Retrospekt", "Rutenett", "Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker", "BevegElse", "Engasjement", "Enkel", "Fliser",
					"Konservativ", "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje", "Smale striper", "Stabler",
					"Transcenderende")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "Avlukker", "BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Puslespill", "Sidelinje", "Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept", "Integraal", "Ion (donker)", "Ion (licht)",
					"Raster", "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet", "Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden", "Beweging", "Blikvanger", "Contrast", "Eenvoudig",
					"Jaarlijks", "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief", "Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging", "Blikvanger", "Contrast", "Eenvoudig",
					"Jaarlijks", "Krijtstreep", "Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra", "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete",
					"Filigrana", "Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral", "Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias", "Conservador", "Contraste", "Exposição",
					"Grade", "Ladrilhos", "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas", "Quebra-cabeça", "Transcend")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Baias", "Conservador", "Contraste", "Exposição",
					"Ladrilhos", "Linha Lateral", "Listras", "Mod", "Pilhas", "Quebra-cabeça", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)", "Jon (mörkt)", "Knippe", "Rutnät",
					"RörElse", "Sektor (ljus)", "Sektor (mörk)", "Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt", "Kontrast", "Kritstreck", "Kuber",
					"Perspektiv", "Plattor", "Pussel", "Rutnät", "RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabetmönster", "Årligt", "Enkelt", "Exponering", "Konservativt", "Kontrast", "Kritstreck",
					"Kuber", "Övergående", "Plattor", "Pussel", "RörElse", "Sidlinje", "Sobert", "Staplat")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion",
						"Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid",
						"Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion",
						"Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
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

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	WriteWordLine 3 0 "Computer Information"
	WriteWordLine 0 1 "General Computer"
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, @{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		$GotComputerItems = $False
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotComputerItems)
	{
		ForEach($Item in $ComputerItems)
		{
			WriteWordLine 0 2 "Manufacturer`t: " $Item.manufacturer
			WriteWordLine 0 2 "Model`t`t: " $Item.model
			WriteWordLine 0 2 "Domain`t`t: " $Item.domain
			WriteWordLine 0 2 "Total Ram`t: $($Item.totalphysicalram) GB"
			WriteWordLine 0 2 ""
		}
	}

	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"
	WriteWordLine 0 1 "Drive(s)"
	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
		$drives = $Results | select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		$GotDrives = $False
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotDrives)
	{
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				WriteWordLine 0 2 "Caption`t`t: " $drive.caption
				WriteWordLine 0 2 "Size`t`t: $($drive.drivesize) GB"
				If(![String]::IsNullOrEmpty($drive.filesystem))
				{
					WriteWordLine 0 2 "File System`t: " $drive.filesystem
				}
				WriteWordLine 0 2 "Free Space`t: $($drive.drivefreespace) GB"
				If(![String]::IsNullOrEmpty($drive.volumename))
				{
					WriteWordLine 0 2 "Volume Name`t: " $drive.volumename
				}
				If(![String]::IsNullOrEmpty($drive.volumedirty))
				{
					WriteWordLine 0 2 "Volume is Dirty`t: " -nonewline
					If($drive.volumedirty)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
				{
					WriteWordLine 0 2 "Volume Serial #`t: " $drive.volumeserialnumber
				}
				WriteWordLine 0 2 "Drive Type`t: " -nonewline
				Switch ($drive.drivetype)
				{
					0	{WriteWordLine 0 0 "Unknown"}
					1	{WriteWordLine 0 0 "No Root Directory"}
					2	{WriteWordLine 0 0 "Removable Disk"}
					3	{WriteWordLine 0 0 "Local Disk"}
					4	{WriteWordLine 0 0 "Network Drive"}
					5	{WriteWordLine 0 0 "Compact Disc"}
					6	{WriteWordLine 0 0 "RAM Disk"}
					Default {WriteWordLine 0 0 "Unknown"}
				}
				WriteWordLine 0 2 ""
			}
		}
	}

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"
	WriteWordLine 0 1 "Processor(s)"
	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
		$Processors = $Results | select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		$GotProcessors = $False
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotProcessors)
	{
		ForEach($processor in $processors)
		{
			WriteWordLine 0 2 "Name`t`t`t: " $processor.name
			WriteWordLine 0 2 "Description`t`t: " $processor.description
			WriteWordLine 0 2 "Max Clock Speed`t: $($processor.maxclockspeed) MHz"
			If($processor.l2cachesize -gt 0)
			{
				WriteWordLine 0 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
			}
			If($processor.l3cachesize -gt 0)
			{
				WriteWordLine 0 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
			}
			If($processor.numberofcores -gt 0)
			{
				WriteWordLine 0 2 "# of Cores`t`t: " $processor.numberofcores
			}
			If($processor.numberoflogicalprocessors -gt 0)
			{
				WriteWordLine 0 2 "# of Logical Procs`t: " $processor.numberoflogicalprocessors
			}
			WriteWordLine 0 2 "Availability`t`t: " -nonewline
			Switch ($processor.availability)
			{
				1	{WriteWordLine 0 0 "Other"}
				2	{WriteWordLine 0 0 "Unknown"}
				3	{WriteWordLine 0 0 "Running or Full Power"}
				4	{WriteWordLine 0 0 "Warning"}
				5	{WriteWordLine 0 0 "In Test"}
				6	{WriteWordLine 0 0 "Not Applicable"}
				7	{WriteWordLine 0 0 "Power Off"}
				8	{WriteWordLine 0 0 "Off Line"}
				9	{WriteWordLine 0 0 "Off Duty"}
				10	{WriteWordLine 0 0 "Degraded"}
				11	{WriteWordLine 0 0 "Not Installed"}
				12	{WriteWordLine 0 0 "Install Error"}
				13	{WriteWordLine 0 0 "Power Save - Unknown"}
				14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
				15	{WriteWordLine 0 0 "Power Save - Standby"}
				16	{WriteWordLine 0 0 "Power Cycle"}
				17	{WriteWordLine 0 0 "Power Save - Warning"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 ""
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"
	WriteWordLine 0 1 "Network Interface(s)"
	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration 
		$Nics = $Results | where {$_.ipenabled -eq $True}
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		$GotNics = $False
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}

	if( $Nics -eq $Null ) 
	{ 
		$GotNics = $False 
	} 
	else 
	{ 
		$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
	} 
	
	If($GotNics)
	{
		ForEach($nic in $nics)
		{
			$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | where {$_.index -eq $nic.index}
			If($ThisNic.Name -eq $nic.description)
			{
				WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
			}
			Else
			{
				WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
				WriteWordLine 0 2 "Description`t`t: " $nic.description
			}
			WriteWordLine 0 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
			WriteWordLine 0 2 "Manufacturer`t`t: " $ThisNic.manufacturer
			WriteWordLine 0 2 "Availability`t`t: " -nonewline
			Switch ($ThisNic.availability)
			{
				1	{WriteWordLine 0 0 "Other"}
				2	{WriteWordLine 0 0 "Unknown"}
				3	{WriteWordLine 0 0 "Running or Full Power"}
				4	{WriteWordLine 0 0 "Warning"}
				5	{WriteWordLine 0 0 "In Test"}
				6	{WriteWordLine 0 0 "Not Applicable"}
				7	{WriteWordLine 0 0 "Power Off"}
				8	{WriteWordLine 0 0 "Off Line"}
				9	{WriteWordLine 0 0 "Off Duty"}
				10	{WriteWordLine 0 0 "Degraded"}
				11	{WriteWordLine 0 0 "Not Installed"}
				12	{WriteWordLine 0 0 "Install Error"}
				13	{WriteWordLine 0 0 "Power Save - Unknown"}
				14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
				15	{WriteWordLine 0 0 "Power Save - Standby"}
				16	{WriteWordLine 0 0 "Power Cycle"}
				17	{WriteWordLine 0 0 "Power Save - Warning"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 "Physical Address`t: " $nic.macaddress
			WriteWordLine 0 2 "IP Address`t`t: " $nic.ipaddress
			WriteWordLine 0 2 "Default Gateway`t: " $nic.Defaultipgateway
			WriteWordLine 0 2 "Subnet Mask`t`t: " $nic.ipsubnet
			If($nic.dhcpenabled)
			{
				$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
				$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
				WriteWordLine 0 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
				WriteWordLine 0 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
				WriteWordLine 0 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
				WriteWordLine 0 2 "DHCP Server`t`t:" $nic.dhcpserver
			}
			If(![String]::IsNullOrEmpty($nic.dnsdomain))
			{
				WriteWordLine 0 2 "DNS Domain`t`t: " $nic.dnsdomain
			}
			If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
			{
				[int]$x = 1
				WriteWordLine 0 2 "DNS Search Suffixes`t:" -nonewline
				ForEach($DNSDomain in $nic.dnsdomainsuffixsearchorder)
				{
					If($x -eq 1)
					{
						$x = 2
						WriteWordLine 0 0 " $($DNSDomain)"
					}
					Else
					{
						WriteWordLine 0 5 " $($DNSDomain)"
					}
				}
			}
			WriteWordLine 0 2 "DNS WINS Enabled`t: " -nonewline
			If($nic.dnsenabledforwinsresolution)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
			{
				[int]$x = 1
				WriteWordLine 0 2 "DNS Servers`t`t:" -nonewline
				ForEach($DNSServer in $nic.dnsserversearchorder)
				{
					If($x -eq 1)
					{
						$x = 2
						WriteWordLine 0 0 " $($DNSServer)"
					}
					Else
					{
						WriteWordLine 0 5 " $($DNSServer)"
					}
				}
			}
			WriteWordLine 0 2 "NetBIOS Setting`t`t: " -nonewline
			Switch ($nic.TcpipNetbiosOptions)
			{
				0	{WriteWordLine 0 0 "Use NetBIOS setting from DHCP Server"}
				1	{WriteWordLine 0 0 "Enable NetBIOS"}
				2	{WriteWordLine 0 0 "Disable NetBIOS"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 "WINS:"
			WriteWordLine 0 3 "Enabled LMHosts`t: " -nonewline
			If($nic.winsenablelmhostslookup)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
			{
				WriteWordLine 0 3 "Host Lookup File`t: " $nic.winshostlookupfile
			}
			If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
			{
				WriteWordLine 0 3 "Primary Server`t`t: " $nic.winsprimaryserver
			}
			If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
			{
				WriteWordLine 0 3 "Secondary Server`t: " $nic.winssecondaryserver
			}
			If(![String]::IsNullOrEmpty($nic.winsscopeid))
			{
				WriteWordLine 0 3 "Scope ID`t`t: " $nic.winsscopeid
			}
		}
	}
	WriteWordLine 0 0 ""
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
	
Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module | % { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
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
	Remove-Variable -Name word -Scope Global
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

Function BuildMultiColumnTable
{
	Param([Array]$xArray, [String]$xType)
	
	#divide by 0 bug reported 9-Apr-2014 by Lee Dehmer 
	#if security group name or OU name was longer than 60 characters it caused a divide by 0 error
	
	#added a second parameter to the function so the verbose message would say whether 
	#the function is processing servers, security groups or OUs.
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	#remove 60 as a hard-coded value
	#60 is the max width the table can be when indented 36 points
	[int]$MaxTableWidth = 60
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	$TableRange = $doc.Application.Selection.Range
	#removed hard-coded value of 60 and replace with MaxTableWidth variable
	[int]$Columns = [Math]::Floor($MaxTableWidth / $MaxLength)
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells  = $xArray.Count
		#reset column count so there are no empty columns
		$Columns   = $xArray.Count 
	}
	ElseIf($Columns -eq 0)
	{
		#divide by 0 bug if this condition is not handled
		#number was larger than $MaxTableWidth so there can only be one column
		#with one cell per row
		[int]$Rows = $xArray.count
		$Columns   = 1
		$MaxCells  = 1
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells  = $Columns
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	[int]$xRow = 1
	[int]$ArrayItem = 0
	While($xRow -le $Rows)
	{
		For($xCell=1; $xCell -le $MaxCells; $xCell++)
		{
			$Table.Cell($xRow,$xCell).Range.Text = $xArray[$ArrayItem]
			$ArrayItem++
		}
		$xRow++
	}
	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$xArray = $Null
}

#Script begins

$script:startTime = Get-Date

#If hardware inventory is requested, make sure user is running the script with domain admin rights
If($Hardware)
{
	Write-Verbose "$(Get-Date): Hardware inventory requested, testing to see if $($env:username) has domain admin rights"
	If(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Domain Admins"))
	{
		#user has domain admin rights
	}
	Else
	{
		#user does not have domain admin rights
		#don't abort script, set $hardware to false
		Write-Warning "Hardware inventory was requested but $($env:username) does not have domain admin rights."
		Write-Warning "Hardware inventory option will be turned off."
		$Hardware = $False
	}
}

#make sure ActiveDirectory module is loaded
If(!(Check-LoadedModule "ActiveDirectory"))
{
	Write-Error "The ActiveDirectory module could not be loaded.`nScript cannot continue."
	Exit
}

If(![String]::IsNullOrEmpty($ComputerName)) 
{
	#get server name
	#first test to make sure the server is reachable
	Write-Verbose "$(Get-Date): Testing to see if $($ComputerName) is online and reachable"
	If(Test-Connection -ComputerName $ComputerName -quiet -EA 0)
	{
		Write-Verbose "$(Get-Date): Server $($ComputerName) is online.  Testing to see if it is a Domain Controller."
		#the server may be online but is it really a domain controller?
		Try
		{
			$Results = Get-ADDomainController $ComputerName -Server $ADForest -EA 0
		}
		
		Catch
		{
			If(!$?)
			{
				Write-Error "$($ComputerName) is not a domain controller.`nScript cannot continue."
				Exit
			}
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Computer $($ComputerName) is offline"
		Write-Error "Computer $($ComputerName) is offline.`nScript cannot continue."
		Exit
	}
}

CheckWordPreReq

#get forest information so output filename can be generated
Write-Verbose "$(Get-Date): Testing to see if $($ADForest) is a valid forest name"
If([String]::IsNullOrEmpty($ComputerName))
{
	Try
	{
		$Forest = Get-ADForest -Identity $ADForest -EA 0
	}
	
	Catch
	{
		Write-Error "Could not find a forest identified by: $($ADForest).`nScript cannot continue."
		Exit
	}
}
Else
{
	Try
	{
		$Forest = Get-ADForest -Identity $ADForest -Server $ComputerName
	}
	
	Catch
	{
		Write-Error "Could not find a forest with the name of $($ADForest).`nScript cannot continue."
		Exit
	}
}
Write-Verbose "$(Get-Date): $($ADForest) is a valid forest name"

[string]$ForestName = $Forest.Name
[string]$Title      = "Inventory Report for the $($ForestName) Forest"
[string]$filename1  = "$($pwd.path)\$($ForestName).docx"
If($PDF)
{
	[string]$filename2 = "$($pwd.path)\$($XDForestName).pdf"
}

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

[int]$WordVersion = [int] $Word.Version
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
Switch ($PSUICulture.Substring(0,3))
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

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Company Name : $CompanyName"
Write-Verbose "$(Get-Date): Cover Page   : $CoverPage"
Write-Verbose "$(Get-Date): User Name    : $UserName"
Write-Verbose "$(Get-Date): Save As PDF  : $PDF"
Write-Verbose "$(Get-Date): HW Inventory : $Hardware"
Write-Verbose "$(Get-Date): Forest Name  : $ADForest"
Write-Verbose "$(Get-Date): Title        : $Title"
Write-Verbose "$(Get-Date): Filename1    : $filename1"
If($PDF)
{
	Write-Verbose "$(Get-Date): Filename2    : $filename2"
}
Write-Verbose "$(Get-Date): OS Detected  : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture  : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture    : $PSCulture"
Write-Verbose "$(Get-Date): Word version : $WordProduct"
Write-Verbose "$(Get-Date): Word language: $($Word.Language)"
Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

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
	#word 2010/2013
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

#cannot continue if cover page does not exist
If(!$CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Error "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist.  Script cannot continue."
	Write-Verbose "$(Get-Date): Closing Word"
	AbortScript
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

If($BuildingBlocksExist)
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
ForEach ($footer in $footers) 
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

######################START OF BUILDING REPORT

#Forest information

#set naming context
$ConfigNC = (Get-ADRootDSE -Server $ADForest).ConfigurationNamingContext

Write-Verbose "$(Get-Date): Writing forest data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Forest Information"

Switch ($Forest.ForestMode)
{
	"Windows2000Forest"        {$ForestMode = "Windows 2000"}
	"Windows2003InterimForest" {$ForestMode = "Windows Server 2003 interim"}
	"Windows2003Forest"        {$ForestMode = "Windows Server 2003"}
	"Windows2008Forest"        {$ForestMode = "Windows Server 2008"}
	"Windows2008R2Forest"      {$ForestMode = "Windows Server 2008 R2"}
	"Windows2012Forest"        {$ForestMode = "Windows Server 2012"}
	"Windows2012R2Forest"      {$ForestMode = "Windows Server 2012 R2"}
	"UnknownForest"            {$ForestMode = "Unknown Forest Mode"}
	Default                    {$ForestMode = "Unable to determine Forest Mode: $($Forest.ForestMode)"}
}

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 2
[int]$Rows = 6
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = $myHash.Word_TableGrid
$table.Borders.InsideLineStyle = 0
$table.Borders.OutsideLineStyle = 0
$Table.Cell(1,1).Range.Text = "Forest name"
$Table.Cell(1,2).Range.Text = $Forest.Name
$Table.Cell(2,1).Range.Text = "Forest mode"
$Table.Cell(2,2).Range.Text = $ForestMode
$Table.Cell(3,1).Range.Text = "Root domain"
$Table.Cell(3,2).Range.Text = $Forest.RootDomain
$Table.Cell(4,1).Range.Text = "Domain naming master"
$Table.Cell(4,2).Range.Text = $Forest.DomainNamingMaster
$Table.Cell(5,1).Range.Text = "Schema master"
$Table.Cell(5,2).Range.Text = $Forest.SchemaMaster
$Table.Cell(6,1).Range.Text = "Partitions container"
$Table.Cell(6,2).Range.Text = $Forest.PartitionsContainer
$Table.Rows.SetLeftIndent(0,1)
$table.AutoFitBehavior(1)

#return focus back to document
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
$selection.EndKey($wdStory,$wdMove) | Out-Null

Write-Verbose "$(Get-Date): `tApplication partitions"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "Application partitions: " -NoNewLine
$AppPartitions = $Forest.ApplicationPartitions | Sort
If($AppPartitions -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 1
	If($AppPartitions -is [array])
	{
		[int]$Rows = $AppPartitions.Count
	}
	Else
	{
		[int]$Rows = 1
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	[int]$xRow = 0
	ForEach($AppPartition in $AppPartitions)
	{
		$xRow++
		$Table.Cell($xRow,1).Range.Text = $AppPartition
	}
	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
}

Write-Verbose "$(Get-Date): `tCross forest references"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "Cross forest references: " -NoNewLine
$CrossForestReferences = $Forest.CrossForestReferences | Sort
If($CrossForestReferences -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 1
	If($CrossForestReferences -is [array])
	{
		[int]$Rows = $CrossForestReferences.Count
	}
	Else
	{
		[int]$Rows = 1
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	[int]$xRow = 0
	ForEach($CrossForestReference in $CrossForestReferences)
	{
		$xRow++
		$Table.Cell($xRow,1).Range.Text = $CrossForestReference
	}
	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
}

Write-Verbose "$(Get-Date): `tSPN suffixes"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "SPN suffixes: " -NoNewLine
$SPNSuffixes = $Forest.SPNSuffixes | Sort
If($SPNSuffixes -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	BuildMultiColumnTable $SPNSuffixes "SPN SUffixes"
}

Write-Verbose "$(Get-Date): `tUPN suffixes"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "UPN Suffixes: " -NoNewLine
$UPNSuffixes = $Forest.UPNSuffixes | Sort
If($UPNSuffixes -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	BuildMultiColumnTable $UPNSuffixes "UPN Suffixes"
}

Write-Verbose "$(Get-Date): `tDomains in forest"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "Domains in forest: " -NoNewLine
$Domains = $Forest.Domains | Sort
If($Domains -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	#redo list of domains so forest root domain is listed first
	$tmpDomains = @($Forest.RootDomain)
	ForEach($Domain in $Domains)
	{
		If($Domain -ne $Forest.RootDomain)
		{
			$tmpDomains += $Domain
		}
	}
	
	$Domains = $tmpDomains
	
	WriteWordLine 0 0 ""
	BuildMultiColumnTable $Domains "Domains in forest"
}

Write-Verbose "$(Get-Date): `tSites"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "Sites: " -NoNewLine
$Sites = $Forest.Sites | Sort
If($Sites -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	BuildMultiColumnTable $Sites "Sites"
}

Write-Verbose "$(Get-Date): `tDomain controllers"
WriteWordLine 0 0 ""
WriteWordLine 0 0 "Domain Controllers: " -NoNewLine
#get all DCs in the forest
#http://www.superedge.net/2012/09/how-to-get-ad-forest-in-powershell.html
#http://msdn.microsoft.com/en-us/library/vstudio/system.directoryservices.activedirectory.forest.getforest%28v=vs.90%29
$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ADForest) 
$Forest2 = [system.directoryservices.activedirectory.Forest]::GetForest($ADContext)
$AllDCs = $Forest2.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
$AllDCs = $AllDCs | Sort
$ADContext = $Null
$Forest2 = $Null

If($AllDCs -eq $Null)
{
	WriteWordLine 0 0 "<None>"
}
Else
{
	WriteWordLine 0 0 ""
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 3
	If($AllDCs -is [array])
	{
		[int]$Rows = $AllDCs.Count
	}
	Else
	{
		[int]$Rows = 1
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 1
	$table.Borders.OutsideLineStyle = 1
	$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,1).Range.Font.Bold = $True
	$Table.Cell(1,1).Range.Text = "Name"
	$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,2).Range.Font.Bold = $True
	$Table.Cell(1,2).Range.Text = "Global Catalog"
	$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
	$Table.Cell(1,3).Range.Font.Bold = $True
	$Table.Cell(1,3).Range.Text = "Read-only"
	[int]$xRow = 1
	ForEach($DC in $AllDCs)
	{
		$DCName = $DC.SubString(0,$DC.IndexOf("."))
		$SrvName = $DC.SubString($DC.IndexOf(".")+1)
		$xRow++
		$Table.Cell($xRow,1).Range.Text = $DC
		
		$Results = Get-ADDomainController -Identity $DCName -Server $SrvName -EA 0
		
		If($? -and $Results -ne $Null)
		{
			$Table.Cell($xRow,2).Range.Text = $Results.IsGlobalCatalog
			$Table.Cell($xRow,3).Range.Text = $Results.IsReadOnly
		}
		Else
		{
			$Table.Cell($xRow,2).Range.Text = "Unknown"
			$Table.Cell($xRow,3).Range.Text = "Unknown"
		}
	}
	$Table.Rows.SetLeftIndent(36,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
}

#Site information
Write-Verbose "$(Get-Date): Writing sites and services data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Site and Services"

#get site information
#some of the following was taken from
#http://blogs.msdn.com/b/adpowershell/archive/2009/08/18/active-directory-powershell-to-manage-sites-and-subnets-part-3-getting-site-and-subnets.aspx

$tmp = $Forest.PartitionsContainer
$ConfigurationBase = $tmp.SubString($tmp.IndexOf(",") + 1)
$Sites = Get-ADObject -Filter 'ObjectClass -eq "site"' -SearchBase $ConfigurationBase -Properties Name, SiteObjectBl -Server $ADForest -EA 0 | Sort Name
$siteContainerDN = ("CN=Sites," + $configNC)

If($? -and $Sites -ne $Null)
{
	ForEach($Site in $Sites)
	{
		Write-Verbose "$(Get-Date): `tProcessing site $($Site.Name)"
		WriteWordLine 2 0 $Site.Name
		WriteWordLine 0 1 "Subnets: " -NoNewLine
		Write-Verbose "$(Get-Date): `t`tProcessing subnets"
		$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
		$i = 0
		foreach ($subnetDN in $Site.siteObjectBL) 
		{
			$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
			$subnetArray[$i] = $subnetName
			$i++
		}
		$subnetArray = $subnetArray | Sort
		If($subnetArray -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			BuildMultiColumnTable $subnetArray "Subnets"
		}
		WriteWordLine 0 0 
		
		Write-Verbose "$(Get-Date): `t`tProcessing servers"
		WriteWordLine 0 1 "Servers:"
		$siteName = $Site.Name
		
		#build array of connect objects
		Write-Verbose "$(Get-Date): `t`t`tProcessing automatic connection objects"
		$Connections = @()
		$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' -Searchbase $ConfigNC -Property DistinguishedName, fromServer -Server $ADForest -EA 0
		
		If($? -and $ConnectionObjects -ne $Null)
		{
			ForEach($ConnectionObject in $ConnectionObjects)
			{
				$xArray = $ConnectionObject.DistinguishedName.Split(",")
				#server name is 3rd item in array (element 2)
				$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
				$xArray = $ConnectionObject.FromServer.Split(",")
				#server name is 2nd item in array (element 1)
				$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
				#site name is 4th item in array (element 3)
				$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
				$xArray = $Null
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name           -Value "<automatically generated>"
				$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
				$Connections += $obj
			}
		}
		
		Write-Verbose "$(Get-Date): `t`t`tProcessing manual connection objects"
		$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' -Searchbase $ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest -EA 0
		
		If($? -and $ConnectionObjects -ne $Null)
		{
			ForEach($ConnectionObject in $ConnectionObjects)
			{
				$xArray = $ConnectionObject.DistinguishedName.Split(",")
				#server name is 3rd item in array (element 2)
				$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
				$xArray = $ConnectionObject.FromServer.Split(",")
				#server name is 2nd item in array (element 1)
				$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
				#site name is 4th item in array (element 3)
				$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
				$xArray = $Null
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name Name           -Value $ConnectionObject.Name
				$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
				$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
				$Connections += $obj
			}
		}

		If($Connections -ne $Null)
		{
			$Connections = $Connections | Sort Name, ToServer, FromServer
		}
		
		#list each server
		$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
		$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel -filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest | Select DNSHostName, Name | Sort DNSHostName
		If($? -and $SiteServers -ne $Null)
		{
			ForEach($SiteServer in $SiteServers)
			{
				WriteWordLine 0 2 $SiteServer.DNSHostName
				#for each server list each connection object
				If($Connections -ne $Null)
				{
					$Results = $Connections | Where {$_.ToServer -eq $SiteServer.Name}

					If($? -and $Results -ne $Null)
					{
						WriteWordLine 0 3 "Connection Objects to source server $($SiteServer.Name)"
						$TableRange = $doc.Application.Selection.Range
						[int]$Columns = 3
						If($Results -is [array])
						{
							[int]$Rows = $Results.Count + 1
						}
						Else
						{
							[int]$Rows = 2
						}
						[int]$xRow = 1
						$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
						$table.Style = $myHash.Word_TableGrid
						$table.Borders.InsideLineStyle = 1
						$table.Borders.OutsideLineStyle = 1
						$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "From Server"
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "From Site"
						ForEach($Result in $Results)
						{
							$xRow++
							$Table.Cell($xRow,1).Range.Text = $Result.Name
							$Table.Cell($xRow,2).Range.Text = $Result.FromServer
							$Table.Cell($xRow,3).Range.Text = $Result.FromServerSite
						}
						$Table.Rows.SetLeftIndent(108,1)
						$table.AutoFitBehavior(1)

						#return focus back to document
						$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$selection.EndKey($wdStory,$wdMove) | Out-Null
						WriteWordLine 0 0 ""
					}
				}
				Else
				{
					WriteWordLine 0 3 "Connection Objects: <None>"
				}
			}
		}
		ElseIf(!$?)
		{
			Write-Warning "No Site Servers were retrieved."
			WriteWordLine 0 0 "Warning: No Site Servers were retrieved" "" $null 0 $False $True
		}
		Else
		{
			WriteWordLine 0 2 "No servers in this site"
		}
	}
}
ElseIf(!$?)
{
	Write-Warning "No Sites were retrieved."
	WriteWordLine 0 0 "Warning: No Sites were retrieved" "" $null 0 $False $True
}
Else
{
	Write-Warning "There were no sites found to retrieve."
	WriteWordLine 0 0 "There were no sites found to retrieve" "" $null 0 $False $True
}

#domains
Write-Verbose "$(Get-Date): Writing domain data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Domain Information"
$AllDomainControllers = @()
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"

	Try
	{
		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
	}
	
	Catch
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
	}
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $Forest.RootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Switch ($DomainInfo.DomainMode)
		{
			"Windows2000Domain"   {$DomainMode = "Windows 2000"}
			"Windows2003Mixed"    {$DomainMode = "Windows Server 2003 mixed"}
			"Windows2003Domain"   {$DomainMode = "Windows Server 2003"}
			"Windows2008Domain"   {$DomainMode = "Windows Server 2008"}
			"Windows2008R2Domain" {$DomainMode = "Windows Server 2008 R2"}
			"Windows2012Domain"   {$DomainMode = "Windows Server 2012"}
			"Windows2012R2Domain" {$DomainMode = "Windows Server 2012 R2"}
			"UnknownForest"       {$DomainMode = "Unknown Domain Mode"}
			Default               {$DomainMode = "Unable to determine Doamin Mode: $($ADDomain.DomainMode)"}
		}
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
		{
			[int]$Rows = 17
		}
		Else
		{
			[int]$Rows = 16
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Range.Text = "Domain name"
		$Table.Cell(1,2).Range.Text = $DomainInfo.Name
		$Table.Cell(2,1).Range.Text = "Distinguished name"
		$Table.Cell(2,2).Range.Text = $DomainInfo.DistinguishedName
		$Table.Cell(3,1).Range.Text = "NetBIOS name"
		$Table.Cell(3,2).Range.Text = $DomainInfo.NetBIOSName
		$Table.Cell(4,1).Range.Text = "Domain mode"
		$Table.Cell(4,2).Range.Text = $DomainMode
		$Table.Cell(5,1).Range.Text = "DNS root"
		$Table.Cell(5,2).Range.Text = $DomainInfo.DNSRoot
		$Table.Cell(6,1).Range.Text = "Infrastructure master"
		$Table.Cell(6,2).Range.Text = $DomainInfo.InfrastructureMaster
		$Table.Cell(7,1).Range.Text = "PDC Emulator"
		$Table.Cell(7,2).Range.Text = $DomainInfo.PDCEmulator
		$Table.Cell(8,1).Range.Text = "RID Master"
		$Table.Cell(8,2).Range.Text = $DomainInfo.RIDMaster
		$Table.Cell(9,1).Range.Text = "Default computers container"
		$Table.Cell(9,2).Range.Text = $DomainInfo.ComputersContainer
		$Table.Cell(10,1).Range.Text = "Default users container"
		$Table.Cell(10,2).Range.Text = $DomainInfo.UsersContainer
		$Table.Cell(11,1).Range.Text = "Deleted objects container"
		$Table.Cell(11,2).Range.Text = $DomainInfo.DeletedObjectsContainer
		$Table.Cell(12,1).Range.Text = "Domain controllers container"
		$Table.Cell(12,2).Range.Text = $DomainInfo.DomainControllersContainer
		$Table.Cell(13,1).Range.Text = "Foreign security principals container"
		$Table.Cell(13,2).Range.Text = $DomainInfo.ForeignSecurityPrincipalsContainer
		$Table.Cell(14,1).Range.Text = "Lost and Found container"
		$Table.Cell(14,2).Range.Text = $DomainInfo.LostAndFoundContainer
		$Table.Cell(15,1).Range.Text = "Quotas container"
		$Table.Cell(15,2).Range.Text = $DomainInfo.QuotasContainer
		$Table.Cell(16,1).Range.Text = "Systems container"
		$Table.Cell(16,2).Range.Text = $DomainInfo.SystemsContainer
		If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
		{
			$Table.Cell(17,1).Range.Text = "Managed by"
			$Table.Cell(17,2).Range.Text = $DomainInfo.ManagedBy
		}

		$Table.Rows.SetLeftIndent(0,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null

		Write-Verbose "$(Get-Date): `t`tGetting Allowed DNS Suffixes"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Allowed DNS Suffixes: " -NoNewLine
		$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort
		If($DNSSuffixes -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			BuildMultiColumnTable $DNSSuffixes "Allowed DNS suffixes"
		}

		Write-Verbose "$(Get-Date): `t`tGetting Child domains"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Child domains: " -NoNewLine
		$ChildDomains = $DomainInfo.ChildDomains | Sort
		If($ChildDomains -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			BuildMultiColumnTable $ChildDomains "Child domains"
		}

		Write-Verbose "$(Get-Date): `t`tGetting Read-only replica directory servers"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Read-only replica directory servers: " -NoNewLine
		$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort
		If($ReadOnlyReplicas -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			BuildMultiColumnTable $ReadOnlyReplicas "Read-only replica directory servers"
		}

		Write-Verbose "$(Get-Date): `t`tGetting Replica directory servers"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Replica directory servers: " -NoNewLine
		$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort
		If($Replicas -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			BuildMultiColumnTable $Replicas "Replica directory servers"
		}

		Write-Verbose "$(Get-Date): `t`tGetting Subordinate references"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Subordinate references: " -NoNewLine
		$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort
		If($SubordinateReferences -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			If($SubordinateReferences -is [array])
			{
				[int]$Rows = $SubordinateReferences.Count
			}
			Else
			{
				[int]$Rows = 1
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			[int]$xRow = 0
			ForEach($SubordinateReference in $SubordinateReferences)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $SubordinateReference
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}

		Write-Verbose "$(Get-Date): `t`tGetting domain trusts"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Domain trusts: " -NoNewLine
		$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties * -EA 0

		If($? -and $ADDomainTrusts -ne $Null)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			
			If($ADDomainTrusts -is [array])
			{
				[int]$Rows = $ADDomainTrusts.Count * (8) #add an empty row for spacing
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
			
			ForEach($Trust in $ADDomainTrusts) 
			{ 
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Trust.Name 
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Trust.Description
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Created"
				$Table.Cell($xRow,2).Range.Text = $Trust.Created
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Modified"
				$Table.Cell($xRow,2).Range.Text = $Trust.Modified

				$TrustDirectionNumber = $Trust.TrustDirection
				$TrustTypeNumber = $Trust.TrustType
				$TrustAttributesNumber = $Trust.TrustAttributes

				#http://msdn.microsoft.com/en-us/library/cc220955.aspx
				#no values are defined at the above link
				Switch ($TrustTypeNumber) 
				{ 
					1 { $TrustType = "Downlevel"} 
					2 { $TrustType = "Uplevel"} 
					3 { $TrustType = "MIT (non-Windows)"} 
					4 { $TrustType = "DCE (Theoretical)"} 
					Default { $TrustType = $TrustTypeNumber }
				} 
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Type"
				$Table.Cell($xRow,2).Range.Text = $TrustType

				#http://msdn.microsoft.com/en-us/library/cc223779.aspx
				Switch ($TrustAttributesNumber) 
				{ 
					1 { $TrustAttributes = "Non-Transitive"} 
					2 { $TrustAttributes = "Uplevel clients only"} 
					4 { $TrustAttributes = "Quarantined Domain (External)"} 
					8 { $TrustAttributes = "Forest Trust"} 
					16 { $TrustAttributes = "Cross-Organizational Trust (Selective Authentication)"} 
					32 { $TrustAttributes = "Intra-Forest Trust"} 
					64 { $TrustAttributes = "Inter-Forest Trust"} 
					Default { $TrustAttributes = $TrustAttributesNumber }
				} 
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Attributes"
				$Table.Cell($xRow,2).Range.Text = $TrustAttributes
				 
				#http://msdn.microsoft.com/en-us/library/cc223768.aspx
				Switch ($TrustDirectionNumber) 
				{ 
					0 { $TrustDirection = "Disabled"} 
					1 { $TrustDirection = "Inbound"} 
					2 { $TrustDirection = "Outbound"} 
					3 { $TrustDirection = "Bidirectional"} 
					Default { $TrustDirection = $TrustDirectionNumber }
				}
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Direction"
				$Table.Cell($xRow,2).Range.Text = $TrustDirection
				
				#blank line for spacing
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
			#error retrieving domain trusts
			WriteWordLine 0 0 "Error retrieving domain trusts for $Domain" "" $null 0 $False $True
		}
		Else
		{
			#no domain trust data
			WriteWordLine 0 0 "<None>"
		}

		Write-Verbose "$(Get-Date): `t`tProcessing domain controllers"
		Try
		{
			$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot -EA 0 | Sort Name
		}
		
		Catch
		{
			Write-Warning "Error retrieving domain controller data for domain $($Domain)."
		}
		
		If($? -and $DomainControllers -ne $Null)
		{
			$AllDomainControllers += $DomainControllers
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "Domain Controllers: "
			#BuildMultiColumnTable $DomainControllers "Domain controllers"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			If($DomainControllers -is [array])
			{
				[int]$Rows = $DomainControllers.Count
			}
			Else
			{
				[int]$Rows = 1
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			[int]$xRow = 0
			ForEach($DomainController in $DomainControllers)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $DomainController.Name
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
			WriteWordLine 0 0 "Error retrieving domain controller data for domain $($Domain)" "" $null 0 $False $True
		}
		Else
		{
			WriteWordLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" "" $null 0 $False $True
		}
		
		$DomainControllers = $Null
		$LinkedGPOs = $Null
		$SubordinateReferences = $Null
		$Replicas = $Null
		$ReadOnlyReplicas = $Null
		$ChildDomains = $Null
		$DNSSuffixes = $Null
		$First = $False
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
}

#domain controllers
Write-Verbose "$(Get-Date): Writing domain controller data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Domain Controllers in $($ForestName)"
$AllDomainControllers = $AllDomainControllers | Sort Name
$First = $True

ForEach($DC in $AllDomainControllers)
{
	Write-Verbose "$(Get-Date): `tProcessing domain controller $($DC.name)"
	
	If(!$First)
	{
		#put each DC, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	
	WriteWordLine 2 0 $DC.Name
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	If(!$Hardware)
	{
		[int]$Rows = 16
	}
	Else
	{
		[int]$Rows = 11
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = 0
	$table.Borders.OutsideLineStyle = 0
	$Table.Cell(1,1).Range.Text = "Default partition"
	$Table.Cell(1,2).Range.Text = $DC.DefaultPartition
	$Table.Cell(2,1).Range.Text = "Domain"
	$Table.Cell(2,2).Range.Text = $DC.domain
	$Table.Cell(3,1).Range.Text = "Enabled"
	If($DC.Enabled -eq $True)
	{
		$Table.Cell(3,2).Range.Text = "True"
	}
	Else
	{
		$Table.Cell(3,2).Range.Text = "False"
	}
	$Table.Cell(4,1).Range.Text = "Hostname"
	$Table.Cell(4,2).Range.Text = $DC.HostName
	$Table.Cell(5,1).Range.Text = "Global Catalog"
	If($DC.IsGlobalCatalog -eq $True)
	{
		$Table.Cell(5,2).Range.Text = "Yes" 
	}
	Else
	{
		$Table.Cell(5,2).Range.Text = "No"
	}
	$Table.Cell(6,1).Range.Text = "Read-only"
	If($DC.IsReadOnly -eq $True)
	{
		$Table.Cell(6,2).Range.Text = "Yes"
	}
	Else
	{
		$Table.Cell(6,2).Range.Text = "No"
	}
	$Table.Cell(7,1).Range.Text = "LDAP port"
	$Table.Cell(7,2).Range.Text = $DC.LdapPort
	$Table.Cell(8,1).Range.Text = "SSL port"
	$Table.Cell(8,2).Range.Text = $DC.SslPort
	$Table.Cell(9,1).Range.Text = "Operation Master roles"
	$FSMORoles = $DC.OperationMasterRoles | Sort
	If($FSMORoles -eq $Null)
	{
		$Table.Cell(9,2).Range.Text = "<None>"
	}
	Else
	{
		$tmp = ""
		ForEach($FSMORole in $FSMORoles)
		{
			$tmp += ($FSMORole.ToString() + "`n")
		}
		$Table.Cell(9,2).Range.Text = $tmp
	}
	$Table.Cell(10,1).Range.Text = "Partitions"
	$Partitions = $DC.Partitions | Sort
	If($Partitions -eq $Null)
	{
		$Table.Cell(10,2).Range.Text = "<None>"
	}
	Else
	{
		$tmp = ""
		ForEach($Partition in $Partitions)
		{
			$tmp += ($Partition + "`n")
		}
		$Table.Cell(10,2).Range.Text = $tmp
	}
	$Table.Cell(11,1).Range.Text = "Site"
	$Table.Cell(11,2).Range.Text = $DC.Site
	If(!$Hardware)
	{
		$Table.Cell(12,1).Range.Text = "IPv4 Address"
		If([String]::IsNullOrEmpty($DC.IPv4Address))
		{
			$Table.Cell(12,2).Range.Text = "<None>"
		}
		Else
		{
			$Table.Cell(12,2).Range.Text = $DC.IPv4Address
		}
		$Table.Cell(13,1).Range.Text = "IPv6 Address"
		If([String]::IsNullOrEmpty($DC.IPv6Address))
		{
			$Table.Cell(13,2).Range.Text = "<None>"
		}
		Else
		{
			$Table.Cell(13,2).Range.Text = $DC.IPv6Address
		}
		$Table.Cell(14,1).Range.Text = "Operating System"
		$Table.Cell(14,2).Range.Text = $DC.OperatingSystem
		$Table.Cell(15,1).Range.Text = "Service Pack"
		$Table.Cell(15,2).Range.Text = $DC.OperatingSystemServicePack
		$Table.Cell(16,1).Range.Text = "Operating System version"
		$Table.Cell(16,2).Range.Text = $DC.OperatingSystemVersion
	}
	$Table.Rows.SetLeftIndent(0,1)
	$table.AutoFitBehavior(1)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	
	If($Hardware)
	{
		If(Test-Connection -ComputerName $DC.name -quiet -EA 0)
		{
			GetComputerWMIInfo $DC.Name
		}
		Else
		{
			Write-Verbose "$(Get-Date): `t`t$($DC.Name) is offline or unreachable.  Hardware inventory is skipped."
			WriteWordLine 0 0 "Server $($DC.Name) was offline or unreachable at "(get-date).ToString()
			WriteWordLine 0 0 "Hardware inventory was skipped."
		}

	}
	$First = $False
}

#organizational units
Write-Verbose "$(Get-Date): Writing OU data by Domain"
$selection.InsertNewPage()
WriteWordLine 1 0 "Organizational Units"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $Forest.RootDomain)
	{
		WriteWordLine 2 0 "OUs in Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "OUs in Domain $($Domain)"
	}
	#get all OUs for the domain
	Try
	{
		$OUs = Get-ADOrganizationalUnit -filter * -Server $Domain -Properties CanonicalName, DistinguishedName, Name -EA 0 | Select CanonicalName, DistinguishedName, Name | Sort CanonicalName
	}
	
	Catch
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
	}
	
	If($? -and $OUs -ne $Null)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 6
		If($OUs -is [array])
		{
			[int]$Rows = $OUs.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 1
		$table.Borders.OutsideLineStyle = 1
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Name"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Created"
		$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,3).Range.Font.Bold = $True
		$Table.Cell(1,3).Range.Text = "Protected"
		$Table.Cell(1,4).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,4).Range.Font.Bold = $True
		$Table.Cell(1,4).Range.Text = "# Users"
		$Table.Cell(1,5).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,5).Range.Font.Bold = $True
		$Table.Cell(1,5).Range.Text = "# Computers"
		$Table.Cell(1,6).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,6).Range.Font.Bold = $True
		$Table.Cell(1,6).Range.Text = "# Groups"
		[int]$xRow = 1

		ForEach($OU in $OUs)
		{
			$xRow++
			If($xRow % 2 -eq 0)
			{
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray05
			}
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName)"
			#WriteWordLine 3 0 $OUDisplayName
			
			#get data for the individual OU
			Try
			{
				$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * -EA 0
			}
			
			Catch
			{
				Write-Warning "Error retrieving OU data for OU $($OU.CanonicalName)."
			}
			
			If($? -and $OUInfo -ne $Null)
			{
				#get counts of users, computers and groups in the OU
    			Write-Verbose "$(Get-Date): `t`t`tGetting user count"
				
				[int]$UserCount = 0
				[int]$ComputerCount = 0
				[int]$GroupCount = 0
				
				$Results = Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
				If($Results -eq $Null)
				{
					$UserCount = 0
				}
                ElseIf($Results -is [array])
                {
                    $UserCount = $Results.Count
                }
				Else
				{
					$UserCount = 1
				}
    			Write-Verbose "$(Get-Date): `t`t`tGetting computer count"
				$Results = Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
				If($Results -eq $Null)
				{
					$ComputerCount = 0
				}
                ElseIf($Results -is [array])
                {
                    $ComputerCount = $Results.Count
                }
				Else
				{
					$ComputerCount = 1
				}
    			Write-Verbose "$(Get-Date): `t`t`tGetting group count"
				$Results = Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain
				If($Results -eq $Null)
				{
					$GroupCount = 0
				}
                ElseIf($Results -is [array])
                {
                    $GroupCount = $Results.Count
                }
				Else
				{
					$GroupCount = 1
				}
				$UserCount = "{0:N0}" -f $UserCount
				$ComputerCount = "{0:N0}" -f $ComputerCount
				$GroupCount = "{0:N0}" -f $GroupCount

				$Table.Cell($xRow,1).Range.Text = $OUDisplayName
				$Table.Cell($xRow,2).Range.Text = $OUInfo.Created
				If($OUInfo.ProtectedFromAccidentalDeletion -eq $True)
				{
					$Table.Cell($xRow,3).Range.Text = "Yes"
				}
				Else
				{
					$Table.Cell($xRow,3).Range.Text = "No"
				}
				$Table.Cell($xRow,4).Range.Text = $UserCount
				$Table.Cell($xRow,5).Range.Text = $ComputerCount
				$Table.Cell($xRow,6).Range.Text = $GroupCount
			}
			ElseIf(!$?)
			{
				Write-Error "Error retrieving OU data for OU $($OU.CanonicalName)"
			}
			Else
			{
				$Table.Cell($xRow,1).Range.Text = "<None>"
				$Table.Cell($xRow,2).Range.Text = "<None>"
				$Table.Cell($xRow,3).Range.Text = "<None>"
				$Table.Cell($xRow,4).Range.Text = "<None>"
				$Table.Cell($xRow,5).Range.Text = "<None>"
				$Table.Cell($xRow,6).Range.Text = "<None>"
			}
		}
		$Table.Rows.SetLeftIndent(0,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving OU data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No OU data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
	$First = $False
}

#Group information
Write-Verbose "$(Get-Date): Writing group data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Groups"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing groups in domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $Forest.RootDomain)
	{
		WriteWordLine 2 0 "Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "Domain $($Domain)"
	}

	#get all Groups for the domain
	Try
	{
		$Groups = Get-ADGroup -Filter * -Server $Domain -Properties Name, GroupCategory, GroupType -EA 0 | Sort Name
	}
	
	Catch
	{
		Write-Warning "Error retrieving group data for domain $($Domain)."
	}
	
	If($? -and $Groups -ne $Null)
	{
		#get counts
		
		Write-Verbose "$(Get-Date): `t`tGetting counts"
		
		[int]$SecurityCount = 0
		[int]$DistributionCount = 0
		[int]$GlobalCount = 0
		[int]$UniversalCount = 0
		[int]$DomainLocalCount = 0
		
		$Results = $groups | Where {$_.groupcategory -eq "Security"}
		
		If($Results -eq $Null)
		{
			[int]$SecurityCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$SecurityCount = $Results.Count
		}
		Else
		{
			[int]$SecurityCount = 1
		}
		
		$Results = $groups | Where {$_.groupcategory -eq "Distribution"}
		
		If($Results -eq $Null)
		{
			[int]$DistributionCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$DistributionCount = $Results.Count
		}
		Else
		{
			[int]$DistributionCount = 1
		}

		$Results = $groups | Where {$_.groupscope -eq "Global"}

		If($Results -eq $Null)
		{
			[int]$GlobalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$GlobalCount = $Results.Count
		}
		Else
		{
			[int]$GlobalCount = 1
		}

		$Results = $groups | Where {$_.groupscope -eq "Universal"}

		If($Results -eq $Null)
		{
			[int]$UniversalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$UniversalCount = $Results.Count
		}
		Else
		{
			[int]$UniversalCount = 1
		}
		
		$Results = $groups | Where {$_.groupscope -eq "DomainLocal"}

		If($Results -eq $Null)
		{
			[int]$DomainLocalCount = 0
		}
		ElseIf($Results -is [array])
		{
			[int]$DomainLocalCount = $Results.Count
		}
		Else
		{
			[int]$DomainLocalCount = 1
		}

		[int]$TotalCount = "{0:N0}" -f ($SecurityCount + $DistributionCount)
		$SecurityCount = "{0:N0}" -f $SecurityCount
		$DomainLocalCount = "{0:N0}" -f $DomainLocalCount
		$GlobalCount = "{0:N0}" -f $GlobalCount
		$UniversalCount = "{0:N0}" -f $UniversalCount
		$DistributionCount = "{0:N0}" -f $DistributionCount
		
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 6
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Range.Text = "Total Groups"
		$Table.Cell(1,2).Range.Text = $TotalCount
		$Table.Cell(2,1).Range.Text = "`tSecurity Groups"
		$Table.Cell(2,2).Range.Text = $SecurityCount
		$Table.Cell(3,1).Range.Text = "`t`tDomain Local"
		$Table.Cell(3,2).Range.Text = $DomainLocalCount
		$Table.Cell(4,1).Range.Text = "`t`tGlobal"
		$Table.Cell(4,2).Range.Text = $GlobalCount
		$Table.Cell(5,1).Range.Text = "`t`tUniversal"
		$Table.Cell(5,2).Range.Text = $UniversalCount
		$Table.Cell(6,1).Range.Text = "`tDistribution Groups"
		$Table.Cell(6,2).Range.Text = $DistributionCount

		$Table.Rows.SetLeftIndent(0,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		
		#get members of privileged groups
		
		WriteWordLine 0 0 "Privileged Groups"
		Write-Verbose "$(Get-Date): `t`tListing domain admins"
		WriteWordLine 0 1 "Domain Admins:" -NoNewLine
		$Admins = Get-ADGroupMember -Identity "Domain Admins" -Server $Domain -EA 0
		If($? -and $Admins -ne $Null)
		{
			WriteWordLine 0 0 ""
			ForEach($Admin in $Admins)
			{
				WriteWordLine 0 2 $Admin.Name
			}
		}
		Else
		{
				WriteWordLine 0 0 "<None>"
		}

		Write-Verbose "$(Get-Date): `t`tListing enterprise admins"
		WriteWordLine 0 1 "Enterprise Admins:" -NoNewLine
		
		Try
		{
			$Admins = Get-ADGroupMember -Identity "Enterprise Admins" -Server $Domain -EA 0
		}
		
		Catch
		{
			#no enterprise admins in this domain
		}
		
		If($? -and $Admins -ne $Null)
		{
			WriteWordLine 0 0 ""
			ForEach($Admin in $Admins)
			{
				WriteWordLine 0 2 $Admin.Name
			}
		}
		Else
		{
				WriteWordLine 0 0 "<None>"
		}

		Write-Verbose "$(Get-Date): `t`tListing schema admins"
		WriteWordLine 0 1 "Schema Admins:" -NoNewLine
		
		Try
		{
			$Admins = Get-ADGroupMember -Identity "Schema Admins" -Server $Domain -EA 0
		}
		
		Catch
		{
			#no schema admins in this domain
		}
		
		If($? -and $Admins -ne $Null)
		{
			WriteWordLine 0 0 ""
			ForEach($Admin in $Admins)
			{
				WriteWordLine 0 2 $Admin.Name
			}
		}
		Else
		{
				WriteWordLine 0 0 "<None>"
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving Group data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Group data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
	$First = $False
}

#GPOs by domain
Write-Verbose "$(Get-Date): Writing domain group policy data"

$selection.InsertNewPage()
WriteWordLine 1 0 "Group Policies by Domain"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing group policies for domain $($Domain)"

	Try
	{
		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
	}
	
	Catch
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
	}
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $Forest.RootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Write-Verbose "$(Get-Date): `t`tGetting linked GPOs"
		WriteWordLine 0 0 ""
		WriteWordLine 0 0 "Linked Group Policy Objects: " -NoNewLine
		$LinkedGPOs = $DomainInfo.LinkedGroupPolicyObjects | Sort
		If($LinkedGpos -eq $Null)
		{
			WriteWordLine 0 0 "<None>"
		}
		Else
		{
			WriteWordLine 0 0 ""
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			If($LinkedGpos -is [array])
			{
				[int]$Rows = $LinkedGpos.Count
			}
			Else
			{
				[int]$Rows = 1
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 0
			$table.Borders.OutsideLineStyle = 0
			[int]$xRow = 0
			ForEach($LinkedGpo in $LinkedGpos)
			{
				$xRow++
				#taken from Michael B. Smith's work on the XenApp 6.x scripts
				#this way we don't need the GroupPolicy module
				$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
				If($gpObject.DisplayName -eq $Null)
				{
					$p1 = $LinkedGPO.IndexOf("{")
					#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
					$GUID = $LinkedGPO.SubString($p1,38)
					$tmp = "GPO with GUID $($GUID) was not found in this domain"
				}
				Else
				{
					$tmp = $gpObject.DisplayName	### name of the group policy object
				}
				$Table.Cell($xRow,1).Range.Text = $tmp
			}
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
		}
		$LinkedGPOs = $Null
		$First = $False
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
}

#group policies by organizational units
Write-Verbose "$(Get-Date): Writing Group Policy data by Domain by OU"
$selection.InsertNewPage()
WriteWordLine 1 0 "Group Policies by Organizational Unit"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
	If(!$First)
	{
		#put each domain, starting with the second, on a new page
		$selection.InsertNewPage()
	}
	If($Domain -eq $Forest.RootDomain)
	{
		WriteWordLine 2 0 "Group Policies by OUs in Domain $($Domain) (Forest Root)"
	}
	Else
	{
		WriteWordLine 2 0 "Group Policies by OUs in Domain $($Domain)"
	}
	#get all OUs for the domain
	Try
	{
		$OUs = Get-ADOrganizationalUnit -filter * -Server $Domain -Properties CanonicalName, DistinguishedName, Name -EA 0 | Select CanonicalName, DistinguishedName, Name | Sort CanonicalName
	}
	
	Catch
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
	}
	
	If($? -and $OUs -ne $Null)
	{
		ForEach($OU in $OUs)
		{
			$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
			Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName)"
			
			#get data for the individual OU
			Try
			{
				$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * -EA 0
			}
			
			Catch
			{
				Write-Warning "Error retrieving OU data for OU $($OU.CanonicalName)."
			}
			
			If($? -and $OUInfo -ne $Null)
			{
    			Write-Verbose "$(Get-Date): `t`t`tGetting linked GPOs"
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "$($OUDisplayName) Linked Group Policy Objects: " -NoNewLine
				$LinkedGPOs = $OUInfo.LinkedGroupPolicyObjects | Sort
				If($LinkedGpos -eq $Null)
				{
					WriteWordLine 0 0 "<None>"
				}
				Else
				{
					WriteWordLine 0 0 ""
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 1
					If($LinkedGpos -is [array])
					{
						[int]$Rows = $LinkedGpos.Count
					}
					Else
					{
						[int]$Rows = 1
					}
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$table.Style = $myHash.Word_TableGrid
					$table.Borders.InsideLineStyle = 0
					$table.Borders.OutsideLineStyle = 0
					[int]$xRow = 0
					ForEach($LinkedGpo in $LinkedGpos)
					{
						$xRow++
						#taken from Michael B. Smith's work on the XenApp 6.x scripts
						#this way we don't need the GroupPolicy module
						$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
						If($gpObject.DisplayName -eq $Null)
						{
							$p1 = $LinkedGPO.IndexOf("{")
							#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
							$GUID = $LinkedGPO.SubString($p1,38)
							$tmp = "GPO with GUID $($GUID) was not found in this domain"
						}
						Else
						{
							$tmp = $gpObject.DisplayName	### name of the group policy object
						}
						$Table.Cell($xRow,1).Range.Text = $tmp
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
				Write-Error "Error retrieving OU data for OU $($OU.CanonicalName)"
			}
			Else
			{
				$Table.Cell($xRow,1).Range.Text = "<None>"
				$Table.Cell($xRow,2).Range.Text = "<None>"
				$Table.Cell($xRow,3).Range.Text = "<None>"
				$Table.Cell($xRow,4).Range.Text = "<None>"
				$Table.Cell($xRow,5).Range.Text = "<None>"
				$Table.Cell($xRow,6).Range.Text = "<None>"
			}
		}
		$Table.Rows.SetLeftIndent(0,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving OU data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No OU data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
	$First = $False
}

#misc info by domain
Write-Verbose "$(Get-Date): Writing miscellaneous data by domain"

$selection.InsertNewPage()
WriteWordLine 1 0 "Miscellaneous data by Domain"
$First = $True

ForEach($Domain in $Domains)
{
	Write-Verbose "$(Get-Date): `tProcessing miscellaneous data for domain $($Domain)"

	Try
	{
		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
	}
	
	Catch
	{
		Write-Warning "Error retrieving domain data for domain $($Domain)."
	}
	
	If($? -and $DomainInfo -ne $Null)
	{
		If(!$First)
		{
			#put each domain, starting with the second, on a new page
			$selection.InsertNewPage()
		}
		
		If($Domain -eq $Forest.RootDomain)
		{
			WriteWordLine 2 0 "$($Domain) (Forest Root)"
		}
		Else
		{
			WriteWordLine 2 0 $Domain
		}

		Write-Verbose "$(Get-Date): `t`tGathering user misc data"
		
		$Users = Get-ADUser -Filter * -Server $Domain -EA 0
		
		If($? -and $Users -ne $Null)
		{
		
			If($Users -is [array])
			{
				[int]$UsersCount = $Users.Count
			}
			Else
			{
				[int]$UsersCount = 1
			}
			
			Write-Verbose "$(Get-Date): `t`t`tUsers cannot change password"
			$Results = $Users | Where {$_.CannotChangePassword -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersCannotChangePassword = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersCannotChangePassword = $Results.Count
			}
			Else
			{
				[int]$UsersCannotChangePassword = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tDisabled users"
			$Results = $Users | Where {$_.Enabled -eq $False}
		
			If($Results -eq $Null)
			{
				[int]$UsersDisabled = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersDisabled = $Results.Count
			}
			Else
			{
				[int]$UsersDisabled = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tLocked out users"
			$Results = $Users | Where {$_.LockedOut -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersLockedOut = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersLockedOut = $Results.Count
			}
			Else
			{
				[int]$UsersLockedOut = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tPassword expired"
			$Results = $Users | Where {$_.PasswordExpired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordExpired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordExpired = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordExpired = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tPassword never expires"
			$Results = $Users | Where {$_.PasswordNeverExpires -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordNeverExpires = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordNeverExpires = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordNeverExpires = 1
			}

			Write-Verbose "$(Get-Date): `t`t`tPassword not required"
			$Results = $Users | Where {$_.PasswordNotRequired -eq $True}
		
			If($Results -eq $Null)
			{
				[int]$UsersPasswordNotRequired = 0
			}
			ElseIf($Results -is [array])
			{
				[int]$UsersPasswordNotRequired = $Results.Count
			}
			Else
			{
				[int]$UsersPasswordNotRequired = 1
			}
		}
		Else
		{
			[int]$UsersCount = 0
			[int]$UsersDisabled = 0
			[int]$UsersLockedOut = 0
			[int]$UsersPasswordExpired = 0
			[int]$UsersPasswordNeverExpires = 0
			[int]$UsersPasswordNotRequired = 0
			[int]$UsersCannotChangePassword = 0
		}
		
		$TableRange   = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 7
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = 0
		$table.Borders.OutsideLineStyle = 0
		$Table.Cell(1,1).Range.Text = "Total Users"
		$Table.Cell(1,2).Range.Text = $UsersCount
		$Table.Cell(2,1).Range.Text = "Disabled users"
		$Table.Cell(2,2).Range.Text = $UsersDisabled
		$Table.Cell(3,1).Range.Text = "Locked out users"
		$Table.Cell(3,2).Range.Text = $UsersLockedOut
		$Table.Cell(4,1).Range.Text = "Password expired"
		$Table.Cell(4,2).Range.Text = $UsersPasswordExpired
		$Table.Cell(5,1).Range.Text = "Password never expires"
		$Table.Cell(5,2).Range.Text = $UsersPasswordNeverExpires
		$Table.Cell(6,1).Range.Text = "Password not required"
		$Table.Cell(6,2).Range.Text = $UsersPasswordNotRequired
		$Table.Cell(7,1).Range.Text = "Can't change password"
		$Table.Cell(7,2).Range.Text = $UsersCannotChangePassword

		$Table.Rows.SetLeftIndent(0,1)
		$table.AutoFitBehavior(1)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving domain data for domain $($Domain)" "" $null 0 $False $True
	}
	Else
	{
		WriteWordLine 0 0 "No Domain data was retrieved for domain $($Domain)" "" $null 0 $False $True
	}
	$First = $False
}

Write-Verbose "$(Get-Date): Finishing up Word document"
#end of document processing

#Update document properties
If($CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "Active Directory Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp = $doc.CustomXMLParts | Where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract = "Microsoft Active Directory Inventory for $CompanyName"
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
	#Word 2007
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2") -or $RunningOS.Contains("Server 2012"))
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
	Remove-Item $filename1
}
Write-Verbose "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	Write-Verbose "$(Get-Date): $($filename2) is ready for use"
}
Else
{
	Write-Verbose "$(Get-Date): $($filename1) is ready for use"
}
Write-Verbose "$(Get-Date): "

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