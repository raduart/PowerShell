<#########################################################
    .SCRIPT SYNOPSIS 
        Convert CGD bank statement to QIF
        https://en.wikipedia.org/wiki/Quicken_Interchange_Format

        	  
	.Description
        This script inputs a csv statement file 
        Selects Rows to be converted
        And Outputs a converted QIF file.

	.Parameter InHelp
		Optional: This item will display syntax help
		Alias: H

	.Parameter InFileIn
		Optional: This item is the CSV file to be converted
		Alias: I

	.Parameter InFileOut
		Optional: This item is the QIF file convertd
		Alias: O

	.Example
		.\CsvToQif.ps1 -I ".\fileIn.csv" -O ".\fileOut.qif" 

    .Author  
        Rafael Duarte
		Created By Rafael Duarte
		Email raduart@microsoft.com		

    .Credits

    .Notes / Versions / Output
        * Version: 1.2
          Date: April 29th 2019
          Purpose/Change:	
            > Selected data has to be contiguous
            > Display Initial and final balance of selected data
            > funtion ErrorMsgCentral() is restrutured
        * Version: 1.1
          Date: April 29th 2019
          Purpose/Change:	Adding total rows selected
    	* Version: 1.0
          Date: April 22th 2019
          Purpose/Change:	Initial function development
          # Constrains / Pre-requisites:
            > none
          # Output
            > Creates a Transcript File (<ScriptName>_<TrackTimeStamp>.txt)
            > Creates a fileout.qif to be used on MS Money
    .Link
        https://github.com/raduart/PowerShell/tree/master/CsvToQif

#########################################################>


	Param(
	[Parameter(Mandatory=$false)][Alias('H')][Switch]$InHelp,
	[Parameter(Mandatory=$false)][Alias('I')][String]$InFileIn = "",
	[Parameter(Mandatory=$false)][Alias('O')][String]$InFileOut = "")


<####################### Function ########################
    .Function SYNOPSIS - ErrorMsgCentral
      Displays a custom message to console output depending on MsgID
	  
	.Description
	  	This function helps to centralize all custom messages to console output
	  	depending on MsgID selected.

	.Parameter MsgID
		Mandatory: This item idenfify message to be displayed on console output
		Alias: ID

	.Parameter MsgType
		Mandatory: This item idenfify type of message to be displayed on console output 
                    E - Error or End Script
                    W - Warning
                    I - Information
		Alias: Type

	.Parameter MsgData
		Optional: Additional data that can be used when displaying message to console output
		Alias: Data

	.Example
		ErrorMsgCentral -ID 10 -Type "E" -Data "Demo"
		
		This example will output message, type Error that is assigned to ID 10 and may use "Demo" string
        to be added on Message ID selected.

	.Notes
		Created By Rafael Duarte
		Email raduart@microsoft.com		

        Version: 1.1
        Date: April 29th 2019
        Purpose/Change:	
            > funtion ErrorMsgCentral() is restrutured
        Version: 1.0
        Date: April 22th 2019
        Purpose/Change:	Initial function development

#########################################################>

function ErrorMsgCentral{
	Param(
	[Parameter(Mandatory=$True)][Alias('ID')][Int32]$MsgID,
	[Parameter(Mandatory=$True)][Alias('Type')][String]$MsgType,
	[Parameter(Mandatory=$False)][Alias('Data')][String]$MsgData)

    $MsgTxt = "`n<$MsgType$MsgID> - "
    switch ($MsgID) 
    { 
        0   {$MsgTxt += "END SCRIPT`n"}
        5   {$MsgTxt += (-join (
                                "Syntax:`n",
                         "==============`n",
		                        ".\$($ScriptName).ps1 -H | -I fileIn.csv [-O fileOut.qif]"
                               )
                        )
            }
        10  {$MsgTxt += (-join (
                                "Error:`n",
                        "==============`n",
		                        "Missing data !`n",
                                "$MsgData`n"
                               )
                        )
            }
        20  {$MsgTxt += (-join (
                                "Error:`n",
                        "==============`n",
		                        "Invalid data !`n",
                                "$MsgData`n"
                               )
                        )
            }
        90  {$MsgTxt += (-join (
                                "Information\Warning:`n",
                        "============================`n",
		                        "$MsgData`n"
                               )
                        )
            }
   default  {$MsgTxt += (-join (
                                "Error:`n",
                        "==============`n",
		                        "Error unknown !!!`n"
                               )
                        )
             $MsgType = "E"
            }
    }

    #Write-Host "`n<$MsgType$MsgID>" -ForegroundColor Yellow 
    switch ($MsgType)
    {
        I {Write-Host $MsgTxt -ForegroundColor Green}
        W {Write-Host $MsgTxt -ForegroundColor Yellow}
        E {Write-Host $MsgTxt -ForegroundColor Red}
    }

    If ($MsgType -eq "E")
    {
        Write-Host "`n####### End - PoSH script $ScriptName.ps1 #######" -ForegroundColor Green
        Stop-Transcript
    }
}
# ################### End Function ########################

<####################### Function ########################
    .Function SYNOPSIS - PayeeNormalize
      Replaces some Payees names for normalized one's
	  
	.Description
	  	This function receives a Payee name search for a normalized one and replace it.

	.Parameter InPayee
		Mandatory: Payee name to normlize
		Alias: Payee

	.Example
		PayeeNormalize -Payee "Payee Name text"
		
		This example will return a payee name normalized

	.Notes
		Created By Rafael Duarte
		Email raduart@microsoft.com		

		Version: 1.0
		Date: April 22th 2019
		Purpose/Change:	Initial function development

    .Link

#########################################################>

function PayeeNormalize{
	Param(
    [Parameter(Mandatory=$True)][Alias('Payee')][String]$InPayee)
    
    # Table for normalize
    $TNormalize = [ordered]@{ 
                    "LEVANTAMENTO"                               = "LEV"; 
                    "COMPRA CENTRO CLINICO"                      = "CCSAO LUCAS";
                    "COMPRA LA DOLCE VITA"                       = "REST LA DOLCE VITA";
                    "COMPRA 1986 LA DOLCE VITA 1990-203 LISBO"   = "REST LA DOLCE VITA";
                    "COMPRA ORQUESTRA DE P"                      = "REST ADAMASTOR";
                    "COMPRA 1986 ORQUESTRA DE PANELASLISBOA"     = "REST ADAMASTOR";
                    "COMPRA 1986 ORQUESTRA DE PANELASLOTE"       = "REST ADAMASTOR";
                    "COMPRA RARO   ORIGINA"                      = "REST DEL REI";
                    "COMPRA 1986 "                               = "";
                    "COMPRA"                                     = "";
                    "PAGAMENTO"                                  = "PAG";
                    "DESPESAS0614002132400"                      = "PENSAO FILHOS";
                    "BX VALOR 03 TRANSACCO"                      = "VIA VERDE";
                    "R C SANCHES -"                              = "VITAMINAS";
                    "APOS AUTORIZAC"                             = "MBNET";
                    "TRF P2P 966XXX646"                          = "MBNET TRANSF CLAUDIA";
                    "TRF P2P"                                    = "MBNET TRANSF";
                   }
    $InPayee = $InPayee.ToUpper()    
    foreach ($item in $TNormalize.Keys)
    {
        $InPayee = $InPayee.Replace($item, $TNormalize.($item))
    }
    
    $text= ""
    foreach ($item in ($InPayee.split()))
    {
        If ($item -ne "")
        {
            $text += "$item "
        }
    }

    $InPayee = $text.Trim()
    return $InPayee
}
# ################### End Function ########################


### Parameters / Constants ###
## Get Script Name
    # invocation from POSH Command Line
    $ScriptName = $MyInvocation.MyCommand.Name
    if (($ScriptName -eq $null) -or ($ScriptName -eq ""))
    {
        # invocation from POSH ISE Environment
        $ScriptName = ($psISE.CurrentFile.DisplayName).Replace("*","")
    }
    $ScriptName = $ScriptName.Replace(".ps1","")
    $ScriptPath = $psISE.CurrentPowerShellTab.Prompt.Replace("PS ","")
    $ScriptPath = $ScriptPath.Replace("> ","")

## Files / logs / Paths
    $TrackTimeStamp = "$('{0:yyyyMMddHHmmss}_{1,-1}' -f $(Date), $(Get-Random))" 
    $TranscriptPath = ".\"
    if (!(Test-Path -LiteralPath $TranscriptPath -PathType Container)) 
        {Invoke-Command -ScriptBlock {md $TranscriptPath}}
    $LogPath        = ".\"
    if (!(Test-Path -LiteralPath $LogPath -PathType Container)) 
        {Invoke-Command -ScriptBlock {md $LogPath}}
    $TranscriptFile = $TranscriptPath + $ScriptName + "_" + $TrackTimeStamp + ".log"

### Main Script ###

## Track Log
    Start-Transcript $TranscriptFile

## Begin Script
    Write-Host "`n####### Begin - PoSH script $ScriptName.ps1 #######`n" -ForegroundColor Green

## Parameters Validation
    If ($InHelp)
    {
        ErrorMsgCentral -ID 5 -Type "E"
        Throw
    }
    
    $FileIn = $InFileIn
    $FileOut = $InFileOut

    If ($FileIn -eq "")
    {
        ErrorMsgCentral -ID 10 -Type "E" -MsgData "Missing Input CSV file!"
        Throw
    }
    If ($FileIn.Split(".")[-1] -ne "csv")
    {
        ErrorMsgCentral -ID 10 -Type "E" -MsgData "Input file is not CSV extension!"
        Throw
    }

    If ($FileOut -eq "")
    {
        $FileOut = ($FileIn.Split(".")[-2]).Replace("\","") + ".qif"
    }


$NumDataRows=((get-content $FileIn | select-object -skip 7).count)-3
$Header = 'Data mov', 'Data valor', 'Descrição', 'Débito', 'Crédito', 'Saldo contabilístico', 'Saldo disponível', 'Categoria' 
$RawDataFileCsv=(get-content $FileIn | select-object -Skip 7 -first $NumDataRows | ConvertFrom-Csv -Delimiter ";" -Header $Header)

$RowNum = 1
$RawDataFileCsv = ($RawDataFileCsv |
                  Select-Object @{ Name = 'RowNum' ; Expression= {(([ref]$RowNum).Value++)} }, 'Data mov', 'Data valor', 'Descrição', 'Débito', 'Crédito', 'Saldo contabilístico', 'Saldo disponível', 'Categoria' )

$SelectedDataCsv = $RawDataFileCsv | Out-GridView -PassThru -Title "Bank statement"

# Cannot select only one row. One Row is equal to select all rows.
if ($SelectedDataCsv.Count -lt 0)
{
    $SelectedDataCsv = $RawDataFileCsv
    ErrorMsgCentral -ID 90 -Type "I" -MsgData ">>>>> All file content selected !!!"
}
$TotalRows = $SelectedDataCsv.Count
$SelectedDataCsv = $SelectedDataCsv | Sort-Object 'RowNum' 

#Checking that selected Data is contiguous
$ControlRowNum = $SelectedDataCsv[0].RowNum + $SelectedDataCsv.Count - 1 
$EndRowNum = $SelectedDataCsv[-1].RowNum
If ($ControlRowNum -ne $EndRowNum)
{
    ErrorMsgCentral -ID 20 `
                    -Type "E" `
                    -MsgData (-join(
                                    "Selected data isn't contiguous! `n",
                                    "$TotalRows rows selected of $($SelectedDataCsv[-1].RowNum - $SelectedDataCsv[0].RowNum + 1) rows expected!`n",
                                    "From first row #$($SelectedDataCsv[0].RowNum) to last row #$EndRowNum (inclusive)"
                                  ))
                    
    Throw
}


if (Test-Path -LiteralPath $FileOut)
   {
    if (Test-Path -LiteralPath "$FileOut.old") 
    {
        Invoke-Command -ScriptBlock {del "$FileOut.old"}
        ErrorMsgCentral -ID 90 -Type "W" -MsgData "File <$ScriptPath\$FileOut.old> deleted !"
    }
    Invoke-Command -ScriptBlock {ren $FileOut "$FileOut.old"}
    ErrorMsgCentral -ID 90 `
                    -Type "W" `
                    -MsgData (-join(
                                    "File <$ScriptPath\$FileOut> `n",
                                    "renamed to <$ScriptPath\$FileOut.old> !`n"
                                   )
                             )

   }
Add-Content $FileOut "!Type:Bank"

foreach ($Item in $SelectedDataCsv)
{
    $DateMov = ($Item.'Data valor').Replace("-","/")
    $Credit = [double](($Item.'Crédito').Replace(".","")).Replace(",",".")
    $Debit = [double](($Item.'Débito').Replace(".","")).Replace(",",".")
    $Amount = $Credit-$Debit
    $Payee = PayeeNormalize -InPayee $Item.'Descrição'

    Add-Content $FileOut "D$DateMov"
    Add-Content $FileOut ("U{0:f2}" -f $Amount).Replace(",",".")
    Add-Content $FileOut ("T{0:f2}" -f $Amount).Replace(",",".")
    Add-Content $FileOut "P$Payee"
    Add-Content $FileOut "^"
}

ErrorMsgCentral -ID 90 `
                -Type "I" `
                -MsgData (-join(
                                "File <$ScriptPath\$FileOut> generated !`n",
                                "Total rows: $TotalRows`n",
                                "Initial Balance: $($SelectedDataCsv[-1].'Saldo contabilístico')`n",
                                "Final   Balance: $($SelectedDataCsv[0].'Saldo contabilístico')"
                               )
                         )
ErrorMsgCentral -ID 0 -Type "I"
Stop-Transcript
