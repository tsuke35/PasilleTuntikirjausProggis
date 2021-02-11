Param (
        $CustomerListingExcelFile = 'esimerkkiasiakaslista.xlsx',
        $WorkHoursExcelFile = 'esimerkkikirjanpit.xlsx',
        $ExportFileName = 'combinedContent.xlsx'
        )

# Set working location
Set-Location (Split-Path $MyInvocation.MyCommand.Definition)

# Check for required module
try
{
    Import-Module ImportExcel -ErrorAction Stop
}

catch
{
    $errMsg = "Virhe. Moduuli nimeltä ImportExcel vaaditaan. Asennusohje kts. https://www.powershellgallery.com/packages/ImportExcel"
    Write-Error $errMsg; Start-Sleep -Seconds 10; exit 0
}

# Check that files exist and are able to import using Import-Excel cmdlet
try
{
    $customerListingfilecontent = Import-Excel $CustomerListingExcelFile -HeaderName "1,2,3,4,5".Split(",") -ErrorAction Stop 
    Import-Excel $WorkHoursExcelFile -NoHeader -ErrorAction Stop | Out-Null
}

catch
{
    throw "Pakollinen tiedosto puuttuu. Virhe: $_"
}

###
# START OF 'READ CUSTOMERLISTING EXCEL AND PARSE ITS CONTENTS'
###

$customerListCustomObject = @()

# Add empty custom row in the end of customer file content to easify the foreach loop logic (loop depends that file needs to end to an empty row )
$customerListingfilecontent += New-Object psobject -Property @{"1"=$null}

# Dynamically read all customers from file properties
foreach ($row in $customerListingfilecontent)
{
    # When cell value is "Asiakas", continue iteration to next as it is the first row of the file (header row)
    if ($row.1 -eq "Asiakas")
    {
        continue
    }

    # If cell name is not empty, start new customer
    elseif ($row.1)
    {
        $currentCustomerName = $row.1
        $currentCustomerNumber = $row.2
        
        # Initialize new empty custom object in which all projectnumbers for this customer are added
        $currentCustomerProjectNumbers = @()
        
        $projectProperties = @{ 'ProjectNumber' = $row.3
                                'ProjectName' = $row.4
                                'ProjectReferenceNumber' = $row.5
                              }
        $currentCustomerProjectNumbers += New-Object psobject -Property $projectProperties

    }

    # If first cell (Asiakas) is empty but project values are present
    elseif (!$row.1 -and $row.3)
    {
        # Add project values to existing projectnumbers
        $projectProperties = @{ 
                                'ProjectNumber' = $row.3
                                'ProjectName' = $row.4
                                'ProjectReferenceNumber' = $row.5
                              }

        $currentCustomerProjectNumbers += New-Object psobject -Property $projectProperties
    }

    # if blank row (current customer final row processed) but currentCustomerProjectNumbers not empty yet
    elseif (!$row.1 -and !$row.3 -and $currentCustomerProjectNumbers)
    {
        $customerProperties = @{
                                'CustomerName' = $currentCustomerName
                                'CustomerNumber' = $currentCustomerNumber
                                'CustomerProjects' = $currentCustomerProjectNumbers
                               }
        
        $customerListCustomObject += New-Object psobject -Property $customerProperties

        # nullify currentCustomerProjectNumbers
        $currentCustomerProjectNumbers = $null
    }

    # Do nothing and continue iteration on empty rows
    else
    {
        continue    
    }
}

# Create mapping hashtable of all project reference numbers pointing to base customer
$projectMappingHashtable = @{}
$customerListCustomObject | ForEach-Object {

    $customerName = $_.CustomerName
    $customerNumber = $_.CustomerNumber

    foreach ($project in $_.CustomerProjects)
    {
        $customProjectObjectProperties = @{
                        ProjectName = $project.ProjectName
                        ProjectNumber = $project.ProjectNumber
                        CustomerName = $customerName
                        CustomerNumber = $customerNumber
        }
        
        # Add to hashtable using project number as the key and custom object of customer values as value
        $projectMappingHashtable.Add($project.ProjectReferenceNumber, (New-Object psobject -Property $customProjectObjectProperties))
    }
}

###
# END OF 'READ CUSTOMERLISTING EXCEL AND PARSE ITS CONTENTS'
###





###
# START OF 'READ WORKHOURS EXCEL AND PARSE ITS CONTENTS'
###

$workHourMarkings = @()

foreach ($file in $WorkHoursExcelFile)
{
    # Check object type to enable string and FileInfo inputs
    if ($file.GetType().Name -eq "FileInfo")
    {
        $filePath = $file.FullName
    }

    else
    {
        try
        {
            Test-Path $file -ErrorAction Stop | Out-Null
            $filePath = $file
        }

        catch
        {
            Write-Error "Ei pystytty lukemaan tuntikirjaustiedostoa $($file.ToString()) - Virheilmoitus: $_"
            continue
        }
    }
    
    # Get all worksheets in the file
    $workSheets = Get-ExcelSheetInfo $filePath

    foreach ($workSheet in $workSheets)
    {
        $firstIteration = $true

        # Import current worksheet content
        $currentWorkHoursFileContent = Import-Excel -Path $workSheet.Path -WorksheetName $workSheet.Name -NoHeader

        # Initialize variable
        # Map column property name and project reference number (e.g P5 = 01123456) together to easily find the matching project for columns
        # This hashtable will be constructed during processing of row 2 (if P2 equals "Total")
        $projectNumberMappingToSheetsPropertyNames = @{}

        # Start loop for each row in current work sheet
        foreach ($row in $currentWorkHoursFileContent)
        {
            # If first row and it's empty
            if ($firstIteration -eq $true)
            {
                $lastProperty = $row.PSObject.Properties.Name | Sort-Object -Descending | Select-Object -First 1
                $firstIteration = $false

                # If last property on this row is empty - skip
                if (!($row.$($lastProperty))) 
                {
                    continue
                }
            }
            
            # If third row of file, skip
            elseif ($row.P2 -eq "PVM")
            {
                continue
            }

            # This is the second line of file that defines all the project numbers, save mapping hashtable for later use
            elseif ($row.P2 -eq "Total")
            {
                $i = 3
                # Parse all project numbers included in this worksheet
                do 
                {
                    $currentIterationColumnHeaderName = "P" + $i.ToString()

                    # Select project number value with this header name value
                    $currentValue = $row.$($currentIterationColumnHeaderName)

                    # if value not 'selitys' add project number to hashtable with empty hashtable as value
                    if (!($currentValue -eq 'selitys'))
                    {
                        $projectNumberMappingToSheetsPropertyNames.Add($currentIterationColumnHeaderName, $currentValue)
                    }

                    # Increment i value by one to continue to next column
                    $i++
                }
                
                until ($currentValue -eq 'selitys')
            }

            # If value exists in any of P1 (pvm), P2 (total) or the last property (selitys) -> then parse the row contents
            elseif ($row.$($lastProperty) -or $row.P1 -or $row.P2)
            {
                # Select all property values excluding first, second and last
                $iteratableProperties = $row.PSObject.Properties.Name | Where-Object {$_ -notin "P1","P2",$lastProperty}
                
                foreach ($property in $iteratableProperties)
                {
                    $rowValue = $row.$($property)
                
                    # If cell has value
                    if ($rowValue)
                    {
                        # Get project name based on mapping value of hashtable and the column name (propertyname)
                        $projectReferenceNumber = $projectNumberMappingToSheetsPropertyNames[$property]
                        # Get customer information based on project reference number
                        $customerInformation = $projectMappingHashtable[$projectReferenceNumber]

                        $workHourMarkingProperties = [ordered]@{
                                            'Date' = $row.'P1'                
                                            # replace , with . to enable correct number behaviour
                                            'Hours' = [float]($rowValue -replace ",",".")
                                            'ProjectReferenceNumber' = $projectReferenceNumber
                                            'ProjectName' = $customerInformation.ProjectName
                                            'ProjectNumber' = $customerInformation.ProjectNumber
                                            'CustomerName' = $customerInformation.CustomerName
                                            'CustomerNumber' = $customerInformation.CustomerNumber
                                            'Definition' = $row.$($lastProperty)
                                        }
                        
                        $workHourMarkings += New-Object psobject -Property $workHourMarkingProperties

                    }
                }
            }
            
            # Propably an empty row in the middle of content? dunno why this would ever happen :<
            else
            {
                continue    
            }
        }
    }
}

# If content is null for some reason?
if (!$workHourMarkings)
{
    $errMsg = 'Jotain meni vikaan, koska työaikamerkintöjä ei tullut lainkaan lopulliseen muuttujaan! Ei tehdä excel tiedoston exporttia. Kaadetaan skriptin suoritus..'
    Write-Error $errMsg; Start-Sleep -Seconds 10; throw $errMsg
}

###
# END OF 'READ WORKHOURS EXCEL AND PARSE ITS CONTENTS'
###



###
# START OF 'EXPORT COMBINED DATA INTO NEW WORKBOOK'
###

# Get all unique customers in parsed content
$customers = ($workHourMarkings | Select-Object -Property CustomerName -Unique).CustomerName

foreach ($customer in $customers)
{
    # If file exists already -> Only create new worksheet
    if (Test-Path $ExportFileName)
    {
        # If missing Excel object needed for worksheet addition
        if (!$ExcelExportFileObject)
        {
            $ExcelExportFileObject = Open-ExcelPackage $ExportFileName
        }

        # Add new blank worksheet with 'customer' as the name
        #Add-Worksheet -ExcelPackage $ExcelExportFileObject -WorksheetName $customer -ClearSheet | Out-Null
        Export-Excel -Path $ExportFileName -WorksheetName $customer
    }

    # Otherwise create new xlxs file and set the first worksheet name
    else
    {
        Export-Excel -Path $ExportFileName -WorksheetName $customer
    }
}

# Export file as csv that can be easily opened in Excel
$workHourMarkings | Export-Csv -Path $ExportFileName -Encoding UTF8 -Delimiter ";" -NoTypeInformation