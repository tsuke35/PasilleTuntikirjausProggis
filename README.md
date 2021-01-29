# PasilleTuntikirjausProggis

# Vaatimukset
Install-Module -Name ImportExcel

# Käyttöohje
.\Create-CombinedWorkhourDataFile -CustomerListingExcelFile 'esimerkkiasiakaslista.xlsx' -WorkHoursExcelFile 'esimerkkikirjanpit.xlsx' -ExportFileName 'combinedContent.csv'