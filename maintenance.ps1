$componentFiles = Get-ChildItem -Name -Include *.xlsx

Write-output "# Verkada SystemSurveyor Element Profiles`n" | Out-File ./README.md -Force

foreach ($comp in $componentFiles){
	# Import Excel File
	$Rows = Import-Excel -Path ./$comp -StartRow 7 -EndRow 10 -StartColumn 4 -NoHeader -DataOnly

	# Get all column names
	$AllColumns = $Rows[0].psobject.Properties.Name
	$LabelColumn = $AllColumns[0]
	$DeviceColumns = $AllColumns | Select-Object -Skip 1

	# Array to hold all final transposed device objects
	$TransposedDevices = [System.Collections.Generic.List[PSCustomObject]]::new()

	# Loop through each device column dynamically
	foreach ($DeviceCol in $DeviceColumns) {
			# Initialize an ordered dictionary to keep row order intact
			$DeviceData = [ordered]@{}
			
			# Process every row for the current device column
			foreach ($Row in $Rows) {
					$Key = $Row.$LabelColumn
					
					# Skip empty label rows
					if (-not [string]::IsNullOrWhiteSpace($Key)) {
							$DeviceData[$Key] = $Row.$DeviceCol
					}
			}
			
			# Convert the collected properties into an object and add to our list
			$TransposedDevices.Add([PSCustomObject]$DeviceData)
	}
	# Get title case for component from file name
	$header = (Get-Culture).TextInfo.ToTitleCase($comp.ToLower().Replace("_"," ").Replace(".xlsx",""))
	Write-output "* [$header]($comp)" | Out-File ./README.md -Append 

	# Get models in array
	$models = $TransposedDevices | Where-Object {!([string]::IsNullOrEmpty($_.'Component Model #'))} | Select-Object 'Component Model #' -ExpandProperty 'Component Model #'
	foreach ($model in $models){
		Write-output "  * $model" | Out-File ./README.md -Append
	}
}