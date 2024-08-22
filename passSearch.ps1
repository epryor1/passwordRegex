$directoryPath = 'your\file\path' # search location
	$outFile = "your\file\path" # where the report will go
	$passRules = 'your regex'  # password rules you are searching for
	$files = Get-ChildItem -Path $directoryPath -Recurse -File -Include *.docx, *.txt, *.xlsx
	
	Clear-Content -Path $outFile
	
	#check if files are being recived
	if ($files.Count -eq 0) {
		Write-Output "No files found"

	} else {
		Write-Output "$($files.Count) files found"
	}
	
	foreach($file in $files) {
		try  {
			$content = ""
			
			Write-Output "Processing file $($file.FullName)"
			
			# read file based on file extension
			switch ($file.Extension.ToLower()) {
				".docx" {
				#	Write-Output "Attempting to read word documnent..."
					$word = New-Object -ComObject Word.Application
					$word.Visible = $false
					$doc = $word.Documents.Open($file.FullName)
					
					if ($doc -ne $null) {
						$content = $doc.Content.Text
						$doc.Close()
					#	Write-Output "Sucessfully read word document"
					} else {
						Write-Output "Failed to open Word document: $($file.FullName)"
					}
					$word.Quit()
				}
				
				".xlsx" {
				#	Write-Output "Attempting to read excel documnent..."
					$excel = New-Object -ComObject Excel.Application
					$excel.Visible = $false
					$workbook = $excel.Workbooks.Open($file.FullName)
					
					if ($workbook -ne $null) {
						foreach ($sheet in $workbook.Sheets) {
							$range = $sheet.UsedRange
							if ($range -ne $null) {
								$content += $range.Text
							#	Write-Output "Processed sheet: $($sheet.Name). COntent Length: $($range.Text.Length)"
							} else {
								Write-Output "Failed to get range for sheet: $($sheet.Name)"
							}
						}
						
						$workbook.Close()
					#	Write-Output "Sucessfully read excel document"
					} else {
						Write-Output "Failed to open Excel document: $($file.FullName)"
					}
					
					$excel.Quit()
					
				}
				
				".txt" {
				#	Write-Output "Attempting to read text documnent..."
					$content = Get-Content -Path $file.FullName -Raw
				#	Write-Output "Sucessfully read text document"
				}
				
				default {
					Write-Output "Unsuported file type: $($file.Extention)"
					continue
				}
			}
			
			
			
			if ($content -notmatch $urlPattern -and $content -match $passRules) {
				Write-Output "File: $($file.FullName)" |Out-File -FilePath $outFile -Append
				Write-Output "Possible match found: " |Out-File -FilePath $outFile -Append
				$matches |Out-File -FilePath $outFile -Append
				Write-Output "------------------------------------------------------" |Out-File -FilePath $outFile -Append
				Write-Output " " |Out-File -FilePath $outFile -Append
			} 
			
		} catch {
		Write-Output "Error processing file: $($file.FullName). Error: $_" 
		}
	}
	Write-Output "Finished. Results can be found in $($outFile)."
