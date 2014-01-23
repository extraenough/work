# some function magic
function get-aurorareport{
	<#
		.Synopsis
		 Get-AuroraReport
		.Description
		.Example
		 Get-AuroraReport [-Dir <string>] [-Mode <string>][-OutDir <string>]
		 Get-AuroraReport [-File <string>] [-Mode <string>][-OutDir <string>]
		.Notes
		 Mode: month, node, full
	 #>
	[cmdletbinding()]
	param(
		# dir for parse
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelineByPropertyName=$True)]
		#[ValidateNotNullOrEmpty()]
		[string[]]$dir,		
		# switch from $month, $full, $node
		[Parameter(Mandatory=$false)]
		#[ValidateNotNullOrEmpty()]
		[string]$mode = [string](""),
		# if required a report about specific month
		[Parameter(Mandatory=$False)]
		[Int32]$month = [Int32](0),
		# scan completed tasks only
		[Parameter(Mandatory=$False)]
		[switch]$completed = $True,
		# out directory $out-dir
		[Parameter(Mandatory=$False)]
		[string]$outdir = [string](""),
		# for tests
		[Parameter(Mandatory=$False)]
		[switch]$report = $False
	)
	BEGIN{
		$location = Get-Location
		[string]$folder = [string]("")
		# суммарное врем€ за год
		[Int64]$time = 0
#		info = new-object -psobject 
		# врем€ дл€ каждого лог-файла
		[Int32]$logcount = [Int32](0)
		[Int64[]]$timeperlog = @()
		[bool[]]$logscan = @()
		# врем€ дл€ каждой ноды
		[Int32]$nodecount = [Int32](0)
		[Int64[]]$timepernode = @()
		[string[]]$nodenames = @()
		# врем€ дл€ каждого мес€ца
		# bdsm-scripting - time !
		$timem1 = @([Int64[]])
		# имена задач
		[string[]]$tasknames = @()		
		
		[string]$folders = @()
		[bool]$flag = $False		
	}
	PROCESS{
		
		if(-not $flag){
			$folder = $dir
			$flag = $True
		}else{
			$folder = $_			
		}
		$folder
		# $folder = $dir
		if($dir.Length -gt [Int32](1)){	Write-Verbose -message "Folder '$location' scanning." }
		else{ Write-Verbose -message "Folder '$dir' scanning." }
		Write-Verbose -message "Start scanning folders."
		# need some table for Write-Verbose with info about folders
		
		$timeperlog += $logcount
		$timem1 += $logcount		
		$tasknames += $logcount		
		
		# parse name of task
		[Int32]$i = [Int32]($folder.Length)
		[string]$substr = [string]("")							
		while([bool]($substr -ne [string](" "))){
			$i--
			$substr =  $folder.Substring($i, 1)
			$tasknames[$logcount] = $folder.Substring($i, 1)+$tasknames[$logcount]								
		}							
		$tasknames[$logcount] = $tasknames[$logcount].Trim()
		
		Write-Verbose -message "Start scanning .xml-files."
		if($completed){
			$logscan += $logcount
			$logscan[$logcount] = $False
			[xml]$xml = get-childitem $folder | select name | where {$_ -like "*.xml*"} | % {get-content ($folder + "\" + $_.Name)}
			[Int32]$taskscomp = select-xml "//JobInfo/TasksCompleted" $xml | % {$_.node.'#text'}	
			if($taskscomp -ne [Int32](0)){			
				$logscan[$logcount] = $True
			}
		}
		if((($taskscomp -ne [Int32](0)) -and ($completed)) -or ( -not $completed)){						
			$log = get-childitem $folder | select name | where {$_ -like "*.log*"}	| % {get-content ($folder + "\" + $_.Name)}
			
			# array magic		
			for([Int32]$i=0; $i -lt 12; $i++){ $timem1[$logcount] = [Int64[]](0, 0, 0, 0 , 0, 0, 0, 0, 0, 0, 0, 0)}
			
			foreach ($logstr in $log) {
				# get strings with PRG
				if($logstr -like "*PRG*"){			
					$monthnum = [Int64]($logstr.Substring(5, 2) - 1)			
					$timeperlog[$logcount] = [Int64]($timeperlog[$logcount] + [Int64]($logstr.Substring(28, 8)))
					$time = [Int64]($time + [Int64]$logstr.Substring(28, 8))
					$timem1[$logcount][$monthnum] = [Int64]($timem1[$logcount][$monthnum] + [Int64]($logstr.Substring(28, 8)))
					
					# some magic
					[string]$tempstr = [string]($logstr.Substring(51, 10))
					if($nodenames -contains $tempstr){				
						for($i = 0; $i -lt $nodenames.Length; $i++){
							if($nodenames[$i] -eq $tempstr){					
								$timepernode[$i] = [Int64]($timepernode[$i] + [Int64]($logstr.Substring(28, 8)))
							}
						}
					}else{
						$nodenames += $nodecount
						$timepernode += $nodecount
						$nodenames[$nodecount] = $tempstr
						$timepernode[$nodecount] = [Int64]($timepernode[$nodecount] + [Int64]($logstr.Substring(28, 8)))
						$nodecount = [Int32]($nodecount + 1)
					}
				}
			}

		}else{
			if($completed){
				$logscan[$logcount] = $False
				$tasknames[$logcount]+' : TasksCompleted is equals zero.'
			}
		}
		$logcount++

		# try
		# $nodenames, $timeperlog, ($timem1.SyncRoot), $logscan | format-table	
	}
	END{
		$input
		# $logcount
		if(($report) -and ($outdir -ne [string]("")) -and ($input.Count -eq $logcount)){
			Write-Verbose -message "Check MS Excel is installed."		
			$regint = [string](get-childitem HKLM:\Software\Classes -ea 0| ? {$_.PSChildName -match '^\w+\.\w+$' -and (get-itemproperty "$($_.PSPath)\CLSID" -ea 0)} | where {$_ -match "Excel.Application"} | % {$_.Name})		
			if($regint -ne [string]("")){
				Write-Verbose -message "MS Excel is installed."
				Write-Verbose -message "Start creating Aurora report."
				$xlinst = New-Object -ComObject "Excel.Application" 			
				# time report
				# if $month or $full switched
				# if($mode -eq [string]("month") -or $mode -eq [string]("full")){
				$xlinst.SheetsInNewWorkbook = 2		
				$wbook = $xlinst.Workbooks.Add()		
				# 2 reports: year and month
				$wbook.worksheets.item(1).Name = "Full report"
				$wbook.worksheets.item(2).Name = "Month report"
				Write-Verbose -message "Start creating report."
				
				# create year report
				$wsheet = $wbook.worksheets.item("Full report")
				$cells=$wsheet.Cells
				[Int32]$rows = 6
				$cells.item(2, 1) = "Report. 3d max, April 2013."
				$cells.item(2, 1).font.size = 14
				$cells.item(2, 1).font.bold = $True
				$cells.item(4, 2) = "Time (min)"
				$cells.item(4, 3) = "Time (hour)"
				$cells.item(4, 2), $cells.item(4, 3) | foreach{ 
					$_.HorizontalAlignment = 7
					$_.VerticalAlignment = -4108
					$_.Borders.item(9).LineStyle = 1; 		
				}
				$cells.item(5, 1) = "Aurora"	
				$cells.item(5, 1).font.name = "Arial Cyr"
				$cells.item(5, 1).font.italic = $True	
				# full time in minutes
				$cells.item(5, 2) = $time / 60
				# full time in hours
				$cells.item(5, 3) = $time / 3600
				$cells.item(5, 1), $cells.item(5, 2), $cells.item(5, 3) | foreach{		
					$_.font.bold = $True
					$_.font.Size = 11
					$_.VerticalAlignment = -4108
				}
				$cells.item(5, 2), $cells.item(5, 3) | foreach{
					$_.HorizontalAlignment = -4108
					$_.NumberFormat="0,00"
				}
				
				for($i=0; $i -lt $logcount; $i++){
					if((( -not $completed) -and ($logscan[$i])) -or ($completed)){
						$cells.item($rows, 1), $cells.item($rows, 2), $cells.item($rows, 3) | foreach{
								$_.VerticalAlignment = -4108
								$_.font.Size = 10
						}
						$cells.item($rows, 2), $cells.item($rows, 3) | foreach{
							$_.HorizontalAlignment = -4108
							$_.NumberFormat = "0,00"
						}
					}
				}
				for($i=0; $i -lt $logcount; $i++){
					if((( -not $completed) -and ($logscan[$i])) -or ($completed)){
						$cells.item($rows, 1) = $tasknames[$i]
						$cells.item($rows, 1), $cells.item($rows, 2), $cells.item($rows, 3) | foreach{
								$_.VerticalAlignment = -4108
								$_.font.Size = 10
						}
						$cells.item($rows, 2), $cells.item($rows, 3) | foreach{
							$_.HorizontalAlignment = -4108
							$_.NumberFormat = "0,00"
						}
						$cells.item($rows, 2) = $timeperlog[[Int32]($rows - 6)] / 60
						$cells.item($rows, 3) = $timeperlog[[Int32]($rows - 6)] / 3600
						$cells.item($rows, 1), $cells.item($rows, 2), $cells.item($rows, 3) | foreach{
							$_.VerticalAlignment = -4108
							$_.font.Size = 10
						}
						$cells.item($rows, 2), $cells.item($rows, 3) | foreach{
							$_.HorizontalAlignment = -4108
							$_.NumberFormat = "0,00"
						}
						$rows++
					}
				}
				Write-Verbose -message "Creating report complete."							
				Write-Verbose -message "Start creating month report."
				# if $month or $full switched
				# create month report
				if(($mode -eq [string]("month")) -or ($mode -eq [string]("full"))){			
					$wsheet = $wbook.worksheets.item("Month report")
					$cells=$wsheet.Cells
					$cells.item(2, 1) = "Month report. 2013"
					$cells.item(2, 1) | foreach {
						$_.font.size = 14
						$_.font.bold = $True
					}
					$cells.item(4, 1) = "Name"
					$cells.item(4, 1) | foreach {
						$_.font.size = 11
						$_.font.bold = $True
						$_.VerticalAlignment = -4108
					}
					$cols = [Int32](2)
					$rows = [Int32](5)	
					#
					# try switch
					switch ($month){
						1 { $cells.item(4, $cols) = "January" }
						2 { $cells.item(4, $cols) = "February" }
						3 { $cells.item(4, $cols) = "March" }
						4 { $cells.item(4, $cols) = "April" }
						5 { $cells.item(4, $cols) = "May" }
						6 { $cells.item(4, $cols) = "June" }
						7 { $cells.item(4, $cols) = "July" }
						8 { $cells.item(4, $cols) = "August" }
						9 { $cells.item(4, $cols) = "September" }
						10 { $cells.item(4, $cols) = "October" }
						11 { $cells.item(4, $cols) = "November" }
						12 { $cells.item(4, $cols) = "December" }
						0 {
							"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" | foreach {
								$cells.item(4, $cols) = $_
								$cells.item(4, $cols).font.size = 11
								$cells.item(4, $cols).font.bold = $True
								$cells.item(4, $cols).VerticalAlignment = -4108
								$cells.item(4, $cols).HorizontalAlignment = -4108
								$cols++
							}
						}
						default{ Write-Error -message "Wrong -month value." }
					}											
					$cols = [Int32](1)
					write-verbose test1
					switch ($month){						
						0 {
							$correction = [Int32](0)
							write-verbose test2
							for($j=0; $j -lt $logcount; $j++){
								write-verbose -message [string]($logscan[$j])
								if((( -not $completed) -and ($logscan[$j])) -or ($completed)){
									write-verbose test3
									$cells.item($rows - $correction, $cols) = $tasknames[$j]					
									$cells.item($rows - $correction, $cols) | foreach {
										$_.font.size = 10			
										$_.VerticalAlignment = -4108
									}
									for($i=0; $i -lt 12; $i++){						
										$cells.item($rows - $correction, [Int32]($i+2)) = $timem1[[Int32]($rows-5)][[Int32]($i)] / 60
										$cells.item($rows - $correction, [Int32]($i+2)) | foreach {
											$_.font.size = 10
											$_.NumberFormat = "0,00"
											$_.HorizontalAlignment = -4108
											$_.VerticalAlignment = -4108
										}
									}
								}else{
									write-verbose test4
									$correction++
								}
								$rows++
							}
						}
						default{
							write-verbose test5
							if(($month -lt 0) -or ($month -gt 12)){
								write-verbose test6
								Write-Error -message "Wrong -month value." 
							}else{
								write-verbose test7
								$month--
								for($j=0; $j -lt $logcount; $j++){
									if((( -not $completed) -and ($logscan[$j])) -or ($completed)){
										$cells.item($rows - $correction, $cols) = $tasknames[$j]					
										$cells.item($rows - $correction, $cols) | foreach {
											$_.font.size = 10			
											$_.VerticalAlignment = -4108
										}
										$cells.item($rows - $correction, $month) = $timem1[[Int32]($rows-5)][$month] / 60
										$cells.item($rows - $correction, $month) | foreach {
											$_.font.size = 10
											$_.NumberFormat = "0,00"
											$_.HorizontalAlignment = -4108
											$_.VerticalAlignment = -4108
										}
									}else{
										$correction++
									}
									$rows++								
								}
							}
						}
					}
					$wbook.SaveAs($outdir+"\"+"report.xlsx")
					$wbook.Close()
					Write-Verbose -message "Creating month report complete."
				}
				
				# if $node or $full switched
				# create node report
				if(($mode -eq [string]("node")) -or ($mode -eq [string]("full"))){
					Write-Verbose -message "Start creating node report."
					$wbook=$xlinst.Workbooks.Add()
					$wsheet=$wbook.ActiveSheet
					$cells=$wsheet.Cells
					$cells.item(2, 1) = "Node report. 2013"
					$cells.item(2, 1).font.size = 14
					$cells.item(2, 1).font.bold = $True
					$cells.item(4, 1) = "Node name"
					$cells.item(4, 2) = "Time (min)"
					$cells.item(4, 3) = "Time (hour)"
					$cells.item(4, 1), $cells.item(4, 2), $cells.item(4, 3) | foreach{ 
						$_.HorizontalAlignment = -4108
						$_.VerticalAlignment = -4108
						$_.font.size = 11
						$_.Borders.item(9).LineStyle = 1; 		
					}
					[Int32]$rows = 5
					for($i=0; $i -lt $nodenames.Length; $i++){
						$cells.item([Int32]($rows + $i), 1) = $nodenames[$i]
						$cells.item([Int32]($rows + $i), 2) = $timepernode[$i] / 60
						$cells.item([Int32]($rows + $i), 3) = $timepernode[$i] / 3600
						$cells.item([Int32]($rows + $i), 1), $cells.item([Int32]($rows + $i), 2), $cells.item([Int32]($rows + $i), 3) | foreach{
							$_.font.size = 10
							$_.VerticalAlignment = -4108
							$_.HorizontalAlignment = -4108
						}
						$cells.item([Int32]($rows + $i), 2), $cells.item([Int32]($rows + $i), 3) | foreach{ $_.NumberFormat="0,00" }		
					}
					$wbook.SaveAs($outdir+"\"+"node-report.xlsx")
					$wbook.Close()
				}				
				Write-Verbose -message "Creating Aurora report complete."

				$xlinst.Quit()
			}else{
				Write-Verbose -message "MS Excel is not installed."
				Write-Error "Can't create a report files. Please, install the MS Excel."
			}	
		}else{
			if($outdir -eq [string]("")){
				Write-Error "Can't create a report files. Please, check -outdir parameter"
			}
		}
		Set-Location $location	
	}
}