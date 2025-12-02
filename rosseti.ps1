param (
    [string]$baseUrl = "https://rosseti-lenenergo.ru/planned_work/",
    [string]$telegramToken = "7731641851:AAFrlAdw_uPAm11111111111111",
    [string]$errorChatId = "287241111", # Chat ID for errors
    #[string]$groupChatId = "-1002513697785", # Group chat ID for TEST
    #[string]$messageThreadId = "3", # Topic ID for group chat 3 - TEST
    [string]$groupChatId = "-1001896393733", # Group chat ID for PROD
    [string]$messageThreadId = "8183", # Topic ID for group chat 8183 - PROD
    [string[]]$locations = @("Воронцово", "КП Линтулово", "Сабур 2", "КП Огоньково"),
    [string]$checkTime = "16:00",
    [string]$logFile = "C:\infrabot\WebsiteCheckLog.txt" # Log file path # Time to start checking the website
)

# Function to send message to Telegram
function Send-TelegramMessage {
    param (
        [string]$message,
        [string]$chatId,
        [string]$messageThreadId = $null
    )
    "[$(Get-Date)] Sending Telegram message to chat $chatId (thread: $messageThreadId): $message" | Out-File -FilePath $logFile -Append -Encoding UTF8
    $telegramApi = "https://api.telegram.org/bot$telegramToken/sendMessage"
    $body = @{
        chat_id = $chatId
        text = $message
        parse_mode = "Markdown"
    }
    if ($messageThreadId) {
        $body.message_thread_id = $messageThreadId
    }
    $bodyJson = $body | ConvertTo-Json
    try {
        $response = Invoke-RestMethod -Uri $telegramApi -Method Post -Body $bodyJson -ContentType "application/json"
        "[$(Get-Date)] Telegram message sent successfully to chat $chatId" | Out-File -FilePath $logFile -Append -Encoding UTF8
        return $response
    } catch {
        "[$(Get-Date)] Failed to send Telegram message to chat $chatId : $_" | Out-File -FilePath $logFile -Append -Encoding UTF8
        throw $_
    }
}

# Initial check to store previous data
$previousData = @()
$lastCheckDate = $null # Track the last date the check was performed
"[$(Get-Date)] Script started. Initial previousData: $previousData" | Out-File -FilePath $logFile -Append -Encoding UTF8

while ($true) {
    try {
        $currentTime = Get-Date
        $currentDate = $currentTime.Date
        $checkHour = [int]($checkTime.Split(":")[0])
        $checkMinute = [int]($checkTime.Split(":")[1])
        $checkDateTime = Get-Date -Hour $checkHour -Minute $checkMinute -Second 0 -Date $currentDate
        "[$(Get-Date)] Current time: $currentTime, Check time: $checkDateTime, Last check date: $lastCheckDate" | Out-File -FilePath $logFile -Append -Encoding UTF8

        # Check if it's time to run and hasn't run today
        if ($currentTime -ge $checkDateTime -and ($lastCheckDate -eq $null -or $lastCheckDate.Date -ne $currentDate)) {
            "[$(Get-Date)] Current time ($currentTime) is past $checkTime and hasn't run today, checking website" | Out-File -FilePath $logFile -Append -Encoding UTF8
            $currentData = @()
            $messagesToSend = @{}
            
            foreach ($location in $locations) {
                $encodedLocation = [System.Web.HttpUtility]::UrlEncode($location)
                $queryParams = "reg=&city=&date_start=&date_finish=&res=&street=$encodedLocation"
                $filterUrl = "${baseUrl}?$queryParams"
                "[$(Get-Date)] Fetching webpage content from $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8

                if (-not [System.Uri]::IsWellFormedUriString($filterUrl, [System.UriKind]::Absolute)) {
                    "[$(Get-Date)] Invalid URL constructed: $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    throw "Invalid URL constructed: $filterUrl"
                }

                $webContent = Invoke-WebRequest -Uri $filterUrl -UseBasicParsing
                if (-not $webContent) {
                    "[$(Get-Date)] Web content is null for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    throw "Failed to retrieve webpage content for $filterUrl"
                }
                "[$(Get-Date)] Web content retrieved. Status code: $($webContent.StatusCode), Content length: $($webContent.Content.Length)" | Out-File -FilePath $logFile -Append -Encoding UTF8

                # Parse HTML content
                "[$(Get-Date)] Attempting to parse HTML content for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                $html = $webContent.Content
                if (-not $html) {
                    "[$(Get-Date)] HTML content is empty for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    throw "HTML content is empty for $filterUrl"
                }

                # Find the data table by looking for a table with 'Адрес' header
                "[$(Get-Date)] Searching for table in HTML" | Out-File -FilePath $logFile -Append -Encoding UTF8
                $tablePattern = '<table[^>]*>.*?</table>'
                $tableMatches = [regex]::Matches($html, $tablePattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                if ($tableMatches.Count -eq 0) {
                    "[$(Get-Date)] No table found in HTML for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    continue
                }
                $tableHtml = $null
                foreach ($match in $tableMatches) {
                    if ($match.Value -match '<th[^>]*>Адрес</th>') {
                        $tableHtml = $match.Value
                        break
                    }
                }
                if (-not $tableHtml) {
                    "[$(Get-Date)] No table with 'Адрес' header found in HTML for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    continue
                }
                "[$(Get-Date)] Table found. Length: $($tableHtml.Length)" | Out-File -FilePath $logFile -Append -Encoding UTF8

                # Extract data rows (exclude headers and inputs)
                "[$(Get-Date)] Extracting data rows from table" | Out-File -FilePath $logFile -Append -Encoding UTF8
                $rowPattern = '<tr[^>]*>(?:(?!<th|<input).)*?<td[^>]*>(?:[^<]*)</td>.*?</tr>'
                $rows = [regex]::Matches($tableHtml, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                if (-not $rows) {
                    "[$(Get-Date)] No data rows found in table for $filterUrl" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    continue
                }
                "[$(Get-Date)] Found $($rows.Count) data rows" | Out-File -FilePath $logFile -Append -Encoding UTF8

                $rowIndex = 0
                foreach ($row in $rows) {
                    $rowIndex++
                    "[$(Get-Date)] Processing row $rowIndex for $location" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    $rowHtml = $row.Value

                    # Extract cells
                    $cellPattern = '<td[^>]*>(.*?)</td>'
                    $cells = [regex]::Matches($rowHtml, $cellPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
                    if ($cells.Count -gt 2) {
                        "[$(Get-Date)] Row $rowIndex has $($cells.Count) cells" | Out-File -FilePath $logFile -Append -Encoding UTF8
                        $address = $cells[2].Groups[1].Value.Trim()
                        "[$(Get-Date)] Row $rowIndex address: $address" | Out-File -FilePath $logFile -Append -Encoding UTF8

                        # Check if the address contains any of the target locations
                        $containsTarget = $address -like "*КП Линтулово*" -or $address -like "*КП Огоньково*" -or $address -like "*СНТ Воронцово*" -or $address -like "*Сабур 2*"
                        if (-not $containsTarget) {
                            "[$(Get-Date)] Row $rowIndex does not contain target locations (КП Линтулово, КП Огоньково, СНТ Воронцово, Сабур 2)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                            continue
                        }

                        "[$(Get-Date)] Row $rowIndex matches target location criteria" | Out-File -FilePath $logFile -Append -Encoding UTF8

                        # Clean up the address field: split by <br> to preserve multi-word names
                        $addressItems = $address -split '<br>'
                        $cleanAddresses = @()
                        foreach ($item in $addressItems) {
                            $cleanItem = $item -replace '<[^>]+>', '' # Remove HTML tags
                            $cleanItem = $cleanItem.Trim()
                            if ($cleanItem) {
                                $cleanAddresses += $cleanItem
                            }
                        }

                        # Extract other fields
                        $startDate = $cells[3].Groups[1].Value.Trim()
                        $startTime = $cells[4].Groups[1].Value.Trim()
                        $endDate = $cells[5].Groups[1].Value.Trim()
                        $endTime = $cells[6].Groups[1].Value.Trim()

                        # Create a key based on start and end times to group messages
                        $timeKey = "$startDate $startTime - $endDate $endTime"

                        # Add the address to the message for this time slot
                        if (-not $messagesToSend.ContainsKey($timeKey)) {
                            $messagesToSend[$timeKey] = @{
                                StartDate = $startDate
                                StartTime = $startTime
                                EndDate   = $endDate
                                EndTime   = $endTime
                                Addresses = @()
                            }
                        }
                        $messagesToSend[$timeKey].Addresses += $cleanAddresses

                        $rowText = ($cells | ForEach-Object { $_.Groups[1].Value.Trim() }) -join " | "
                        $currentData += $rowText
                        "[$(Get-Date)] Row $rowIndex added to currentData: $rowText" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    } else {
                        "[$(Get-Date)] Row $rowIndex has insufficient cells: $($cells.Count)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    }
                }
            }

            # Process and send messages
            foreach ($timeKey in $messagesToSend.Keys) {
                $msgData = $messagesToSend[$timeKey]
                $combinedAddresses = @()
                foreach ($address in $msgData.Addresses) {
                    $uniqueAddresses = @()
                    foreach ($item in $address) {
                        if ($item -and $uniqueAddresses -notcontains $item) {
                            $uniqueAddresses += $item
                        }
                    }
                    $combinedAddresses += $uniqueAddresses
                }
                $uniqueCombinedAddresses = $combinedAddresses | Sort-Object -Unique
                $formattedAddress = ($uniqueCombinedAddresses | ForEach-Object {
                    if ($locations -contains $_) {
                        "**$_**"  # Bold the target locations
                    } else {
                        $_
                    }
                }) -join ", "

                # Format the message
                $message = "РосСети - Плановые работы в: $formattedAddress`nНачало работ $($msgData.StartDate) с $($msgData.StartTime)`nОкончание работ $($msgData.EndDate) до $($msgData.EndTime)"

                # Check if startDate is greater than or equal to current date
                try {
                    $startDate = [datetime]::ParseExact($msgData.StartDate, "dd-MM-yyyy", $null)
                    $currentDate = Get-Date
                    if ($startDate.Date -ge $currentDate.Date) {
                        if ($previousData -notcontains ($currentData -join "|")) {
                            "[$(Get-Date)] New planned work detected for date $($msgData.StartDate), sending Telegram message to group chat $groupChatId (thread $messageThreadId)" | Out-File -FilePath $logFile -Append -Encoding UTF8
                            Send-TelegramMessage -message $message -chatId $groupChatId -messageThreadId $messageThreadId
                        } else {
                            "[$(Get-Date)] Planned work for date $($msgData.StartDate) already exists in previousData" | Out-File -FilePath $logFile -Append -Encoding UTF8
                        }
                    } else {
                        "[$(Get-Date)] Planned work for date $($msgData.StartDate) is in the past, skipping Telegram message" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    }
                } catch {
                    "[$(Get-Date)] Error parsing date $($msgData.StartDate): $_" | Out-File -FilePath $logFile -Append -Encoding UTF8
                    Send-TelegramMessage -message "Error parsing date $($msgData.StartDate): $_" -chatId $errorChatId
                }
            }

            # Update previousData and lastCheckDate
            "[$(Get-Date)] Updating previousData with currentData" | Out-File -FilePath $logFile -Append -Encoding UTF8
            $previousData = $currentData
            $lastCheckDate = $currentDate
            "[$(Get-Date)] Current iteration complete. Found $($currentData.Count) matching rows" | Out-File -FilePath $logFile -Append -Encoding UTF8
        } else {
            "[$(Get-Date)] Current time ($currentTime) is before $checkTime or already checked today, skipping website check" | Out-File -FilePath $logFile -Append -Encoding UTF8
        }

        # Wait for 1 hour (3600 seconds) before checking time again
        "[$(Get-Date)] Sleeping for 1 hour" | Out-File -FilePath $logFile -Append -Encoding UTF8
        Start-Sleep -Seconds 14400
    } catch {
        "[$(Get-Date)] Error occurred: $_" | Out-File -FilePath $logFile -Append -Encoding UTF8
        Send-TelegramMessage -message "Error fetching data: $_" -chatId $errorChatId
        # Wait for 1 hour before retrying in case of error
        "[$(Get-Date)] Sleeping for 1 hour due to error" | Out-File -FilePath $logFile -Append -Encoding UTF8
        Start-Sleep -Seconds 3600
    }
}
