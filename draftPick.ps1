# Check if the ImportExcel module is installed
$module = Get-Module -ListAvailable -Name ImportExcel

If (-Not $module) {
    Write-Host "ImportExcel module not found. Installing requires elevated privileges..."

    # Check if running as administrator
    If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        # Re-run the script as administrator if not already
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"" + $myInvocation.MyCommand.Definition + "`""
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        Exit
    }

    Try {
        # Install the module with elevated privileges
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
        Write-Host "Module successfully installed."
    }
    Catch {
        Write-Host "Failed to install the module. Please check your network connection or permissions."
        Exit
    }
}
Else {
    start-sleep .1
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#$survey = Import-Excel -Path "$env:userprofile\downloads\CTF Draft League CTFDL 2024-1 UPDATED 08152024.xlsx" -NoHeader -DataOnly

# Define the filename
$fileName = "CTF Draft League CTFDL 2024-1 UPDATED 08162024 1020am 111 Registrations.xlsx"

# Define the file path
$filePath = Join-Path -Path "$env:userprofile\downloads" -ChildPath $fileName

# Check if the file exists
If (-Not (Test-Path $filePath)) {
    [System.Windows.Forms.MessageBox]::Show(
        "The file '$fileName' was not found in the Downloads folder. 
        Please select the file manually or ensure the file is downloaded and placed in the Downloads folder.", 
        "File Not Found", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    # Create a new OpenFileDialog object
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = "$env:userprofile\downloads"
    $openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $openFileDialog.Title = "Select the Excel file"
    $openFileDialog.Multiselect = $false

    # Show the file picker dialog
    $result = $openFileDialog.ShowDialog()

    # Check if the user selected a file
    If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Update the file path with the selected file
        $filePath = $openFileDialog.FileName
    } Else {
        # Exit if the user cancels the dialog
        Write-Host "No file selected. Exiting script."
        Exit
    }
}

# Proceed with importing the Excel file using the selected or default path
$survey = Import-Excel -Path $filePath -NoHeader -DataOnly

# Parse the survey results into variables/array
$surveyResults = foreach ($entry in $survey[1..($survey.Count - 1)]){
    $prop = [ordered]@{
        'Alias' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[0])"].Value
        'DiscordAgree' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[1])"].Value
        'DiscordJoined' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[2])"].Value
        'DiscordName' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[3])"].Value
        'RulesAgree' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[4])"].Value
        'HelpNewPlayer' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[5])"].Value
        'NewPlayerListen' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[6])"].Value
        'Classes' = @($entry.Psobject.Properties["P8"].Value, $entry.Psobject.Properties["P9"].Value, $entry.Psobject.Properties["P10"].Value, $entry.Psobject.Properties["P11"].Value, $entry.Psobject.Properties["P12"].Value, $entry.Psobject.Properties["P13"].Value, $entry.Psobject.Properties["P14"].Value, $entry.Psobject.Properties["P15"].Value, $entry.Psobject.Properties["P16"].Value, $entry.Psobject.Properties["P17"].Value, $entry.Psobject.Properties["P18"].Value, $entry.Psobject.Properties["P19"].Value, $entry.Psobject.Properties["P20"].Value)
        'DatesNotAvailable' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[20])"].Value
        'Captain' = $entry.Psobject.Properties["$($entry.Psobject.Properties.name[21])"].Value
    }
    New-Object psobject -Property $prop
}
#$surveyResults

<#
# Rules
Write-Host "Did not agree to rules: " -ForegroundColor Cyan
Write-Host $(($surveyResults | ?{$_.rulesagree -eq 'no'} |select alias).alias) -ForegroundColor Red
Write-Host "Did not agree to help new players: " -ForegroundColor Cyan
$aliases = $surveyResults | Where-Object { $_.HelpNewPlayer -eq 'no' } | Select-Object -ExpandProperty alias
foreach ($alias in $aliases) {
    Write-Host $alias -ForegroundColor Red
}

# New player
Write-Host "Did not agree to listen as new player: " -ForegroundColor Cyan
Write-Host $(($surveyResults | ?{$_.NewPlayerListen -eq 'no'} |select alias).alias) -ForegroundColor Red

Write-Host "Listen as new player results: " -ForegroundColor Cyan
$data = $surveyResults.NewPlayerListen | Group-Object | Sort-Object -Property Count -Descending
$data | ForEach-Object {
    Write-Host ("{0,-20} {1,5}" -f $_.Name, $_.Count) -ForegroundColor Red
}

# Captains
Write-Host "Captains: " -ForegroundColor Cyan
$captains = $surveyResults | Where-Object { $_.Captain -ne 'no' } | Select-Object -ExpandProperty alias
foreach ($captain in $captains) {
    Write-Host $captain -ForegroundColor Gray
}

Write-Host Class preference breakdown:
# Define color mappings for each class type
$classColors = @{
    "Infantry" = "Red"
    "Heavy" = "Cyan"
    "Medic" = "Yellow"
    "Squad Leader" = "Green"
    "Infiltrator" = "Magenta"
    "Engineer" = "DarkYellow"
    "JT" = "Gray"
}
#>
# Define alternating colors for counts
$countColors = @([ConsoleColor]::White, [ConsoleColor]::DarkGray)

# Function to get the appropriate color for a class
function Get-ClassColor($className) {
    foreach ($key in $classColors.Keys) {
        if ($className -match $key) {
            return [ConsoleColor]::$($classColors[$key])
        }
    }
    return [ConsoleColor]::Gray
}

# Flatten all the Classes arrays into a single list
$allClasses = $surveyResults | ForEach-Object { $_.Classes } | ForEach-Object { $_ } | Where-Object { $_ -ne $null }

# Group by unique class and count
$groupedClasses = $allClasses | Group-Object | Sort-Object Count -Descending

<#
# Write the unique class counts to the host in a table format
Write-Host "Class".PadRight(30) "Count" -ForegroundColor Cyan
Write-Host "-----------------------------" "-----" -ForegroundColor Cyan

$useColorIndex = 0
$groupedClasses | ForEach-Object {
    $classColor = Get-ClassColor $_.Name
    $countColor = $countColors[$useColorIndex]
    Write-Host ($_.Name.PadRight(30)) -ForegroundColor $classColor -NoNewline
    Write-Host ($_.Count) -ForegroundColor $countColor
    $useColorIndex = ($useColorIndex + 1) % $countColors.Length
}
#>

$textFromGoogleSheet = @"
Player	Overall	Infantry	Heavy	JT	Medic	SL	Eng	Infil	Comms	Notes
Quimak	11	11							A+	
Sov	10	10	10					10	A	
Kev	10		7	10		9	8		A	
Angelus	10	10							A	
Smoka	10	10							A	
Dinobot	9	6	9			8			D	
Herthbul	9	9		8		9*			B	
Iron	9	9							C	
Silly Wanker	9	9							A	
Rocky	9					9			A	
Designer	8		8			8	8		C	
Keyser	8	8	7			3			D	
Tempest	8	8				8	8	7	D	
baal	8	8							D	
Evade	8	8							C	
Greed	8	8							B	
Revenge	8	8							B	
Shadeaux	8	8							D	
MajaN	8	8								
bt	8	7					8		C	
Zmn	8	6		5	8		8	8	A	big JT derank
bonds	8				8	8			C	
Amok	8				8				C	
Mixm	8				7	8	8		C	
JACKIE	7	7	7			7		7	B	Also pre-miner
Sorrido	7	6	7						D	
Tactical	7		7						C	
TyraeL	7		7						C	
Waylander	7	7			7				C	
aei.own.u	7	7							B	
anjro	7	7							D	
Distrikt	7	7							C	
Flair	7	7							D	
Ruler	7	7							F	
Raekwon	7	6				7			B	
posty	7	4			7				B	
Dilatory	7				7				D	
Breakdance	7					7	7		C	
Mugen	6	6	6		6	6	6		D	
Mountain	6		6						D	
Rue	6	6				3	6		C	about to be deranked
Bosnia	6	6				3			F	
kylauf	6	6				3			D	
Agh	6	6							C	
Debris	6	6							C	
goose	6	6							C	
metal	6	6							C	7 is very generous
Novice	6	6							D	
Soup	6	6							D	
Worth	6	6							C	
Planner	6				6	6			D	
Captain Ax	6					6			F	
Mecca	5	5	5	5					B	
Bozza	5		5						D	
DarkBomer	5	5			6				D	
juetnihilia	5	5				3			C	
A Big Deal	5	5							C	
An RR User	5	5							C	
billayyyy	5	5							C	
Cara Tank	5	5							F	
Chev Rising	5	5							D	
EarthwormJim	5	5		2					C	
England	5	5							D	
Ghost Bomber	5	5							D	
Mights	5	5								
PJL90	5	5							D	
thegreatchompy	5	5							C	
VinEzl	5	5							D	
Kaizer	5	5								
XXXXX	5	4				5			D	
Got Tsolvy?	5				5				D	
Jaguar	5					5			C	
Decker	5						5		C	ovd pre-miner, pre-mades on mix
Infinite	5						5		D-	
Rph	4		4						D	
Badmagik	4	4							F	
Bap	4	4							D	
Britney's Spear	4	4							D	
Epyon	4	4							D	
Katana	4	4							D	
Leak	4	4							D	
Typhoon	4	4							D	
Rbz	4	4								
Gungirl	4				4				C	
Lamerboi	4					4			D	
Juve	3		3			3	3		D	
Exo-	3		3						F	
Panic	3		3							
dp	3	3			3				C	
1992	3	3							D	Manual flanker on D - "daanngggggg"
Astro / FOFO / Congee	3	3							F	
Marqui2156	3	3							D	
Yami Moto Kenichi	3	3							D	
Loaf Pincher	3				3				D	
tpblah	3					3			D	
Chuck Schuldiner	3			3					D	Next birdchest
LoneZodiac	2	2							D	
Cortana	2	1						2	D	
ItzDozier	2				2				D	
Esoteric	1		1						D	
Caso Prime	1	1							F	
Eva	1				1				D	
DarrenDinosaurs	1					1			D	
Snapplers1	1			1						Made aei rq
Galaxy										
Gallet										
Gravity										
jimbo8000										
Korean										
Lingo										
Lobas										
Meep										
Mikhail										
Moose										
My Condolences										
Mysticgohan~										
pamncake										
Patel										
Power										
r										
Sabotage										
scout258										
Seifer										
shh / Scobra										
SickBrain										
SMEG										
Snam										
Spark										
Super lovers										
Terminator										
thatsnice										
Torch										
typ										
UnborN										
Vince Carter										
Walt / Witness										
Wongman.										
Yosh										
zblu										
"@

# Step 1: Split the text into an array of rows
$rows = $textFromGoogleSheet -split "`r`n"

# Step 2: Extract the header row to get the property names
$headers = $rows[0] -split "`t"

# Step 3: Process each subsequent row and create a custom object with properties
$playerRatings = $rows[1..($rows.Length - 1)] | ForEach-Object {
    $values = $_ -split "`t"
    $object = [pscustomobject]@{}

    for ($i = 0; $i -lt $headers.Length; $i++) {
        # Assign each value to the corresponding header as a property
        $object | Add-Member -NotePropertyName $headers[$i] -NotePropertyValue $values[$i]
    }

    $object
}


# Lets build out rough elo schematic
$combinedResults =@()
 foreach ($surveyEntry in $surveyResults) {
    # Find the matching entry in $newDataset based on Alias/Player
    $playerRanking = $playerRatings | Where-Object { $_.Player -eq $surveyEntry.Alias }
    #Write-Host Player Ranking: $playerRanking -ForegroundColor Green
    if ($playerRanking) {
        # Add the ranking properties to the survey entry
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Overall' -Value $playerRanking.Overall
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Infantry' -Value $playerRanking.Infantry
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Heavy' -Value $playerRanking.Heavy
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'JT' -Value $playerRanking.JT
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Medic' -Value $playerRanking.Medic
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'SL' -Value $playerRanking.SL
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Eng' -Value $playerRanking.Eng
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Infil' -Value $playerRanking.Infil
        $surveyEntry | Add-Member -MemberType NoteProperty -Name 'Comms' -Value $playerRanking.Comms
    }
    
    # Output the combined entry
    $combinedResults += $surveyEntry
}
$combinedResults = $combinedResults | ?{$_.alias -notmatch 'Open-ended response'}

# Now $combinedResults contains the combined data with rankings
#$combinedResults

# Define the required and recommended roles for each team
$defenseRequired = @{ Medic = 1; Eng = 1 }
$defenseRecommended = @( 
    @{ Infantry = 2; Heavy = 1 }, 
    @{ Infantry = 3 } 
)


$offenseRequired = @{ SL = 1 }
$offenseRecommended = @( 
    @{ Infantry = 2; JT = 1; Infil = 1 },
    @{ Infantry = 3; Heavy = 1 },
    @{ Infantry = 3; JT = 1 },
    @{ Infantry = 4 }, 
    @{ Infantry = 3; Infil = 1 },
    @{ Infantry = 2; JT = 1; Heavy = 1 },
    @{ Infantry = 2; JT = 1; Heavy = 1 }
)

# Function to evaluate and select players for a role
function Select-PlayersForRole {
    param (
        [array]$players,
        [string]$role,
        [int]$count
    )
    
    # Filter players who have a valid preference for the role
    $validPlayers = $players | Where-Object { $_.$role -ne '' }

    # Randomly select the required number of players
    return $validPlayers | Get-Random -Count $count
}

# Function to build a team
function Build-Team {
    param (
        [array]$players,
        [hashtable]$requiredRoles,
        [array]$recommendedRoles
    )
    
    $team = @()

    # Select required roles
    foreach ($role in $requiredRoles.Keys) {
        $selected = Select-PlayersForRole $players $role $requiredRoles[$role]
        $team += $selected
        # Mark selected players' class
        foreach ($player in $selected) {
            $player | Add-Member -MemberType NoteProperty -Name SelectedClass -Value $role -Force
        }
        # Remove selected players from the pool
        $players = $players | Where-Object { $selected -notcontains $_ }
    }

    # Select recommended roles
    foreach ($roleSet in $recommendedRoles) {
        $remainingPlayers = $players
        $tempTeam = @()

        foreach ($role in $roleSet.Keys) {
            $selected = Select-PlayersForRole $remainingPlayers $role $roleSet[$role]
            $tempTeam += $selected
            foreach ($player in $selected) {
                $player | Add-Member -MemberType NoteProperty -Name SelectedClass -Value $role -Force
            }
            $remainingPlayers = $remainingPlayers | Where-Object { $selected -notcontains $_ }
        }

        # Calculate the total roles needed
        $totalRolesNeeded = ($roleSet.Values | Measure-Object -Sum).Sum
        if ($tempTeam.Count -eq $totalRolesNeeded) {
            $team += $tempTeam
            $players = $remainingPlayers
            break
        }
    }

    # Fill remaining slots with any players
    while ($team.Count -lt 5) {
        $availablePlayer = $players | Where-Object { $_.SelectedClass -ne $null } | Select-Object -First 1
        if ($availablePlayer) {
            $team += $availablePlayer
            $players = $players | Where-Object { $availablePlayer -notcontains $_ }
        } else {
            break
        }
    }

    return $team
}

# Function to display the team with class and ratings
function Display-Team {
    param (
        [array]$team,
        [string]$teamName
    )

    Write-Host "$teamName Team:" -ForegroundColor Cyan
    $overallRating = 0

    foreach ($player in $team) {
        if ($player.Alias -eq "Open-Ended Response" -or $player.SelectedClass -eq $null) {
            continue  # Skip invalid entries and those without a selected class
        }

        $classColor = Get-ClassColor $player.SelectedClass

        Write-Host ($player.Alias.PadRight(20)) -ForegroundColor Gray -NoNewline
        Write-Host ("Class: " + $player.SelectedClass).PadRight(20) -ForegroundColor $classColor -NoNewline
        Write-Host ("Overall: " + $player.Overall) -ForegroundColor $classColor

        $overallRating += [int]$player.Overall
    }

    Write-Host "Overall Team Rating: $overallRating" -ForegroundColor Yellow
    Write-Host ""
}

function Build-Teams {
    if ($combinedResults -and $combinedResults.Count -gt 0) {
        # Initialize the SelectedClass property
        $combinedResults | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name SelectedClass -Value $null -Force
        }

        # Filter out invalid entries before building teams
        $filteredResults = $combinedResults | Where-Object { $_.Alias -ne "Open-Ended Response" }

        # Build the defense team first
        $defenseTeam = Build-Team $filteredResults $defenseRequired $defenseRecommended

        # Remove defense team members from the available players for offense team
        $remainingPlayers = $filteredResults | Where-Object { $defenseTeam -notcontains $_ }

        # Build the offense team with the remaining players
        $offenseTeam = Build-Team $remainingPlayers $offenseRequired $offenseRecommended

        # Display the teams with classes and ratings
        Display-Team $defenseTeam "Defense"
        Display-Team $offenseTeam "Offense"
    } else {
        Write-Host "Error: $combinedResults is null or empty." -ForegroundColor Red
    }
}

#Build-Teams

$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        Title="Team Draft" Height="800" Width="1000" Background="White">
    <Grid Background="LightGray">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2*" />
            <ColumnDefinition Width="1*" />
        </Grid.ColumnDefinitions>

        <!-- Teams Section -->
        <StackPanel Grid.Column="0" Margin="10" Background="Gainsboro">
            <TextBlock Text="Teams" FontSize="16" FontWeight="Bold" Margin="0,0,0,10" />

            <!-- Dropdown for Team Selection -->
            <ComboBox Name="TeamSelector" Width="200" Margin="0,0,10,10">
                <ComboBoxItem Content="Team 1" />
                <ComboBoxItem Content="Team 2" />
                <ComboBoxItem Content="Team 3" />
                <ComboBoxItem Content="Team 4" />
                <ComboBoxItem Content="Team 5" />
                <ComboBoxItem Content="Team 6" />
            </ComboBox>
            
            <!-- Placeholder for teams -->
            <StackPanel Name="TeamsPanel"></StackPanel>
        </StackPanel>

        <!-- Available Players Section -->
        <StackPanel Grid.Column="1" Margin="10" Background="Gainsboro">
            <TextBlock Text="Available Players" FontSize="16" FontWeight="Bold" Margin="0,0,0,10" />

            <!-- Sorting Options -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,0,0,10">
                <TextBlock Text="Sort by:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <ComboBox Name="SortByComboBox" Width="150" SelectedIndex="0">
                    <ComboBoxItem Content="Alias (A-Z)" />
                    <ComboBoxItem Content="Alias (Z-A)" />
                    <ComboBoxItem Content="Rating (High to Low)" />
                    <ComboBoxItem Content="Rating (Low to High)" />
                </ComboBox>
            </StackPanel>

            <!-- Available Players ListBox -->
            <ListBox Name="AvailablePlayersListBox" Height="600" Background="White">
                <ListBox.ItemTemplate>
                    <DataTemplate>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="{Binding Alias}" Width="150" />
                            <TextBlock Text="{Binding Overall}" Width="50" />
                        </StackPanel>
                    </DataTemplate>
                </ListBox.ItemTemplate>
            </ListBox>

            <Button Name="DraftPlayerButton" Content="Draft Player" Width="150" Height="30" HorizontalAlignment="Center" Margin="0,10,0,10"/>
            <Button Name="UndoDraftButton" Content="Undo Last Draft" Width="150" Height="30" HorizontalAlignment="Center" />
            <Button Name="ExportButton" Content="Export Teams" Width="150" Height="30" HorizontalAlignment="Center" Margin="0,10,0,10"/>
        </StackPanel>
    </Grid>
</Window>
"@

$stream = [System.IO.MemoryStream]::new()
$writer = [System.IO.StreamWriter]::new($stream)
$writer.Write($XAML)
$writer.Flush()
$stream.Position = 0

$reader = [System.Xml.XmlReader]::Create($stream)
$window = [Windows.Markup.XamlReader]::Load($reader)

$classColors = @{
    "Infantry" = "Red"
    "Heavy" = "Cyan"
    "Medic" = "Yellow"
    "Squad Leader" = "Green"
    "Infiltrator" = "Magenta"
    "Engineer" = "DarkYellow"
    "JT" = "Gray"
}

# Function to generate the export text
function Export-TeamsToFile {
    param (
        [string]$filePath
    )

    $output = ""

    # Part 1: Team ratings
    for ($i = 1; $i -le 6; $i++) {
        $teamRating = $teamTotalRatings[$i]
        $teamName = $teamNames[$i - 1]
        $output += "$teamName (Rating= $teamRating)`r`n"
    }

    $output += "`r`n"  # Add a blank line between sections

    # Part 2: Team players
    for ($i = 1; $i -le 6; $i++) {
        $teamDataGrid = $window.FindName("Team${i}DataGrid")
        $teamName = $teamNames[$i - 1]
        $output += "$teamName`r`n"
        $output += "________`r`n"

        foreach ($player in $teamDataGrid.Items) {
            $output += "$($player.Alias) - $($player.Overall) - $($player.PreferredClass)`r`n"
        }

        $output += "`r`n"  # Add a blank line between teams
    }

    # Save the output to the file
    [System.IO.File]::WriteAllText($filePath, $output)
}

# Attach the Export functionality to the Export button
$exportButton = $window.FindName("ExportButton")
$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
    $saveFileDialog.DefaultExt = "txt"
    $saveFileDialog.AddExtension = $true
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $saveFileDialog.FileName
        Export-TeamsToFile -filePath $filePath
        Write-Host "Teams exported to $filePath" -ForegroundColor Green
    } else {
        Write-Host "Export cancelled." -ForegroundColor Yellow
    }
})

$teamsPanel = $window.FindName("TeamsPanel")
$teamSelector = $window.FindName("TeamSelector")
$availablePlayersListBox = $window.FindName("AvailablePlayersListBox")
$draftPlayerButton = $window.FindName("DraftPlayerButton")
$undoDraftButton = $window.FindName("UndoDraftButton")
$sortByComboBox = $window.FindName("SortByComboBox")

$captains = @("posty", "Thiz", "Mixm", "Decker", "Verb", "metal")
$teamTotalRatings = @{}
$teamNames = @("Coming From Behind", "Seventh Plague", "Fangs Out", "Care Bear Countdown", "We Are All Friends Here", "Takuteko Draculakushyon")

function Get-PreferredClass {
    param (
        [pscustomobject]$player
    )

    $classRatings = @{
        "Infantry" = [int]$player.Infantry
        "Heavy" = [int]$player.Heavy
        "JT" = [int]$player.JT
        "Medic" = [int]$player.Medic
        "SL" = [int]$player.SL
        "Eng" = [int]$player.Eng
        "Infil" = [int]$player.Infil
    }

    $validClasses = $classRatings.GetEnumerator() | Where-Object { $_.Value -gt 0 }

    if ($validClasses.Count -gt 0) {
        $preferredClass = $validClasses | Sort-Object -Property Value -Descending | Select-Object -First 1
        return $preferredClass.Name
    } else {
        return $null
    }
}

function Add-TeamUI {
    param (
        [int]$teamNumber,
        [string]$teamName
    )

    # Create a StackPanel for the team
    $teamStackPanel = New-Object Windows.Controls.StackPanel

    # Add a TextBlock for the team's total rating
    $totalRatingTextBlock = New-Object Windows.Controls.TextBlock
    $totalRatingTextBlock.Name = "Team${teamNumber}TotalRating"
    $totalRatingTextBlock.Text = "        Total Rating: [0]"
    $totalRatingTextBlock.FontSize = 14
    $totalRatingTextBlock.FontWeight = 'Bold'
    $totalRatingTextBlock.Foreground = 'Black'

    $teamNameBlock = New-Object Windows.Controls.TextBlock
    $teamNameBlock.Text = $teamName
    $teamNameBlock.FontSize = 14
    $teamNameBlock.FontWeight = 'Bold'
    $teamNameBlock.Foreground = 'Black'

    $teamRatingBlock = New-Object Windows.Controls.TextBlock
    $teamRatingBlock.Text = "        Total Rating: [0]"
    $teamRatingBlock.FontSize = 14
    $teamRatingBlock.FontWeight = 'Bold'
    $teamRatingBlock.Foreground = 'Blue'

    $window.RegisterName("Team${teamNumber}TotalRating", $teamRatingBlock)

    # Create the Expander for the team
    $expander = New-Object Windows.Controls.Expander
    $expander.Header = $teamName
    $expander.IsExpanded = $false
    $expander.Name = "Team${teamNumber}Expander"
    $expander.Background = "LightGray"

    $stackPanel = New-Object Windows.Controls.StackPanel

    $textBlockPlayers = New-Object Windows.Controls.TextBlock
    $textBlockPlayers.Text = "Players:"
    $textBlockPlayers.FontWeight = 'Bold'
    $stackPanel.Children.Add($textBlockPlayers)

    $dataGrid = New-Object Windows.Controls.DataGrid
    $dataGrid.Name = "Team${teamNumber}DataGrid"
    $dataGrid.AutoGenerateColumns = $false
    $dataGrid.Height = 250
    $dataGrid.IsReadOnly = $true
    $dataGrid.Background = "White"

    $columnPlayer = New-Object Windows.Controls.DataGridTextColumn
    $columnPlayer.Header = "Player"
    $columnPlayer.Binding = [Windows.Data.Binding]::new("Alias")
    $columnPlayer.Width = '*'
    $dataGrid.Columns.Add($columnPlayer)

    $columnOverall = New-Object Windows.Controls.DataGridTextColumn
    $columnOverall.Header = "Overall"
    $columnOverall.Binding = [Windows.Data.Binding]::new("Overall")
    $columnOverall.Width = '*'
    $dataGrid.Columns.Add($columnOverall)

    $columnClass = New-Object Windows.Controls.DataGridTextColumn
    $columnClass.Header = "Preferred Class"
    $columnClass.Binding = [Windows.Data.Binding]::new("PreferredClass")
    $columnClass.Width = '*'
    $dataGrid.Columns.Add($columnClass)

    $dataGrid.LoadingRow.Add({
        param ($sender, $args)
        $player = $args.Row.Item
        if ($player -and $player.PreferredClass -and $classColors.ContainsKey($player.PreferredClass)) {
            $rowColor = [Windows.Media.Brushes]::$(classColors[$player.PreferredClass])
            $args.Row.Background = $rowColor
        }
    })

    $stackPanel.Children.Add($dataGrid)

    $expander.Content = $stackPanel
    $teamStackPanel.Children.Add($expander)
    $teamStackPanel.Children.Add($teamRatingBlock)
    $teamsPanel.Children.Add($teamStackPanel)
    $window.RegisterName("Team${teamNumber}DataGrid", $dataGrid)

    # Initialize team rating
    $teamTotalRatings[$teamNumber] = 0

    # Sync the Expander and Team Selector
    $expander.Add_Expanded({
        $teamSelector.SelectedIndex = $teamNumber - 1
        $totalRatingTextBlock.Text = "        Total Rating: [$($teamTotalRatings[$teamNumber])]"
    }.GetNewClosure())
}

# Add teams UI and load captains into their respective teams
for ($i = 1; $i -le 6; $i++) {
    Add-TeamUI -teamNumber $i -teamName $teamNames[$i - 1]

    # Load the captains into their respective teams
    $captain = $combinedResults | Where-Object { $_.Alias.Trim().ToLower() -eq $captains[$i - 1].Trim().ToLower() }
    if ($captain) {
        $captain.PreferredClass = Get-PreferredClass -player $captain

        # Add captain to the respective team
        $teamDataGrid = $window.FindName("Team${i}DataGrid")
        $teamDataGrid.Items.Add($captain)

        # Update the team rating with the captain's overall
        $teamTotalRatings[$i] += [int]$captain.Overall

        # Update the total rating display
        $totalRatingTextBlock = $window.FindName("Team${i}TotalRating")
        if ($totalRatingTextBlock) {
            $totalRatingTextBlock.Text = "        Total Rating: [$($teamTotalRatings[$i])]"
        }

        Write-Host "Captain $($captain.Alias) added to Team $i"
    } else {
        Write-Host "Captain $($captains[$i - 1]) not found in player list!" -ForegroundColor Red
    }
}

# Ensure PreferredClass property exists for each player
$combinedResults | ForEach-Object {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name "PreferredClass" -Value $null -Force
}

$availablePlayersList = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

foreach ($player in $combinedResults) {
    if ($captains -notcontains $player.Alias) {
        $player.PreferredClass = Get-PreferredClass -player $player
        Write-Host "Loading Player: $($player.Alias) with Overall: $($player.Overall) and Preferred Class: $($player.PreferredClass)"
        $availablePlayersList.Add([pscustomobject]@{
            Alias = $player.Alias
            Overall = $player.Overall
            PreferredClass = $player.PreferredClass
            PlayerObject = $player
        })
    }
}

# Bind to the AvailablePlayersListBox
$availablePlayersListBox.ItemsSource = $availablePlayersList

$lastDraft = @{}
#$window.FindName("Team1Expander").IsExpanded = $true

# Handle sorting when selection in the ComboBox changes
$sortByComboBox.Add_SelectionChanged({
    $selectedIndex = $sortByComboBox.SelectedIndex

    # Sort based on the selected option
    $sortedPlayersList = switch ($selectedIndex) {
        0 { $availablePlayersList | Sort-Object Alias }                # Alias (A-Z)
        1 { $availablePlayersList | Sort-Object Alias -Descending }     # Alias (Z-A)
        2 { $availablePlayersList | Sort-Object Overall -Descending }   # Rating (High to Low)
        3 { $availablePlayersList | Sort-Object Overall }               # Rating (Low to High)
        default { $availablePlayersList }                               # Default (no change)
    }

    # Update the list with the sorted items
    $availablePlayersList.Clear()
    foreach ($player in $sortedPlayersList) {
        $availablePlayersList.Add($player)
    }

    # Keep focus on the Available Players ListBox
    $availablePlayersListBox.Focus()
})

# Double-click to draft player
$availablePlayersListBox.Add_MouseDoubleClick({
    $draftPlayerButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
})

# Create a global variable for tracking the last draft
$global:lastDraft = $null

# Initialize the default team selection
$teamSelector.SelectedIndex = 0

# Set up the UI elements
$teamSelector.Add_SelectionChanged({
    $selectedTeamIndex = $teamSelector.SelectedIndex + 1
    $selectedTeamExpander = $window.FindName("Team${selectedTeamIndex}Expander")
    if ($selectedTeamExpander) {
        $selectedTeamExpander.IsExpanded = $true
    }
})

# Draft Player Button Click Logic
$draftPlayerButton.Add_Click({
    $selectedItem = $availablePlayersListBox.SelectedItem
    if ($selectedItem) {
        $selectedTeamIndex = $teamSelector.SelectedIndex + 1
        $teamDataGrid = $window.FindName("Team${selectedTeamIndex}DataGrid")

        if ($teamDataGrid -eq $null) {
            Write-Host "Error: Could not find DataGrid for Team $selectedTeamIndex" -ForegroundColor Red
            return
        }

        $selectedPlayer = $selectedItem.PlayerObject

        # Add the player to the selected team's DataGrid and remove from available players list
        $availablePlayersList.Remove($selectedItem)
        $teamDataGrid.Items.Add($selectedPlayer)

        # Update the team's total rating
        $teamTotalRatings[$selectedTeamIndex] += [int]$selectedPlayer.Overall
        $totalRatingTextBlock = $window.FindName("Team${selectedTeamIndex}TotalRating")
        if ($totalRatingTextBlock) {
            $totalRatingTextBlock.Text = "        Total Rating: [$($teamTotalRatings[$selectedTeamIndex])]"
        }

        # Track the last draft for undo
        $global:lastDraft = @{
            Player = $selectedPlayer
            TeamIndex = $selectedTeamIndex
        }

        # Move to the next team in the dropdown
        $nextTeamIndex = ($selectedTeamIndex % 6)
        $teamSelector.SelectedIndex = $nextTeamIndex

        # Keep the focus on the Available Players ListBox
        $availablePlayersListBox.Focus()
    }
})

# Undo Last Draft Logic
$undoDraftButton.Add_Click({
    if ($global:lastDraft) {
        $teamDataGrid = $window.FindName("Team$($global:lastDraft.TeamIndex)DataGrid")
        $playerToRemove = $global:lastDraft.Player
        $itemToRemove = $teamDataGrid.Items | Where-Object { $_.Alias -eq $playerToRemove.Alias } | Select-Object -First 1

        if ($itemToRemove) {
            $teamDataGrid.Items.Remove($itemToRemove)
            $teamTotalRatings[$global:lastDraft.TeamIndex] -= [int]$playerToRemove.Overall
            $totalRatingTextBlock = $window.FindName("Team${global:lastDraft.TeamIndex}TotalRating")
            if ($totalRatingTextBlock) {
                $totalRatingTextBlock.Text = "        Total Rating: [$($teamTotalRatings[$global:lastDraft.TeamIndex])]"
            }

            $availablePlayersList.Add([pscustomobject]@{
                Alias = $playerToRemove.Alias
                Overall = $playerToRemove.Overall
                PreferredClass = $playerToRemove.PreferredClass
                PlayerObject = $playerToRemove
            })

            # Sort the available players list based on current sorting
            $sortByComboBox.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent))

            # Clear the last draft to prevent multiple undos
            $global:lastDraft = $null
        }

        $availablePlayersListBox.Focus()
    } else {
        Write-Host "No draft to undo." -ForegroundColor Red
    }
})

# Show the window
$window.ShowDialog()
