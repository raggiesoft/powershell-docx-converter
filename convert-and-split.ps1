<#
.SYNOPSIS
    A PowerShell script to convert structured Word documents into a multi-file Markdown project.

.DESCRIPTION
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program. If not, see <https://www.gnu.org/licenses/>.

.NOTES
    Author: Michael Ragsdale
    Copyright (c) 2025 Michael Ragsdale
#>

# --- SCRIPT TO CONVERT DOCX AND SPLIT INTO A NESTED, NUMBERED FOLDER STRUCTURE ---

# The 'param' block defines the parameters (or command-line options) the script accepts.
# This makes the script flexible and configurable without having to edit the code directly.
param(
    # A specific .docx file to process. If not provided, the script will process all .docx files.
    [Parameter(Mandatory=$false)]
    [string]$DocxFile,

    # The number of leading zeros for file/folder prefixes (e.g., 3 means '001', '002').
    [Parameter(Mandatory=$false)]
    [int]$Padding = 3,

    # A switch parameter. If present (-PurgeOutputFolder), it deletes the output folder's contents before running.
    [Parameter(Mandatory=$false)]
    [switch]$PurgeOutputFolder,

    # A switch parameter. If present (-ObsidianFriendlyLinks), it creates simpler [[wiki-style]] links.
    [Parameter(Mandatory=$false)]
    [switch]$ObsidianFriendlyLinks,

    # A switch parameter for showing help. The Alias '?' allows users to type -? instead of -Help.
    [Parameter(Mandatory=$false)]
    [Alias("?")]
    [switch]$Help
)

# --- HELP FUNCTION ('MAN PAGE') ---
# This block checks if the -Help parameter was used or if the script was run with no parameters at all.
# If so, it prints a detailed help message formatted like a Unix MAN page and then exits.
if ($Help -or (($PSBoundParameters.Count -eq 0) -and (-not $DocxFile))) {
    Write-Host "`nNAME" -ForegroundColor Green
    Write-Host "    Convert-And-Split.ps1 - Converts Word DOCX files into a structured, multi-file Markdown project."
    
    Write-Host "`nSYNOPSIS" -ForegroundColor Green
    Write-Host "    & `".\Convert-And-Split.ps1`" [-DocxFile `"YourStory.docx`"] [-Padding <int>] [-PurgeOutputFolder] [-ObsidianFriendlyLinks] [-Help | -?]"
    
    Write-Host "`nDESCRIPTION" -ForegroundColor Green
    Write-Host "    This script uses Pandoc to convert .docx files structured with 'Heading 1', 'Heading 2',"
    Write-Host "    and 'Heading 3' styles into a nested folder structure suitable for Obsidian."
    Write-Host "    It can also read a YAML metadata block from the top of the Word document."
    Write-Host "    If -DocxFile is omitted, it processes ALL .docx files in the current directory."
    
    Write-Host "`nPARAMETERS" -ForegroundColor Green
    Write-Host "    -DocxFile <string>" -ForegroundColor Yellow
    Write-Host "        (Optional) The name of a specific Word document to process. If omitted, all .docx files are processed."
    Write-Host "    -Padding <int>" -ForegroundColor Yellow
    Write-Host "        (Optional) The number of leading zeros for folder/file prefixes. Defaults to 3."
    Write-Host "    -PurgeOutputFolder" -ForegroundColor Yellow
    Write-Host "        (Optional) If present, deletes all files/folders in the output directory before creating new ones."
    Write-Host "    -ObsidianFriendlyLinks" -ForegroundColor Yellow
    Write-Host "        (Optional) If present, creates simple [[filename]] links that rely on Obsidian's internal linking."
    Write-Host "    -Help, -?" -ForegroundColor Yellow
    Write-Host "        (Optional) Displays this help message and exits."

    Write-Host "`nEXAMPLES" -ForegroundColor Green
    Write-Host "    # Process a single file with default settings" -ForegroundColor Gray
    Write-Host "    .\Convert-And-Split.ps1 -DocxFile 'MyNovel.docx'"
    Write-Host ""
    Write-Host "    # Process all DOCX files in the folder, purging the output folders first" -ForegroundColor Gray
    Write-Host "    .\Convert-And-Split.ps1 -PurgeOutputFolder"

    Write-Host "`nNOTES" -ForegroundColor Green
    Write-Host "    This script requires Pandoc to be installed and accessible in the system's PATH environment variable."
    Write-Host "    The script expects a document structure of H1 for Book, H2 for Chapter, and H3 for Part."

    Write-Host "`nAUTHOR" -ForegroundColor Green
    Write-Host "    Written by Michael R."
    Write-Host ""
    return # Exit the script after showing help.
}

# --- CONFIGURATION ---
$pandocPath = "pandoc" # Assumes pandoc is in the system's PATH.

# --- SCRIPT LOGIC ---
$scriptDirectory = $PSScriptRoot # An automatic variable that gets the directory where the script itself is located.
$filesToProcess = @() # Initialize an empty array to hold the file(s) we need to process.

# Check if the user specified a single file or if we should find all .docx files.
if (-not [string]::IsNullOrWhiteSpace($DocxFile)) {
    # User provided a filename. Build the full path and check if it exists.
    $fullPath = Join-Path -Path $scriptDirectory -ChildPath $DocxFile
    if (-not (Test-Path $fullPath)) { Write-Error "Source file not found at: $fullPath"; return }
    $filesToProcess += Get-Item -Path $fullPath
}
else {
    # No filename provided. Get all .docx files in the script's directory.
    $filesToProcess = Get-ChildItem -Path $scriptDirectory -Filter *.docx
    Write-Host "No specific file provided. Found $($filesToProcess.Count) DOCX files to process." -ForegroundColor Cyan
}

# This is the main loop. It will run once for each DOCX file found.
foreach ($file in $filesToProcess) {
    # --- STATE RESET ---
    # Reset all counters and temporary arrays. This is CRITICAL to ensure that data from
    # a previously processed file doesn't "bleed over" into the next one.
    $bookCounter = 0; $chapterCounter = 0; $partCounter = 0
    $lastH1 = ""; $lastH2 = ""
    $allParts = @(); $fileInfos = @()

    $sourceDocxFile = $file.FullName
    
    # --- DYNAMIC HEADER CREATION ---
    # This block creates a professional-looking, full-width header in the console for each file being processed.
    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $separatorLine = "-" * $consoleWidth
    $processingText = " Processing: $($file.Name) "
    $totalPadding = $consoleWidth - $processingText.Length
    if ($totalPadding -lt 0) { $totalPadding = 0 } # Prevents errors if the filename is very long.
    $leftPadding = " " * [Math]::Floor($totalPadding / 2)
    $centeredText = "$leftPadding$processingText".PadRight($consoleWidth)

    Write-Host "`n$separatorLine" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host $centeredText -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host "$separatorLine`n" -ForegroundColor White -BackgroundColor DarkBlue

    # --- FILE & FOLDER SETUP ---
    # Sanitize the DOCX filename to create a clean root folder name (e.g., "My Novel.docx" -> "my-novel").
    $kebabCaseFolderName = [System.IO.Path]::GetFileNameWithoutExtension($sourceDocxFile).ToLower() -replace '[^a-z0-9\s-]', '' -replace '\s+', '-'
    $rootOutputDirectory = Join-Path -Path $scriptDirectory -ChildPath $kebabCaseFolderName
    # Create the root folder if it doesn't exist.
    if (-not (Test-Path $rootOutputDirectory)) { New-Item -Path $rootOutputDirectory -ItemType Directory | Out-Null }
    # If the -PurgeOutputFolder switch was used, delete everything inside the output folder.
    if ($PurgeOutputFolder) { Get-ChildItem -Path $rootOutputDirectory | Remove-Item -Recurse -Force }
    # Define a temporary file for the Pandoc conversion.
    $intermediateMdFile = Join-Path -Path $scriptDirectory -ChildPath "temp_conversion.md"

    # --- STEP 1: PANDOC CONVERSION ---
    # Use a try/catch block for graceful error handling.
    try {
        # Execute Pandoc.
        # -t "gfm+yaml_metadata_block" tells it to convert to GitHub Flavored Markdown and to recognize YAML blocks.
        # --wrap=none prevents it from adding hard line breaks in paragraphs.
        # 2> $null redirects any error messages from Pandoc so we can handle them cleanly.
        & $pandocPath "$sourceDocxFile" -t "gfm+yaml_metadata_block" --wrap=none -o "$intermediateMdFile" 2> $null
        Write-Host "Pandoc conversion successful." -ForegroundColor Green
    }
    catch { 
        # This block runs if Pandoc fails.
        if ($_.Exception.Message -like "*Permission denied*") {
            Write-Host "WARNING: Could not process '$($file.Name)'. The file is likely open. Please close it and try again. Skipping." -ForegroundColor Black -BackgroundColor Yellow
        } else {
            Write-Host "WARNING: Pandoc execution failed for '$($file.Name)'. Error: $($_.Exception.Message). Skipping." -ForegroundColor Black -BackgroundColor Yellow
        }
        continue # Skip to the next file in the loop.
    }
    
    # Check if the conversion actually produced a file.
    if (-not (Test-Path $intermediateMdFile)) { 
        Write-Host "WARNING: Intermediate file not created for '$($file.Name)'. This may be due to a file lock or a Pandoc error. Skipping." -ForegroundColor Black -BackgroundColor Yellow
        continue 
    }

    # --- STEP 2: PARSE MARKDOWN & COLLECT DATA ---
    # Read the entire temporary markdown file into memory.
    $allLines = Get-Content -Path $intermediateMdFile
    
    # Replace Microsoft Word's curly "smart quotes" with standard straight quotes to prevent YAML errors.
    $allLines = $allLines -replace '“', '"' -replace '”', '"' -replace '‘', "'" -replace '’', "'"

    $customYamlContent = @()
    $mainContentStartIndex = 0
    
    # Check for and extract the custom YAML block from the top of the file.
    if ($allLines[0] -eq "---") {
        for ($j = 1; $j -lt $allLines.Count; $j++) {
            if ($allLines[$j] -eq "---" -or $allLines[$j] -eq "...") {
                $mainContentStartIndex = $j + 1 # Mark the line where the main content begins.
                break
            }
            $customYamlContent += $allLines[$j] # Add the line to our custom YAML array.
        }
    }
    
    # Isolate the actual story content, skipping the YAML block.
    $contentLines = $allLines | Select-Object -Skip $mainContentStartIndex
    
    $currentH1 = ""; $currentH2 = ""; $currentPartContent = $null
    $partH1 = ""; $partH2 = ""

    # Loop through the story content line by line to split it into parts based on headings.
    foreach ($line in $contentLines) {
        if ($line.StartsWith("### ")) { # A new Part (H3) is found.
            if ($null -ne $currentPartContent) {
                # Save the previously collected part, associating it with the headings that were active when it started.
                $allParts += @{ H1 = $partH1; H2 = $partH2; Content = $currentPartContent }
            }
            # Start a new part.
            $currentPartContent = @($line)
            # Capture the current Book (H1) and Chapter (H2) state for this new part.
            $partH1 = $currentH1; $partH2 = $currentH2
        } elseif ($line.StartsWith("## ")) { # A new Chapter (H2) is found.
            $currentH2 = $line
        } elseif ($line.StartsWith("# ")) { # A new Book (H1) is found.
            $currentH1 = $line; $currentH2 = ""
        } elseif ($null -ne $currentPartContent) {
            # This is a regular line of text, add it to the current part's content.
            $currentPartContent += $line
        }
    }
    # After the loop, save the very last part that was being collected.
    if ($null -ne $currentPartContent) {
        $allParts += @{ H1 = $partH1; H2 = $partH2; Content = $currentPartContent }
    }

    # --- VALIDATION STEP ---
    # If no '###' headings were found, the document is not structured correctly.
    if ($allParts.Count -eq 0) {
        Write-Host "WARNING: No parts found in '$($file.Name)'. Ensure the document is structured with '###' for parts. Skipping." -ForegroundColor Black -BackgroundColor Yellow
        Remove-Item -Path $intermediateMdFile -ErrorAction SilentlyContinue
        continue
    }

    # --- STEP 3: GENERATE FILE INFO & METADATA ---
    $formatString = "{0:D$($Padding)}" # Creates the format for padded numbers (e.g., "001").
    
    # Loop through the collected parts to generate final file info, including correct numbering.
    for ($i = 0; $i -lt $allParts.Count; $i++) {
        $part = $allParts[$i]
        
        # Increment book, chapter, and part counters based on heading changes.
        if ($part.H1 -ne $lastH1) {
            $bookCounter++; $chapterCounter = 1; $partCounter = 1
            $lastH1 = $part.H1; $lastH2 = $part.H2
        } elseif ($part.H2 -ne "" -and $part.H2 -ne $lastH2) {
            $chapterCounter++; $partCounter = 1
            $lastH2 = $part.H2
        } else {
            $partCounter++
        }

        $paddedBook = $formatString -f $bookCounter
        $paddedChapter = $formatString -f $chapterCounter
        $paddedPart = $formatString -f $partCounter

        # Sanitize headings to create clean, URL-friendly folder and file names.
        $bookFolderName = "$($paddedBook)-$($part.H1 -replace '#+\s*' | ForEach-Object { $_.ToLower() -replace '[^a-z0-9\s-]', '' -replace '\s+', '-' } | ForEach-Object { $_.Trim('-') })"
        $chapterFolderName = if (-not [string]::IsNullOrWhiteSpace($part.H2)) {
            "$($paddedChapter)-$($part.H2 -replace '#+\s*', '' | ForEach-Object { $_.ToLower() -replace '[^a-z0-9\s-]', '' -replace '\s+', '-' } | ForEach-Object { $_.Trim('-') })"
        } else { "$($paddedChapter)-chapter-$paddedChapter" }
        
        $partHeading = $part.Content[0].Trim()
        $kebabCaseHeading = ($partHeading -replace '#+\s*', '').ToLower() -replace '[^a-z0-9\s-]', '' -replace '\s+', '-' | ForEach-Object { $_.Trim('-') }
        $fileName = "$($paddedPart)-$($kebabCaseHeading).md"
        
        # Store all the generated information for this part in our fileInfos array.
        $fileInfos += @{
            FileName = $fileName; RelativePath = "$($bookFolderName)/$($chapterFolderName)/$($fileName)";
            BookName = $part.H1 -replace '#+\s*', ''; ChapterName = $part.H2 -replace '#+\s*', '';
            PartName = $partHeading -replace '#+\s*', ''; Content = $part.Content;
        }
    }

    # --- STEP 4: WRITE FINAL MARKDOWN FILES ---
    $titleName = (Get-Culture).TextInfo.ToTitleCase([System.IO.Path]::GetFileNameWithoutExtension($sourceDocxFile))
    
    # Loop through our processed file information one last time to write the actual files.
    for ($i = 0; $i -lt $fileInfos.Count; $i++) {
        $currentInfo = $fileInfos[$i]
        $currentFolderPath = (Split-Path $currentInfo.RelativePath -Parent)
        $fullFolderPath = Join-Path -Path $rootOutputDirectory -ChildPath $currentFolderPath
        # Create the nested Book/Chapter folders if they don't exist.
        if (-not (Test-Path $fullFolderPath)) { New-Item -Path $fullFolderPath -ItemType Directory | Out-Null }
        
        # Generate the 'previous' and 'next' navigation links for the YAML.
        $previousLink = '""'; $nextLink = '""'
        if ($ObsidianFriendlyLinks) {
            if ($i -gt 0) { $previousLink = """[[$(($fileInfos[$i-1].FileName -replace '.md$', ''))|$(($fileInfos[$i-1].PartName))]]""" }
            if (($i + 1) -lt $fileInfos.Count) { $nextLink = """[[$(($fileInfos[$i+1].FileName -replace '.md$', ''))|$(($fileInfos[$i+1].PartName))]]""" }
        } else {
            if ($i -gt 0) { $previousLink = """[[$(($fileInfos[$i-1].RelativePath -replace '\\', '/'))|$(($fileInfos[$i-1].PartName))]]""" }
            if (($i + 1) -lt $fileInfos.Count) { $nextLink = """[[$(($fileInfos[$i+1].RelativePath -replace '\\', '/'))|$(($fileInfos[$i+1].PartName))]]""" }
        }
        
        # Build the final YAML block by merging the custom YAML from the Word doc with the script-generated YAML.
        $generatedYaml = @(
            "title: ""$titleName""", "book: ""$($currentInfo.BookName)""",
            "chapter: ""$($currentInfo.ChapterName)""", "part: ""$($currentInfo.PartName)""",
            "previous: $previousLink", "next: $nextLink"
        )
        $finalYaml = @("---") + $customYamlContent + $generatedYaml + @("---")
        
        # Build the main content of the file, adding the Part heading back in.
        $mainContent = @(); $mainContent += "# $($currentInfo.PartName)"; $mainContent += $currentInfo.Content | Select-Object -Skip 1
        $rawFileContent = ($finalYaml + $mainContent) -join [System.Environment]::NewLine
        # A final cleanup step to fix any stray backslashes Pandoc might have added.
        $finalFileContent = $rawFileContent -replace "\\'", "'"
        
        $fullOutputPath = Join-Path -Path $rootOutputDirectory -ChildPath $currentInfo.RelativePath
        Set-Content -Path $fullOutputPath -Value $finalFileContent
        Write-Host "Created file: $($currentInfo.RelativePath)"
    }

    # Clean up the temporary file.
    Remove-Item -Path $intermediateMdFile
    Write-Host "Finished processing '$($file.Name)'." -ForegroundColor White -BackgroundColor DarkGreen
}

Write-Host "`nAll operations complete." -ForegroundColor Cyan
