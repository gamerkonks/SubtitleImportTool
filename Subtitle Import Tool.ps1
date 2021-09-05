# Subtitle Import Tool
#===================================================#
# Parameters										#
#===================================================#

[String]$srcBaseDir = 'Z:\downloads\complete_movies\'
[String]$destBaseDir = 'Z:\Media\Movies\'
[String[]]$subFilter = @('*english.srt', '*turkish.srt', '*czech.srt', '*slovak.srt')
[String[]]$extFilter = @('*.mkv', '*.mp4')
[Int]$downloadedDays = -14

# Define Class
Class subData {
    [String]$movieTitle
    [String]$srcMovieDir
    [String]$srcMovieDirWC
    [String]$destMovieDir
    [String]$destMovieDirWC
    [String]$destMovieDirFilter
    [Bool]$destMovieDirExists
    [String]$destMovieFileName
    [String[]]$subLang
    [String[]]$srcSubDir
    [String[]]$destSubDir
}

# Instanciate Variables
$subObject = @()

#===================================================#
# Logic     										#
#===================================================#
Write-Host -backgroundcolor "white" -foregroundcolor "black" 'Copying Subtitles from ' $srcBaseDir ' to ' $destBaseDir
[System.Console]::ReadLine()

# Build array of directorys ending in "-RARBG"
Get-ChildItem $srcBaseDir -Filter *-RARBG | Where-Object {
    $_.LastWriteTime -gt (Get-Date).AddDays($downloadedDays) } | ForEach-Object { 
    $subObject += New-Object -TypeName subData -Property @{
        movieTitle    = $_.Name
        srcMovieDir   = $_.FullName
        srcMovieDirWC = $_.FullName + '\*'
    }
    Write-Host 'Found movie at' $_.FullName
}

# Process pretty movie title
foreach ($movie in $subObject) {
    $match = [regex]::Match($movie.movieTitle, '([.][1-2][09][0-9][0-9][.])')
    if ($match.Success) {
        $Index = $Match.Index
        $tmp = $movie.movieTitle
        $tmp = $tmp.Substring(0, $Index + 5)
        $tmp = $tmp.Replace('.', ' ')
        $tmp = $tmp.Insert($Index + 1, '(')
        $tmp = $tmp.Insert($Index + 6, ')')
        $movie.movieTitle = $tmp
    }
    else {
        $movie.movieTitle = 'Failed to Process Movie Title'
        Write-Host -backgroundcolor "red" 'Failed to Process Movie Title for ' $movie.movieTitle
    }
    
    # Find Relevant Subtitles
    Get-ChildItem -Path $movie.srcMovieDirWC -Include $subFilter -Recurse | ForEach-Object {
        $movie.srcSubDir += $_.FullName 
        $movie.subLang += $_.Name
        Write-Host 'Found Subtitles at' $_.FullName
    } 

    # Generate regex filter to find movie folder in destination directory
    $tmp = $movie.movieTitle
    $index = $tmp.Length - 6
    $tmp = $tmp.Insert($Index + 5, '\)')
    $tmp = $tmp.Insert($Index, '\')
    $tmp = $tmp.Insert(0, '(')    
    $tmp = $tmp.Replace(' ', ').*(')
    $movie.destMovieDirFilter = $tmp

    # Generate Destination Path
    Get-ChildItem -Path $destBaseDir -Directory | Where-Object { $_.Name -match $movie.destMovieDirFilter } | ForEach-Object { 
        $movie.destMovieDir = $_.FullName
        $movie.destMovieDirWC = $_.FullName + '\*'
    }

    # Verify Destination Directory Exists
    if ($null -eq $movie.destMovieDir) {
        Write-Host -BackgroundColor "red" 'No destination path for movie ' $movie.movieTitle
    }
    else {
        $movie.destMovieDirExists = Test-Path -Path $movie.destMovieDir
    }

    if ($movie.destMovieDirExists) {
        # Get destination movie filename
        Get-ChildItem -Path $movie.destMovieDirWC -Include $extFilter | ForEach-Object { $movie.destMovieFileName = $_.Name 
            $movie.destMovieFileName = $movie.destMovieFileName -replace '....$' }

        # Generate destination subtitle directories / rename   
        [int]$enCount = 0 
        [int]$czCount = 0
        [int]$skCount = 0
        [int]$trCount = 0
        for ($i = 0; $i -lt $movie.srcSubDir.Count; $i++) {
            $lang = $movie.subLang[$i].ToLower()
            $lang = $lang.Substring($lang.IndexOf('_') + 1)
            switch ($lang) {
                'english.srt' {
                    if ($enCount -gt 0) {
                        $enCount++
                        $lang = 'en' + $enCount + '.srt'
                    }
                    else {
                        $lang = 'en.srt' 
                        $enCount++
                    }
                }
                'czech.srt' {
                    if ($czCount -gt 0) {
                        $czCount++
                        $lang = 'cz' + $czCount + '.srt'
                    }
                    else {
                        $lang = 'cz.srt' 
                        $czCount++
                    }
                }
                'slovak.srt' {
                    if ($skCount -gt 0) {
                        $skCount++
                        $lang = 'sk' + $skCount + '.srt'         
                    }
                    else {
                        $lang = 'sk.srt'
                        $skCount++
                    }
                }
                'turkish.srt' {
                    if ($trCount -gt 0) {
                        $trCount++
                        $lang = 'tr' + $trCount + '.srt'
                    }
                    else {
                        $lang = 'tr.srt'  
                        $trCount++
                    }
                }
            }
            $movie.destSubDir += $movie.destMovieDir + '\' + $movie.destMovieFileName + '.' + $lang
        }
    }
}

# WhatIf Copy Subtitle Files to Corresponding Media Directory
foreach ($movie in $subObject) {
    if ($movie.destMovieDirExists) {
        for ($i = 0; $i -lt $movie.srcSubDir.Count; $i++) {
            Copy-Item -Path $movie.srcSubDir[$i] -Destination $movie.destSubDir[$i] -WhatIf
        }
    }
    else {
        Write-Host -BackgroundColor "red" "Not coping subtitles for '" $movie.movieTitle "', Destination directory doesn't exist."
    }
}

# Copy Subtitle Files to Corresponding Media Directory
Write-Host -backgroundcolor "white" -foregroundcolor "black" 'Confirm Copy? Type "y" to copy: ' -NoNewline
[string]$conf = Read-Host

if ($conf.ToLower() -eq 'y') {
    foreach ($movie in $subObject) {
        if ($movie.destMovieDirExists) {
            for ($i = 0; $i -lt $movie.srcSubDir.Count; $i++) {
                Copy-Item -Path $movie.srcSubDir[$i] -Destination $movie.destSubDir[$i]
                Write-Host 'Copied ' $movie.srcSubDir[$i] ' to ' $movie.destSubDir[$i]
            }
        }
        else {
            Write-Host -BackgroundColor "red" "Didn't copy subtitles for '" $movie.movieTitle "', Destination Directory doesn't exist."
        }
    }
}
else {
    Write-Host 'Copy Cancelled!'
}

[System.Console]::ReadLine()
