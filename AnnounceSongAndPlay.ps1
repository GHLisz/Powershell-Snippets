$dir = "D:\Music\Classical_Music"

Function Main(){
    Get-ChildItem -Path $dir -Recurse -Include *.aac,*.flac,*.m4a,*.mp3,*.wav,*.wma | Sort-Object {Get-Random} | % {
        $path = $_.FullName
        Play-Music $path
        Write-Host "Played before is: ", $path
        Speak-Sth "Played before is: "
        Speak-Sth ([System.Io.Path]::GetFileNameWithoutExtension($path))
    }
}

Function Play-Music($path){
    Add-Type -AssemblyName presentationCore
    $player = New-Object System.Windows.Media.MediaPlayer
    $player.Open($path)
    Start-Sleep 1
    $duration = $player.NaturalDuration.TimeSpan.TotalSeconds
    $player.Play()
    Start-Sleep $duration
    $player.Stop()
    $player.Close()
}

Function Speak-Sth($words){
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
    # $synthesizer.GetInstalledVoices().VoiceInfo  # ref to this for available voice or language
    $synthesizer.Volume = 100
    $synthesizer.SelectVoice('Microsoft Zira Desktop')
    $synthesizer.Speak($words)
    $synthesizer.Dispose()
}

. Main
