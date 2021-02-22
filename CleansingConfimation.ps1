Function Show-Msgbox {
    Param([string]$message=$(Throw "You must specify a default message"),
        [string]$button="okonly",
        [string]$icon="critical",
        [string]$title="Message Box"
    )

    # Buttons: OkOnly, OkCancel, AbortRetryIgnore, YesNoCancel, YesNo, RetryCancel
    # Icons: Critical, Question, Exclamation, Information
    [reflection.assembly]::loadwithpartialname("microsoft.visualbasic") | Out-Null
    [microsoft.visualbasic.interaction]::Msgbox($message,"$button,$icon",$title)

}

#Ensure to change location of purgeable folder within the host device.
$question_result=Show-Msgbox -message "Are you ready to Proceed with Path relocation and Cleansing?" -icon "exclamation" -button "YesNo" -title "Confirmation of Path Cleansing"
Switch ($question_result) {
"Yes" { Get-Content "C:\Users\mmorales\Documents\Output.txt"| ForEach-Object { Move-Item -Path $_ -Destination "C:\Users\mmorales\Purgeable Folder" -Verbose } }
"No" { break }
}