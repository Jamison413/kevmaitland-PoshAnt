function new-meetingRoom(){
    [cmdletbinding()]
    Param (
        [parameter(Mandatory = $true)]
        [string]$meetingRoomName 
        ,[parameter(Mandatory = $true)]
        [string]$meetingRoomCity 
        ,[parameter(Mandatory = $true)]
        [string]$meetingRoomCountryCode 
        ,[parameter(Mandatory = $false)]
        [string]$meetingRoomCapacity 
        ,[parameter(Mandatory = $false)]
        [string]$meetingRoomPhoneNumber 
        )
    $meetingRoomCompositeName = "Meeting Room ($meetingRoomCountryCode, $meetingRoomCity) - $meetingRoomName"
    try{$meetingRoom = Get-Mailbox $meetingRoomCompositeName -ErrorAction SilentlyContinue}
    Catch{}
    if(!$meetingRoom){
        $meetingRoom = New-Mailbox -Name $meetingRoomCompositeName -Room -ResourceCapacity $meetingRoomCapacity -Phone $meetingRoomPhoneNumber -Office "$meetingRoomCountryCode - $meetingRoomCity"
        }
    Start-Sleep -Seconds 15
    Set-CalendarProcessing $meetingRoom.ExternalDirectoryObjectId -AutomateProcessing AutoAccept -BookingWindowInDays 365 -MaximumDurationInMinutes 10080 -DeleteSubject $false -EnforceSchedulingHorizon $false

    }

$meetingRoomName = "P3"
$meetingRoomCity = "Manlleu"
$meetingRoomCountryCode = "ESP"
$meetingRoomCapacity = "12"
$meetingRoomPhoneNumber = ""

new-meetingRoom -meetingRoomName $meetingRoomName -meetingRoomCity $meetingRoomCity -meetingRoomCountryCode $meetingRoomCountryCode -meetingRoomCapacity $meetingRoomCapacity -meetingRoomPhoneNumber $meetingRoomPhoneNumber
