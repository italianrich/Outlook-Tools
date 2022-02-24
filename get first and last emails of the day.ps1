# open outlook and open the pst in outlook before starting
$Outlook = New-Object -ComObject 'Outlook.Application' -ErrorAction 'Stop' 
$pstRootFolder =$outlook.GetNameSpace('MAPI').Stores|?{("C:\Richard\outlook tools\test.pst" -eq [string]$_.FilePath)}|%{$_.GetRootFolder()} 
$AllEmail =$pstRootFolder.Folders|?{$_.FolderPath -match 'Sent Items'}|%{$_.items} 
$SaveTheseEmail =@()
$SaveTheseEmail += $AllEmail| select senton, sendername, to, cc, bcc, Subject 

#Remember to set the $Today variable to the last day of the 30 days period you are interested in. The format is MM/DD/YYYY

$Today = [Datetime]("02/24/2022") 
$startdate = $SaveTheseEmail |sort-object SentOn | select-object -first 1 
$Startfrom1 = $startdate.SentOn
write-output "Start For Loop" 

$Results = for($i = $Today; $i -ge $Startfrom1.AddDays(-1); $i = $i.AddDays(-1))

 
    { $SaveTheseEmail | Where-Object {
    $_.SentOn -le $i -and $_.SentOn -gt $i.AddDays(-1) } | select-object -first 1
    
    # need to run a similar loop to grab the last emails of the day

     $SaveTheseEmail | Where-Object {
    $_.SentOn -le $i -and $_.SentOn -gt $i.AddDays(-1) } | select-object -last 1
     }

    # Scripts results are in the $Results variable
    # My Pc locale is French, so set value separator to semicolon...
    $Results | Export-Csv .\output.csv -NoTypeInformation -Delimiter ";"