Import-Module pwps_dab -DisableNameChecking
Import-Module C:\Scripts\PWPowerShell\PW-Login.psm1
PW-LoginLochnerCloud

# Saved Search looking for Documents that meet our Requirements
$SavedSearch = 'Checked Out'

# Get all checked out documents.  
$PWDocuments = Get-PWDocumentsBySearch -SearchName $SavedSearch -GetAttributes 


#Put usernames and file paths in array and organize them by username
$DataTable = @() 

foreach($PWDoc in $PWDocuments){
    $User = $PWdoc.DocumentOutToName
    $FilePath = $PWdoc.FullPath
    $CheckedOutDate = $PWdoc.StatusChangeDate

    $item = New-Object PSObject  

    $item | add-member -force -type NoteProperty -Name "Username" -Value $User
    $item | add-member -force -type NoteProperty -Name "FilePath" -Value $FilePath
    $item | add-member -force -type NoteProperty -Name "CheckedOutDate" -Value $CheckedOutDate

    $DataTable += $item
}
    #Create list of files checked out for more than 3 days
    $OldFiles = $DataTable | Where-Object "CheckedOutDate" -lt (Get-Date).AddDays(-3)

    #Create a sorted table of users and files checked out
    $SortedTable = $OldFiles | Sort-object "Username"

    #Create list of email addresses
    $EmailAddresses = $SortedTable | Sort-object "Username" -unique | Select-Object -ExpandProperty "Username"


        Foreach ($Email in $EmailAddresses){
            $FilesCheckedOut = $SortedTable | Where-Object "Username" -eq "bchung@hwlochner.com" |Select-Object -ExpandProperty "FilePath"
        
        
        $HtmlTable ="<table border='1' align='Center' cellpadding='2' cellspacing='0' style='color:black;font-family:arial,helvetica,sans-serif;text-align:left;'>
        <tr style ='font-size:12px;font-weight: normal;background: #FFFFFF'>
        <th align=left><b>FilePath</b></th>
        </tr>"

        foreach ($Row in $FilesCheckedout){
            $HtmlTable +="<tr style='font-size:12px;background-color:#FFFFFF'>
             <td>" + $Row + "</td>
            </tr>"
        }
        
            $EmailBody = @"
                <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {behavior:url(#default#VML);}
                o\:* {behavior:url(#default#VML);}
                w\:* {behavior:url(#default#VML);}
                .shape {behavior:url(#default#VML);}
                </style><![endif]--><style><!--
                /* Font Definitions */
                @font-face
	            {font-family:"Cambria Math";
	            panose-1:2 4 5 3 5 4 6 3 2 4;}
                @font-face
	            {font-family:Calibri;
	            panose-1:2 15 5 2 2 2 4 3 2 4;}
                /* Style Definitions */
                p.MsoNormal, li.MsoNormal, div.MsoNormal
            	{margin:0in;
	            font-size:11.0pt;
	            font-family:"Calibri",sans-serif;}
                span.EmailStyle17
            	{mso-style-type:personal-compose;
            	font-family:"Calibri",sans-serif;
	            color:windowtext;}
                .MsoChpDefault
            	{mso-style-type:export-only;
	            font-family:"Calibri",sans-serif;}
                @page WordSection1
	            {size:8.5in 11.0in;
	            margin:1.0in 1.0in 1.0in 1.0in;}
                div.WordSection1
	            {page:WordSection1;}
                --></style><!--[if gte mso 9]><xml>
                <o:shapedefaults v:ext="edit" spidmax="1026" />
                </xml><![endif]--><!--[if gte mso 9]><xml>
                <o:shapelayout v:ext="edit">
                <o:idmap v:ext="edit" data="1" />
                </o:shapelayout></xml><![endif]--></head><body lang=EN-US link="#0563C1" vlink="#954F72" 
                style='word-wrap:break-word'><div class=WordSection1>
                <p class=MsoNormal>You are receiving this email because you left one or more Projectwise documents checked out for 3 or more days. We ask that all files are checked in at the end of each work day 
                so that we are able to perform backups of all latest work. If they are not checked in, the latest version of the file only exists on your C drive which we are unable to back up. <o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Below is a list of files you have checked out on the LochnerCloud data source for more than 3 days. Please be sure to check them in when you are finished working 
                on them so that we can best manage ProjectWise data.<o:p></o:p><br>
                <br><b> <table> $HtmlTable </table> <br/></b>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Note that if you would like to see all checked out files and easily check them all back in, you can use the local document organizer to do so<o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>First, open the document organizer located at the top of Projectwise<o:p></o:p></p>
                <p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><img width=383 height=117 style='width:3.9895in;height:1.2187in' <img src="cid:Attachment1" /><o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Then, highlight and right click on any files you would like to check in, and select &#8220;check in&#8221;<o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><img width=681 height=451 style='width:7.0937in;height:4.6979in' <img src="cid:Attachment2" /><o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p>
                </p><p class=MsoNormal>If you have any questions or issues, please contact PWHelp@hwlochner.com<o:p></o:p>
                </p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Thank you,<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal><b><i>The Lochner Projectwise Team</i></b><o:p></o:p></p>
                
                </div></body></html>
"@ 

        $smtpServer = "smtp.hwlochner.com"
        $msg = New-Object Net.Mail.MailMessage
        $smtp = new-object Net.Mail.SmtpClient($smtpServer)

        $msg.from = "LochnerPWAdmin@hwlochner.com"
        $msg.To.Add("sbooth@hwlochner.com")
        $msg.subject = "LochnerCloud Projectwise Documents Checked Out Over 3 Days"
        $msg.IsBodyHtml = $True
        $msg.Body = $EmailBody

        $attachment1 = New-Object System.Net.Mail.Attachment -ArgumentList "C:\Temp\pic1.png"
        $attachment2 = New-Object System.Net.Mail.Attachment -ArgumentList "C:\Temp\pic2.png"

        $attachment1.ContentDisposition.Inline = $True
        $attachment2.ContentDisposition.Inline = $True

        $attachment1.ContentDisposition.DispositionType = "Inline"
        $attachment2.ContentDisposition.DispositionType = "Inline"

        $attachment1.ContentType.MediaType = "image/png"
        $attachment2.ContentType.MediaType = "image/png"

        $attachment1.ContentID = "Attachment1"
        $attachment2.ContentID = "Attachment2"

        $msg.Attachments.add($attachment1)
        $msg.Attachments.add($attachment2)

        $smtp.send($msg)
        $attachment1.Dispose()
        $attachment2.Dispose()
    }

    $SortedTable | Export-Csv -path "C:\Scripts\Table.csv"
  
    $MailMessage = @{
        To = "sbooth@hwlochner.com"
        From = "LochnerPWAdmin@hwlochner.com"
        Subject = "List of Users With Files Checked Out More Than 3 Days"
        Body = "See attached list of users"
        SMTPServer = "smtp.hwlochner.com"
        Attachment = "C:\Scripts\Table.csv"
    }
        Send-MailMessage @MailMessage


