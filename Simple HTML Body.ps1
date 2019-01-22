$ol = New-Object -comObject Outlook.Application
$mail = $ol.CreateItem(0)
$mail.SentOnBehalfOfName = "ivanenev@mysolutionshop.com"
$mail.Subject = "Top demand apps-SOURCE CLARIFICATION"
$mail.HTMLBody="<head><title></title><STYLE TYPE=""text/css"">
<!--
body 
{
	font-size: 10pt;
	font-family: arial;
}
-->
</style>
</head>
<body style=""FONT-FAMILY: Arial; FONT-SIZE: 10pt; CURSOR: auto"">
<p>&nbsp;</p>
<p><span style=""font-size: 10pt; font-weight: bold;"">Clay Gillespie, </span><span style=""font-size: 8pt; font-weight: bold;"">BBA CFP CIM FCSI</span><br /><span style=""color: black; font-size: 10pt;"">Financial Advisor &amp; Portfolio Manager</span><br /><span style=""color: black; font-size: 10pt;"">Managing Director</span></p>
<p><span style=""color: black; font-size: 10pt;""><span style=""font-size: 8pt;""><strong>Rogers Group Financial (RGF)|</strong> tel: 604.732.6551 | fax: 604.732.6553 | tollfree: 800.784.6066 | </span><a href=""mailto:cgillespie@rogersgroup.com""><span style=""font-size: 8pt;"">cgillespie@rogersgroup.com</span></a><br /><span style=""font-size: 8pt;""><span style=""font-size: 8pt;"">1701 West Broadway&nbsp;| Vancouver, BC, V6J&nbsp;1Y3 </span>| </span><span style=""font-size: 8pt;""><a href=""http://www.rogersgroup.com/claygillespie"">www.rogersgroup.com/claygillespie</a></span></span></p>
<p><span style=""color: black; font-size: 10pt;""><span style=""font-size: 8pt; font-weight: bold;"">Financial Planning Team:</span></span></p>
<table class=""null"" style=""width: 781px; height: 94px;"" border=""1"" cellspacing=""1"" cellpadding=""0"">
<tbody>
<tr>
<th style=""width: 172.16px;"">
<p><span style=""font-size: 8pt;"">John Hale</span><span style=""font-size: 8pt;"">, </span><span style=""font-size: 6pt;"">CPA CGA CFP</span><br /><span style=""font-family: Arial; font-size: 8pt;"">Financial&nbsp;Advisor</span><br /><span style=""font-family: Arial; font-size: 8pt;"">Direct:&nbsp;604.737.6762</span><br /><a title=""mailto:jhale@rogersgroup.com

Please press CTRL+Click to open link"" href=""mailto:jhale@rogersgroup.com""><span style=""font-size: 8pt;"">jhale@rogersgroup.com</span></a></p>
</th>
<th style=""width: 213.21px;"">
<p><span style=""font-size: 8pt;"">Linson Chen, </span><span style=""font-size: 6pt;"">BA </span><span style=""font-size: 6pt;"">CFP </span><span style=""font-size: 6pt;"">CIM CLU FCSI</span><br /><span style=""font-size: 8pt;"">Financial&nbsp;Advisor &amp; Associate Portfolio Manager</span><br /><span style=""font-size: 8pt;"">Direct:&nbsp;604.737.6784</span><br /><a href=""mailto:jhale@rogersgroup.com""><span style=""font-size: 8pt;"">lchen@rogersgroup.com</span></a></p>
</th>
<th style=""width: 193.33px;""><span style=""font-size: 8pt;"">Lorraine Watson</span><br /><span style=""font-size: 8pt;"">Executive Assistant</span><br /><span style=""font-size: 8pt;"">Direct: 604.737.6787</span><br /><a href=""mailto:lwatson@rogersgroup.com""><span style=""font-size: 8pt;"">lwatson@rogersgroup.com</span></a></th>
<th style=""width: 187.3px;"">
<p><span style=""font-size: small;""><span style=""font-size: 8pt;"">Carly O'Connell, </span><span style=""font-size: 6pt;"">BA</span><br /></span><span style=""color: black; font-size: 8pt;"">Administrative Assistant</span><br style=""color: black; font-size: 7pt;"" /><span style=""color: black; font-size: 8pt;"">Direct: 604.737.6752</span><br style=""color: black; font-size: 8pt;"" /><span style=""font-size: 8pt;""><a title=""mailto:coconnell@rogersgroup

Please press CTRL+Click to open link"" href=""mailto:coconnell@rogersgroup.com"">coconnell@rogersgroup.com</a></span></p>
</th>
</tr>
</tbody>
</table>
<table style=""width: 220px;"" border=""0"" cellspacing=""10"" cellpadding=""0"">
<tbody>
<tr>
<td align=""left"" valign=""middle""><a title=""http://www.rogersgroup.com/ Please press CTRL+Click to open link"" href=""http://www.rogersgroup.com/"" target=""""><img src=""http://www.rogersgroup.com/Portals/0/RGFLOGO_emailtemplate.jpg"" alt="""" border=""0"" /></a></td>
</tr>
</tbody>
</table>
<p style=""line-height: 14pt;""><span style=""font-size: xx-small;""><span style=""font-size: 7pt;""><span style=""font-size: 8pt; font-weight: bold;"">Rogers Group Financial Advisors Ltd. | Rogers Group Investment Advisors Ltd.</span></span></span></p>
<p>IMPORTANT NOTICE: This email message and all attachments are intended solely for the use of the addressee and may contain confidential information. If the reader is not the intended recipient, you are hereby notified that any dissemination, distribution, copying, or other use of this message or its attachments is strictly prohibited. If this message is received in error, please notify the sender immediately and delete it from your computer. Should you wish to stop receiving all email communications from Rogers Group Financial, you may <a href=""mailto:consent@rogersgroup.com?subject=Withdraw%20Consent&amp;body=I%20withdraw%20my%20consent%20to%20receive%20electronic%20communications%20from%20Rogers%20Group%20Financial."">withdraw your consent</a> at any time.&nbsp;</p>
<p><span style=""font-size: xx-small;""><span style=""font-size: xx-small;""><span style=""font-size: 7pt;"">&nbsp;<br /><br />PLEASE NOTE that all incoming e-mails will be automatically scanned to eliminate unsolicited promotional e-mails. This could result in deletion of a legitimate e-mail before it is read by its intended recipient. Please tell us if you have concerns about this automatic filtering</span>.</span></span></p>
</body>
"
$mail.save()

$inspector = $mail.GetInspector
$inspector.Display()