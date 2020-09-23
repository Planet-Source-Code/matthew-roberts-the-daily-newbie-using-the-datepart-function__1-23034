<div align="center">

## The Daily Newbie \- Using the DatePart Function


</div>

### Description

Explains how to use the DatePart Function in Visual Basic.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-the-daily-newbie-using-the-datepart-function__1-23034/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type"
content="text/html; charset=iso-8859-1">
<title>Daily Newbie - 05/01/2001</title>
</head>
<body bgcolor="#FFFFFF">
<p> </p>
<p class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoTitle"><font size="7"><strong>The
Daily Newbie</strong></font></p>
<p align="center" class="MsoTitle"><strong>&#8220;To Start Things
Off Right&#8221;</strong></p>
<p align="center" class="MsoTitle"><font size="1">
May 8,
2001
</font></p>
<p align="center" class="MsoTitle"><img width="100%" height="3"
v:shapes="_x0000_s1027"></p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p align="center" class="MsoNormal" style="text-align:center"> </p>
<p class="MsoNormal"><font face="Arial"></font></p>
<p class="MsoNormal"><font size="2" face="Arial"></font></p>
<p class="MsoNormal"><font size="2" face="Arial"></font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Today&#8217;s Keyword:</strong>
        </font><font
size="4" face="Arial"> DatePart()</font></p>
<p class="MsoNormal"
style="margin-left:135.0pt;text-indent:-135.0pt"><font size="2"
face="Arial"><strong>Name Derived
From:  </strong>   </font>
 <font size="2" face="Arial">"Part of a Date"</a></i> </em></font></p>
 </p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Used for: </strong>
Getting a part of a date value (i.e. Day, Month, Year, etc.).</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>VB Help Description: </strong>  Returns a Variant (Integer) containing the specified part of a given date.
</font></p>
<font size="2" face="Arial"><strong>Plain
English: </strong>Lets you get only one part of a date/time value. For example you can determine what weekday a certain date falls on.<br><br>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Syntax:  </strong>       Val=DatePart(Part, Date)</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Usage:  </strong>        intWeekDay = DatePart("w","01/12/2000")</font></p>
<p class="MsoNormal"
style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"><font
size="2" face="Arial"><strong>Parameters:  </strong>
<br>
<font face = "arial" size="2">
<li><b>Part</b> - The part of the date you want returned . This can be:
	<blockquote>
		<blockquote>
	<li>s - Seconds
	<li>n - Minutes
	<li>h - Hours
	<li>d - Days
	<li>y - Day of Year
	<li>w - Weekday
	<li>w - Week
	<li>m - Months
	<li>q - Quarter
	<li>yyyy - Year
		</blockquote>
	</blockquote>
<li><b>Date</b> - The date that the part is derived from.
Example: <br>
<br>
To get the current week within the current year (What week is this for the year? 1-52 )):
<br><br>
<blockquote>
<code><font size="2">MsgBox DatePart("ww", Date)</font></code>
</blockquote>
</font>
</font></p>
If you have not read the Daily Newbie on how VB stores date format, you may want to review it now <a href="http://www.planetsourcecode.com/xq/ASP/txtCodeId.22876/lngWId.1/qx/vb/scripts/ShowCode.htm"> by clicking here.</a>
 <br><br>
<br>
Today's code snippet returns the Julian date for today.
</font></p>
<p class="MsoNormal"
style="margin-left:135.35pt;text-indent:-135.35pt"><font size="2"
face="Arial"><strong>Copy & Paste Code:</strong></font></p>
  <p class="MsoNormal"
  style="margin-left:135.35pt;text-indent:-135.35pt"><font
  size="2" face="Arial"></font></p>
    <pre>
<font size="2" face="Arial"><code></code></font></pre>
    <pre
    style="margin-left:1.25in;text-indent:.35pt;tab-stops:45.8pt 91.6pt 183.2pt 229.0pt 274.8pt 320.6pt 366.4pt 412.2pt 458.0pt 503.8pt 549.6pt 595.4pt 641.2pt 687.0pt 732.8pt"><font
size="3" face="Arial"><code>
<br><br>
MsgBox "Today's Julian Date is: " & DatePart("y",Date) & "/" & DatePart("yyyy",Date)
<br><br>
				</code></font></pre>
 <p class="MsoNormal"
 style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;
margin-left:135.0pt;text-indent:-135.0pt"> </p>
<br>
<br>
</body>
</html>

