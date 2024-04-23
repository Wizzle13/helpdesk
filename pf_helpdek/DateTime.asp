	<SCRIPT LANGUAGE=Javascript>
<!--
/*
 Java Script Calendar is a small script for web pages which displays current Day of the week, Month, Date and Year along with Holidays notification. Just like this working demo below. It will work with Netscape 2++, Microsoft 3++ and Opera 3++ web browsers.

                                    Tuesday, November 24, 1998    

About 70 holidays from all over the world are displayed by Calendar (if you know other ones, let me know, I will add them). It even calculates Easter Sunday date! To see how it works change the date on your computer to January, 1,
for example and re-load the page.

Installation is very simple. Just copy the script below, everything between SCRIPT and SCRIPT tags (including tags themselves) and paste in the place you want it to appear on your page. You can customize the way
 it looks by changing document.write function arguments just after month names definitions. Java Script Calendar is free without any limitations. Copyright © Eugene Vassiltsov. But please, let me know if you will use it.
*/
 calendar = new Date();
 day = calendar.getDay();
 month = calendar.getMonth();
 date = calendar.getDate();
 year = calendar.getYear();
 if (year < 1000)
 year+=1900
 cent = parseInt(year/100);
 g = year % 19;
 k = parseInt((cent - 17)/25);
 i = (cent - parseInt(cent/4) - parseInt((cent - k)/3) + 19*g + 15) % 30;
 i = i - parseInt(i/28)*(1 - parseInt(i/28)*parseInt(29/(i+1))*parseInt((21-g)/11));
 j = (year + parseInt(year/4) + i + 2 - cent + parseInt(cent/4)) % 7;
 l = i - j;
 emonth = 3 + parseInt((l + 40)/44);
 edate = l + 28 - 31*parseInt((emonth/4));
 emonth--;
 var dayname = new Array ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
 var monthname = 
 new Array ("January","February","March","April","May","June","July","August","September","October","November","December" );

 document.write(dayname[day] + ", ");
 document.write(monthname[month] + " ");
 if (date< 10) document.write("0" + date + ", ");
         else document.write(date + ", ");

 document.write(year);
 
 document.write("<br>");
//-->
 </SCRIPT>