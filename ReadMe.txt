Title: imMXRecord Class
Description: This class make you know the MXrecords for a given email address or domain. The based code is not mine. It is based on some code from the web, specially in modMXQuery.bas from vbSendMail by Dean Dusenbery.
It wont freeze if you give a non-existent domain or DNS server, an improvement from original code. 
Please see readme.txt file for references of based codes.

Based on:
modMXQuery.bas from vbSendMail by Dean Dusenbery
  http://www.freevbcode.com/ShowCode.Asp?ID=109

MX Lookup Control (UserControl) by Gregg Housh
  http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11306&lngWId=1

MX Query by Saurabh Gupta
  http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=34016&lngWId=1

Lookup MX records by Jason Martin
 http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11006&lngWId=1

and more and more code from Planet-Source-Code and FreeVBCode

thanks for all

Jorge Colaccini
Informas.com
software<AT>informas.com
Rosario, Argentina, August, 2005

This is the original description from vbSendMail documentation

MXQuery

Public Function MXQuery(optional IPDomain as string ) As String

Performs a name server (DNS) search for mail exchange (MX) records for the passed IP domain. If blank the IP Domain is assumed to be the home domain of  the local host.  If successful, a host name is returned that can be used to set the SMTPHost Property. If the lookup fails, a null string is returned.  If the local host is part of a sub domain and no domain is specified, the search will start at the sub domain level and continue up to the root domain until a MX host is found. If nothing is found at the root level or a DNS server is not available, a null string is returned. Searching foreign domains is permitted unless the local host is located behind a firewall. Searching through a firewall is not supported. 

Host names returned by the MXQuery function are the names provided by the local DNS server. The DNS server may or may not have MX records for the closest or 'best' SMTP host depending on the network admin's  policies. Most networks that operate DNS servers will have the local SMTP host identified but there is no requirement that they do so.  For instance, the @home network does not (at least in my location)  have any MX records for local SMTP hosts. An MXQuery in this situation will return the primary  SMTP host at the root domain level (home.com).   

The MXQuery Method may not return a valid SMTP host in all scenarios. If your app uses this method it would be prudent to provide an alternative method for the user to enter the SMTP host in the event the MXQuery Method is not able to resolve a usable host.  


