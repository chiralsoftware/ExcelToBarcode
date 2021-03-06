# ExcelToBarcode
This simple Spring project converts an Excel spreadsheet into standard format labels, with text or barcode.
It uses Apache Poi to parse the Excel spreadsheet and uses iText to generate the PDF file suitable for
printing.  This is a compact and simple demonstration of the powerful capabilities
of these components, and how, with only a little bit of code, you can create a useful, real web application.
You can use this system live at:
https://exceltobarcode.com/

Please try it out. You can also download and install into any Java web server, such as Tomcat or Jetty.

License: GNU Public License Version 3.

Contact Chiral Software at https://chiralsoftware.com/ for licensing or consulting.

# Building and installing

Build using mvn package.  This results in a WAR file.  Copy the WAR file to the deployment
directory and start using it.  There's no configuration needed.  We have tested this
on Apache Tomcat 8.0.x, but it should work on any modern Java web server.

# Installing with an Nginx reverse proxy

Run it on Tomcat as usual. Set up the site with the nginx config example file.
