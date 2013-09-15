Excel Geocoding Tool
=====================

Easy to use Geocoding Tool for Excel. Download, enable macros, and add your own location data. Click Geocode All and you're done.

Requirements
------------
* Windows XP/Vista/7 (32bit/64bit) OR Mac OS X 10.5.8 or later (Intel)
* Excel 2003/2007/2010 OR Excel 2011 for Mac

Installation
------------
Simply download (https://github.com/maxrice/excel-geocoding-tool/blob/master/excel-geocoding-tool.xls) and run the Excel file. Make sure to enable macros and enter a proxy address if necessary.

Getting Started
---------------
See the excel file for basic instructions.


---------------
###Changelog

###3.6 - 2013-09-15
* Feature - Mac compatibility returns! Use Excel for Mac 2011 or greater
* Tweak - Refactor for easier maintainability
* Tweak - Greatly improved error handling

###3.5.1 - 2013-05-02
* Fix - Fixed issue with error handling

###3.5 - 2013-04-21
* Fix - Use Bing for geocoding now that Yahoo's PlaceFinder API was discontinued

###3.4.2 - 2012-07-15
* Feature - Added debug mode
* Tweak - Removed string cache, as it was causing a fatal error in some Excel versions
* Tweak - Refactored some code in preparation for v3.5 release
* Fix - fixed url encoding bug that affected accuracy of locations

###3.4.1 - 2012-05-17
* Feature - Proxy support on Mac
* Tweak - Code readability and variable declaration
* Fix - fixed curl url encoding bug on mac
* Misc - Added MIT License notice

###3.4 - 2012-05-12
* Feature - Now works on Mac! (proxy support on mac coming in next version)
* Tweak - Simpler proxy setup
* Tweak - New instructions
* Fix - Removed Create KML functionality

###3.3 - 2012-03-28
* Feature - Added macro to clear all data entry fields
* Feature - Added Geocode Not Found macro to only retry not found locations
* Feature - Added Google Maps link generation
* Feature - Added Proxy traversal
* Feature - Ability to geocode place names (ex: "The White House") or ZIP codes via free-form location format
* Feature - Ability to geocode international locations
* Tweak - Modified Geocode Selected Row macro to clear lat data, enabling it to run again
* Tweak - Modified Geocode All macro to clear entered data
* Tweak - Removed Google Earth auto-start on export

###3.2 - 2012-03-27
* Tweak - Removed juice analytics logo and misc. extraneous code
* Tweak - Removed beep on geocode
* Tweak - Removed geocoder.us
* Fix - Changed Yahoo API to Placefinder API
* Fix - Removed mClipboard module to make compatible with 64bit systems
* Fork - Initial fork (http://www.juiceanalytics.com/writing/excel-geocoding-tool-v2/)

----------

##Want to contribute?

1) Fork this repository
2) Make your changes to the worksheets / modules
3) Export any modules changed / added
4) Commit and send a pull request

__Contributors: maxrice,juiceinc,switchman2210__