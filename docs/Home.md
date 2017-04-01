**Project Description**
Tool for downloading the Visual Studio 2012/2013/2015 help packages for offline use

# Overview

Allows Visual Studio 2012/2013/2015 packages to be downloaded to an offline cache location before importing them into Microsoft Help Viewer 2.0/2.1/2.2. 

If the cache is kept following the import, then only changes will be downloaded on subsequent occasions.
 
![](Home_screenshot.jpg)

# Quick Guide

* Select the version of Visual Studio to download the help for
* Select your language from the drop down list
* Press "Load Books" to retrieve the list of available books. Books that are already in the cache (partially or fully) will automatically be checked. Note that because packages are shared between different books, you may get extra items checked automatically.
* Check (or uncheck) book that you want to download. The "Download Size" and "Num Downloads" columns indicate an approximate amount of data that needs to be downloaded for the book based on what is already in the cache.
* Press "Download" to start downloading the requested books. Packages for books no longer selected will be deleted at this point.
* When the download is complete, import the new books into "Microsoft Help Viewer" using the "Manage Content" tab.

# Credits

This project is based on (and shares some code with) the earlier project  [Visual Studio Help Downloader](http://vshd.codeplex.com/) by [nop](http://www.codeplex.com/site/users/view/nop). 