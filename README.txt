Last Update: 20.07.2011 - Zeno R.R. Davatz

= Spreadsheet

http://spreadsheet.rubyforge.org

The Mailing List can be found here:

http://groups.google.com/group/rubyspreadsheet

The code can be found here:

http://spreadsheet.rubyforge.org

For a viewable directory of all recent changes, please see:

http://scm.ywesee.com/?p=spreadsheet/.git;a=summary

To get a graphical overview of the Library please see

http://spreadsheet.rubyforge.org/spreadsheet.jpeg

For Non-GPLv3 commercial licencing, please see:

http://www.spreadsheet.ch


== Description

The Spreadsheet Library is designed to read and write Spreadsheet Documents.
As of version 0.6.0, only Microsoft Excel compatible spreadsheets are
supported. Spreadsheet is a combination/complete rewrite of the
Spreadsheet::Excel Library by Daniel J. Berger and the ParseExcel Library by
Hannes Wyss. Spreadsheet can read, write and modify Spreadsheet Documents.


== What's new?

* Supported outline (grouping) functions
* Significantly improved memory-efficiency when reading large Excel Files
* Limited Spreadsheet modification support
* Improved handling of String Encodings


== Roadmap

0.7.0:: Improved Format support/Styles
0.7.1:: Document Modification: Formats/Styles
0.8.0:: Formula Support
0.8.1:: Document Modification: Formulas
0.9.0:: Write-Support: BIFF5
1.0.0:: Ruby 1.9 Support;
        Remove backward compatibility code


== Dependencies

* ruby 1.8
* ruby-ole[http://code.google.com/p/ruby-ole/]


== Examples

Have a look at the GUIDE[link://files/GUIDE_txt.html].


== Installation

Using RubyGems[http://www.rubygems.org]:

* sudo gem install spreadsheet

If you don't like RubyGems[http://www.rubygems.org], let me know which
installation solution you prefer and I'll include it in the future.

If you can use 'rake' and 'hoe' library is also installed, you can 
build a gem package as follows:

* rake gem 

The gem package is built in pkg directory.


== Authors

Original Code:

Spreadsheet::Excel:
Copyright (c) 2005 by Daniel J. Berger (djberg96@gmail.com)

ParseExcel:
Copyright (c) 2003 by Hannes Wyss (hannes.wyss@gmail.com)

New Code:
Copyright (c) 2010 ywesee GmbH (mhatakeyama@ywesee.com, zdavatz@ywesee.com)


== License

This library is distributed under the GPLv3.
Please see the LICENSE[link://files/LICENSE_txt.html] file.

