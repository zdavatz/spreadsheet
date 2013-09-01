= Getting Started with Spreadsheet
This guide is meant to get you started using Spreadsheet. By the end of it,
you should be able to read and write Spreadsheets.


== Reading is easy!
First, make sure all that code is loaded:

  require 'spreadsheet'

Worksheets come in various Encodings. You need to tell Spreadsheet which
Encoding you want to deal with. The Default is UTF-8

  Spreadsheet.client_encoding = 'UTF-8'

Let's open a workbook:

  book = Spreadsheet.open '/path/to/an/excel-file.xls'

We can either access all the Worksheets in a Workbook...

  book.worksheets

...or access them by index or name (encoded in your client_encoding)

  sheet1 = book.worksheet 0
  sheet2 = book.worksheet 'Sheet1'

Now you can either iterate over all rows that contain some data. A call to
Worksheet.each without argument will omit empty rows at the beginning of the
Worksheet:

  sheet1.each do |row|
    # do something interesting with a row
  end

Or you can tell Worksheet how many rows should be omitted at the beginning.
The following starts at the 3rd row, regardless of whether or not it or the
preceding rows contain any data:

  sheet2.each 2 do |row|
    # do something interesting with a row
  end

Or you can access rows directly, by their index (0-based):

  row = sheet1.row(3)

To access the values stored in a Row, treat the Row like an Array.

  row[0]

-> this will return a String, a Float, an Integer, a Formula, a Link or a Date
or DateTime object - or nil if the cell is empty.

More information about the formatting of a cell can be found in the Format
with the equivalent index

  row.format 2


== Writing is easy
As before, make sure you have Spreadsheet required and the client_encoding
set. Then make a new Workbook:

  book = Spreadsheet::Workbook.new

Add a Worksheet and you're good to go:

  sheet1 = book.create_worksheet

This will create a Worksheet with the Name "Worksheet1". If you prefer another
name, you may do either of the following:

  sheet2 = book.create_worksheet :name => 'My Second Worksheet'
  sheet1.name = 'My First Worksheet'

Now, add data to the Worksheet, using either Worksheet#[]=,
Worksheet#update_row, or work directly on Row using any of the Array-Methods
that modify an Array in place:

  sheet1.row(0).concat %w{Name Country Acknowlegement}
  sheet1[1,0] = 'Japan'
  row = sheet1.row(1)
  row.push 'Creator of Ruby'
  row.unshift 'Yukihiro Matsumoto'
  sheet1.row(2).replace [ 'Daniel J. Berger', 'U.S.A.',
                          'Author of original code for Spreadsheet::Excel' ]
  sheet1.row(3).push 'Charles Lowe', 'Author of the ruby-ole Library'
  sheet1.row(3).insert 1, 'Unknown'
  sheet1.update_row 4, 'Hannes Wyss', 'Switzerland', 'Author'

Add some Formatting for flavour:

  sheet1.row(0).height = 18

  format = Spreadsheet::Format.new :color => :blue,
                                   :weight => :bold,
                                   :size => 18
  sheet1.row(0).default_format = format

  bold = Spreadsheet::Format.new :weight => :bold
  4.times do |x| sheet1.row(x + 1).set_format(0, bold) end

And finally, write the Excel File:

  book.write '/path/to/output/excel-file.xls'


== Modifying an existing Document

Spreadsheet has some limited support for modifying an existing Document. This
is done by copying verbatim those parts of an Excel-document which Spreadsheet
can't modify (yet), recalculating relevant offsets, and writing the data that
can be changed.
Here's what should work:
* Adding, changing and deleting cells.
* You should be able to fill in Data to be evaluated by predefined Formulas
Limitations:
* Spreadsheet can only write BIFF8 (Excel97 and higher). The results of
  modifying an earlier version of Excel are undefined.
* Spreadsheet does not modify Formatting at present. That means in particular
  that if you set the Value of a Cell to a Date, it can only be read as a
  Date if its Format was set correctly prior to the change.
* Although it is theoretically possible, it is not recommended to write the
  resulting Document back to the same File/IO that it was read from.

And here's how it works:

  book = Spreadsheet.open '/path/to/an/excel-file.xls'
  sheet = book.worksheet 0
  sheet.each do |row|
    row[0] *= 2
  end
  book.write '/path/to/output/excel-file.xls'


== Date and DateTime
Excel does not know a separate Datatype for Dates. Instead it encodes Dates
into standard floating-point numbers and recognizes a Date-Cell by its
formatting-string:

  row.format(3).number_format

Whenever a Cell's Format describes a Date or Time, Spreadsheet will give you
the decoded Date or DateTime value. Should you need to access the underlying
Float, you may do the following:

  row.at(3)

If for some reason the Date-recognition fails, you may force Date-decoding:

  row.date(3)
  row.datetime(3)

When you set the value of a Cell to a Date, Time or DateTime, Spreadsheet will
try to set the cell's number-format to a corresponding value (one of Excel's
builtin formats). If you have already defined a Date- or DateTime-format,
Spreadsheet will use that instead. If a format has already been applied to
a particular Cell, Spreadsheet will leave it untouched:

  row[4] = Date.new 1975, 8, 21
  # -> assigns the builtin Date-Format: 'M/D/YY'
  book.add_format Format.new(:number_format => 'DD.MM.YYYY hh:mm:ss')
  row[5] = DateTime.new 2008, 10, 12, 11, 59
  # -> assigns the added DateTime-Format: 'DD.MM.YYYY hh:mm:ss'
  row.set_format 6, Format.new(:number_format => 'D-MMM-YYYY')
  row[6] = Time.new 2008, 10, 12
  # -> the Format of cell 6 is left unchanged.

== Outline (Grouping) and Hiding
Spreadsheet supports outline (grouping) and hiding functions from version 
0.6.5. In order to hide rows or columns, you can use 'hidden' property. 
As for outline, 'outline_level' property is also available. You can use 
both 'hidden' and 'outline_level' at the same time.

You can create a new file with outline and hiding rows and columns as
follows:

  require 'spreadsheet'
  
  # create a new book and sheet
  book = Spreadsheet::Workbook.new
  sheet = book.create_worksheet
  5.times {|j| 5.times {|i| sheet[j,i] = (i+1)*10**j}}
  
  # column
  sheet.column(2).hidden = true
  sheet.column(3).hidden = true
  sheet.column(2).outline_level = 1
  sheet.column(3).outline_level = 1
  
  # row
  sheet.row(2).hidden = true
  sheet.row(3).hidden = true
  sheet.row(2).outline_level = 1
  sheet.row(3).outline_level = 1
  
  # save file
  book.write 'out.xls'

Also you can read an existing file and change the hidden and outline
properties. Here is the example below:

  require 'spreadsheet'
  
  # read an existing file
  file = ARGV[0]
  book = Spreadsheet.open(file, 'rb')
  sheet= book.worksheet(0)
  
  # column
  sheet.column(2).hidden = true
  sheet.column(3).hidden = true
  sheet.column(2).outline_level = 1
  sheet.column(3).outline_level = 1
  
  # row
  sheet.row(2).hidden = true
  sheet.row(3).hidden = true
  sheet.row(2).outline_level = 1
  sheet.row(3).outline_level = 1
  
  # save file
  book.write "out.xls"

Notes
  * The outline_level should be under 8, which is due to the Excel data format.

== More about Encodings
Spreadsheet assumes it's running on Ruby 1.8 with Iconv-support. It is your
responsibility to handle Conversion Errors, or to prevent them e.g. by using
the Iconv Transliteration and Ignore flags:
Spreadsheet.client_encoding = 'LATIN1//TRANSLIT//IGNORE'


== Backward Compatibility
Spreadsheet is designed to be a drop-in replacement for both ParseExcel and
Spreadsheet::Excel. It provides a number of require-paths for backward
compatibility with its predecessors. If you have been working with ParseExcel,
you have probably used one or more of the following:

  require 'parseexcel'
  require 'parseexcel/parseexcel'
  require 'parseexcel/parser'

Either of the above will define the ParseExcel.parse method as a facade to
Spreadsheet.open. Additionally, this will alter Spreadsheets behavior to define
the ParseExcel::Worksheet::Cell class and fill each parsed Row with instances
thereof, which in turn provide ParseExcel's Cell#to_s(encoding) and Cell#date
methods.
You will have to manually uninstall the parseexcel library.

If you are upgrading from Spreadsheet::Excel, you were probably using
Workbook#add_worksheet and Worksheet#write, write_row or write_column.
Use the following to load the code which provides them:

  require 'spreadsheet/excel'

Again, you will have to manually uninstall the spreadsheet-excel library.

If you perform fancy formatting, you may run into trouble as the
Format implementation has changed considerably. If that is the case, please
drop me a line at hannes.wyss@gmail.com and I will try to help you. Don't
forget to include the offending code-snippet!

All compatibility code is deprecated and will be removed in version 1.0.0
