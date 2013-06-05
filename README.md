## PyODConverter
=============

PyODConverter (for Python OpenDocument Converter) is a Python script that
automates office document conversions from the command line using
LibreOffice or OpenOffice.org.

The script was written as a simpler alternative to
[JODConverter](http://code.google.com/p/jodconverter/) for command line usage.

## Usage
-----

PyODConverter requires LibreOffice/OpenOffice.org to be running as a service
and listening on port (by default) 2002; this can be achieved e.g. by starting
it from the command line as

	$ soffice --headless --nofirststartwizard "--accept=socket,host=localhost,port=2002;urp;StarOffice.Service"

The script expects exactly 2 parameters: an input and an output file names.
The document formats are inferred from the file extensions.

Since it uses the Python/UNO bridge, the script requires the UNO modules to be
already present in your Python installation. Most of the time this means you
need to use the Python version installed with OpenOffice.org, e.g. on Windows

    > "C:\Program Files\OpenOffice.org 3.1\program\python" DocumentConverter.py test.odt test.pdf

or on Linux

    $ /opt/openoffice.org3.1/program/python DocumentConverter.py test.odt test.pdf

### Possible options are the following:

* **Page Orientation** - Set the page orientation `-o LANDSCAPE` or `--page-orientation LANDSCAPE` (defaults to PORTRAIT)
* **Page Size** - Set the page size `-s A3` or `--page-size A3` (defaults to A4)

If you want to write your own scripts in Python, PyODConverter can also act as
a Python module, exporting a DocumentConverter class with a very simple API.

## ChangeLog
---------
v1.6 - 2013-06-05
* Fix support to print all sheets
* Fix parameters to initialize SOffice service

v1.5 - 2013-01-07
* Adding method to be able to get file base name
* Improvement files export from Presentation to Images. Now for each
  slide, an image will be created.

v1.4 - 2013-01-03

* Improvement the toProperties method to be able add array Uno properties
* Adding the Overwrite and IsSkipEmptyPages options.
* Update the README.

v1.3 - 2013-01-02

* Adding new docx format support.
* Adding paper size and orientation capable.
* Updated the options parser.
* Updated contributors.

v1.2 - 2012-03-10

* Changed default port to 2002
* Moved to GitHub

v1.1 - 2009-11-14

* Fixed HTML import issues by adding FAMILY\_WEB
* Support for specifying input formats and options
* Support for passing filter options to output formats
* Added CSV and TXT as input and output formats
* Support for overriding Page Style properties, especially useful for specifying
  how spreadsheets should fit into pages when exporting to PDF

v1.0.0 - 2008-05-05

* Let OOo determine the input document type, rather than using the file
  extension. This means all OOo-supported input types should now be accepted
  without any additional configuration.

## Contributors ##

* __mirkonasato__ <mirko.nasato@gmail.com>
* __marcelaraujo__ <admin@marcelaraujo.me>
