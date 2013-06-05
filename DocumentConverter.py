#
# PyODConverter (Python OpenDocument Converter) v1.5 - 2013-01-07
#
# This script converts a document from one office format to another by
# connecting to an OpenOffice.org instance via Python-UNO bridge.
#
# Copyright (C) 2008-2013 Mirko Nasato
# Licensed under the GNU LGPL v2.1 - http://www.gnu.org/licenses/lgpl-2.1.html
# - or any later version.
#
DEFAULT_OPENOFFICE_PORT = 2002

import uno

from os.path import abspath, isfile, splitext
from com.sun.star.awt import Size
from com.sun.star.beans import PropertyValue
from com.sun.star.view.PaperFormat import USER
from com.sun.star.view.PaperOrientation import PORTRAIT, LANDSCAPE
from com.sun.star.task import ErrorCodeIOException
from com.sun.star.connection import NoConnectException

FAMILY_TEXT = "Text"
FAMILY_WEB = "Web"
FAMILY_SPREADSHEET = "Spreadsheet"
FAMILY_PRESENTATION = "Presentation"
FAMILY_DRAWING = "Drawing"

#---------------------#
# Configuration Start #
#---------------------#

'''
See http://www.openoffice.org/api/docs/common/ref/com/sun/star/view/PaperFormat.html
'''
PAPER_SIZE_MAP = {
    "A5": Size(14800,21000),
    "A4": Size(21000,29700),
    "A3": Size(29700,42000),
    "LETTER": Size(21590,27940),
    "LEGAL": Size(21590,35560),
    "TABLOID": Size(27900,43200)
}

'''
See http://www.openoffice.org/api/docs/common/ref/com/sun/star/view/PaperOrientation.html
'''
PAPER_ORIENTATION_MAP = {
    "PORTRAIT": PORTRAIT,
    "LANDSCAPE": LANDSCAPE
}

'''
See http://wiki.services.openoffice.org/wiki/Framework/Article/Filter
most formats are auto-detected; only those requiring options are defined here
'''
IMPORT_FILTER_MAP = {
    "txt": {
        "FilterName": "Text (encoded)",
        "FilterOptions": "utf8"
    },
    "csv": {
        "FilterName": "Text - txt - csv (StarCalc)",
        "FilterOptions": "44,34,0"
    }
}

'''
The filter options to export PDF files can be viewed on URL below
http://wiki.openoffice.org/wiki/API/Tutorials/PDF_export#General_properties
'''
EXPORT_FILTER_MAP = {
    "pdf": {
        FAMILY_TEXT: {
            "FilterName": "writer_pdf_Export",
            "FilterData": {
                "IsSkipEmptyPages": True
            },
            "Overwrite": True
        },
        FAMILY_WEB: {
            "FilterName": "writer_web_pdf_Export",
            "FilterData": {
                "IsSkipEmptyPages": True
            },
            "Overwrite": True
        },
        FAMILY_SPREADSHEET: {
            "FilterName": "calc_pdf_Export",
            "FilterData": {
                "IsSkipEmptyPages": True
            },
            "Overwrite": True
        },
        FAMILY_PRESENTATION: {
            "FilterName": "impress_pdf_Export",
            "FilterData": {
                "IsSkipEmptyPages": True
            },
            "Overwrite": True
        },
        FAMILY_DRAWING: {
            "FilterName": "draw_pdf_Export",
            "FilterData": {
                "IsSkipEmptyPages": True
            },
            "Overwrite": True
        }
    },
    "html": {
        FAMILY_TEXT: {
            "FilterName": "HTML (StarWriter)",
            "Overwrite": True
        },
        FAMILY_SPREADSHEET: {
            "FilterName": "HTML (StarCalc)",
            "Overwrite": True
        },
        FAMILY_PRESENTATION: {
            "FilterName": "impress_html_Export",
            "Overwrite": True
        }
    },
    "odt": {
        FAMILY_TEXT: {
            "FilterName": "writer8",
            "Overwrite": True
        },
        FAMILY_WEB: {
            "FilterName": "writerweb8_writer",
            "Overwrite": True
        }
    },
    "doc": {
        FAMILY_TEXT: {
            "FilterName": "MS Word 97",
            "Overwrite": True
        }
    },
    "docx": {
        FAMILY_TEXT: {
            "FilterName": "MS Word 2007 XML",
            "Overwrite": True
        }
    },
    "rtf": {
        FAMILY_TEXT: {
            "FilterName": "Rich Text Format",
            "Overwrite": True
        }
    },
    "txt": {
        FAMILY_TEXT: {
            "FilterName": "Text",
            "FilterOptions": "utf8",
            "Overwrite": True
        }
    },
    "ods": {
        FAMILY_SPREADSHEET: {
            "FilterName": "calc8",
            "Overwrite": True
        }
    },
    "xls": {
        FAMILY_SPREADSHEET: {
            "FilterName": "MS Excel 97",
            "Overwrite": True
        }
    },
    "csv": {
        FAMILY_SPREADSHEET: {
            "FilterName": "Text - txt - csv (StarCalc)",
            "FilterOptions": "44,34,0",
            "Overwrite": True
        }
    },
    "odp": {
        FAMILY_PRESENTATION: {
            "FilterName": "impress8",
            "Overwrite": True
        }
    },
    "ppt": {
        FAMILY_PRESENTATION: {
            "FilterName": "MS PowerPoint 97",
            "Overwrite": True
        }
    },
    "pptx": {
        FAMILY_PRESENTATION: {
            "FilterName": "Impress MS PowerPoint 2007 XML",
            "Overwrite": True
        }
    },
    "swf": {
        FAMILY_DRAWING: {
            "FilterName": "draw_flash_Export",
            "Overwrite": True
        },
        FAMILY_PRESENTATION: {
            "FilterName": "impress_flash_Export",
            "Overwrite": True
        }
    },
    "png": {
        FAMILY_PRESENTATION: {
            "FilterName": "impress_png_Export",
            "Overwrite": True
        },
        FAMILY_DRAWING: {
            "FilterName": "draw_png_Export",
            "Overwrite": True
        }
    },
    "gif": {
        FAMILY_PRESENTATION: {
            "FilterName": "impress_gif_Export",
            "Overwrite": True
        },
        FAMILY_DRAWING: {
            "FilterName": "draw_gif_Export",
            "Overwrite": True
        }
    },
    "jpg": {
        FAMILY_PRESENTATION: {
            "FilterName": "impress_jpg_Export",
            "Overwrite": True
        },
        FAMILY_DRAWING: {
            "FilterName": "draw_jpg_Export",
            "Overwrite": True
        }
    }
}

PAGE_STYLE_OVERRIDE_PROPERTIES = {
    FAMILY_SPREADSHEET: {
        #--- Scale options: uncomment 1 of the 3 ---
        # a) 'Reduce / enlarge printout': 'Scaling factor'
        "PageScale": 100,
        # b) 'Fit print range(s) to width / height': 'Width in pages' and 'Height in pages'
        #"ScaleToPagesX": 1, "ScaleToPagesY": 1000,
        # c) 'Fit print range(s) on number of pages': 'Fit print range(s) on number of pages'
        #"ScaleToPages": 1,
        "PrintGrid": False
    }
}

IMAGES_MEDIA_TYPE = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif"
}

#-------------------#
# Configuration End #
#-------------------#

class DocumentConversionException(Exception):

    def _get_message(self): 
        return self._message
    
    def _set_message(self, message): 
        self._message = message
    
    message = property(_get_message, _set_message)


class DocumentConverter:
    
    def __init__(self, port=DEFAULT_OPENOFFICE_PORT):
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        try:
            self.context = resolver.resolve("uno:socket,host=localhost,port=%s;urp;StarOffice.ComponentContext" % port)
        except NoConnectException:
            raise DocumentConversionException, "failed to connect to OpenOffice.org on port %s" % port
        self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)

    def convert(self, inputFile, outputFile, paperSize, paperOrientation):
        
        if PAPER_SIZE_MAP.has_key(paperSize) is False:
            raise Exception("The paper size given doesn't exist.")
        else:
            paperSize = PAPER_SIZE_MAP[paperSize]
        
        if PAPER_ORIENTATION_MAP.has_key(paperOrientation) is False:
            raise Exception("The paper orientation given doesn't exist.")
        else:
            paperOrientation = PAPER_ORIENTATION_MAP[paperOrientation]

        inputUrl = self._toFileUrl(inputFile)
        outputUrl = self._toFileUrl(outputFile)

        loadProperties = { "Hidden": True }
        
        inputExt = self._getFileExt(inputFile)
        outputExt = self._getFileExt(outputFile);
        
        if IMPORT_FILTER_MAP.has_key(inputExt):
            loadProperties.update(IMPORT_FILTER_MAP[inputExt])
        
        document = self.desktop.loadComponentFromURL(inputUrl, "_blank", 0, self._toProperties(loadProperties))
        try:
            document.refresh()
        except AttributeError:
            pass

        family = self._detectFamily(document)
        
        try:
            '''
            If you wish convert a document to an image, so each page needs be converted to a individual image.
            '''
            if IMAGES_MEDIA_TYPE.has_key(outputExt):
                
                drawPages = document.getDrawPages()
                pagesTotal = drawPages.getCount()
                mediaType = IMAGES_MEDIA_TYPE[outputExt]
                fileBasename = self._getFileBasename(outputUrl)
                graphicExport = self.context.ServiceManager.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", self.context)
                
                for pageIndex in xrange(pagesTotal):
                    
                    page = drawPages.getByIndex(pageIndex)
                    fileName = "%s-%d.%s" % (self._getFileBasename(outputUrl), pageIndex, outputExt)
                    
                    graphicExport.setSourceDocument( page )
                    
                    props = {
                        "MediaType": mediaType,
                        "URL": fileName
                    }
                    
                    graphicExport.filter( self._toProperties( props ) )
            else:
                
                self._overridePageStyleProperties(document, family)
            
                storeProperties = self._getStoreProperties(document, outputExt)
                
                printConfigs = {
                    'AllSheets': True,
                    'Size': paperSize,
                    'PaperFormat': USER,
                    'PaperOrientation': paperOrientation
                }
                
                document.setPrinter( self._toProperties( printConfigs ) )
            
                document.storeToURL(outputUrl, self._toProperties(storeProperties))
        finally:
            document.close(True)

    def _overridePageStyleProperties(self, document, family):
        if PAGE_STYLE_OVERRIDE_PROPERTIES.has_key(family):
            styleFamilies = document.getStyleFamilies()
            if styleFamilies.hasByName('PageStyles'):
                properties = PAGE_STYLE_OVERRIDE_PROPERTIES[family]
                pageStyles = styleFamilies.getByName('PageStyles')
                for styleName in pageStyles.getElementNames():
                    pageStyle = pageStyles.getByName(styleName)
                    for name, value in properties.items():
                        pageStyle.setPropertyValue(name, value)
        
    def _getStoreProperties(self, document, outputExt):
        family = self._detectFamily(document)
        try:
            propertiesByFamily = EXPORT_FILTER_MAP[outputExt]
        except KeyError:
            raise DocumentConversionException, "unknown output format: '%s'" % outputExt
        try:
            return propertiesByFamily[family]
        except KeyError:
            raise DocumentConversionException, "unsupported conversion: from '%s' to '%s'" % (family, outputExt)
    
    def _detectFamily(self, document):
        if document.supportsService("com.sun.star.text.WebDocument"):
            return FAMILY_WEB
        if document.supportsService("com.sun.star.text.GenericTextDocument"):
            # must be TextDocument or GlobalDocument
            return FAMILY_TEXT
        if document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return FAMILY_SPREADSHEET
        if document.supportsService("com.sun.star.presentation.PresentationDocument"):
            return FAMILY_PRESENTATION
        if document.supportsService("com.sun.star.drawing.DrawingDocument"):
            return FAMILY_DRAWING
        raise DocumentConversionException, "unknown document family: %s" % document

    def _getFileExt(self, path):
        ext = splitext(path)[1]
        if ext is not None:
            return ext[1:].lower()
    
    def _getFileBasename(self, path):
        name = splitext(path)[0]
        if name is not None:
            return name    
    
    def _toFileUrl(self, path):
        return uno.systemPathToFileUrl(abspath(path))

    def _toProperties(self, options):
        props = []
        for key in options:
            if isinstance(options[key], dict):
                property = PropertyValue(key, 0, uno.Any("[]com.sun.star.beans.PropertyValue", (self._toProperties(options[key]))), 0)
            else:
                property = PropertyValue(key, 0, options[key], 0)
            props.append(property)
        return tuple(props)

    def _dump(self, obj):
        for attr in dir(obj):
            print "obj.%s = %s\n" % (attr, getattr(obj, attr))

if __name__ == "__main__":
    
    from sys import exit
    from optparse import OptionParser
    
    parser = OptionParser(usage="usage: python %prog [options] <input-file> <output-file>", version="%prog 1.5")
    parser.add_option("-s", "--paper-size", default="A4", action="store", type="string", dest="paper_size", help="defines the paper size to converter that can be A3, A4, A5.")
    parser.add_option("-o", "--paper-orientation", default="PORTRAIT", action="store", type="string", dest="paper_orientation", help="defines the paper orientation to converter that can be PORTRAIT or LANDSCAPE.")
    
    (options, args) = parser.parse_args()
    
    if len(args) != 2:
        parser.error("wrong number of arguments")
    
    if not isfile(args[0]):
        print "No such input file: %s" % args[0]
        exit(1)
        
    try:
        converter = DocumentConverter()    
        converter.convert(args[0], args[1], options.paper_size, options.paper_orientation)
    except DocumentConversionException, exception:
        print "ERROR! " + str(exception)
        exit(1)
    except ErrorCodeIOException, exception:
        print "ERROR! ErrorCodeIOException %d" % exception.ErrCode
        exit(1)

