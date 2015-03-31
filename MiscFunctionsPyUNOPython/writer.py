# encoding=UTF-8
__author__ = 'mlechner'
# LibreOffice PyUNO Writer examples.
#
# Based on code from:
#   PyODConverter (Python OpenDocument Converter) v1.0.0 - 2008-05-05
#   Copyright (C) 2014 Marco Lechner <marco@marcolechner.de>
#   Licensed under the GNU LGPL v2.1 - or any later version.
#   http://www.gnu.org/licenses/lgpl-2.1.html
#

import os
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Size
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK, LINE_BREAK
from com.sun.star.awt.FontWeight import BOLD as FW_BOLD

from runner import LibreOfficeRunner


class WriterExamples(object):
    def run(self):
        oor = LibreOfficeRunner()
        self.desktop, self.graphicprovider = oor.connect()
        url = unohelper.systemPathToFileUrl('/tmp/test.odt')
        self.dpi = 150
        self.edit_doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
        self.cursor = self.edit_doc.Text.createTextCursor()
        self.cursor.gotoEnd(False)
        self.cl_ID = 0
        # do what ever you like from the following functions
        self.do_heading('First Heading')
        self.do_paragraph('Lorem ipsum', 14)
        # ... go on ...
        oor.shutdown()

    # search and replace text
    def do_search_and_replace(self, searchtext=None, replacetext=None):
        doc = self.edit_doc
        if searchtext is None or replacetext is None:
            return
        search = doc.createSearchDescriptor()
        search.SearchString = searchtext
        found = doc.findFirst(search)
        while found:
            found.String = found.String.replace(searchtext, replacetext)
            found = doc.findNext(found.End, search)

    def add_paragraph_break(self):
        doc = self.edit_doc
        doc.Text.insertControlCharacter(self.cursor, PARAGRAPH_BREAK, 0)

    def add_line_break(self):
        doc = self.edit_doc
        doc.Text.insertControlCharacter(self.cursor, LINE_BREAK, 0)

    # add a pagebreak
    # if you have a pagetemplate changing the page to landscape u can use it here
    def do_page_break(self, vorlage='Standard'):
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        doc.Text.insertControlCharacter(self.cursor, 4, 0)
        cursor.BreakType = 4
        if vorlage:
            try:
                cursor.setPropertyValue("PageDescName", vorlage)
            except:
                print("Setting PageDescName failed. May be Vorlage " + str(vorlage) + " does not exist in template.")

    # Heading
    def do_heading(self, titel='Titel', parastyle='Heading 1', numberinglevel=1):
        self.cl_ID += 1
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.ParaStyleName = parastyle
        cursor.NumberingLevel = numberinglevel
        cursor.ListId = self.cl_ID
        doc.Text.insertString(cursor, titel, 0)
        self.add_paragraph_break()

    def do_paragraph(self, text, cweight):
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.CharWeight = cweight
        doc.Text.insertString(cursor, text, 0)
        self.add_paragraph_break()

    def do_text(self, text, cweight):
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.CharWeight = cweight
        doc.Text.insertString(cursor, text, 0)

    # format a tables column width
    # input: table and list of positions of columnseparators in percent(float)
    def do_table_colwidth(self, table=None, collist=None):
        if table and collist:
            try:
                if table.Columns.Count == (len(collist)+1):
                    tablecolseps = table.TableColumnSeparators
                    pos = 0
                    for i in collist:
                        tablecolseps[pos].Position = int(i*100)
                        pos += 1
                    table.TableColumnSeparators = tablecolseps
            except Exception as e:
                print(e)

    def add_linked_image(self, url=None, width=None, height=None):
        dpi = self.dpi
        scale = 1000 * 2.54 / float(dpi)
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        try:
            fileurl = unohelper.systemPathToFileUrl(url)
            thisgraphicobject = doc.createInstance("com.sun.star.text.TextGraphicObject")
            thisgraphicobject.Surround = 'NONE'
            thisgraphicobject.GraphicURL = fileurl
            if thisgraphicobject.Graphic.SizePixel is None:
                # Then we're likely dealing with vector graphics. Then we try to
                # get the "real" size, which is enough information to
                # determine the aspect ratio
                original_size = thisgraphicobject.Graphic.Size100thMM
            else:
                original_size = thisgraphicobject.Graphic.SizePixel
            if width and height:
                size = Size(int(width * scale), int(height * scale))
            elif width:
                size = Size(int(width * scale), int((float(width)/original_size.Width) * original_size.Height * scale))
            elif height:
                size = Size(int((float(height)/original_size.Height) * original_size.Width * scale), int(height * scale))
            else:
                size = Size(int(original_size.Width * scale), original_size.Height * scale)
            thisgraphicobject.setSize(size)
            doc.Text.insertTextContent(cursor, thisgraphicobject, False)
        except Exception as e:
            print(e)

    def add_embedded_image(self, url=None, width=None, height=None, paraadjust=None):
        dpi = self.dpi
        scale = 1000 * 2.54 / float(dpi)
        doc = self.edit_doc
        cursor = doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        try:
            fileurl = unohelper.systemPathToFileUrl(url)
            graphic = self.graphicprovider.queryGraphic((PropertyValue('URL', 0, fileurl, 0), ))
            if graphic.SizePixel is None:
                # Then we're likely dealing with vector graphics. Then we try to
                # get the "real" size, which is enough information to
                # determine the aspect ratio
                original_size = graphic.Size100thMM
            else:
                original_size = graphic.SizePixel
            graphic_object_shape = doc.createInstance('com.sun.star.drawing.GraphicObjectShape')
            graphic_object_shape.Graphic = graphic
            if width and height:
                size = Size(int(width * scale), int(height * scale))
            elif width:
                size = Size(int(width * scale), int((float(width)/original_size.Width) * original_size.Height * scale))
            elif height:
                size = Size(int((float(height)/original_size.Height) * original_size.Width * scale), int(height * scale))
            else:
                size = Size(int(original_size.Width * scale), original_size.Height * scale)
            graphic_object_shape.setSize(size)
            # doc.Text.insertTextContent(cursor, graphic_object_shape, False)
            thisgraphicobject = doc.createInstance("com.sun.star.text.TextGraphicObject")
            thisgraphicobject.Graphic = graphic_object_shape.Graphic
            thisgraphicobject.setSize(size)
            if paraadjust:
                oldparaadjust = cursor.ParaAdjust
                cursor.ParaAdjust = paraadjust
            doc.Text.insertTextContent(cursor, thisgraphicobject, False)
            os.unlink(url)
            if paraadjust:
                cursor.ParaAdjust = oldparaadjust
        except Exception as e:
            print(e)

    #Table
    def do_table(self, cols=1, rows=1, data=None, **kwargs):
        doc = self.edit_doc
        data = data if data else {'A1': 'C1', 'A2': 'C2'}
        ncols = cols
        nrows = rows
        table = doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(nrows, ncols)
        doc.Text.insertTextContent(doc.Text.getEnd(), table, True)
        # change col width before merging any cells
        # set col1 = 50%,, col2 = 25%, col3 = 25% of table width
        collist = [50, 75]
        self.do_format_table(table=table, collist=collist)

        #Merge Cells#################################
        a = table.createCursorByCellName("A1")
        a.gotoCellByName("B1", False)
        a.goRight(1, True)
        a.mergeRange()
        a.gotoCellByName("B2", False)
        a.goDown(1, True)
        a.mergeRange()

        #Insert fixed Strings###############################
        table.getCellByName("A1").setString("A1")
        table.createCursorByCellName("A1").ParaAdjust = CENTER
        table.createCursorByCellName("A1").CharWeight = FW_BOLD
        table.getCellByName("B1").setString("merged B1 C1")
        table.getCellByName("B2").setString("merged B2 B3")
        for k, v in data:
            table.getCellByName(k).setString(k)
            table.getCellByName(v).setString(v)


if __name__ == '__main__':
    writerexamples = WriterExamples()
    writerexamples.run()
