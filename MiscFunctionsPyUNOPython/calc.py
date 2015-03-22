import unohelper

from runner import OORunner

class CalcExamples(object):
    def run(self):
        oor = OORunner()
        self.desktop, self.graphicprovider = oor.connect()
        url = unohelper.systemPathToFileUrl('/tmp/test.ods')
        self.document = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
        self.sheet = self.document.getSheets().getByIndex(0)
        # do what ever you like from the following functions
        self.do_fill((4, 3), 'data')
        self.do_formula(5)
        # ... go on ...
        oor.shutdown()

    # search and replace text
    def do_fill(self, size, data):
        cols, rows = size
        for cid in range(cols):
            for rid in range(rows):
                self.sheet.getCellByPosition(cid, rid).setString(data)

    def do_formula(self, row):
        self.sheet.getCellRangeByName('A%s' % (row,)).setValue(1.1)
        self.sheet.getCellRangeByName('B%s' % (row,)).setValue(2.2)
        self.sheet.getCellRangeByName('C%s' % (row,)).setFormula('=SUM(A%s:B%s)' % (row, row))


if __name__ == '__main__':
    calcexmples = CalcExamples()
    calcexmples.run()
