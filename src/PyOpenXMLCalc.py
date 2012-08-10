import re
import time
import sys
from xml.dom.minidom import *
from zipfile import *

__author__ = "peckto"
__version__ = "1.0"
__status__ = "productive"

class OP(object):
    types = dict()
    types['docProps/app.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
    types['docProps/core.xml'] = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
    types['xl/workbook.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    types['worksheets/sheetX.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    types['styles.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
    types['sharedStrings.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
    types['../tables/tableX.xml'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
    contentTypes = dict()
    contentTypes['/xl/theme/themeX.xml']= 'application/vnd.openxmlformats-officedocument.theme+xml'
    contentTypes['/xl/styles.xml']= 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
    contentTypes['rels']= 'application/vnd.openxmlformats-package.relationships+xml'
    contentTypes['xml']= 'application/xml'
    contentTypes['/xl/workbook.xml']= 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    contentTypes['/docProps/app.xml']= 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
    contentTypes['/xl/worksheets/sheetX.xml']= 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    contentTypes['/docProps/core.xml']= 'application/vnd.openxmlformats-package.core-properties+xml'
    contentTypes['/xl/sharedStrings.xml']= 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
    contentTypes['/xl/tables/tableX.xml'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'

    def __init__(self):
        self.rId = None

    def set_rId(self,id_):
        self.rId = id_

    def get_contentType(self,partName):
        """replace eventual number with 'X' first, then ask the dictionary"""
        number = self.get_number(partName)
        if number:
            partName = partName.replace(number,'X.')
        return self.contentTypes[partName]

    def get_Type(self,target):
        """replace eventual number with 'X' first, then ask the dictionary"""
        number = self.get_number(target)
        if number:
            target = target.replace(number,'X.')
        return self.types[target]

    def get_number(self,ref):
        out = re.search('\d+\.?\d*',ref)
        if out:
            return out.group(0)
        else:
            return None

    def get_text(self,ref):
        return ref.replace(self.get_number(ref),'')

    def toxml(self,encoding):
        return self.root.toxml(encoding=encoding)

    def toprettyxml(self,encoding):
        return self.root.toprettyxml(encoding=encoding)
    
    def getInt4CN(self,CN):
        """translate CN in integer"""
        l = list(CN)
        l.reverse()
        all = 0
        for i in range(len(l)):
            all+=ord(l[i])*(i+1)
        return all

    def compCN(self,CN1,CN2):
        """compare self.startCN to CN.

        CN > CN2 """
        i1 = self.getInt4CN(CN2)
        i2 = self.getInt4CN(CN1)
        if i2 > i1:
            return True
        else:
            return False

class Content_Types(OP):
    def __init__(self,f=None):
        if f:
            self._open(f)
        else:
            self.root = Document()
            self.types = types = self.root.createElement('Types')
            self.types.setAttribute('xmlns','http://schemas.openxmlformats.org/package/2006/content-types')
            self.root.appendChild(self.types)
            self.new_default('rels')
            self.new_default('xml')

    def _open(self,f):
        """parse a axisting Content_Types xml"""
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.types = self.root.getElementsByTagName('Types')[0]

    def getOverrides(self):
        overrides = list()
        for override in self.types.getElementsByTagName('Override'):
            type_ = override.getAttribute('ContentType')
            partName = override.getAttribute('PartName')
            overrides.append([type_,partName])
        return overrides

    def new_default(self,extensio):
        default = self.root.createElement('Default')
        default.setAttribute('Extension',extensio)
        default.setAttribute('ContentType',self.get_contentType(extensio))
        self.types.appendChild(default)

    def new_Override(self,partName):
        default = self.root.createElement('Override')
        default.setAttribute('PartName',partName)
        default.setAttribute('ContentType',self.get_contentType(partName))
        self.types.appendChild(default)

    def new_Sheet(self,id_):
        sheetPath = '/xl/worksheets/sheet%s.xml' %id_
        self.new_Override(sheetPath)

class Relationships(OP):
    def __init__(self,f=None):
        if f:
            self._open(f)
        else:
            self.id_ = 0
            self.root = Document()
            self.relationships = self.root.createElement('Relationships')
            self.relationships.setAttribute('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
            self.root.appendChild(self.relationships)

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.relationships = self.root.getElementsByTagName('Relationships')[0]
        
    def new_relationship(self,target):
        rId = self.get_NewID()
        relationship = self.root.createElement('Relationship')
        relationship.setAttribute('Id','rId%s'%rId)
        relationship.setAttribute('Type',self.get_Type(target))
        relationship.setAttribute('Target',target)
        self.relationships.appendChild(relationship)
        return rId

    def get_NewID(self):
        self.id_+= 1
        return self.id_

    def getTarget(self,rel):
        """return Target for rId"""
        for relation in self.relationships.getElementsByTagName('Relationship'):
            if relation.getAttribute('Id') == rel:
                return relation.getAttribute('Target')
        return -1

class App(OP):
    def __init__(self,company='',f=None):
        if f:
            self._open(f)
        else:
            self.countTables = 0
            self.root = Document()
            self.properties = self.root.createElement('Properties')
            self.properties.setAttribute('xmlns','http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
            self.properties.setAttribute('xmlns:vt','http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
            self.set_TotalTime(0)
            self.set_Application('Microsoft Excel')
            self.set_DocSecurity(0)
            self.set_ScaleCrop('false')
    #        self.headingPairs = self.root.createElement('HeadingPairs')
    #        self.properties.appendChild(self.headingPairs)
    #        vt_vector = self.root.createElement('vt:vector')
    #        vt_vector.setAttribute('size','2')
    #        vt_vector.setAttribute('baseType','variant')
    #        self.headingPairs.appendChild(vt_vector)
    #        vt_variant = self.root.createElement('vt:variant')
    #        vt_vector.appendChild(vt_variant)
    #        vt_lpstr = self.new_ElementWithText('vt:lpstr',sheetName)
    #        vt_variant.appendChild(vt_lpstr)
    #        vt_variant = self.root.createElement('vt:variant')
    #        vt_vector.appendChild(vt_variant)
    #        vt_i4 = self.new_ElementWithText('vt:i4',0)
    #        vt_variant.appendChild(vt_i4)
    #        vt_variant = self.root.createElement('vt:variant')
    #        vt_i4 = self.new_ElementWithText('vt:i4',self.countTables)
    #        vt_variant.appendChild(vt_i4)
            self.titlesOfParts = self.root.createElement('TitlesOfParts')
            self.properties.appendChild(self.titlesOfParts)
            vt_vector = self.root.createElement('vt:vector')
            vt_vector.setAttribute('size','0')
            vt_vector.setAttribute('baseType','lpstr')
            self.titlesOfParts.appendChild(vt_vector)
            self.company = self.new_ElementWithText('Company',company)
            self.properties.appendChild(self.company)
            self.linksUpToDate = self.new_ElementWithText('LinksUpToDate','false')
            self.properties.appendChild(self.linksUpToDate)
            self.sharedDoc = self.new_ElementWithText('SharedDoc','false')
            self.properties.appendChild(self.sharedDoc)
            self.hyperlinksChanged = self.new_ElementWithText('HyperlinksChanged','false')
            self.properties.appendChild(self.hyperlinksChanged)
            self.appVersion = self.new_ElementWithText('AppVersion','12.0000')
            self.properties.appendChild(self.appVersion)
            self.root.appendChild(self.properties)

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.properties = self.root.getElementsByTagName('Properties')[0]
        self.totalTime = self.properties.getElementsByTagName('TotalTime')[0]
        self.application = self.properties.getElementsByTagName('Application')[0]
        self.docSecurity = self.properties.getElementsByTagName('DocSecurity')[0]
        self.scaleCrop = self.properties.getElementsByTagName('ScaleCrop')[0]
        self.titlesOfParts = self.properties.getElementsByTagName('TitlesOfParts')[0]
        company = self.properties.getElementsByTagName('Company')
        if company:
            self.company = company[0]
        else:
            self.company = self.root.createElement('Company')
            self.properties.appendChild(self.company)
        self.linksUpToDate = self.properties.getElementsByTagName('LinksUpToDate')[0]
        self.sharedDoc = self.properties.getElementsByTagName('SharedDoc')[0]
        self.hyperlinksChanged = self.properties.getElementsByTagName('HyperlinksChanged')[0]
        self.appVersion = self.properties.getElementsByTagName('AppVersion')[0]

    def new_Table(self,tableName):
#        vt_variant = self.headingPairs.getElementsByTagName('vt:variant')[1]
#        vt_i4 = vt_variant.getElementsByTagName('vt:i4')[0]
        self.countTables+=1
#        newChild = self.new_ElementWithText('vt:i4',self.countTables)
#        vt_variant.replaceChild(newChild,vt_i4)
        vt_vector = self.titlesOfParts.getElementsByTagName('vt:vector')[0]
        vt_vector.setAttribute('size',str(self.countTables))
        vt_lpstr = self.new_ElementWithText('vt:lpstr',tableName)
        vt_vector.appendChild(vt_lpstr)

    def new_ElementWithText(self,name,text):
        newElement = self.root.createElement(name)
        if type(text) != str and type(text) != unicode:
            text = str(text)
        newElement_text = self.root.createTextNode(text)
        newElement.appendChild(newElement_text) 
        return newElement

    def set_TotalTime(self,time):
        self.totalTime = self.new_ElementWithText('TotalTime',time)
        self.properties.appendChild(self.totalTime)
       
    def set_Application(self,app):
        self.application = self.new_ElementWithText('Application',app)
        self.properties.appendChild(self.application)

    def set_DocSecurity(self,security):
        self.docSecurity = self.new_ElementWithText('DocSecurity',security)
        self.properties.appendChild(self.docSecurity)

    def set_ScaleCrop(self,scaleCrop):
        self.scaleCrop = self.new_ElementWithText('ScaleCrop',scaleCrop)
        self.properties.appendChild(self.scaleCrop)

    def getNextTableID(self):
        return self.countTables+1

class Sheet(OP):
    def __init__(self,f=None):
        if f:
            self._open(f)
        else:
            self.root = Document()
            self.worksheet = self.root.createElement('worksheet')
            self.worksheet.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            self.worksheet.setAttribute('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            self.root.appendChild(self.worksheet)
            self.dimension = self.root.createElement('dimension')
            self.dimension.setAttribute('ref','A1')
            self.worksheet.appendChild(self.dimension) 
            self.sheetViews = self.root.createElement('sheetViews')
            self.worksheet.appendChild(self.sheetViews)
            self.sheetViews.appendChild(self.newSheetView(1,0))
            self.sheetFormatPr = self.root.createElement('sheetFormatPr')
            self.sheetFormatPr.setAttribute('baseColWidth',str(10))
            self.sheetFormatPr.setAttribute('defaultRowHeight',str(15))
            self.worksheet.appendChild(self.sheetFormatPr)
            self.cols = None
            self.rows = dict()
            self.tableParts = None
            self.cursor = Ref('A1')
            self.sheetData = self.root.createElement('sheetData')
            self.worksheet.appendChild(self.sheetData)
    #        self.pageMargins = self.root.createElement('pageMargins')
    #        self.pageMargins.setAttribute('left','0.7')
    #        self.pageMargins.setAttribute('right','0.7')
    #        self.pageMargins.setAttribute('top','0.78740157499999996')
    #        self.pageMargins.setAttribute('bottom','0.78740157499999996')
    #        self.pageMargins.setAttribute('header','0.3')
    #        self.pageMargins.setAttribute('footer','0.3')
    #        self.worksheet.appendChild(self.pageMargins)
            self.writeEngine = 'inlineStr'
            self.dimensionRef = Sqref('A1')

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.worksheet = self.root.getElementsByTagName('worksheet')[0]
        self.dimension = self.worksheet.getElementsByTagName('dimension')[0]
        self.sheetViews = self.worksheet.getElementsByTagName('sheetViews')[0]
        self.sheetFormatPr = self.worksheet.getElementsByTagName('sheetFormatPr')[0]
        e = self.worksheet.getElementsByTagName('cols')
        if e:
            self.cols = e[0]
        self.rows = dict()
        self.tableParts = None
        self.sheetData = self.worksheet.getElementsByTagName('sheetData')[0]
        self.writeEngine = 'sharedStrings'
        self.dimensionRef = Sqref(*self.dimension.getAttribute('ref').split(':'))
        tableParts = self.worksheet.getElementsByTagName('tableParts')
        if tableParts:
            self.tableParts = tableParts[0]
        else:
            self.tableParts = self.root.createElement('tableParts')
            self.worksheet.appendChild(self.tableParts)
        for row in self.sheetData.getElementsByTagName('row'):
            self.rows[int(row.getAttribute('r'))] = row
        self.cursor = self.getSelectedCell()

    def newSheetView(self,tabSelected,workbookViewId):
        sheetView = self.root.createElement('sheetView')
#        sheetView.setAttribute('tabSelected',str(tabSelected))
        sheetView.setAttribute('workbookViewId',str(workbookViewId))
        return sheetView

    def selectedTab(self):
        sheetView = self.sheetViews.getElementsByTagName('sheetView')[0]
        sheetView.setAttribute('tabSelected',"1")

    def selectCell(self,ref):
        """Select a cell
        <selection activeCell="B2" sqref="B2"/>"""
        sheetView = self.sheetViews.getElementsByTagName('sheetView')[0]
        selection = sheetView.getElementsByTagName('selection')
        if not selection:
            selection = self.root.createElement('selection')
            sheetView.appendChild(selection)
        else:
            selection = selection[0]
        selection.setAttribute('activeCell',ref.ref)
        selection.setAttribute('sqref',ref.ref)

    def getSelectedCell(self):
        """get the selected Cell Ref in sheet"""
        sheetView = self.sheetViews.getElementsByTagName('sheetView')[0]
        selection = sheetView.getElementsByTagName('selection')
        if not selection:
            activeCell = 'A1'
        else:
            activeCell = selection[0].getAttribute('activeCell')
        return Ref(activeCell)
    
    def hideColume(self,min_,max_):
        """Hide a Colume col
        <worksheet>
        <cols>
            <col min="2" max="2" width="0" hidden="1" customWidth="1"/>
        </cols>
        </worksheet>
        """
        if not self.cols:
            self.cols = self.root.createElement('cols')
            self.worksheet.insertBefore(self.cols,self.sheetData)
        col = self.root.createElement('col')
        col.setAttribute('min',str(min_))
        col.setAttribute('max',str(max_))
        col.setAttribute('width','0')
        col.setAttribute('hidden','1')
        col.setAttribute('customWidth','1')
        self.cols.appendChild(col)

    def addTablePart(self,rID):
        """<tableParts count="1">
            <tablePart r:id="rId2"/>
            </tableParts>"""
        if not self.tableParts:
            self.tableParts = self.root.createElement('tableParts')
            self.worksheet.appendChild(self.tableParts)
#        self.tableParts.setAttribute('count',str(count))
        tablePart = self.root.createElement('tablePart')
        tablePart.setAttribute('r:id','rId%s' %rID)
        self.tableParts.appendChild(tablePart)
#        self.updateCount(self.tableParts)

    def get_row(self,ref):
        if ref.rowID in self.rows:
            return self.rows[ref.rowID]
        else:
            return self.new_row(ref)

    def new_row(self,ref):
        """create a new row and append it on the right position
           * append new Row in the right position
           * search row that comes after the new row
           * insertBefore that row"""
        row = self.root.createElement('row')
#        row.setAttribute('spans','1:1') # <- ignore
        row.setAttribute('r',str(ref.rowID))
        nextRowID = self.getNextRowID(ref)
        if nextRowID == -1:
            self.sheetData.appendChild(row)
        else:
            self.sheetData.insertBefore(row,self.rows[nextRowID])
        self.rows[ref.rowID] = row
        return row

    def getNextRowID(self,ref):
        """return the next higher existing row ID if there is no, return -1"""
        keys = self.rows.keys()
        keys.sort()
        for i in keys:
            if i >= ref.rowID:
                return i
        return -1
        
        
    def getC(self,ref,row):
        """get c tag from row by ref or create a new"""
        for c in row.getElementsByTagName('c'):
            if c.getAttribute('r') == ref.ref:
                return c
        c = self.root.createElement('c')
        c.setAttribute('r',ref.ref)
        row.appendChild(c)
        return c


    def writeLine(self,ref,line):
        """write a line(list) to Sheet at ref"""
        ref2 = Ref(ref.ref)
        for cell in line:
            self.write(ref2,cell)
            ref2.walk('right')

    def write(self,ref,text):
        self.dimensionRef.end.update(ref)
        if not text:
            text = ''
        if type(text) != int:
            try:
                text = int(text)
            except ValueError:
                pass
        if type(text) == int:
            self.writeInt(ref,text)
        else:
            if self.writeEngine == 'inlineStr':
                self.new_row_inlineStr(ref,text)
            else:
                self.new_row_sharedStr(ref,text)
            
    def writeInt(self,ref,i):
        """write integer in cell
        <row r="1" spans="1:2">
           <c r="A1">
             <v>1</v>
           </c>"""
        row = self.get_row(ref)
        c = self.getC(ref,row)
        if 't' in c.attributes.items():
            c.removeAttribute('t')
        vs = c.getElementsByTagName('v')
        if vs: 
            v = vs[0]
        else:
            v = self.root.createElement('v')
            v_text = self.root.createTextNode('')
            v.appendChild(v_text)
            c.appendChild(v)
        v.firstChild.replaceWholeText(str(i))

    def new_row_inlineStr(self,ref,text):
        """<row r="3" spans="1:1">
            <c r="A3" t="inlineStr">
            <is>
            <t>hjkffioj</t>
            </is>
            </c>
            </row>"""
        if not text:
            return
        if type(text) == int:
            text = str(text)
        row = self.get_row(ref)
        c = self.getC(ref,row)
        if c.getAttribute('r') == 's':
            c.removeChild('v')
        else:
            is_ = c.getElementsByTagName('is')
            if is_:
                c.removeChild(is_[0])
        if text.startswith('='):
            c.setAttribute('t','str')
            f = self.root.createElement('f')
            text = text.replace('=','',1)
            f_text = self.root.createTextNode(text)
            f.appendChild(f_text)
            c.appendChild(f)
        else:
            c.setAttribute('t','inlineStr')
            is_ = self.root.createElement('is')
            t = self.root.createElement('t')
            t_text = self.root.createTextNode(text)
            t.appendChild(t_text)
            is_.appendChild(t)
            c.appendChild(is_)
#        self.update_spans(row) # <- ignore

    def new_row_sharedStr(self,r,text):
        text = str(text)
#        global rowID
        if not text:
            return
    #    print str(rowID)
    #    print text
        """<si>
            <t>substituted</t>
            </si>"""
        if text.startswith('='):
            formel = True
        else:
            formel = False
        textID = self.OP['xl/sharedStrings.xml'].newString(text)
        row = self.get_row(self.activeSheet.sheetData,r.startRowID)
        root = self.activeSheet.root
        c = root.createElement('c')
        c.setAttribute('r',r.start)
        if formel:
            c.setAttribute('t','str')
            f = root.createElement('f')
            text = text.replace('=','',1)
            f_text = root.createTextNode(text)
            f.appendChild(f_text)
            c.appendChild(f)
        else:
            c.setAttribute('t','s')
            v = root.createElement('v')
            v_text = root.createTextNode(str(textID))
            v.appendChild(v_text)
            c.appendChild(v)
#            self.rowID+=1 # "Pointer" to next cell
        row.appendChild(c)
        self.update_spans(row)

    def get_c4ref(self,ref,row=None):
        if not row:
#            row = self.get_row(ref)
            row = self.get_row4ref(ref)
        if not row:
            return None
        for c in row.getElementsByTagName('c'):
            if c.getAttribute('r') == ref.ref:
                return c
#        print 'no cell with ref = %s found!' %ref.ref
        return None

    def get_row4ref(self,ref):
        rows = self.sheetData.getElementsByTagName('row')
        for row in rows:
            if ref.rowID == int(row.getAttribute('r')):
                return row

    def read(self,ref):
        """read cell content"""
#        if ref not in self.range:
#            return None
        if type(ref) == str:
            ref = Ref(ref)
        c = self.get_c4ref(ref)
        if not c:
            return None
        t = c.getAttribute('t')
        if not t:
            self.readInt(ref)
        elif t == 'inlineStr':
            return self.get_inlineStr(c)
        elif t == 's':
            if_ = None
            return self.getStringFromSharedStings(id_)
        elif t == 'str':
            f = c.getElementsByTagName('f')[0]
            return self.readExpressin(f)

    def readLine(self,ref=None):
        """read the holse line on ref.rowID, default self.cursor
        length is constant = dimensionRef.countColumes()"""
#        print 'Read line at RowID: %s' %ref.rowID
        length = self.dimensionRef.count_cols()
        if not ref:
            ref = self.cursor
        if ref.rowID > self.dimensionRef.end.rowID:
#            print 'Table END'
            return None
        line = list()
        if ref.rowID not in self.rows:
#            print 'empty row'
            cell = dict()
            cell['t'] = None
            cell['v'] = ''
            for i in range(length):
                line.append(cell)
            return line
        else:
            row = self.rows[ref.rowID]
            rowRef = Ref('@%s' %ref.rowID)
#            print 'Row found'
        for i in range(length):
            rowRef.walk('right')
#            print 'get C for: %s' %rowRef.start
            cell = dict()
            c = self.get_c4ref(rowRef,row)
            if not c:
#                print 'No <c> Tag found'
                cell['t'] = None
                cell['v'] = ''
                line.append(cell)
                continue
            vs = c.getElementsByTagName('v')
            if vs:
                v = vs[0]
            else:
#                print 'No <v> Tag found'
                cell['t'] = None
                cell['v'] = ''
                line.append(cell)
                continue
            if c.attributes.has_key('t'):
#                print '<v> Tag found!'
                t = c.getAttribute('t')
                cell['t'] = t
                if t == 's':
                    cell['v'] = int(v.firstChild.nodeValue)
                elif t == 'inlineStr':
                    pass
            else:
                cell['t'] = 'int'
                cell['v'] = unicode(v.firstChild.nodeValue)
            line.append(cell)
        return line

    def readExpressin(self,f):
        """get expression from cell"""
        return f.firstChild.nodeValue
    
    def getSharedStringId(self,ref):
        """get String ID from SharedStrings by Ref"""
        c = self.get_c4ref(ref)
        v = c.getElementsByTagName('v')[0]
        return int(v.firstChild.nodeValue)
    
    def get_inlineStr(self,c):
        """get cell contend from inlineStr
        <row r="3" spans="1:1">
            <c r="A3" t="inlineStr">
            <is>
            <t>test</t>
            </is>
            </c>
         </row>"""
        text = ''
        for t in c.getElementsByTagName('t'):
            text+=t.firstChild.nodeValue
        return text
    
    def getTableSice(self,ref):
        """get the size of the table starting at Ref in both directions, rows and columns. return Sqref"""
        c = self.get_c4ref(ref)
        ref2 = Ref(ref.ref)
        while self.get_c4ref(ref2):
            ref2.walk('down')
        ref2.walk('up')
        endRowID = ref2.rowID
        ref2 = Ref(ref.ref)
        while self.get_c4ref(ref2):
            ref2.walk('right')
        ref2.walk('left')
        endCN = ref2.CN
        end = Ref('%s%s'%(endRowID,endCN))
        return Sqref(ref,end)
        
    def readRow(self,ref):
        """read and return the first row of ref"""
        ref2 = Ref(ref.ref)
        row = list()
        while self.dimensionRef.end.CN != ref2.CN:
            row.append(self.read(ref2))
            ref2.walk('right')
        row.append(self.read(ref2))
        return row

    def new_conditionalFormatting(self,sqref):
        conditionalFormatting = self.root.createElement('conditionalFormatting')
        conditionalFormatting.setAttribute('sqref',sqref.ref)
        if self.tableParts:
            self.worksheet.insertBefore(conditionalFormatting,self.tableParts)
        else:
            self.worksheet.appendChild(conditionalFormatting)
        return conditionalFormatting
        
    def get_conditionalFormatting(self,sqref):
        conditionalFormattingS = self.worksheet.getElementsByTagName('conditionalFormatting')
        if len(conditionalFormattingS) == 0:
            return self.new_conditionalFormatting(sqref)
        else:
            for i in conditionalFormattingS:
                if i.getAttribute('sqref') == sqref.ref:
                    return i
            return self.new_conditionalFormatting(sqref)

    def add_conditionalFormatting(self,type_,sqref,dxfId,priority,format_=None,operator=None,text=None):
        """create a conditional formating for sheet
            <conditionalFormatting sqref="D1:D1048576">
                <cfRule type="beginsWith" dxfId="1" priority="1" operator="beginsWith" text="OK">
                    <formula>LEFT(D1,2)="OK"</formula>
                </cfRule>
                <cfRule type="expression" dxfId="5" priority="4"> # dxfId = style ID
        	    <formula>IF($F6="Default",C6=$B6,$C6="&lt;ns&gt;")</formula>
		</cfRule>
            </conditionalFormatting>"""
        conditionalFormatting = self.get_conditionalFormatting(sqref)
#        cfRule = Document().createElement('cfRule') # <--
        cfRule = self.root.createElement('cfRule')
        cfRule.setAttribute('type',type_)
        cfRule.setAttribute('dxfId',str(dxfId))
        cfRule.setAttribute('priority',str(priority))
        if type_ == 'beginsWith':
            format_ = 'LEFT(%s,%s)="%s"' %(sqref.start.ref,len(text),text)
            cfRule.setAttribute('operator',operator)
            cfRule.setAttribute('text',text)
        formula = self.root.createElement('formula')
        formula_text = self.root.createTextNode(format_)
        formula.appendChild(formula_text)
        cfRule.appendChild(formula)
        conditionalFormatting.appendChild(cfRule)

    def import_list(self,ref,l,cetAutoFit=True):
        """import a table as list to the active Sheet"""
        ref2 = Ref(ref.ref)
        for line in l:
            self.writeLine(ref2,line)
            ref2.walk('down')
    
    def add_frozen_row(self,columnID):
        """create a frozen row in active Sheet by columnID
           <sheetViews>
           <sheetView tabSelected="1" workbookViewId="0">
              <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
              <selection pane="bottomLeft"/>
           </sheetView>"""
        sheetView = self.sheetViews.getElementsByTagName('sheetView')[0]
        pane = self.root.createElement('pane')
        pane.setAttribute('ySplit',str(columnID))
        pane.setAttribute('topLeftCell','A%s' %(columnID+1))
        pane.setAttribute('activePane','bottomLeft')
        pane.setAttribute('state','frozen')
        sheetView.appendChild(pane)
        selection = self.root.createElement('selection')
        selection.setAttribute('pane',"bottomLeft")
        sheetView.appendChild(selection)

    def set_cel_width(self,columns):
        """"set the column width to fitt the largest cell contend
        <cols>
         <col min="1" max="1" width="4.7109375" bestFit="1" customWidth="1"/>
         <col min="2" max="2" width="12" bestFit="1" customWidth="1"/>
         <col min="3" max="3" width="13.42578125" bestFit="1" customWidth="1"/>
         </cols> """
        if not self.cols:
            self.cols = self.root.createElement('cols')
            self.worksheet.insertBefore(self.cols,self.sheetData)
        i = 0
        for column in columns:
            i+=1
            col = self.root.createElement('col')
            col.setAttribute('min',str(i))
            col.setAttribute('max',str(i))
            col.setAttribute('width',str(column))
            col.setAttribute('bestFit',str(1))
            col.setAttribute('customWidth',str(1))
            self.cols.appendChild(col)

    def update_dimension(self):
        """update the dimension of the Sheet using self.dimension"""
        self.dimension.setAttribute('ref',self.dimensionRef.ref)

class Workbook(OP):
    def __init__(self,appName='',lastEdited=4,lowestEdited=4,rupBuild=4506,defaultThemeVersion=124226,calcId=125725,f=None):
        if f:
            self._open(f)
        else:
            self.root = Document()
            self.workbook = self.root.createElement('workbook')
            self.workbook.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            self.workbook.setAttribute('xmlns:r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            self.root.appendChild(self.workbook)
            self.fileVersion = self.root.createElement('fileVersion')
            self.fileVersion.setAttribute('appName',appName)
            self.fileVersion.setAttribute('lastEdited',str(lastEdited))
            self.fileVersion.setAttribute('lowestEdited',str(lowestEdited))
            self.fileVersion.setAttribute('rupBuild',str(rupBuild))
            self.workbook.appendChild(self.fileVersion)
            self.workbookPr = self.root.createElement('workbookPr')
            self.workbookPr.setAttribute('defaultThemeVersion',str(defaultThemeVersion))
            self.workbook.appendChild(self.workbookPr)
            self.bookViews = self.root.createElement('bookViews')
            self.workbook.appendChild(self.bookViews)
            self.workbookView = self.root.createElement('workbookView')
            self.workbookView.setAttribute('xWindow','240')
            self.workbookView.setAttribute('yWindow','105')
            self.workbookView.setAttribute('windowWidth','18795')
            self.workbookView.setAttribute('windowHeight','12270')
            self.bookViews.appendChild(self.workbookView)
            self.sheets = self.root.createElement('sheets')
            self.workbook.appendChild(self.sheets)
            self.calcPr = self.root.createElement('calcPr')
            self.calcPr.setAttribute('calcId',str(calcId))
            self.workbook.appendChild(self.calcPr)

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.workbook = self.root.getElementsByTagName('workbook')[0]
        self.fileVersion = self.workbook.getElementsByTagName('fileVersion')[0]
        self.workbookPr = self.workbook.getElementsByTagName('workbookPr')[0]
        self.bookViews = self.workbook.getElementsByTagName('bookViews')[0]
        self.sheets = self.workbook.getElementsByTagName('sheets')[0]
        self.calcPr = self.workbook.getElementsByTagName('calcPr')[0]
        self.workbookView = self.bookViews.getElementsByTagName('workbookView')[0]

    def new_sheet(self,name,sheetId,rId):
        newSheet = self.root.createElement('sheet')
        newSheet.setAttribute('name',name)
        newSheet.setAttribute('sheetId',str(sheetId))
        newSheet.setAttribute('r:id','rId%s'%rId)
        self.sheets.appendChild(newSheet)

    def set_activeTab(self,sheetName):
        i  = 0
        for sheet in self.sheets.getElementsByTagName('sheet'):
            if sheet.getAttribute('name') == sheetName:
                self.workbookView.setAttribute('activeTab',str(i))
            i+=1

    def getRId4Sheet(self,sheetName):
        """return the rId for the Sheet sheetName, else -1"""
        for sheet in self.sheets.getElementsByTagName('sheet'):
            if sheet.getAttribute('name') == sheetName:
                return sheet.getAttribute('r:id')
        return -1

    def listSheets(self):
        """retrun a dict of all sheets in workbook with sheet name and sheetId"""
        sheets = dict()
        for sheet in self.sheets.getElementsByTagName('sheet'):
            name = sheet.getAttribute('name')
            sheetID = sheet.getAttribute('sheetId')
            sheets[name] = sheetID
        return sheets
        
class Core(OP):
    def __init__(self,creator='',f=None):
        if f:
            self._open(f)
        else:
            self.root = Document()
            self.cp_coreProperties = self.root.createElement('cp:coreProperties')
            self.cp_coreProperties.setAttribute('xmlns:cp','http://schemas.openxmlformats.org/package/2006/metadata/core-properties')
            self.cp_coreProperties.setAttribute('xmlns:dc','http://purl.org/dc/elements/1.1/')
            self.cp_coreProperties.setAttribute('xmlns:dcterms','http://purl.org/dc/terms/')
            self.cp_coreProperties.setAttribute('xmlns:dcmitype','http://purl.org/dc/dcmitype/')
            self.cp_coreProperties.setAttribute('xmlns:xsi','http://www.w3.org/2001/XMLSchema-instance')
            self.root.appendChild(self.cp_coreProperties) 
            self.dc_creator = self.root.createElement('dc:creator')
            self.cp_coreProperties.appendChild(self.dc_creator) 
            self.set_creator(creator)
            self.cp_lastModifiedBy = self.root.createElement('cp:lastModifiedBy')
            self.cp_coreProperties.appendChild(self.cp_lastModifiedBy)
            self.set_lastModifiedBy(creator)
            self.dcterms_created = self.root.createElement('dcterms:created')
            self.dcterms_created.setAttribute('xsi:type','dcterms:W3CDTF')
            self.cp_coreProperties.appendChild(self.dcterms_created)
            self.set_dcterms_created()
            self.dcterms_modified = self.root.createElement('dcterms:modified')
            self.dcterms_modified.setAttribute('xsi:type','dcterms:W3CDTF')
            self.cp_coreProperties.appendChild(self.dcterms_modified)
            self.set_dcterms_modified()

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.cp_coreProperties = self.root.getElementsByTagName('cp:coreProperties')[0]
        dc_creator = self.cp_coreProperties.getElementsByTagName('dc:creator')
        if dc_creator:
            self.dc_creator = dc_creator[0]
        else:
            self.dc_creator = self.root.createElement('dc:creator')
            self.cp_coreProperties.appendChild(self.dc_creator)
        self.cp_lastModifiedBy = self.cp_coreProperties.getElementsByTagName('cp:lastModifiedBy')[0]
        self.dcterms_created = self.cp_coreProperties.getElementsByTagName('dcterms:created')[0]
        self.dcterms_modified = self.cp_coreProperties.getElementsByTagName('dcterms:modified')[0]
        
    def set_creator(self,creator):
        self.dc_creator_text = self.root.createTextNode(creator)
        if self.dc_creator.firstChild:
            self.dc_creator.removeChild(self.dc_creator.firstChild)
        self.dc_creator.appendChild(self.dc_creator_text)

    def set_lastModifiedBy(self,modifier):
        self.cp_lastModifiedBy_text = self.root.createTextNode(modifier)
        if self.cp_lastModifiedBy.firstChild:
            self.cp_lastModifiedBy.removeChild(self.cp_lastModifiedBy.firstChild)
        self.cp_lastModifiedBy.appendChild(self.cp_lastModifiedBy_text)

    def set_dcterms_created(self):
        date = time.strftime('%Y-%m-%dT%XZ')
        self.dcterms_created_text = self.root.createTextNode(date)
        if self.dcterms_created.firstChild:
            self.dcterms_created.removeChild(self.dcterms_created.firstChild)
        self.dcterms_created.appendChild(self.dcterms_created_text)

    def set_dcterms_modified(self):
        date = time.strftime('%Y-%m-%dT%XZ')
        self.dcterms_modified_text = self.root.createTextNode(date)
        if self.dcterms_modified.firstChild:
            self.dcterms_modified.removeChild(self.modified.firstChild)
        self.dcterms_modified.appendChild(self.dcterms_modified_text)

class Styles(OP):
    def __init__(self,f=None):
        self.countFonts = 0
        self.countFills = 0
        self.countBorders = 0
        self.countCellStyleXfs = 0
        self.countCellXfs = 0
        self.countCellStyles = 0
        if f:
            self._open(f)
        else:
            self.root = Document()
            self.styleSheet = self.root.createElement('styleSheet')
            self.styleSheet.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            self.root.appendChild(self.styleSheet)
            self.fonts = self.root.createElement('fonts')
            self.styleSheet.appendChild(self.fonts)
            self.new_font(11,1,'Calibri',2,'minor')
            self.fills = self.root.createElement('fills')
            self.styleSheet.appendChild(self.fills)
            self.new_fill('none')
            self.new_fill('gray125')
            self.borders = self.root.createElement('borders')
            self.styleSheet.appendChild(self.borders)
            self.new_border({})
            self.cellStyleXfs = self.root.createElement('cellStyleXfs')
            self.styleSheet.appendChild(self.cellStyleXfs)
            self.new_cellStyleXfs(0,0,0,0)
            self.cellXfs = self.root.createElement('cellXfs')
            self.styleSheet.appendChild(self.cellXfs)
            self.new_cellXfs(0,0,0,0)
            self.cellStyles = self.root.createElement('cellStyles')
            self.styleSheet.appendChild(self.cellStyles)
            self.new_cellStyle('Standard',0,0)
            self.dxfs = self.root.createElement('dxfs')
            self.dxfs.setAttribute('count','0')
            self.styleSheet.appendChild(self.dxfs)
    #        self.tableStyles = self.root.createElement('tableStyles')
    #        self.tableStyles.setAttribute('count','0')
    #        self.tableStyles.setAttribute('defaultTableStyle','TableStyleMedium9')
    #        self.tableStyles.setAttribute('defaultPivotStyle','PivotStyleLight16')
    #        self.styleSheet.appendChild(self.tableStyles)
            self.xdfs = None
            self.countDxfs = 0

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.styleSheet = self.root.getElementsByTagName('styleSheet')[0]
        self.fonts = self.styleSheet.getElementsByTagName('fonts')[0]
        self.fills = self.styleSheet.getElementsByTagName('fills')[0]
        self.borders = self.styleSheet.getElementsByTagName('borders')[0]
        self.cellStyleXfs = self.styleSheet.getElementsByTagName('cellStyleXfs')[0]
        self.cellXfs = self.styleSheet.getElementsByTagName('cellXfs')[0]
        self.cellStyles = self.styleSheet.getElementsByTagName('cellStyles')[0]
        self.dxfs = self.styleSheet.getElementsByTagName('dxfs')[0]

    def new_cellStyle(self,name,xfId,builtinId):
        cellStyle = self.root.createElement('cellStyle')
        cellStyle.setAttribute('name',name)
        cellStyle.setAttribute('xfId',str(xfId))
        cellStyle.setAttribute('builtinId',str(builtinId))
        self.cellStyles.appendChild(cellStyle)
        self.countCellStyles+=1
        self.cellStyles.setAttribute('count',str(self.countCellStyles))

    def new_xf(self,numFmtId,fontId,fillId,borderId,xfId=None):
        xf = self.root.createElement('xf')
        xf.setAttribute('numFmtId',str(numFmtId))
        xf.setAttribute('fontId',str(fontId))
        xf.setAttribute('fillId',str(fillId))
        xf.setAttribute('borderId',str(borderId))
        if xfId:
            xf.setAttribute('xfId',str(xfId))
        return xf

    def new_cellStyleXfs(self,numFmtId,fontId,fillId,borderId):
        self.cellStyleXfs.appendChild(self.new_xf(numFmtId,fontId,fillId,borderId))
        self.countCellStyleXfs+=1
        self.cellStyleXfs.setAttribute('count',str(self.countCellStyleXfs))

    def new_cellXfs(self,numFmtId,fontId,fillId,borderId):
        self.cellXfs.appendChild(self.new_xf(numFmtId,fontId,fillId,borderId,0))
        self.countCellXfs+=1
        self.cellXfs.setAttribute('count',str(self.countCellXfs))

    def new_border(self,borderStyle):
        """{'left':{'style':style,'color':color}, 'right':{ }} """
        border = self.root.createElement('border')
        for position in ('left','right','top','bottom','diagonal'):
            newPos = self.root.createElement(position) 
            if position in borderStyle:
                if 'style' in borderStyle[position]:
                    newPos.setAttribute('style',borderStyle[position]['style'])
                border.appendChild(newPos)
                newColor = self.root.createElement('color')
                if 'indexed' in borderStyle[position]:
                    newColor.setAttribute('indexed',borderStyle[position]['indexed'])
                newPos.appendChild(newColor)
            border.appendChild(newPos)
        self.borders.appendChild(border)
        self.countBorders+=1
        self.borders.setAttribute('count',str(self.countBorders))
        
    def new_fill(self,patternType):
        fill = self.root.createElement('fill')
        patternFill = self.root.createElement('patternFill')
        patternFill.setAttribute('patternType',patternType)
        fill.appendChild(patternFill)
        self.fills.appendChild(fill)
        self.countFills+=1
        self.fills.setAttribute('count',str(self.countFills))

    def new_font(self,sz_val,color_theme,name_val,family_val,scheme_val):
        font = self.root.createElement('font')
        sz = self.root.createElement('sz')
        sz.setAttribute('val',str(sz_val))
        font.appendChild(sz)
        color = self.root.createElement('color')
        color.setAttribute('theme',str(color_theme))
        font.appendChild(color)
        name = self.root.createElement('name')
        name.setAttribute('val',name_val)
        font.appendChild(name)
        family = self.root.createElement('family')
        family.setAttribute('val',str(family_val))
        font.appendChild(family)
        scheme = self.root.createElement('scheme')
        scheme.setAttribute('val',scheme_val)
        font.appendChild(scheme)
        self.countFonts+=1
        self.fonts.setAttribute('count',str(self.countFonts))
        self.fonts.appendChild(font)

    def get_dxfId(self,rgb):
        """search a style with background color = rbg and return its dxfId"""
        if not self.dxfs:
            return self.new_dxfs(rgb)
        dxfId = -1
        for dxf in self.dxfs.getElementsByTagName('dxf'):
            dxfId+=1
            fillS = dxf.getElementsByTagName('fill')
            if not fillS:
                continue
            patternFillS = fillS[0].getElementsByTagName('patternFill')
            if not patternFillS:
                continue
            bgColorS = patternFillS[0].getElementsByTagName('bgColor')
            if not bgColorS:
                continue
            if type(rgb) == dict:
                i=0
                for key in rgb.keys():
                    if not bgColorS[0].attributes.has_key(key):
                        break
                    value = bgColorS[0].getAttribute(key)
                    if value == rgb[key]: 
                        i+=1
                if i == len(rgb):
                    return dxfId
            else:
                if not bgColorS[0].attributes.has_key('rgb'):
                    continue
                if bgColorS[0].getAttribute('rgb') == rgb:
                    return dxfId
        return self.new_dxfs(rgb)
            
    def new_dxfs(self,rgb):
        """create a new style with background color = rgb and return its dxfId
    <dxfs>
        <dxf>
		<fill>
			<patternFill>
				<bgColor theme="1" tint="0.499984740745262"/>
			</patternFill>
		</fill>
	</dxf>
	...
    </dxfs>"""
        if not self.dxfs:
            self.dxfs = self.root.createElement('dxfs')
            self.styleSheet.appendChild(self.dxfs)
        dxf = self.root.createElement('dxf')
        fill = self.root.createElement('fill')
        patternFill = self.root.createElement('patternFill')
        bgColor = self.root.createElement('bgColor')
        if type(rgb) == type(dict()):
            bgColor.setAttribute('theme',rgb['theme'])
            if 'tint' in rgb.keys():
                bgColor.setAttribute('tint',rgb['tint'])
        else:
            bgColor.setAttribute('rgb',rgb)
        patternFill.appendChild(bgColor)
        fill.appendChild(patternFill)
        dxf.appendChild(fill)
        self.dxfs.appendChild(dxf)
        dxfsId = self.countDxfs
        self.countDxfs+=1
        return dxfsId

class SharedStrings(OP):
    def __init__(self,f=None):
        if f:
            self._open(f)
        else:
            self.count = 0
            self.uniqueCount = 0
            self.root = Document()
            self.sst = self.root.createElement('sst')
            self.sst.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            self.sst.setAttribute('count',str(self.count))
            self.sst.setAttribute('uniqueCount',str(self.uniqueCount))
            self.root.appendChild(self.sst)
            self.length = 0 # number of strings in SharedStrings.xml

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.sst = self.root.getElementsByTagName('sst')[0]
        self.sis = self.sst.getElementsByTagName('si')
    def _getID4text(self,text):
        """get the ID of a text in SharedStings.xml
        if no entry is found, -1 is returned
        searching only after <t>text<t> tag, because it seams that every <si> tag has only one <t> Tag"""
        id_ = 0
        for t in self.sst.getElementsByTagName('t'):
            if t.firstChild.nodeValue == text:
                return id_
            id_+=1
        return -1

    def newString(self,text):
        text = str(text)
        id_ = self._getID4text(text)
        if id_ == -1:
            si = self.root.createElement('si')
            t = self.root.createElement('t')
            t_text = self.root.createTextNode(text)
            t.appendChild(t_text)
            si.appendChild(t)
            self.sst.appendChild(si)
            id_ = self.length
            self.length +=1
        return id_
    
    def read(self,id_):
        si = self.sis[id_]
        text = ''
#        print si.toprettyxml()
        for t in si.getElementsByTagName('t'):
            child = t.firstChild
            if not child:
                continue
            text+=child.nodeValue
        return text

class Ref(OP):
    """Represents a referenc to a single cell.
    has the volloring properys:
        * ref (A1)
        * RowID (1)
        * CN (A)"""
    
    def __init__(self,ref):
        self._fixRowID = False
        self._fixCN = False
        self.setRef(ref)

    def getRef(self):
        if self._fixCN:
            return '$%s%s' %(self._CN,self._rowID)
        else:
            return '%s%s' %(self._CN,self._rowID)
    def setRef(self,value):
        if value.startswith('$'):
            self._fixCN = True
            value = value.replace('$','',1)
        else:
            self._fixCN = False
        self._rowID = int(self.get_number(value))
        self._CN = self.get_text(value)
    ref = property(getRef,setRef)
    def getRowID(self):
        return self._rowID
    def setRowID(self,value):
        if value <= 0:
            self._rowID = 1
        else:
            self._rowID = int(value)
    rowID = property(getRowID,setRowID)
    def getCN(self):
        return self._CN
    def setCN(self,value):
        if value.startswith('$'):
            self._fixCN = True
        else:
            self._fixCN = False
        self._CN = value
    CN = property(getCN,setCN)
            
    def incChr(self,c):
        """simple increment a chr by i"""
        return chr(ord(c)+1)

    def decChr(self,c):
        return chr(ord(c)-1)
    
    def incCol(self,CN):
        """increase the column by one (A+1=B; ZZ+1=AAA)"""
        CN = list(CN)
        CN.reverse()
        i = 0
        inc = False
        while i < len(CN):
            if inc:
                if CN[i] != 'Z':
                    CN[i] = self.incChr(CN[i])
                    inc = False
                    break
                else:
                    CN[i] = 'A'
                    i+=1
                    continue
            if CN[i] == 'Z':
                CN[i] = 'A'
                inc = True
                i+=1
                continue
            else:
                CN[i] = self.incChr(CN[i])
                break
        if inc:
            CN.append('A')
        CN.reverse()
        return ''.join(CN)

    def decCol(self,CN):
        """decrease the column by one (B-1=A; AA-1=Z)"""
        if CN == 'A':
            return 'A'
        CN = list(CN)
        CN.reverse()
        i = 0
        dec = False
        while i < len(CN):
            if dec:
                if CN[i] != 'A':
                    CN[i] = self.decChr(CN[i])
                    dec = False
                    break
                else:
                    CN[i] = 'Z'
                    i+=1
                    continue
            if CN[i] == 'A':
                CN[i] = 'Z'
                dec = True
                i+=1
                continue
            else:
                CN[i] = self.decChr(CN[i])
                break
        if dec:
            CN.pop()
        CN.reverse()
        return ''.join(CN)
    
#    def appendColumns(self,CN,cols):
#        """CN+=cols"""
#        CN = list(CN)
#        i = 0
#        while i < len(CN):
#            if CN[i] == 'Z':
#                if i == len(CN)-1:
#                    if cols > 26:
#                        cols = cols-26
#                        endCN.append('Z')
#                        self._endCN = ''.join(endCN)
#                        return self.appendColumns(cols)
#                    else:
#                        endCN.append(chr(cols+64))
#                        self._endCN = ''.join(endCN)
#                        break
#                else:
#                    i+=1
#                    continue
#            elif ord(endCN[i])+cols > 90:
#                cols = cols - (90-ord(endCN[i]))
#                endCN[i] = 'Z'
#                self._endCN = ''.join(endCN)
#                self.appendColumns(cols)
#                break
#            else:
#                endCN[i] = chr(ord(endCN[i])+cols)
#                self._endCN = ''.join(endCN)
#                break
#            i+=1

    def walk(self,d):
        """move the Ref one cell in a direction(up,down,left,right)"""
        if d == 'right':
            self._CN = self.incCol(self._CN)
        elif d == 'left':
            self._CN = self.decCol(self._CN)
        elif d == 'up':
            self._rowID -=1
        elif d == 'down':
            self._rowID +=1

    def update(self,ref):
        """if ref > self then update self.startRowID/self.startCN"""
        if self.compCN(ref._CN,self._CN):
            self._CN = ref._CN
        if ref._rowID > self._rowID:
            self._rowID = ref.rowID

class Sqref(OP):
    """class to represent a square ref like A1:B3"""
    def __init__(self,start,end=None):
        if type(start) != Ref:
            if ':' in start:
                sq = start.split(':')
                start = sq[0]
                end = sq[1]
            self.start = Ref(start)
        else:
            self.start = start
        if end:
            if type(end) != Ref:
                self.end = Ref(end)
            else:
                self.end = end
        else:
            self.end = Ref(self.start.ref)

    def getRef(self):
        return '%s%s:%s%s' %(self.start._CN,self.start._rowID,self.end._CN,self.end._rowID)
    def setRef(self,value):
        ref = value.split(':')
        self.start.setRef(ref[0])
        if len(ref) == 1:
            self.end.setRef(ref[0])
        elif len(ref) == 2:
            self.end.setRef(ref[1])
    ref = property(getRef,setRef)

    def count_rows(self):
        return self.end.rowID - self.start.rowID +1

    def count_cols(self):
        return self.getInt4CN(self.end.CN) - self.getInt4CN(self.start.CN) +1
    
class Table(OP):
    def __init__(self,id_='',name='',sqref='',header='',displayName=None,totalsRowShown=0,tableStyle='TableStyleLight16',f=None):
        if f:
            self._open(f)
        else:
            if not displayName:
                displayName = name
            if not tableStyle:
                self.tableStyle = 'TableStyleLight16'
            else:
                self.tableStyle = tableStyle
            self.root = Document()
            self.table = self.root.createElement('table')
            self.table.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            self.table.setAttribute('id',str(id_))
            self.table.setAttribute('name',name)
            self.table.setAttribute('displayName',displayName)
            self.table.setAttribute('ref',sqref.ref)
            self.table.setAttribute('totalsRowShown',str(totalsRowShown))
            self.new_table(sqref,name,header)

    def _open(self,f):
        if type(f) == str:
            self.root = parseString(f)
        else:
            self.root = parse(f)
        self.table = self.root.getElementsByTagName('table')[0]
        self.autoFilter = self.table.getElementsByTagName('autoFilter')[0]
        self.tableColumns = self.table.getElementsByTagName('tableColumns')[0]
        self.tableStyleInfo = self.table.getElementsByTagName('tableStyleInfo')[0]

    def new_table(self,sqref,name,header):
        """
        ref: Ref # Start cell
        name : str # name of table
        header : list # name of columns in the first row (header)
        --------------------
        Ref('A1'),'table1',['Name','first Name']
        --------------------
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tabelle1" displayName="Tabelle1" ref="A1:D101" totalsRowShown="0">
                    <autoFilter ref="A1:D101"/>
                    <tableColumns count="4">
                            <tableColumn id="1" name="Column 1"/>
                            <tableColumn id="2" name="Column 2"/>
                            <tableColumn id="3" name="Column 3"/>
                            <tableColumn id="4" name="Column 4"/>
                    </tableColumns>
                    <tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
            </table>"""
#        ref.appendColumns(len(cols))
        self.count = sqref.count_cols()
        self.autoFilter = self.root.createElement('autoFilter')
        self.autoFilter.setAttribute('ref',sqref.ref)
        self.table.appendChild(self.autoFilter)
        self.tableColumns = self.root.createElement('tableColumns')
        self.tableColumns.setAttribute('count',str(self.count))
        i=0
        for name in header:
            #add the header row
            i+=1
#            ref3 = column + ref.startRowID
#            name = self.get_inlineStr(ref3,sheet)
#            if not name:
#                name = 'Column %s' %i
#                r = chr(ord(ref2['refStartCN'])+(i-1))+'1'
#                sheetData = sheet.getElementsByTagName('sheetData')[0]
#                self.new_row_sharedStr(r,name)
            newTableColumn = self.new_tableColumn(str(i),name)
    #        print newTableColumn.toprettyxml()
            self.tableColumns.appendChild(newTableColumn)
#            column = chr(ord(column) + 1) #?
        self.tableStyleInfo = Element('tableStyleInfo')
        self.tableStyleInfo.setAttribute('name',self.tableStyle)
        self.tableStyleInfo.setAttribute('showFirstColumn',"0")
        self.tableStyleInfo.setAttribute('showLastColumn',"0")
        self.tableStyleInfo.setAttribute('showColumnStripes',"0")
        self.tableStyleInfo.setAttribute('showRowStripes',"1")
        self.table.appendChild(self.tableColumns)
        self.table.appendChild(self.tableStyleInfo)
        self.root.appendChild(self.table)
        
    def new_tableColumn(self,id_,name):
        tableColumn = Element('tableColumn')
        tableColumn.setAttribute('id',id_)
        tableColumn.setAttribute('name',name)
        return tableColumn

class Calc(OP):
    def __init__(self,company='',userName='',f=None,sheets=True):
        """create a new blank OpenXML Calc workbook without any sheets
        docProps/app.xml <-- new
        [Content_Types].xml <-- extend
        _rels/.rels <-- extend
        """
        if f:
            self._open(f,sheets=sheets)
        else:
            self.OP = dict()
            self.OP['[Content_Types].xml'] = Content_Types()
            self.OP['_rels/.rels'] = Relationships()
            self.OP['docProps/app.xml'] = App(company)
            self.OP['[Content_Types].xml'].new_Override('/docProps/app.xml')
            self.OP['docProps/app.xml'].set_rId(self.OP['_rels/.rels'].new_relationship('docProps/app.xml'))
            self.OP['docProps/core.xml'] = Core(userName)
            self.OP['[Content_Types].xml'].new_Override('/docProps/core.xml')
            self.OP['docProps/core.xml'].set_rId(self.OP['_rels/.rels'].new_relationship('docProps/core.xml'))
            self.OP['xl/styles.xml'] = Styles()
            self.OP['xl/workbook.xml'] = Workbook('xl')
            self.OP['[Content_Types].xml'].new_Override('/xl/workbook.xml')
            self.OP['xl/workbook.xml'].set_rId(self.OP['_rels/.rels'].new_relationship('xl/workbook.xml'))
            self.OP['xl/sharedStrings.xml'] = SharedStrings()
            self.OP['xl/_rels/workbook.xml.rels'] = Relationships()
#            self.OP['xl/worksheets/sheet1.xml'] = Sheet()
            self.OP['xl/styles.xml'].set_rId(self.OP['xl/_rels/workbook.xml.rels'].new_relationship('styles.xml'))
            self.OP['[Content_Types].xml'].new_Override('/xl/styles.xml')
            self.OP['xl/sharedStrings.xml'].set_rId(self.OP['xl/_rels/workbook.xml.rels'].new_relationship('sharedStrings.xml'))
            self.OP['[Content_Types].xml'].new_Override('/xl/sharedStrings.xml')
            self.tables = 0
            self.myZIP = False

    def __del__(self):
#        print 'Del'
        if self.myZIP:
            self.myZIP.close()

    def newSheet(self,name=None):
        """create an new Sheet in workbook
        xl/worksheets/sheet1.xml <-- new
        [Content_Types].xml <-- sheet1.xml register
        docProps/app.xml <-- Table register
        xl/workbook.xml <-- app.xml quote reference
        xl/_rels/workbook.xml.rels <-- register !!
        """
#        rId = self.OP['xl/_rels/workbook.xml.rels'].get_NewID
        newTableID = self.OP['docProps/app.xml'].getNextTableID()
        if not name:
            name = 'Tabelle%s' %newTableID
        self.OP['xl/worksheets/sheet%s.xml' %newTableID] = Sheet()
        self.OP['[Content_Types].xml'].new_Sheet(newTableID)
        self.OP['docProps/app.xml'].new_Table(name)
        rId = self.OP['xl/_rels/workbook.xml.rels'].new_relationship('worksheets/sheet%s.xml' %newTableID)
        self.OP['xl/workbook.xml'].new_sheet(name,newTableID,rId)
        self.activeSheet = self.OP['xl/worksheets/sheet%s.xml' %newTableID]
        self.selectSheet(name)
#        print 'New Sheet: %s %s' %(name,newTableID)

    def save(self,path):
        """save the Workbook"""
        self.activeSheet.update_dimension()
        xlsx = ZipFile(path,'w',ZIP_DEFLATED)
    #    xlsx = ZipFile('C:\\temp\\test.xlsx','w')
        for i in self.OP:
            xlsx.writestr(i,self.OP[i].toxml(encoding='UTF-8'))
        xlsx.close()
        
    def getrowID(self,r): # <- odd
        m = re.search('\d+\.?\d*',r)
        return m.group(0)

    def update_spans(self,row): # <- odd
        spanEnd = len(row.getElementsByTagName('c'))
    #    spanEnd+=1
        spans = '1:%s' %spanEnd
        row.setAttribute('spans',spans)
    def update_sst_count(self,sheet): # <- odd
        sst = sheet.getElementsByTagName('sst')[0]
        count = len(list(sst.getElementsByTagName('si')))
        sst.setAttribute('count',str(count))
        sst.setAttribute('uniqueCount',str(count))

    def write(self,ref,text,engine='inlineStr'):
        """write text into a cell like Ref('A1')
        default engine is to use inline string rather then SharedStrings,
        because it is simpler and Excel will convert this automatically to SharedStrings when saving the sheet"""
        if type(ref) == str:
            ref = Ref(ref)
        self.activeSheet.write(ref,text)

    def read(self,ref):
        """read the contend of a cell (Ref) from the active Sheet"""
        if type(ref) == str:
            ref = Ref(ref)
        if self.activeSheet.writeEngine == 'sharedStrings':
            id_ = self.activeSheet.getSharedStringId(ref)
            return self.OP['xl/sharedStrings.xml'].read(id_)
        return self.activeSheet.read(ref)

    def readLine(self,ref=None):
        """read a hole line in activeSheet"""
        if not ref:
            ref = self.activeSheet.cursor
        line = list()
        cells = self.activeSheet.readLine(ref)
        if not cells:
            return None
        for cell in cells:
            if cell['t'] == None:
                value = ''
            elif cell['t'] == 'int':
                value = cell['v']
            elif cell['t'] == 's':
                value = self.OP['xl/sharedStrings.xml'].read(cell['v'])
            line.append(value)
        self.activeSheet.cursor.walk('down')
        return line

###########################
## conditional formating ##
###########################

    def add_conForm_expression(self,sqref,dxfId,format_,priority):
        """add a conditional formating by expression to active Sheet"""
        self.activeSheet.add_conditionalFormatting('expression',sqref,dxfId,priority,format_)
    def add_conForm_beginWith(self,sqref,dxfId,priority,text):
        """add a conditional formating by begin with to active Sheet"""
        self.activeSheet.add_conditionalFormatting('beginsWith',sqref,dxfId,priority,operator='beginsWith',text=text)

    def getStyle(self,rgb):
        return self.OP['xl/styles.xml'].get_dxfId(rgb)
    
############
## Styles ##
############

    def set_cel_width(self,columns):
        """"set the column width to fitt the largest cell contend"""
        self.activeSheet.set_cel_width(columns)

    def add_style_fills(self,patternType,rgb): # <- move to Style class
        """Adds format template(<fills> Tag) to xl/styles.xml and returns its ID"""
        styleSheet = self.OP['xl/styles.xml'].styleSheet
        root = self.activeSheet.root
        fills = styleSheet.getElementsByTagName('fills')[0]
        count = fills.getAttribute('count')
        count = str(int(count)+1)
        fills.setAttribute('count',count)
        fill = root.createElement('fill')
        patternFill = root.createElement('patternFill')
        patternFill.setAttribute('patternType',patternType)
        fgColor = root.createElement('fgColor')
        fgColor.setAttribute('rgb',rgb)
        patternFill.appendChild(fgColor)
        fill.appendChild(patternFill)
        fills.appendChild(fill)
        return str(int(count)-1)

    def add_style_cellXfs(self,fillId,borderId=0,xfId=0,applyFill=1,numFmtId=0,fontId=0): # <- move to Style class
        styleSheet = self.OP['xl/styles.xml'].styleSheet
        root = self.activeSheet.root
        cellXfs = styleSheet.getElementsByTagName('cellXfs')[0]
        count = cellXfs.getAttribute('count')
        count = str(int(count)+1)
        xf = root.createElement('xf')
        xf.setAttribute('fillId',str(fillId))
        xf.setAttribute('borderId',str(borderId))
        xf.setAttribute('xfId',str(xfId))
        xf.setAttribute('applyFill',str(applyFill))
        xf.setAttribute('numFmtId',str(numFmtId))
        xf.setAttribute('fontId',str(fontId))
        cellXfs.appendChild(xf)
        cellXfs.setAttribute('count',count)
        id_ = str(int(count)-1)
        return id_

    def add_cell_color(self,rgb,ref):
        """set the background color of a cell (Ref) in active Sheet"""
        fillId = self.add_style_fills('solid',rgb)
        s = self.add_style_cellXfs(fillId)
        worksheet = self.activeSheet.worksheet
        sheetData = worksheet.getElementsByTagName('sheetData')[0]
        c = self.get_c4ref(ref)
        c.setAttribute('s',s)

    def selectSheet(self,sheetName):
        """select a sheet by name and make to self.activeSheet"""
        self.OP['xl/workbook.xml'].set_activeTab(sheetName)
        rId = self.OP['xl/workbook.xml'].getRId4Sheet(sheetName)
        sheetPath = self.OP['xl/_rels/workbook.xml.rels'].getTarget(rId)
        sheetPath = 'xl/%s' %sheetPath
        if sheetPath not in self.OP:
            if sys.version_info < (2, 6):
                self.OP[sheetPath] = Sheet(f=self.myZIP.read(sheetPath))
            else:
                self.OP[sheetPath] = Sheet(f=self.myZIP.open(sheetPath))
        self.activeSheet = self.OP[sheetPath]

    def hideColume(self,min_,max_):
        """hide a column by columnID"""
        self.activeSheet.hideColume(min_,max_)

    def formatTable(self,ref,name,tableStyle=None):
        """create a Table with default formating settings

        xl/printerSettings/printerSettings1.bin # <- new but ignore
        
        xl/tables/table1.xml # <- new 1.
        
        xl/worksheets/sheet1.xml # <- 2.
            <pageSetup paperSize="9" orientation="portrait" r:id="rId1"/> # <- ignore
            
            <tableParts count="1">
            
                <tablePart r:id="rId2"/>
                
            </tableParts>
            
        xl/worksheets/_rels/sheet1.xml.rels # <- new 3.
        
        [Content_Types].xml # <- 4.
        
            <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>  # <- ignore
            
            <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
        """
        if type(ref) == str:
            ref = Ref(ref)
        sqref = self.activeSheet.getTableSice(ref)
        self.tables+=1 # counter for tables, needed for id of the new table
        # id_,name,ref,header
        header = self.activeSheet.readRow(ref)
        self.OP['xl/tables/table%s.xml' %self.tables] = Table(self.tables,name,sqref,header,tableStyle=tableStyle) # 1.
        self.OP['[Content_Types].xml'].new_Override('/xl/tables/table%s.xml' %self.tables) # 4.
        self.OP['xl/worksheets/_rels/sheet1.xml.rels'] = Relationships() # 3.
        rId = self.OP['xl/worksheets/_rels/sheet1.xml.rels'].new_relationship('../tables/table%s.xml' %self.tables) # 3.1
        self.activeSheet.addTablePart(rId) # 2.
        return sqref

    def import_list(self,ref,l):
        """import a list"""
        if type(ref) == str:
            ref = Ref(ref)
        self.activeSheet.import_list(ref,l)

    def import_csv(self,ref,text,sep=';'):
        """import a table from a CSV file"""
        if type(ref) != Ref:
            ref = Ref(ref)
        out = list()
        for line in text.split('\n'):
            if line.startswith(sep):
                continue
            out.append(line.split(sep))
        return self.import_list(ref,out)

    def add_frozen_row(self,columnID):
        """freeze row columnID"""
        self.activeSheet.add_frozen_row(columnID)

    def selectCell(self,ref):
        if type(ref) != Ref:
            ref = Ref(ref)
        self.activeSheet.selectCell(ref)

    def _open(self,name,sheets=True):
        """open an existing xlsx file"""
        self.OP = dict()
        self.myZIP = ZipFile(name,'r')
        if True:
#        with ZipFile(name,'r') as myZIP:
            if sys.version_info < (2, 6):
                f = self.myZIP.read('[Content_Types].xml')
            else:
                f = self.myZIP.open('[Content_Types].xml')
            self.OP['[Content_Types].xml'] = Content_Types(f)
            for override in self.OP['[Content_Types].xml'].getOverrides():
#                print 'Open: %s' %override[1]
                if override[0] == 'application/vnd.openxmlformats-officedocument.extended-properties+xml':
                    if sys.version_info < (2, 6):
                        self.OP['docProps/app.xml'] = App(f=self.myZIP.read('docProps/app.xml'))
                    else:
                        self.OP['docProps/app.xml'] = App(f=self.myZIP.open('docProps/app.xml'))
                elif override[0] == 'application/vnd.openxmlformats-package.core-properties+xml':
                    if sys.version_info < (2, 6):
                        self.OP['docProps/core.xml'] = Core(f=self.myZIP.read('docProps/core.xml'))
                    else:
                        self.OP['docProps/core.xml'] = Core(f=self.myZIP.open('docProps/core.xml'))
                elif override[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml':
                    if sys.version_info < (2, 6):
                        self.OP['xl/workbook.xml'] = Workbook(f=self.myZIP.read('xl/workbook.xml'))
                    else:
                        self.OP['xl/workbook.xml'] = Workbook(f=self.myZIP.open('xl/workbook.xml'))
                elif override[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml':
                    if sys.version_info < (2, 6):
                        self.OP['xl/styles.xml'] = Styles(f=self.myZIP.read('xl/styles.xml'))
                    else:
                        self.OP['xl/styles.xml'] = Styles(f=self.myZIP.open('xl/styles.xml'))
                elif override[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml':
                    if sys.version_info < (2, 6):
                        self.OP['xl/sharedStrings.xml'] = SharedStrings(f=self.myZIP.read('xl/sharedStrings.xml'))
                    else:
                        self.OP['xl/sharedStrings.xml'] = SharedStrings(f=self.myZIP.open('xl/sharedStrings.xml'))
                elif override[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml':
                    if sheets:
                        name = override[1][1::]
                        if sys.version_info < (2, 6):
                            self.OP[name] = Sheet(f=self.myZIP.read(name))
                        else:
                            self.OP[name] = Sheet(f=self.myZIP.open(name))
                elif override[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml':
                    name = override[1][1::]
                    if sys.version_info < (2, 6):
                        self.OP[name] = Table(f=self.myZIP.read(name))
                    else:
                        self.OP[name] = Table(f=self.myZIP.open(name))
            for i in self.myZIP.infolist():
                if i.filename.endswith('.rels'):
 #                   print 'Open: %s' %i.filename
                    if sys.version_info < (2, 6):
                        self.OP[i.filename] = Relationships(f=self.myZIP.read(i.filename))
                    else:
                        self.OP[i.filename] = Relationships(f=self.myZIP.open(i.filename))

    def writeLine(self,ref,line):
        if type(ref) != Ref:
            ref = Ref(ref)
        self.activeSheet.writeLine(ref,line)

    def listSheets(self):
        return self.OP['xl/workbook.xml'].listSheets()
    
class Expr(object):
    """Class to represent a expression(formula)"""
    def __init__(self,expr,extra=None):
        self.expr = expr
        self.extra = extra
        self.parse()

    def parse(self):
        if not self.extra:
            self.str = self.expr
            return
        d = dict()
        for key in self.extra:
            if type(self.extra[key]) == Ref or type(self.extra[key]) == Sqref:
                d[key] = self.extra[key].ref
            elif type(self.extra[key]) == Expr:
                d[key] = self.extra[key].str
            else:
                d[key] = self.extra[key]
        self.str = self.expr %d

    def __str__(self):
        return self.str
