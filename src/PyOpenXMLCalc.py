import re
import time
from xml.dom.minidom import *
from zipfile import *

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

    def get_number(self,text):
        m = re.search('\d+\.?\d*',text)
        if m:
            return m.group(0)
        else:
            return None

    def get_Text(self,text):
        return re.search('[A-Z]+',text).group(0)

    def toxml(self,encoding):
        return self.root.toxml(encoding=encoding)

    def toprettyxml(self,encoding):
        return self.root.toprettyxml(encoding=encoding)

class Content_Types(OP):
    def __init__(self):
        self.root = Document()
        self.types = types = self.root.createElement('Types')
        self.types.setAttribute('xmlns','http://schemas.openxmlformats.org/package/2006/content-types')
        self.root.appendChild(self.types)
        self.new_default('rels')
        self.new_default('xml')

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
    def __init__(self):
        self.id_ = 0
        self.root = Document()
        self.relationships = types = self.root.createElement('Relationships')
        self.relationships.setAttribute('xmlns','http://schemas.openxmlformats.org/package/2006/relationships')
        self.root.appendChild(self.relationships)
        
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

class App(OP):
    def __init__(self,company):
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
    def __init__(self):
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
        self.dimensionRef = Ref('A1')

    def newSheetView(self,tabSelected,workbookViewId):
        sheetView = self.root.createElement('sheetView')
#        sheetView.setAttribute('tabSelected',str(tabSelected))
        sheetView.setAttribute('workbookViewId',str(workbookViewId))
        return sheetView

    def selectedTab(self):
        sheetView = self.sheetViews.getElementsByTagName('sheetView')[0]
        sheetView.setAttribute('tabSelected',"1")

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
        if ref.startRowID in self.rows:
            return self.rows[ref.startRowID]
        else:
            return self.new_row(ref)

    def new_row(self,ref):
        """create a new row and append it on the right position
           * append new Row in the right position
           * search row that comes after the new row
           * insertBefore that row"""
        row = self.root.createElement('row')
#        row.setAttribute('spans','1:1') # <- ignore
        row.setAttribute('r',str(ref.startRowID))
        nextRowID = self.getNextRowID(ref)
        if nextRowID == -1:
            self.sheetData.appendChild(row)
        else:
            self.sheetData.insertBefore(row,self.rows[nextRowID])
        self.rows[ref.startRowID] = row
        return row

    def getNextRowID(self,ref):
        """return the next higher existing row ID"""
        keys = self.rows.keys()
        keys.sort()
        for i in keys:
            if i >= ref.startRowID:
                return i
        return -1
        
        
    def getC(self,ref,row):
        """get c tag from row by ref or create a new"""
        for c in row.getElementsByTagName('c'):
            if c.getAttribute('r') == ref.start:
                return c
        c = self.root.createElement('c')
        c.setAttribute('r',ref.start)
        row.appendChild(c)
        return c


    def writeLine(self,ref,line):
        """write a line(list) to Sheet at ref"""
        ref2 = Ref(ref.ref)
        for cell in line:
            self.write(ref2,cell)
            ref2.walk('right')

    def write(self,ref,text):
        self.dimensionRef.max(ref)
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

    def get_c4ref(self,ref):
        rows = self.sheetData.getElementsByTagName('row')
        for row in rows:
            if ref.startRowID == int(row.getAttribute('r')):
                for c in row.getElementsByTagName('c'):
                    if c.getAttribute('r') == ref.start:
                        return c
#        print 'no cell with ref = %s found!' %ref.start
        return None

    def read(self,ref):
        """read cell content"""
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

    def readExpressin(self,f):
        """get expression from cell"""
        return f.firstChild.nodeValue
    
    def getStringFromSharedStings(self,id_):
        """get String from SharedStrings by ID"""
        pass
    
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
        """get the size of the table starting at startRef in both directions, rows and columns."""
        c = self.get_c4ref(ref)
        while self.get_c4ref(Ref(ref.end)):
#        while self.read(ref.end):
            ref.extend('down')
        ref.extend('up')
        while self.get_c4ref(Ref(ref.end)):
#        while self.read(ref.end):
            ref.extend('right')
        ref.extend('left')
        
    def readRow(self,ref):
        """read and return the first row of ref"""
        ref2 = Ref(ref.ref)
        row = list()
        while ref.endCN != ref2.startCN:
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
            format_ = 'LEFT(%s,%s)="%s"' %(sqref.start,len(text),text)
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
    def __init__(self,appName,lastEdited=4,lowestEdited=4,rupBuild=4506,defaultThemeVersion=124226,calcId=125725):
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
        
class Core(OP):
    def __init__(self,creator):
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
    def __init__(self):
        self.countFonts = 0
        self.countFills = 0
        self.countBorders = 0
        self.countCellStyleXfs = 0
        self.countCellXfs = 0
        self.countCellStyles = 0
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
    def __init__(self):
        self.count = 0
        self.uniqueCount = 0
        self.root = Document()
        self.sst = self.root.createElement('sst')
        self.sst.setAttribute('xmlns','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        self.sst.setAttribute('count',str(self.count))
        self.sst.setAttribute('uniqueCount',str(self.uniqueCount))
        self.root.appendChild(self.sst)
        self.length = 0 # number of strings in SharedStrings.xml

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
    
    def getStringFromSharedStings(self,id_):
        si = self.sst.getElementsByTagName('si')[id_]
        text = ''
        for t in si.getElementsByTagName('t'):
            text+=t.firstChild.nodeValue
        return text

class Ref(OP):
    def __init__(self,ref):
        """ref format:
        A1:C4
        A3    --> A3:A3
        """
        self._startF = False
        self._endF = False
        self.setRef(ref)

    def getStart(self):
        if self._startF:
            return '$%s%s' %(self._startCN,self._startRowID)
        else:
            return '%s%s' %(self._startCN,self._startRowID)
    def setStart(self,value):
        if value.startswith('$'):
            self._startF = True
            value = value.replace('$','',1)
        else:
            self._startF = False
        self._startRowID = int(self.get_number(value))
        self._startCN = self.get_text(value)
    start = property(getStart,setStart)
    def getEnd(self):
        if self._endF:
            return '$%s%s' %(self._endCN,self._endRowID)
        else:
            return '%s%s' %(self._endCN,self._endRowID)
    def setEnd(self,value):
        if value.startswith('$'):
            self._endF = True
            value = value.replace('$','',1)
        else:
           self._endF = False
        self._endRowID = int(self.get_number(value))
        self._endCN = self.get_text(value)
    end = property(getEnd,setEnd)
    def getRef(self):
        return '%s%s:%s%s' %(self._startCN,self._startRowID,self._endCN,self._endRowID)
    def setRef(self,value):
        ref = value.split(':')
        self.setStart(ref[0])
        if len(ref) == 1:
            self.setEnd(self.getStart())
        elif len(ref) == 2:
            self.setEnd(ref[1])
    ref = property(getRef,setRef)
    def getStartRowID(self):
        return self._startRowID
    def setStartRowID(self,value):
        if value <= 0:
            self._startRowID = 1
        else:
            self._startRowID = int(value)
    startRowID = property(getStartRowID,setStartRowID)
    def getEndRowID(self):
        return str(self._endRowID)
    def setEndRowID(self,value):
        if value <= 0:
            self._endRowID = 1
        else:
            self._endRowID = int(value)
    endRowID = property(getEndRowID,setEndRowID)
    def getStartCN(self):
        return self._startCN
    def setStartCN(self,value):
        if value.startswith('$'):
            self._startF = True
        else:
            self._startF = False
        self._startCN = value
    startCN = property(getStartCN,setStartCN)
    def getEndCN(self):
        return self._endCN
    def setEndCN(self,value):
        if value.startswith('$'):
            self._endF = True
        else:
            self._endF = False
        self.endCN = value
    endCN = property(getEndCN,setEndCN)
            
    def get_number(self,ref):
        return re.search('\d+\.?\d*',ref).group(0)

    def get_text(self,ref):
        return ref.replace(self.get_number(ref),'')

    def count_rows(self):
        return self.endRowID - self.startRowID +1
    
    def count_cols(self):
        return ord(self.endCN) -ord(self.startCN) +1

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
    
    def appendColumns(self,cols):
        """self.endCN+=cols"""
        endCN = list(self._endCN)
        i = 0
        while i < len(endCN):
            if endCN[i] == 'Z':
                if i == len(endCN)-1:
                    if cols > 26:
                        cols = cols-26
                        endCN.append('Z')
                        self._endCN = ''.join(endCN)
                        return self.appendColumns(cols)
                    else:
                        endCN.append(chr(cols+64))
                        self._endCN = ''.join(endCN)
                        break
                else:
                    i+=1
                    continue
            elif ord(endCN[i])+cols > 90:
                cols = cols - (90-ord(endCN[i]))
                endCN[i] = 'Z'
                self._endCN = ''.join(endCN)
                self.appendColumns(cols)
                break
            else:
                endCN[i] = chr(ord(endCN[i])+cols)
                self._endCN = ''.join(endCN)
                break
            i+=1

    def extend(self,d):
        """walk with ref.end"""
        if d == 'right':
            self._endCN = self.incCol(self.getEndCN())
        elif d == 'left':
            self._endCN = self.decCol(self.getEndCN())
        elif d == 'up':
            self._endRowID -=1
        elif d == 'down':
            self._endRowID +=1

    def walk(self,d):
        """move the Ref one cell in a direction(up,down,left,right)"""
        if d == 'right':
            self._startCN = self.incCol(self.getStartCN())
        elif d == 'left':
            self._startCN = self.decCol(self.getStartCN())
        elif d == 'up':
            self._startRowID -=1
        elif d == 'down':
            self._startRowID +=1

    def getInt4CN(self,CN):
        """translate CN in integer"""
        l = list(CN)
        l.reverse()
        all = 0
        for i in range(len(l)):
            all+=ord(l[i])*(i+1)
        return all

    def compCN(self,CN1,CN2):
        """compare self.startCN to CN
        CN > CN2 """
        i1 = self.getInt4CN(CN2)
        i2 = self.getInt4CN(CN1)
        if i2 > i1:
            return True
        else:
            return False

    def max(self,ref):
        """if ref > self then update self.startRowID/self.startCN"""
        if self.compCN(ref.startCN,self.endCN):
            self._endCN = ref.startCN
        if ref.startRowID > self.endRowID:
            self._endRowID = ref.startRowID
    
class Table(OP):
    def __init__(self,id_,name,ref,header,displayName=None,totalsRowShown=0,tableStyle='TableStyleLight16'):
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
        self.table.setAttribute('ref',ref.ref)
        self.table.setAttribute('totalsRowShown',str(totalsRowShown))
        self.new_table(ref,name,header)

    def new_table(self,ref,name,header):
        """
        ref: Ref # Start cell
        name : str # name of table
        cols : list # name of columns (header)
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
        self.count = ref.count_cols()
        self.autoFilter = self.root.createElement('autoFilter')
        self.autoFilter.setAttribute('ref',ref.ref)
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
    def __init__(self,company,userName):
        """create a new blank OpenXML Calc workbook without any sheets
        docProps/app.xml <-- new
        [Content_Types].xml <-- extend
        _rels/.rels <-- extend
        """
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
        self.OP['xl/worksheets/sheet1.xml'] = Sheet()
#        self.new_Table()
#        self.OP['xl/worksheets/sheet2.xml'] = self.new_sheet()
#        self.OP['xl/worksheets/sheet3.xml'] = self.new_sheet()
#        self.rowID = 0
#        self.activeSheet = self.OP['xl/worksheets/sheet1.xml']
        self.OP['xl/styles.xml'].set_rId(self.OP['xl/_rels/workbook.xml.rels'].new_relationship('styles.xml'))
        self.OP['[Content_Types].xml'].new_Override('/xl/styles.xml')
        self.OP['xl/sharedStrings.xml'].set_rId(self.OP['xl/_rels/workbook.xml.rels'].new_relationship('sharedStrings.xml'))
        self.OP['[Content_Types].xml'].new_Override('/xl/sharedStrings.xml')
        self.tables = 0

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
        return self.activeSheet.read(ref)
        
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
        """select a sheet by name"""
        self.OP['xl/workbook.xml'].set_activeTab(sheetName)

    def hideColume(self,min_,max_):
        """hide a column by columnID"""
        self.activeSheet.hideColume(min_,max_)

    def formatTable(self,ref,name,tableStyle=None):
        """create a Table with default formating settings
        xl/printerSettings/printerSettings1.bin << new but ignore
        xl/tables/table1.xml << new 1.
        xl/worksheets/sheet1.xml << 2.
            <pageSetup paperSize="9" orientation="portrait" r:id="rId1"/> # <- ignore
            <tableParts count="1">
                <tablePart r:id="rId2"/>
            </tableParts>
        xl/worksheets/_rels/sheet1.xml.rels << new 3.
        [Content_Types].xml << 4.
            <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>  # <- ignore
            <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
        """
        if type(ref) == str:
            ref = Ref(ref)
        self.activeSheet.getTableSice(ref)
        self.tables+=1 # counter for tables, needed for id of the new table
        # id_,name,ref,header
        header = self.activeSheet.readRow(ref)
        self.OP['xl/tables/table%s.xml' %self.tables] = Table(self.tables,name,ref,header,tableStyle=tableStyle) # 1.
        self.OP['[Content_Types].xml'].new_Override('/xl/tables/table%s.xml' %self.tables) # 4.
        self.OP['xl/worksheets/_rels/sheet1.xml.rels'] = Relationships() # 3.
        rId = self.OP['xl/worksheets/_rels/sheet1.xml.rels'].new_relationship('../tables/table%s.xml' %self.tables) # 3.1
        self.activeSheet.addTablePart(rId) # 2.
        return ref

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

class Expr(object):
    """Class to represent a expression(formula)"""
    def __init__(self,expr,extra=None):
        self.expr = expr
        self.extra = extra
        self.parse()

    def parse(self):
        d = dict()
        for key in self.extra:
            if type(self.extra[key]) == Ref:
                d[key] = self.extra[key].start
            elif type(self.extra[key]) == Expr:
                d[key] = self.extra[key].str
            else:
                d[key] = self.extra[key]
        self.str = self.expr %d
    def __str__(self):
        return self.str
