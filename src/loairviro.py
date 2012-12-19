# -*- coding: iso-8859-15 -*-
import uno
import unohelper
from com.sun.star.task import XJobExecutor
import pdb

from pyAirviro.edb.edb import Edb
from pyAirviro.edb.sourcedb import Source,Sourcedb,SubstEmis,SubgrpEmis
from pyAirviro.edb.sourcedb import ActivityEmis
from pyAirviro.edb.subdb import Subdb
from pyAirviro.edb.subgrpdb import Subgrpdb
from pyAirviro.edb.emfacdb import Emfacdb

class LoadEdb(XJobExecutor,unohelper.Base):
    """
    Load an Airviro edb into a number of spreadsheets
    """
    def __init__(self,ctx):
        self.ctx = ctx
        self.desktop=self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx )
        self.doc=self.desktop.getCurrentComponent()
        self.sheets={"sources":None,"substance groups":None,
                     "emfac":None,"substances":None, "units":None}
        self.edb=None
        self.subdb=None
        self.subgrpdb=None

    def setEdb(self,domainName,userName,edbName):
        self.edb=Edb(domainName,userName,edbName)

    def loadSubdb(self):
        self.sheets["substances"] = self.createSheet("substances")
        self.subdb=Subdb(self.edb)
        self.subdb.readSubstances(
            filename="/local_disk/dvlp/airviro/loairviro/test/substances.out",
            keepEmpty=True)        
        out=(("Index","Substance name"),)
        indices=sorted(self.subdb.substNames.keys())
        for i,ind in enumerate(indices):
            substName=self.subdb.substNames[ind]
            out+=((ind,substName),)

        cellRange=self.sheets[
            "substances"].getCellRangeByPosition(0,0,1,len(out)-1)
        cellRange.setDataArray(out)

    def loadSubgrpdb(self):
        self.sheets["substance groups"] = self.createSheet("substance groups")
        self.subgrpdb=Subgrpdb(self.edb)
        self.subgrpdb.read(
            filename="/local_disk/dvlp/airviro/loairviro/test/subgrp.out")
        
        substances=[]
        for subgrpInd,subgrp in self.subgrpdb.subgrps.items():
            for subst in subgrp.substances:
                substName=self.subdb.substNames[subst]
                if substName not in substances:
                    substances.append(substName)

        header1=("","")
        header2=("Index","Name")
        for substName in substances:
            header1+=(substName,"","")
            header2+=("Slope","Offset","Unit")

        out=(header1,header2)
        for subgrpInd,subgrp in self.subgrpdb.subgrps.items():
            #list names of substances in subgrp
            substInSubgrp=[self.subdb[s] for s in subgrp.substances]
            row=(subgrp.index,subgrp.name)
            for substName in substances:
                if substName in substInSubgrp:
                    substIndex=self.subdb.substIndices[substName]
                    slope=subgrp.substances[substIndex]["slope"]
                    offset=subgrp.substances[substIndex]["offset"]
                    unit=subgrp.substances[substIndex]["unit"]
                else:
                    slope=""
                    offset=""
                    unit=""
                row+=(slope,offset,unit)
            out+=(row,)
                    
        cellRange=self.sheets[
            "substance groups"].getCellRangeByPosition(0,0,
                                                       len(header1)-1,len(out)-1)
        cellRange.setDataArray(out)

    def loadEmfacdb(self):
        self.sheets["emission factors"] = self.createSheet("emission factors")
        self.emfacdb=Emfacdb(self.edb)
        self.emfacdb.readAscii(
            filename="/local_disk/dvlp/airviro/loairviro/test/emfac.out")
        pdb.set_trace()
        nvars=[len(emfac.vars) for emfac in self.emfacdb.activities.values()]
        nvars=max(nvars)

        firstRow=0
        for emfacInd,emfac in self.emfacdb.activities.items():
            out=[]
            out.append(("Index:",emfacInd))
            out.append(("Name:",emfac.name))
            out.append(("Substance:",self.subdb[emfac.subst]))
            varIndices=[]
            varNames=[]
            varTypes=[]
            formula=emfac.formula
            for i in range(1,len(emfac.vars)+1):
                formula.replace("X%i" %i,emfac.vars[i].name)
                varIndices.append(str(i))
                varNames.append(emfac.vars[i].name)
                varTypes.append(emfac.vars[i].type)
            out.append(("Formula",formula))
            out.append(("Variable index:",)+tuple(varIndices))
            out.append(("Variable names:",)+tuple(varNames))
            out.append(("Variable types:",)+tuple(varTypes))
            out.append(("",))
            out.append(("",))

            lastRow=firstRow+len(out)-1
            lastCol=len(varIndices)

            #Make so all rows have the same length
            #Needed to use setDataArray
            for i in range(len(out)):
                completeRow=("",)*(lastCol-(len(out[i])-1))
                out[i]+=completeRow

            cellRange=self.sheets[
                "emission factors"].getCellRangeByPosition(0,firstRow,
                                                           lastCol,lastRow)
            firstRow=lastRow+1
            cellRange.setDataArray(tuple(out))
        
                
    def createSheet(self,name):
        try:
            sheets = self.doc.getSheets()
        except Exception:
            raise TypeError("Model retrived was not a spreadsheet")

        #TODO: warning dialogue before removing sheet
        if sheets.hasByName(name):
            sheets.removeByName(name)
            
        pos = sheets.getCount()
        sheets.insertNewByName(name, pos)
        return sheets.getByIndex(pos)
        
    def loadSources(self):        
        self.sheets["sources"] = self.createSheet("sources")
        sourcedb=Sourcedb(self.edb)
        #Read one source at a time
        sourcedb._batch=1
        filename="/local_disk/dvlp/airviro/loairviro/test/TR_ships.out"

        substEmis={}
        substAlob={}
        subgrpEmis={}
        subgrpAlob={}
        activityEmis={}
        activityAlob={}
        srcAlobs={}

        #Reading sources
        batchInd=0
        while sourcedb.readBatch(filename=filename,accumulate=True):
            print "Read batch %i" %batchInd
            batchInd+=1
        #To present sources in a table with a nice header, it is necessary
        #to list all substances, alobs, subgroups, emfacs and variables that are
        #used in the edb. This is because each alob, substance etc. should be
        #shown only once in the header
        for src in sourcedb.sources:
            #Accumulate all alobs into a list
            for alob in src.ALOBOrder:
                srcAlobs[alob]=None

            #store all substances and their alobs in a dict
            for substInd,emis in src.subst_emis:
                if substInd not in substEmis:
                    substEmis[substInd]={"alob":{}}
                for alob in emis.ALOBOrder:
                    substEmis[substInd]["alob"][alob]=None

            #store all substance groups and their alobs in a dict
            for subgrpInd,emis in src.subgrp_emis.items():
                if subgrpInd not in subgrpEmis:
                    subgrpEmis[subgrpInd]={"alob":{}}
                for alob in emis.ALOBOrder:
                    subgrpEmis[subgrpInd]["alob"][alob]=None

            #Accumulate all activities and included alobs
            for emfacInd,emis in src.activity_emis.items():
                if emfacInd not in activityEmis:
                    activityEmis[emfacInd]={"alob":{},"var":{}}
                for varInd,varVal in emis["VARLIST"]:
                    #vars should also be indexed
                    activityEmis[emfacInd]["var"][varInd]=None
                for alob in emis.ALOBOrder:
                    activityEmis[emfacInd]["alob"][alob]=None

        #Writing header
        header0=()
        header=()
        srcAlobInd={}

        for parName in src.parOrder:
            header+=(parName,)
        header0+=("Static parameters",)
        header0+=(len(header)-len(header0))*("",)
        alobKeys=srcAlobs.keys()
        alobKeys.sort()
        for alobKey in alobKeys:
            header+=(alobKey,)
        if src["ALOB"]>0:
            header0+=("ALOBs",)
            header0+=(len(header)-len(header0))*("",)

        for substInd in substEmis:
            substName=self.subdb.substNames[substInd]
            header+=("Emission","Time variation","Unit","Macro","Activity code")
            header0+=("Substance:",substName)
            header0+=(len(header)-len(header0))*("",)
            alobKeys=substEmis[substInd]["alob"].keys()
            alobKeys.sort()
            for alobKey in alobKeys:
                row+=[src.ALOB.get(alobKey,"")]

            for alob in alobKeys:
                header+=(alob,)
            if len(alobKeys)>0:
                header0+=("ALOBs",)
                header0+=(len(header)-len(header0))*("",)

        for subgrpInd in subgrpEmis:
            subgrp=self.subgrpdb.subgrps[subgrpInd]
            header+=("Activity","Time variation","Unit","Activity code")
            header0+=("Substance group:",subgrp.name)
            header0+=(len(header)-len(header0))*("",)

            alobKeys=subgrpEmis[subgrpInd]["alob"].keys()
            alobKeys.sort()
            for alobKey in alobKeys:
                header+=(alobKey,)
            if subgrpEmis[subgrpInd]["alob"]>0:
                header0+=("ALOBs",)
                header0+=(len(header)-len(header0))*("",)

                
        for emfacInd in activityEmis:
            emfac=self.emfacdb[emfacInd]
            header+=("Time variation",)
            for varInd,var in emfac.vars.items():
                header+=(var.name,)
            header0+=("Emfac:",emfac.name)
            header0+=(len(header)-len(header0))*("",)
                
            header+=("Activity code",)
            alobKeys=activityEmis[emfacInd]["alob"].keys()
            alobKeys.sort()
            for alobKey in alobKeys:
                header+=(alobKey,)
            if len(alobKeys):
                header0+=("ALOBs",)
                header0+=(len(header)-len(header0))*("",)
        header0+=(len(header)-len(header0))*("",)

        firstCol=0
        for colInd,val in enumerate(header0[1:]):
            if val !="":
                bottomCellRange=self.sheets[
                    "sources"].getCellRangeByPosition(firstCol,1,colInd-2,1)
                
                bottomBorder = bottomCellRange.BottomBorder
                bottomBorder.OuterLineWidth = 30
                bottomCellRange.BottomBorder = bottomBorder
                firstCol=colInd-2
                
        out=[header0,header]
        for src in sourcedb.sources:
            row=[]
            for par in src.parOrder:
                row+=[unicode(src[par])]

            #Write alobs for sources
            alobKeys=srcAlobs.keys()
            alobKeys.sort()
            for alobKey in alobKeys:
                row+=[src.ALOB.get(alobKey,"")]

            #write substance emissions with alobs
            for substInd in substEmis:
                alobKeys=substEmis[substInd]["alob"].keys()
                alobKeys.sort()
                if substInd in src.subst_emis:
                    emis=src.subst_emis[substInd]
                    row+=[emis["EMISSION"],
                          emis["TIMEVAR"],
                          emis["UNIT"],
                          emis["MACRO"],
                          emis["ACTCODE"]]
                    for alobKey in alobKeys:
                        row+=[emis.ALOB.get(alobKey,"")]
                else:
                    row+=["","","","",""] #empty cells for substance
                    row+=[""]*len(alobKeys)

            #write substance group emissions with alobs
            for subgrpInd in subgrpEmis:
                alobKeys=subgrpEmis[subgrpInd]["alob"].keys()
                alobKeys.sort()
                if subgrpInd in src.subst_emis:
                    emis=src.subgrp_emis[subgrpInd]
                    row+=[emis["ACTIVITY"],
                          emis["TIMEVAR"],
                          emis["UNIT"],
                          emis["ACTCODE"]]
                    for alobKey in alobKeys:
                        row+=[emis.ALOB.get(alobKey,"")]
                else:
                    row+=["","","",""] #empty cells for substance group
                    row+=[""]*len(alobKeys)

            #write emfac emissions with variables and alobs
            for emfacInd in activityEmis:
                alobKeys=activityEmis[emfacInd]["alob"].keys()
                alobKeys.sort()
                varKeys=activityEmis[emfacInd]["var"].keys()
                varKeys.sort()
                if emfacInd in src.activity_emis:
                    emis=src.activity_emis[emfacInd]
                    row+=[emis["TIMEVAR"]]
                    varVals = [var[1] for var in emis["VARLIST"]]
                    row+=varVals
                    row+=[emis["ACTCODE"]]
                    for alobKey in alobKeys:
                        row+=[emis.ALOB.get(alobKey,"")]
                else:
                    row+=2*[""]+[""]*len(varKeys)+[""]*len(alobKeys)
            out.append(tuple(row))
        cellRange=self.sheets[
            "sources"].getCellRangeByPosition(0,0,len(header)-1,len(out)-1)
        cellRange.setDataArray(tuple(out))
        
    def trigger(self,args=''):
        """Called by addon's UI controls or service:URL"""
        self.setEdb("shipair","airviro","TR_2011_v3")
        print "EDB set"
        self.loadSubdb()
        print "Loaded subdb"
        self.loadSubgrpdb()
        print "Loaded subgrpdb"
        self.loadEmfacdb()
        print "Loaded emfacdb"
        self.loadSources()
        print "Loaded sources"
        


## addon-implementation:
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(LoadEdb,
                        "se.smhi.airviro.loairviro.LoadEdb",
                        ("com.sun.star.task.Job",),)

## call listening office:
if __name__ == "__main__":
    print "Testing"
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    button = LoadEdb(ctx)
    button.trigger()

    #    oView = getCurrentController(ctx)
    ##     sh = getCurrentController(ctx).getActiveSheet()
    ##     for c in range(0,255,2):
    ##         sh.getColumns().getByIndex(c).IsVisible=False
    ##     v = sh.queryVisibleCells()
    ##     print v.getCount(), v.getByIndex(0).getRangeAddress(), v.getByIndex(1).getRangeAddress()
    #from pyXray import XrayBox
    #h = ConfigProvider(ctx,)
    #p= h.HelpPath
    #print p
    #f = h.getHelpURL('SpecialCells.sxw')
    #print f   
    #h.trigger(f +'#_DialogCellFormatRanges')
    #D1 = DialogCellFormatRanges(ctx)
    #D1.trigger(())
    #D2 = DialogCellContents(ctx)
    #D2.trigger(())

