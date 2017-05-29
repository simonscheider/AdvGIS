#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      simon
#
# Created:     30/05/2016
# Copyright:   (c) simon 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import csv
import os.path



class dataf():
    def __init__(self, spath="", file="MOT18.csv", exit1= 'ASSEN01', exit2 = "MARUM01"):
        self.path = spath
        self.descpath = os.path.normpath(os.path.join(spath,file))
        self.scriptpath = ''#os.path.normpath(os.path.join(spath,"MOT_C1.flg"))
        self.exit1 = exit1
        self.exit2 = exit2
    pointfile = ""
    exitnamefield =""
    roadfile = ""
    roadfield = ''
    numalt = 0
    numfields = 0
    alternativenames ={}
    speed = ''
    access =""
    routenamefield =""
    alternativetimefields =[]
    TTCurrentForth = ""
    TTCurrentBack = ""
    TTAltForth = []
    TTAltBack = []
    grp = ''

    allf = [pointfile,    exitnamefield,    roadfile,    roadfield,    numalt,    alternativenames,    speed,    access, routenamefield,    alternativetimefields,TTCurrentForth,TTCurrentBack,TTAltForth,TTAltBack,grp]

    def __str__(self):
        s =""
        for i in self.allf:
            s= s+(str(i)+' ')
        return s

    #table = pself.read_excel('sales.xlsx'
    def dequote(self,s):
        """
        If a string has single or double quotes around it, remove them.
        Make sure the pair of quotes match.
        If a matching pair of quotes is not found, return the string unchanged.
        """
        if (s[0] == s[-1]) and s.startswith(("'", '"')):
            return s[1:-1]
        return s

    def read(self):
        with open(self.descpath) as csvfile:
            reader = csv.reader(csvfile, delimiter=',')
            for row in reader:
                print row
                if row[0].find('Input Data Folder')!=-1:
                    self.grp =(row[1])[-2:]
                    print self.grp
                if row[0].find('Name Shape File containing BeginPoint, EndPoint and all MidPoints ')!=-1:
                    self.pointfile = row[1]
                    print self.pointfile
                elif row[0].find('Fieldname Unique Identifier in Points shape file')!=-1:
                    self.exitnamefield =row[1]
                    print self.exitnamefield
                elif row[0].find('Name Shape File containing integrated roadnetwork (current plus alternatives)')!=-1:
                    self.roadfile = row[1]
                    print self.roadfile
                elif row[0].find('Fieldname Unique Identifier in integrated road shape file')!=-1:
                    self.roadfield = row[1]
                    print self.roadfile
                elif row[0].find('Number of alternative routes (incl straight line)')!=-1:
                    if (row[1].find('.')==-1):
                        self.numalt = int(row[1])
                    else:
                        self.numalt = float(row[1])
                    self.numfields = 5+(self.numalt*2)
                    print str(self.numalt) +' '+str(self.numfields)
                elif row[0].find('Reference Name of alternative')!=-1:
                    self.alternativenames[(row[0])[((row[0].find("alternative "))+12)]] = row[1]
                    print str(self.alternativenames)
                elif row[0].find('Travel speed by car in kmph for all roads')!=-1:
                    self.speed = self.dequote(row[1])
                    print self.dequote(row[1])
                elif row[0].find('Midway access to all road segments')!=-1:
                    self.access = self.dequote(row[1])
                    print self.dequote(row[1])
                elif row[0].find('Reference name current or alternative road')!=-1:
                    self.routenamefield = self.dequote(row[1])
                    print self.dequote(row[1])
                elif row[0].find('Travel time Forth in current situation only')!=-1:
                    self.TTCurrentForth = self.dequote(row[1])
                    print self.dequote(row[1])
                elif row[0].find('Travel time Back in current situation only')!=-1:
                    self.TTCurrentBack = self.dequote(row[1])
                    print self.dequote(row[1])
                elif row[0].find('Travel time Forth')!=-1:
                    self.TTAltForth.append(self.dequote(row[1]))
                    print self.dequote(row[1])
                elif row[0].find('Travel time Back')!=-1:
                    self.TTAltBack.append(self.dequote(row[1]))
                    print self.dequote(row[1])
        try:
            if self.numalt >4 or self.numalt <3:
                raise ValueError("Too many or few alternatives!!")
        except ValueError as err:
            print(err.args)
            return

    def replaceall(self, script = "MOT_C1.flg"):
        #load the script template (local file)
        if script == 'MOT_C2b.flg':
            if (self.numalt== 3):
                scripttemplate = os.path.splitext(script)[0]+'_origin'+os.path.splitext(script)[1]
                f1=open(scripttemplate,'r')
            elif (self.numalt== 4):
                scripttemplate = os.path.splitext(script)[0]+'_origin4'+os.path.splitext(script)[1]
                f1=open(scripttemplate,'r')
        else:
            scripttemplate = os.path.splitext(script)[0]+'_origin'+os.path.splitext(script)[1]
            f1=open(scripttemplate,'r')

        filedata = f1.read()
        f1.close()
        filedata = filedata.replace("D:\Temp\Nordland",self.path)
        #filedata = filedata.replace("e:\data\nordland","H:\nordland")
        filedata = filedata.replace("PointsXX.shp",self.pointfile)
        filedata = filedata.replace("LABELYY",self.exitnamefield)
        filedata = filedata.replace("RoadsXX.shp",self.roadfile)
        filedata = filedata.replace("RoadsXX.dbf",self.roadfile[0:-3]+'dbf')
        filedata = filedata.replace("BN2000_YY",self.roadfield)
        filedata = filedata.replace("XX",self.grp)
        #filedata = filedata.replace("11YY",str(self.numfields))

        import StringIO
        filedata = StringIO.StringIO(filedata)

        #write to new script file at path location
        self.scriptpath= os.path.normpath(os.path.join(self.path,script))
        print "Write script: "+str(self.scriptpath)
        f2=open(self.scriptpath,'w')
        delete = False
        det = False
        line = filedata.readline()
        while line:
            if ( script =="MOT_C2a.flg" ):
                line = line.replace("H:\ADVGIS\Group18",self.path)
                line = line.replace('H:\nordland\MY_ROADS.006','C:\Users\simon\Documents\GitHub\AdvGIS\MY_ROADS.006')
                line = line.replace('POINT18','POINT'+str(self.grp))
                line = line.replace('ROADS18','ROADS'+str(self.grp))
                line = line.replace('ACCESS',self.access)
                line = line.replace("Route", self.routenamefield)
                line = line.replace("Group 18", "Group "+str(self.grp))
                line = line.replace("G18_2a1.jpg", "G18_2a1.jpg".replace('18',self.grp))
                line = line.replace("G18_2a2.jpg", "G18_2a2.jpg".replace('18',self.grp))
                line = line.replace('TIME_FORTH',self.TTCurrentForth)
                line = line.replace('TIME_BACK',self.TTCurrentBack)
                f2.write(line)

            if( script =="MOT_C2b.flg" ):
                if (self.numalt==3):
                    line = line.replace('POINT18','POINT'+str(self.grp))
                    line = line.replace('ROADS18','ROADS'+str(self.grp))
                    line = line.replace("Group 18", "Group "+str(self.grp))
                    line = line.replace("H:\ADVGIS\Group18",self.path)
                    line = line.replace('LABEL',self.exitnamefield)
                    line = line.replace('ACCESS',self.access)
                    line = line.replace('TIME_FORTH',self.TTCurrentForth)
                    line = line.replace('TIME_BACK',self.TTCurrentBack)
                    line = line.replace('ASSEN01',self.exit1)
                    line = line.replace('MARUM01',self.exit2)
                    if (line.find('ECON_FORTH')!=-1):
                        line = line.replace('ECON_FORTH',self.TTAltForth[0])
                    if (line.find('ECON_BACK')!=-1):
                        line= line.replace('ECON_BACK',self.TTAltBack[0])
                    if (line.find('NOIS_FORTH')!=-1):
                        line = line.replace('NOIS_FORTH',self.TTAltForth[1])
                    if (line.find('NOIS_BACK')!=-1):
                        line= line.replace('NOIS_BACK',self.TTAltBack[1])
                    if (line.find('TOUR_FORTH')!=-1):
                        line = line.replace('TOUR_FORTH',self.TTAltForth[2])
                    if (line.find('TOUR_BACK')!=-1):
                        line= line.replace('TOUR_BACK',self.TTAltBack[2])

                if (self.numalt==4):
                    line = line.replace("U:\ADVGIS\Route22",self.path)
                    line = line.replace('POINT22','POINT'+str(self.grp))
                    line = line.replace('ROADS22','ROADS'+str(self.grp))
                    line = line.replace("Group 22", "Group "+str(self.grp))
                    line = line.replace('LABEL',self.exitnamefield)
                    line = line.replace('ACCESS',self.access)
                    line = line.replace('TIME_FORTH',self.TTCurrentForth)
                    line = line.replace('TIME_BACK',self.TTCurrentBack)
                    line = line.replace('Dokkum02',self.exit1)
                    line = line.replace('Nyega',self.exit2)
                    if (line.find('NAT_FORTH')!=-1):
                        line = line.replace('NAT_FORTH',self.TTAltForth[0])
                    if (line.find('NAT_BACK')!=-1):
                        line= line.replace('NAT_BACK',self.TTAltBack[0])
                    if (line.find('MONU_FORTH')!=-1):
                        line = line.replace('MONU_FORTH',self.TTAltForth[1])
                    if (line.find('MONU_BACK')!=-1):
                        line= line.replace('MONU_BACK',self.TTAltBack[1])
                    if (line.find('TOUR_FORTH')!=-1):
                        line = line.replace('TOUR_FORTH',self.TTAltForth[2])
                    if (line.find('TOUR_BACK')!=-1):
                        line= line.replace('TOUR_BACK',self.TTAltBack[2])
                    if (line.find('TOUR2_FORTH')!=-1):
                        line = line.replace('TOUR2_FORTH',self.TTAltForth[3])
                    if (line.find('TOUR2_BACK')!=-1):
                        line= line.replace('TOUR2_BACK',self.TTAltBack[3])
                f2.write(line)

            if ( script =="MOT_C1.flg" ):
                if "11YY" in line:
                    line= line.replace("11YY",str(self.numfields))
                #filedata = filedata.replace("11YY",str(self.numfields))
                if  "# END COPY FIELDS SECTION" in line:
                    delete = False
                    det = False
                if delete == True and det == False:
                    f2.write( self.speed+"\n")
                    f2.write( self.TTCurrentForth+"\n")
                    f2.write( self.TTCurrentBack+"\n")
                    f2.write( self.access+"\n")
                    f2.write( self.routenamefield+"\n")
                    c = 0
                    for a in self.TTAltForth:
                        f2.write( a+"\n")
                        f2.write( self.TTAltBack[c]+"\n")
                        c= c+1
                    det = True
                elif delete and det:
                    pass
                else:
                    f2.write(line)
                if "# LIST OF SELECTED FIELDS:" in line:
                    delete = True

            line = filedata.readline()

        f2.close()

import xlrd
import csv

def xls_to_csv(workbook= 'data.xls'):
    x =  xlrd.open_workbook(workbook)
    x1 = x.sheet_by_index(0)
    csvfile = open(os.path.splitext(workbook)[0]+'.csv', 'wb')
    writecsv = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL)

    for rownum in xrange(x1.nrows):
        writecsv.writerow(x1.row_values(rownum))
    csvfile.close()

def main():
    #xls_to_csv(workbook="P:\AdvGIS\Group10\MOT10.xls")
    #d = dataf(spath='P:\AdvGIS\Group10', file='MOT10.csv', exit1 = "STARTDOKKUM", exit2= "ENDMARUM")
    xls_to_csv(workbook="P:\AdvGIS\Group19\MOTxx.xls")
    d = dataf(spath='P:\AdvGIS\Group19', file='MOTxx.csv', exit1 = "DOKKUMEXIT", exit2= "GRONINGENEXIT")

    d.read()
    d.replaceall(script = "MOT_C1.flg")
    d.replaceall(script = "MOT_C2a.flg")
    d.replaceall(script = "MOT_C2b.flg")
    #print d

if __name__ == '__main__':
    main()
