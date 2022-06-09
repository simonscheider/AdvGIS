#-------------------------------------------------------------------------------
# Name:        preparescript
# Purpose:  Test the correctness of roadfiles for the course Advanced GIS (U Utrecht)
#
# Author:      simon scheider
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
                print(row)
                if row != []:
                    if row[0].lower().find('Input Data Folder'.lower())!=-1:
                        self.grp =(row[1])[-2:]
                        print(self.grp)
                    if row[0].lower().find('Name Shape File containing BeginPoint, EndPoint and all MidPoints'.lower())!=-1:
                        self.pointfile = row[1]
                        print(self.pointfile)
                    elif row[0].lower().find('Fieldname Unique Identifier in Points shape file'.lower())!=-1:
                        self.exitnamefield =row[1]
                        print(self.exitnamefield)
                    elif row[0].lower().find('Name Shape File containing integrated roadnetwork'.lower())!=-1:
                        self.roadfile = row[1]
                        print(self.roadfile)
                    elif row[0].lower().find('Fieldname Unique Identifier in integrated road shape file'.lower())!=-1:
                        self.roadfield = row[1]
                        print(self.roadfield)
                    elif row[0].lower().find('Number of alternative routes (incl straight line)'.lower())!=-1:
                        if (row[1].find('.')==-1):
                            self.numalt = int(row[1])
                        else:
                            self.numalt = float(row[1])
                        self.numfields = 5+(self.numalt*2)
                        print(str(self.numalt) +' '+str(self.numfields))
                    elif row[0].lower().find('Reference Name of alternative'.lower())!=-1:
                        self.alternativenames[(row[0])[((row[0].find("alternative "))+12)]] = row[1]
                        print(str(self.alternativenames))
                    elif row[0].lower().find('Travel speed by car in kmph for all roads'.lower())!=-1:
                        self.speed = self.dequote(row[1])
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Midway access to all road segments'.lower())!=-1:
                        self.access = self.dequote(row[1])
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Reference name current or alternative road'.lower())!=-1:
                        self.routenamefield = self.dequote(row[1])
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Travel time Forth in current situation only'.lower())!=-1:
                        self.TTCurrentForth = self.dequote(row[1])
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Travel time Back in current situation only'.lower())!=-1:
                        self.TTCurrentBack = self.dequote(row[1])
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Travel time Forth'.lower())!=-1:
                        self.TTAltForth.append(self.dequote(row[1]))
                        print(self.dequote(row[1]))
                    elif row[0].lower().find('Travel time Back'.lower())!=-1:
                        self.TTAltBack.append(self.dequote(row[1]))
                        print(self.dequote(row[1]))
        print("Alternative fieldnames: ")
        print(self.TTAltForth)
        print(self.TTAltBack)
        print("number of alternatives: "+str(self.numalt))
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
        filedata = filedata.replace(r"D:\Temp\Nordland",self.path)
        #filedata = filedata.replace("e:\data\nordland","H:\nordland")
        filedata = filedata.replace("PointsXX.shp",self.pointfile)
        filedata = filedata.replace("LABELYY",self.exitnamefield)
        filedata = filedata.replace("RoadsXX.shp",self.roadfile)
        filedata = filedata.replace("RoadsXX.dbf",self.roadfile[0:-3]+'dbf')
        filedata = filedata.replace("BN2000_YY",self.roadfield)
        filedata = filedata.replace("XX",self.grp)
        #filedata = filedata.replace("11YY",str(self.numfields))

        import io
        filedata = io.StringIO(filedata)

        #write to new script file at path location
        self.scriptpath= os.path.normpath(os.path.join(self.path,script))
        print("Write script: "+str(self.scriptpath))
        f2=open(self.scriptpath,'w')
        delete = False
        det = False
        line = filedata.readline()
        while line:
            if ( script =="MOT_C2a.flg" ):
                line = line.replace(r"H:\ADVGIS\Group18",self.path)
                line = line.replace(r'C:\Users\simon\Documents\GitHub\AdvGIS\MY_ROADS.006',r'C:\Users\schei008\.matplotlib\Documents\github\AdvGIS\MY_ROADS.006')
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
                    line = line.replace('ROADs18','ROADS'+str(self.grp))
                    line = line.replace("Group 18", "Group"+str(self.grp))
                    line = line.replace("H:\ADVGIS\Group18",self.path)
                    #line = line.replace('LABEL',self.exitnamefield)
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
                    line = line.replace('ROADs22','ROADS'+str(self.grp))
                    line = line.replace("Group 22", "Group "+str(self.grp))
                    #line = line.replace('LABEL',self.exitnamefield)
                    line = line.replace('ACCESS',self.access)
                    line = line.replace('TIME_FORTH',self.TTCurrentForth)
                    line = line.replace('TIME_BACK',self.TTCurrentBack)
                    line = line.replace('Dokkum02',self.exit1)
                    line = line.replace('Nyega',self.exit2)
                    if (line.find('NNN_FORTH')!=-1):
                        line = line.replace('NNN_FORTH',self.TTAltForth[0])
                    if (line.find('NNN_BACK')!=-1):
                        line= line.replace('NNN_BACK',self.TTAltBack[0])
                    if (line.find('MONU_FORTH')!=-1):
                        line = line.replace('MONU_FORTH',self.TTAltForth[1])
                    if (line.find('MONU_BACK')!=-1):
                        line= line.replace('MONU_BACK',self.TTAltBack[1])
                    if (line.find('TTT_FORTH')!=-1):
                        line = line.replace('TTT_FORTH',self.TTAltForth[2])
                    if (line.find('TTT_BACK')!=-1):
                        line= line.replace('TTT_BACK',self.TTAltBack[2])
                    if (line.find('TTT2_FORTH')!=-1):
                        line = line.replace('TTT2_FORTH',self.TTAltForth[3])
                    if (line.find('TTT2_BACK')!=-1):
                        line= line.replace('TTT2_BACK',self.TTAltBack[3])
                f2.write(line)

            if ( script =="MOT_C1.flg" ):
                if "11YY" in line:
                    line= line.replace("11YY",str(int(self.numfields)))
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
    csvfile = open(os.path.splitext(workbook)[0]+'.csv', 'w')
    writecsv = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL)

    for rownum in range(x1.nrows):
        writecsv.writerow(x1.row_values(rownum))
    csvfile.close()

def main():
    ##    xls_to_csv(workbook="C:\AdvGIS\Group19\MOT19.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group19', file='MOT19.csv', exit1 = "BEILEN01", exit2= "WOLVEGA02")
    # xls_to_csv(workbook="C:\AdvGIS\Group20\MOT20.xls")
    # d = dataf(spath='C:\AdvGIS\Group20', file='MOT20.csv', exit1 = "ExitBeilen", exit2= "ExitSteenw")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group05\MOT05.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group05', file='MOT05.csv', exit1 = "exit_S", exit2= "exit_A")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group06\MOT06.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group06', file='MOT06.csv', exit1 = "JOURE01", exit2= "AKKRUM01")
    # xls_to_csv(workbook="C:\AdvGIS\Group12\MOT12.xls")
    # d = dataf(spath='C:\AdvGIS\Group12', file='MOT12.csv', exit1 = "EMMEN02", exit2= "ASSEN03")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group02\MOT02.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group02', file='MOT02.csv', exit1 = "DOKKUM01", exit2= "DONGJUM01")
    # xls_to_csv(workbook="C:\AdvGIS\Group08\MOT08.xls")
    # d = dataf(spath='C:\AdvGIS\Group08', file='MOT08.csv', exit1 = "Begin_Exit", exit2= "End_Exit")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group01\MOT01.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group01', file='MOT01.csv', exit1 = "BOLSWARD02", exit2= "DONGJUM01")
    # xls_to_csv(workbook="C:\AdvGIS\Group03\MOT03.xls")
    # d = dataf(spath='C:\AdvGIS\Group03', file='MOT03.csv', exit1 = "DOKKUM01", exit2= "DRONRYP02")
    # xls_to_csv(workbook="C:\AdvGIS\Group07\MOT07.xls")
    # d = dataf(spath='C:\AdvGIS\Group07', file='MOT07.csv', exit1 = "VLEDDERVEEN", exit2= "ZUIDBROEK02")
    # xls_to_csv(workbook="C:\AdvGIS\Group09\MOT09.xls")
    # d = dataf(spath='C:\AdvGIS\Group09', file='MOT09.csv', exit1 = "Vledderveen (End)", exit2= "Tynaarlo (Begin)")
    xls_to_csv(workbook="C:\AdvGIS\Group10\MOT10.xls")
    d = dataf(spath='C:\AdvGIS\Group10', file='MOT10.csv', exit1 = "Vleddervee", exit2= "Rhee")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group11\MOT11.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group11', file='MOT11.csv', exit1 = "REDUZUM01", exit2= "BEESTERZWAAG01")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group13\MOT13.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group13', file='MOT13.csv', exit1 = "DOKKUM02", exit2= "GRONINGEN01")
    # xls_to_csv(workbook="C:\AdvGIS\Group14\MOT14.xls")
    # d = dataf(spath='C:\AdvGIS\Group14', file='MOT14.csv', exit1 = "emmeloord", exit2= "zwolle")
    # xls_to_csv(workbook="C:\AdvGIS\Group15\MOT15.xls")
    # d = dataf(spath='C:\AdvGIS\Group15', file='MOT15.csv', exit1 = "EntryEmmeloord", exit2= "ExitStaphorst")
    # xls_to_csv(workbook="C:\AdvGIS\Group17\MOT17.xls")
    # d = dataf(spath='C:\AdvGIS\Group17', file='MOT17.csv', exit1 = "WOLVEGA02", exit2= "ASSEN01")
    # xls_to_csv(workbook="C:\AdvGIS\Group18\MOT18.xls")
    # d = dataf(spath='C:\AdvGIS\Group18', file='MOT18.csv', exit1 = "DOKKUM01", exit2= "DONGJUM01")
    # xls_to_csv(workbook="C:\AdvGIS\Group16\MOT16.xls")
    # d = dataf(spath='C:\AdvGIS\Group16', file='MOT16.csv', exit1 = "ExitAssen", exit2= "ExitBeester")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group04\MOT04.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group04', file='MOT04.csv', exit1 = "LOENGA", exit2= "DRONRY")
    # xls_to_csv(workbook="C:\AdvGIS\Group24\MOT_24.xlsx")
    # d = dataf(spath='C:\AdvGIS\Group24', file='MOT_24.csv', exit1 = "Exit_1", exit2= "Exit_2")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group25\MOT25.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group25', file='MOT25.csv', exit1 = "NIEBERT02", exit2= "RHEE01")
    # xls_to_csv(workbook="C:\AdvGIS\Group22\MOT22_Thomas.xls")
    # d = dataf(spath='C:\AdvGIS\Group22', file='MOT22_Thomas.csv', exit1 = "StartNode", exit2= "EndNode")
    # xls_to_csv(workbook="C:\AdvGIS\Group21\MOT21.xls")
    # d = dataf(spath='C:\AdvGIS\Group21', file='MOT21.csv', exit1 = "SNEEK01", exit2= "AKKRUM01")
    ##    xls_to_csv(workbook="C:\AdvGIS\Group23\MOT23.xls")
    ##    d = dataf(spath='C:\AdvGIS\Group23', file='MOT23.csv', exit1 = "END_EXIT", exit2= "BEG_EXIT")















    d.read()
    d.replaceall(script = "MOT_C1.flg")
    d.replaceall(script = "MOT_C2a.flg")
    d.replaceall(script = "MOT_C2b.flg")
    #print d

if __name__ == '__main__':
    main()
