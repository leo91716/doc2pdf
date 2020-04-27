import glob, csv
from excel2doc import Reader, Writer_Track, Writer_Fluent
class GetDistribution():
    def __init__(self,path,Writer):
        self.path=path
        self.Writer=Writer
    def buildNorm(self,Reader):
        path=self.path
        self.data={}
        self.data['table2']=[]
        self.data['table3']=[]
        #self.files=[]
        for file in glob.glob(path):
            with open(file,newline='') as csvfile:
                reader=Reader(list(csv.reader(csvfile)))
                self.addToNorm(reader)
                #self.files.append(list(csv.reader(csvfile)))
        for item in self.data['table2']:
            item.sort()
        
        print(self.data)
    def addToNorm(self,reader):
        
        #table2data=[reader.getData('Task1','Complete_Time'),reader.getData('Task2','Complete_Time'),reader.getData('Task3','Complete_Time'),reader.getData('Task4','Complete_Time'),reader.getData('Task5','Complete_Time')]
        table2data=self.Writer.getBasicMeasureI2(reader)
        table2=self.data['table2']    
        if not table2:
            table2data=[[item] for item in table2data]
            table2.extend(table2data)
            print(table2)
            print('')
        else:
            #print('else')
            self.addSingleTable(table2,table2data)
        '''
        table3data=[table2data[1]+table2data[2],   table2data[0]-table2data[1],  table2data[3]-table2data[1],   table2data[3]-table2data[2],  table2data[3]-table2data[2]-table2data[2], table2data[3]-table2data[4]]
        table3=self.data[1]        
        self.addSingleTable(table3,table3data)
        '''



    def addSingleTable(self,table,tabledata):
        i=0
        while i<len(tabledata):
            table[i].append(tabledata[i])
            i+=1

        









if __name__=='__main__':
    dist=GetDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合\DFTest_*.csv",Writer_Fluent)
    reader=Reader
    dist.buildNorm(Reader)