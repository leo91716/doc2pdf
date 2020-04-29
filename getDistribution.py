import glob, csv
from excel2doc import Reader, Writer_Track, Writer_Fluent
from scipy.stats import norm
import copy
import numpy as np
import pickle
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
                self.addToNorm(self.Writer.getBasicMeasureI2(reader),self.data['table2'])
                #self.files.append(list(csv.reader(csvfile)))
        print('raw data: ',self.data['table2'])
        rawData=copy.deepcopy(self.data['table2'])
        self.buildScaleTable(self.data['table2'],self.Writer.getNormReverse())
        '''
        for item in self.data['table2']:
            item.sort(reverse=self.Writer.getNormReverse())
            newItem=[]
            for subItems in item:
                cumulativePercentage=(item.index(subItems)+item.count(subItems))/len(item)
                scaleScore=10+3*norm.ppf((item.index(subItems)+item.count(subItems))/len(item))
                newItem.append((subItems,scaleScore,cumulativePercentage)) # get Cumulative percentage (0~1)
            item.clear()
            item.extend(newItem)
        '''
        
        print('raw data in process:', rawData)
        for index, column in enumerate(rawData):
            for itemIndex,item in enumerate(column):
                for distrib in self.data['table2'][index]:
                    if distrib[0]==item:
                        column[itemIndex]=distrib
        rawData=np.array(rawData)
        rawData=rawData.transpose((1,2,0)).tolist()
        # print('\n\nnew raw data: ',rawData)
        # print('len',len(rawData))
        for file in rawData:
            column=file[1]
            self.addToNorm(self.Writer.getMoreMeasureI2(column),self.data['table3'])

        self.buildScaleTable(self.data['table3'],False)
        print('table3', self.data['table3'])
        print('\n\ntable2: ',self.data['table2'])
        file=open('norm.pickle','wb')
        pickle.dump(self.data,file)
        file.close()
        '''
        #perFile=[]
        i=0
        while i<len(rawData[0]):
            #perFile.append([rawData[0][i],rawData[1],rawData[2],rawData[3]])
            
            self.addToNorm([rawData[1][i][1]+rawData[2][i][1],
            rawData[0][i][1]-rawData[1][i][1], 
            rawData[3][i][1]-rawData[1][i][1],
            rawData[3][i][1]-rawData[2][i][1],
            rawData[3][i][1]-rawData[1][i][1]-rawData[2][i][1],
            rawData[3][i][1]-rawData[4][i][1]
            ], self.data['table3']) #rawData first is column, second is item, third is scale score
            #print('rawData[3][i][1]',rawData[3][i][1])
            #print('rawData[3][i][1]',rawData[3][i][2])
            i+=1
        print('\n\n table3: ',self.data['table3'])
        
        
        print('')
        print('table2[0] len: ',len(self.data['table2'][0]))

            


        print('table2 len: ',len(self.data['table2']))
        
        print(self.data['table2'])
        print(rawData)
        '''
    def buildScaleTable(self,source,reverse):
        for item in source:
            item.sort(reverse=reverse)
            newItem=[]
            for subItems in item:
                cumulativePercentage=(item.index(subItems)+item.count(subItems))/len(item)
                if cumulativePercentage>0.99:
                    cumulativePercentage=0.99
                scaleScore=10+3*norm.ppf(cumulativePercentage)
                pr=int(cumulativePercentage*100)

                newItem.append((subItems,scaleScore,pr,cumulativePercentage)) # get Cumulative percentage (0~1)
            item.clear()
            item.extend(newItem)
    def addToNorm(self,source,dest):
        
        #table2data=[reader.getData('Task1','Complete_Time'),reader.getData('Task2','Complete_Time'),reader.getData('Task3','Complete_Time'),reader.getData('Task4','Complete_Time'),reader.getData('Task5','Complete_Time')]  
        if not dest:
            source=[[item] for item in source]
            dest.extend(source)
            print(dest)
            print('')
        else:
            #print('else')
            i=0
            while i<len(source):
                dest[i].append(source[i])
                i+=1
            
            #self.addSingleTable(dest,source)
        



    #def addSingleTable(self,table,tabledata):
        # i=0
        # while i<len(tabledata):
        #     table[i].append(tabledata[i])
        #     i+=1

        









if __name__=='__main__':
    dist=GetDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合\TMTest_*.csv",Writer_Track)
    #dist=GetDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合\DFTest_*.csv",Writer_Fluent)
    reader=Reader
    dist.buildNorm(Reader)