import glob, csv
from excel2doc import Reader, Writer_Track, Writer_Fluent
from scipy.stats import norm
import copy
import numpy as np
import pickle
class TableDistribution():
    def __init__(self,path,Writer,backup,reverse,wayTogetData,source=None):
        self.path=path
        self.Writer=Writer
        data=[]
        rawData=None
        if source==None:
            self.getDataFromFile(Reader,data,wayTogetData)
        else:
            for file in source:
                column=file[1]
                self.addToNorm(wayTogetData(column),data)
        print('just read data', data)
        if backup:
            rawData=copy.deepcopy(data)
        self.buildScaleTable(data,reverse)
        if backup:
            for index, column in enumerate(rawData):
                for itemIndex,item in enumerate(column):
                    for distrib in data[index]:
                        if distrib[0]==item:
                            column[itemIndex]=distrib
            rawData=np.array(rawData)
            rawData=rawData.transpose((1,2,0)).tolist()
        self.data=data
        self.rawData=rawData


    def getDataFromFile(self,Reader,dest,wayTogetData):
        path=self.path
        for file in glob.glob(path):
            with open(file,newline='') as csvfile:
                reader=Reader(list(csv.reader(csvfile)))
                self.addToNorm(wayTogetData(reader),dest)

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
                newItem.append((subItems,scaleScore,pr,cumulativePercentage)) # get Cumulative percentage (0~0.99)
            
            print('len new item', len(newItem))
            newItem=list(set(newItem))
            print('len new item after set:', len(newItem))
            newItem.sort(reverse=reverse)
            item.clear()
            item.extend(newItem)
    def addToNorm(self,source,dest):
        
        #table2data=[reader.getData('Task1','Complete_Time'),reader.getData('Task2','Complete_Time'),reader.getData('Task3','Complete_Time'),reader.getData('Task4','Complete_Time'),reader.getData('Task5','Complete_Time')]  
        if not dest:
            source=[[item] for item in source]
            dest.extend(source)
            # print(dest)
            # print('')
        else:
            #print('else')
            i=0
            while i<len(source):
                dest[i].append(source[i])
                i+=1


class GetDistribution():
    def __init__(self,Writer,name):
        data={}
        table2=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=True, wayTogetData=Writer.getBasicMeasureI2 ,reverse=Writer.getNormReverse())
        i=2
        data['table'+str(i)]=table2.data
        # print('table2 raw data: ',table2.rawData)
        # print('\n\ntable2 data', table2.data)
        table3=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=False,source=table2.rawData, wayTogetData=Writer.getMoreMeasureI2 ,reverse=False)
        # print('table3 raw data: ',table3.rawData)
        # print('\n\ntable3 data', table3.data)
        i+=1
        data['table'+str(i)]=table3.data
        if name=='DFTest':
            table4=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=False, wayTogetData=Writer.getOptionalTableI2 ,reverse=False)
            # print('table3 raw data: ',table3.rawData)
            # print('\n\ntable3 data', table3.data)
            print('enter table4')
            i+=1
            data['table'+str(i)]=table4.data
            # print('data[table4]',data['table4'])


        file=open(name+'_norm.pickle','wb')
        pickle.dump(data,file)
        file.close()










if __name__=='__main__':
    #track1=GetDistribution(Writer_Track,'TMTest')
    track1=GetDistribution(Writer_Fluent,'DFTest')