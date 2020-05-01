import glob, csv
from excel2doc import Reader, Writer_Track, Writer_Fluent
from scipy.stats import norm
import copy
import numpy as np
import pickle
class TableDistribution():
    def __init__(self):
        self.norm=[]
        self.withData=[]
        self.withScale=[]
    def transpose(self,data,change=(1,0)):
        data=np.array(data)
        return data.transpose(change).tolist()
    def create(self,path,Writer,backup,reverse,wayTogetData,source=None):
        self.path=path
        self.Writer=Writer
        print('source:',source)
        data=[]
        rawData=[]
        # if type(source)==list:
        #     self.transpose(source,(1,0,2))
        if source==None:
            self.getDataFromFile(Reader,data,wayTogetData)
        else:
            for file in source:
                self.addToNorm(wayTogetData(file[1]),data)
        withData=copy.deepcopy(data)
        withData=self.transpose(withData)
        self.withData.extend(withData)
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
            rawData=self.transpose(rawData,(2,1,0))
            # rawData=np.array(rawData)
            # rawData=rawData.transpose((1,2,0)).tolist()
        self.norm=data
        self.withScale.extend(rawData)


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
                if pr==0:
                    pr=1
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
        path=r"E:\執行功能output3\EFs_dta\dta_csv集合/"
        #path,Writer,backup,reverse,wayTogetData,source=None)
        tableNumber={'TMTest':2,'DFTest':4}
        tableList=[]
        for i in range(tableNumber[name]):
            tableList.append(TableDistribution())
        tableArgList={
        'TMTest':[{'path':path+name+"_*.csv",'Writer':Writer,'backup':True, 'wayTogetData':Writer.getBasicMeasureI2 ,'reverse':Writer.getNormReverse()},
                {'path':path+name+"_*.csv",'Writer':Writer,'backup':False,'source':tableList[0].withScale, 'wayTogetData':Writer.getMoreMeasureI2 ,'reverse':False}
            ],
        'DFTest':[{'path':path+name+"_*.csv",'Writer':Writer,'backup':True, 'wayTogetData':Writer.getBasicMeasureI2 ,'reverse':Writer.getNormReverse()},
                {'path':path+name+"_*.csv",'Writer':Writer,'backup':False,'source':tableList[0].withScale, 'wayTogetData':Writer.getMoreMeasureI2 ,'reverse':False},
                {'path':path+name+"_*.csv",'Writer':Writer,'backup':False, 'wayTogetData':Writer.getOptionalTableI2 ,'reverse':False},
                # {'path':path+name+"_*.csv",'Writer':Writer,'backup':False, 'wayTogetData':Writer.getOptionalTableI2End,'source':[[tableList[0],0],tableList[1]] ,'reverse':False},
            ],
        
        }
        data={}
        for index,arg in enumerate(tableArgList[name]):
            print('ttttable',index)
            tableList[index].create(**arg)
            data['table'+str(index)]=tableList[index].norm
            print('tableList[index].withScale',tableList[index].withScale)
            print('tableList[index].withData',tableList[index].withData)
            print('tableList[index].norm',tableList[index].norm)
            # print('table.rawData',tableList[index].rawData)

        
        file=open(name+'_norm.pickle','wb')
        pickle.dump(data,file)
        file.close()

        '''
        data={}


        table2=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=True, wayTogetData=Writer.getBasicMeasureI2 ,reverse=Writer.getNormReverse())
        i=2
        data['table'+str(i)]=table2.data
        print('table2 raw data: ',table2.rawData)
        # print('\n\ntable2 data', table2.data)
        table3=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=False,source=table2.rawData, wayTogetData=Writer.getMoreMeasureI2 ,reverse=False)
        # print('table3 raw data: ',table3.rawData)
        # print('\n\ntable3 data', table3.data)
        i+=1
        data['table'+str(i)]=table3.data
        if name=='DFTest':
            table4_1=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=True, wayTogetData=Writer.getOptionalTableI2 ,reverse=False)
            print('enter table4')
            print('\n\ntable4_1.data',table4_1.data)
            print('\n\ntable4_1.rawData',table4_1.rawData)
            #table4_2=TableDistribution(r"E:\執行功能output3\EFs_dta\dta_csv集合/"+name+"_*.csv",Writer,backup=False, wayTogetData=Writer_Fluent.getOptionalTableI2End,source= ,reverse=False)

            i+=1
            data['table'+str(i)]=table4_1.data
        

        file=open(name+'_norm.pickle','wb')
        pickle.dump(data,file)
        file.close()
        '''









if __name__=='__main__':
    #track1=GetDistribution(Writer_Track,'TMTest')
    track1=GetDistribution(Writer_Fluent,'DFTest')