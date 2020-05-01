import pickle
import numpy as np
def getScaleAndPr(self,table,data,reverse,norm):
        scale=[]
        norm=norm[table]
        for index, column in enumerate(norm):
            for itemIndex,item in enumerate(column):
                if (reverse and data[index]>=item[0]) or (not reverse and data[index]<=item[0]):
                    if data[index]==item[0]:
                        scale.append(item[1:])
                    else:
                        if itemIndex==0:
                            print('smallest')
                        else:
                            result=[]
                            # (rrr-column[itemIndex-1][1])/column[itemIndex][1]-column[itemIndex-1][1]=abs(data[index]-column[itemIndex-1][0])/abs(column[itemIndex][0]-column[itemIndex-1][0])
                            # scaleResult=abs(data[index]-column[itemIndex-1][0])/abs(column[itemIndex][0]-column[itemIndex-1][0])*(column[itemIndex][1]-column[itemIndex-1][1])+column[itemIndex-1][1]
                            scaleResult=getInnerValue(data[index],column[itemIndex-1][0],column[itemIndex][0],column[itemIndex-1][1],column[itemIndex][1])
                            result=result+[scaleResult]
                            ############################need to edit#########################################################
                            cumulativePercentage=getInnerValue(data[index],column[itemIndex-1][0],column[itemIndex][0],column[itemIndex-1][3],column[itemIndex][3])
                            pr=int(cumulativePercentage*100)
                            if pr==0:
                                pr=1
                            result=result+[pr]
                            result=result+[cumulativePercentage]
                            scale.append(result)
                    break
                elif itemIndex==len(column)-1: #if you didn't catch anything
                    print('biggest')
        # print('\n\nscale: ',scale)
        scale=np.array(scale)
        scale=scale.transpose().tolist()
        #scale[1]=list(map(int,scale[1]))
        return scale


def getInnerValue(data,low,high,getLow,getHigh):
    return abs(data-low)/abs(high-low)*(getHigh-getLow)+getLow

with open('TMTest_norm.pickle', 'rb') as file:
    norm=pickle.load(file)
    print(getScaleAndPr(1,'table2',[60.11,100,100,100,100],True,norm))