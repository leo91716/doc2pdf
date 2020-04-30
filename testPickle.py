import pickle
import sys

# python testPickle.py DFTest

with open(sys.argv[1]+'_norm.pickle', 'rb') as file:
    norm=pickle.load(file)


# i=0
# print(f'\nnorm["table2"][{i}]',norm['table2'][i])
# print(f'\nlen(norm["table2"][{i}])',len(norm['table2'][i]))
# i+=1
# print(f'\nnorm["table2"][{i}]',norm['table2'][i])
# i+=1

# i=0
# print(f'\nnorm["table3"][{i}]',norm['table3'][i])
# i+=1
# print(f'\nnorm["table3"][{i}]',norm['table3'][i])
# i+=1


for key in norm:
    for index,item in enumerate(norm[key]):
        print(f'\n\nnorm[{key}][{index}]',item)
        print('type',type(item[0][1]))

