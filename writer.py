class Writers():
    def createDoc(self,reader,norm):
        self.norm=norm
        self.reader=reader
        self.doc = Document()
        self.doc.styles['Normal'].font.name = u'標楷體'
        self.doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
        self.doc.styles['Normal'].font.bold=True
        self.doc.styles['Normal'].font.size=Pt(12)
        styles = self.doc.styles
        new_heading_style = styles.add_style('New Heading', WD_STYLE_TYPE.PARAGRAPH)

        font = new_heading_style.font
        font.bold=True
        font.name = u'標楷體'
        font.size = Pt(16)

        new_heading_style._element.rPr.rFonts.set(qn('w:eastAsia'), u'標楷體')
    @staticmethod
    def getNormReverse():
        talbe2=[True,True,True,True,True]
        talbe3=[False,False,False,False,False,False]
        tables=[talbe2,talbe3]
        return tables
    def add_title(self,title):
        paragraph = self.doc.add_paragraph(title, style='New Heading')
        paragraph.alignment = 1

    def add_paragraph(self,context):
        self.doc.add_paragraph(context)
    def save(self,path):
        self.doc.save(path)
    def show(self):
        pass



class table():
    def __init__(self,doc):
        self.doc=doc
    def getData(*args):
        pass
    def create(title,reader):
        i1=['']+title
        i2=['原始\n分數']+self.getData(reader)
        print(i2)
        i3=['量尺\n分數']+self.getScaleAndPr('table2', i2[1:],reverse=self.getNormReverse()[0])[0]
        i4=['PR值']+self.getScaleAndPr('table2', i2[1:],reverse=self.getNormReverse()[0])[1]
        self.tableList2=[i1,i2,i3,i4]
        self.add_table([4,6],self.tableList2)
    def getScaleAndPr(self,table,data,reverse):
        scale=[]
        norm=self.norm[table]
        for index, column in enumerate(norm):
            for itemIndex,item in enumerate(column):
                if (reverse[index] and data[index]>=item[0]) or (not reverse[index] and data[index]<=item[0]):
                    if data[index]==item[0]:
                        scale.append(item[1:])
                    else:
                        if itemIndex==0:
                            scale.append(item[1:])
                        else:
                            result=[]
                            scaleResult=self.getInnerValue(data[index],column[itemIndex-1][0],column[itemIndex][0],column[itemIndex-1][1],column[itemIndex][1])
                            result=result+[scaleResult]
                            cumulativePercentage=self.getInnerValue(data[index],column[itemIndex-1][0],column[itemIndex][0],column[itemIndex-1][3],column[itemIndex][3])
                            pr=int(cumulativePercentage*100)
                            if pr==0:
                                pr=1
                            result=result+[pr]
                            result=result+[cumulativePercentage]
                            scale.append(result)
                    break
                elif itemIndex==len(column)-1: #if you didn't catch anything
                    scale.append(item[1:])
    def add_table(self,size,context, merge=[]):
        table = self.doc.add_table(rows = size[0], cols = size[1], style='TableGrid')
        for index, item in enumerate(context):
            for index2, item2 in enumerate(item):
                if isinstance(item2,float):
                    table.cell(index,index2).text=f'{item2:.3f}'
                else:
                    table.cell(index,index2).text=str(item2)
                
        for col in table.columns:
            for cell in col.cells:
                cell.paragraphs[0].alignment = 1
        if merge:
            a=table.cell(*merge[0])
            b=table.cell(*merge[1])
            a.merge(b)
