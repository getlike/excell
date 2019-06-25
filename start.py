import xlrd
from docxtpl import DocxTemplate
edf=xlrd.open_workbook('inner.xlsx')
sheet=edf.sheet_by_name('inner')
arrData=[]
row_nuber = sheet.nrows
#work with word



#number,fio,ident,address=''


if row_nuber > 0:
    for conut in range(0, row_nuber):
        #arrData.append(str(sheet.row(conut)[1]).replace("number:",""))

        context = {'nuber': int(float(str(sheet.row(conut)[0]).replace("number:",""))),
                    'fio': str(sheet.row(conut)[1]).replace("text:","").replace("'"," "),
                    'id': int(float(str(sheet.row(conut)[2]).replace("number:",""))),
                    'address': str(sheet.row(conut)[3]).replace("text:","").replace("'"," ")}

        doc = DocxTemplate("my_template.docx")
        doc.render(context)
        doc.save('generated/%s %s .docx'%(conut,str(sheet.row(conut)[1]).replace("text:","").replace("'"," ")))
        del doc
        print(conut)
    #print(len(arrData))
# тестовый принт
# for i in range(0,10):
#     print(arrData[i])print(int(float(arrData[0])))





