#coding=gb18030
import xlrd
import xlwings as xw
import os
import scandir

class Links(object):
    def __init__(self):
        pass

    def get_ecr_files(self):
        dic={}
        ar=[r'Z:\THIS WEEK',r'Z:\OLD(2019)']
        txt=u'http://intranet.cclmotors.com/Quality/CSD/DCC/Shared%20Documents/文檔資料庫/非受控文件/ECR/'
        for a in ar:
            for rootdir,subfolder,file in scandir.walk(a):
                for f in file:
                    if f.startswith('HECRAP'):
                        dic_value=os.path.join(rootdir,f).replace('Z:\\',txt)
                        dic[f[:-4]]=dic_value
        return dic

    def xl(self):
        xls=xlrd.open_workbook(r'\\Sjstorage\dept_operaton_cml_pe\APE\APE Report File\ECR & PCR 編號跟進.XLS\2019\ECR编号跟进-2019.xls')
        sheet=xls.sheet_by_name('ECR')
        rows=sheet.nrows
        ar=sheet.col_values(0,1,rows)
        return ar

if __name__=="__main__":
    ob=Links()
    files_dic=ob.get_ecr_files()
    xl_read=ob.xl()

    path=r'\\Sjstorage\dept_operaton_cml_pe\APE\APE Report File\ECR & PCR 編號跟進.XLS\2019\ECR编号跟进-2019.xls'
    app=xw.App(visible=False,add_book=False)
    wb=app.books.open(path)

    for x in xl_read:
        if x in files_dic:
            range_address='A{}'.format(xl_read.index(x)+2)
            link_address=files_dic.get(x)
            wb.sheets['ECR'].range(range_address).add_hyperlink(link_address,text_to_display=x,screen_tip=link_address)
            print('Create links ',x)
    wb.save()
    wb.close()
    app.quit()
    app.kill()

'''
    def get_xls_source(self):
        path=r'd:\ECR.xlsx'
        data=pd.read_excel(path,sheetname='ECR') #,index_col=[0]) #,index=u'ECR 编号')
        d=data.iloc[:,0:1]
        return d
        #d=data[u'ECR 编号']
        #return d

if __name__=="__main__":
    ob=Links()
    files_dic=ob.get_ecr_files()
    source=ob.get_xls_source()
    app=xw.App(visible=False,add_book=False)
    wb=app.books.open(r'd:\ECR.xlsx')
    print source.head(5)
    #print wb.sheets['ECR'].range('A5').value
    for pdf_link in files_dic:
        if files_dic.has_key(pdf_link):
            print source[source==pdf_link].index[0]
            #cell_value=
            #wb.sheets['ECR'].range('a5').value='hello'
    wb.save()
    wb.close()
    app.quit()
    data_range=workbook.sheets('ECR').range('a5')
    data_range.add_hyperlink('baidu',screen_tip='oktest')
    workbook.save()
    workbook.close()
    test=Links()
    x=test.get_ecr_files()
    #for dd in d:
     #   if x.has_key(dd):
      #      print dd
    wb.save()
    wb.close()
    app.quit()
'''