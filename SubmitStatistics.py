# author 周治淦
# updateTime：2022.5.24
import openpyxl
import os
import re
import shutil
class CollectionHomeWork:
    #当前文件路径
    __CurrentPath=os.path.dirname(os.path.abspath(__file__))
    # 名单文件名
    __XLSXFileName="nameList.xlsx"
    #原始实验报告存档文件夹
    __ResourceDirName="实验报告"
    #提交人员名单
    __SubmitList=[]

    # 打开xlsx文件，并返回包含学生学号和姓名的列表
    def OpenXLSX(self): 
        wb=openpyxl.load_workbook(self.__CurrentPath+os.sep+self.__XLSXFileName)
        sheet=wb.worksheets[0]
        names=sheet["B"]
        xuehao=sheet["A"]
        InfoList=[]
        for name,xue in zip(names,xuehao):
            InfoList.append(StudentInfo(name.value, xue.value))
        return InfoList
    
    # 根据学生学号和姓名，创建文件夹（仅当该文件夹不存在时创建）
    def CreateFolder(self,InfoList):
        for info in InfoList:
            folderName=info.toString()
            if not self.__isFolderExists(folderName):
                os.mkdir(os.path.join(self.__CurrentPath,self.__ResourceDirName,folderName))
    
    # 根据正则匹配到学生，并将文件转移到对应文件夹内
    def findStudentAndMove(self,StudentInfoList):
        files = os.listdir(os.path.join(self.__CurrentPath,self.__ResourceDirName))
        for i in files:
            if os.path.isfile(os.path.join(self.__CurrentPath,self.__ResourceDirName,i)):
               for name in StudentInfoList:
                    m=re.search(name.Name, i)
                    if m is not None:
                        if m.group(0) not in self.__SubmitList:
                            self.__SubmitList.append(m.group(0))
                        shutil.move(os.path.join(self.__CurrentPath,self.__ResourceDirName,i), os.path.join(self.__CurrentPath,self.__ResourceDirName,name.toString()))
                        
                                    
    
    # 检查学生文件夹是否存在     
    def __isFolderExists(self,StudentInfo):
       return  os.path.exists(os.path.join(self.__CurrentPath,self.__ResourceDirName,StudentInfo))
    def DeleteUnSubmitFolder(self,Unsubmit,InfoList):
        for i in InfoList:
            if i.Name in Unsubmit:
                shutil.rmtree(os.path.join(self.__CurrentPath,self.__ResourceDirName,i.toString()))
        
   # 统计提交人
    def CountSubmit(self):
        print("已提交人：")
        print(self.__SubmitList)
        print("已提交人数：{}".format(len(self.__SubmitList)))
    # 统计未提交人
    def CountUnSubmit(self,InfoList):
        UnSubmit=[]
        for i in InfoList:
            if i.Name not in self.__SubmitList:
                UnSubmit.append(i.Name)
        print("未提交人：")
        print(UnSubmit)
        print("未提交人数：{}".format(len(UnSubmit)))
        return UnSubmit
    # 调试用，将转移到文件夹的文件，转移出来
    def Debug(self):
        files=os.listdir(os.path.join(self.__CurrentPath,self.__ResourceDirName))
        for i in files:
            if os.path.isdir(os.path.join(self.__CurrentPath,self.__ResourceDirName,i)):
                files1=os.listdir(os.path.join(self.__CurrentPath,self.__ResourceDirName,i))
                for j in files1:
                    shutil.move(os.path.join(self.__CurrentPath,self.__ResourceDirName,i,j), os.path.join(self.__CurrentPath,self.__ResourceDirName))
     # 主调函数
    def main(self):
        self.Debug()
        InfoList=self.OpenXLSX()
        self.CreateFolder(InfoList)
        self.findStudentAndMove(InfoList)
        self.CountSubmit()
        print()
        print()
        unsubmit= self.CountUnSubmit(InfoList)
        self.DeleteUnSubmitFolder(unsubmit, InfoList)
# 学生信息类
class StudentInfo:
    def __init__(self,Name,Id):
        self.Name=Name
        self.Id=Id
    def toString(self):
        return self.Id+"-"+self.Name
if __name__ == '__main__':
    mainClass=CollectionHomeWork()
    mainClass.main()

    