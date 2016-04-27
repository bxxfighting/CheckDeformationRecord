import os
import sys
import shutil
import time
import random
from collections import OrderedDict

from PyQt4.QtGui import *
from PyQt4.QtCore import *
from PyQt4 import uic
import win32com.client as win32

import wei_ui

class Record(QDialog):
    def __init__(self, master=None):
        super(Record, self).__init__(master)
        self.ui = wei_ui.Ui_Dialog()
        self.ui.setupUi(self)

        self.templateFile = '' #模板文件
        self.saveDir = ''      #生成文件存储目录
        self.saveFileName = '' #生成文件名
        self.resultB = ''
        self.resultD = ''
        self.incr = [0 for x in range(6)]
        self.fixedvalue = []

        self.connect(self.ui.chooseTemplateFileBN, SIGNAL('clicked()'), self.chooseTemplateFile)
        self.connect(self.ui.chooseSaveDirBN, SIGNAL('clicked()'), self.chooseSaveDir)
        self.connect(self.ui.startBN, SIGNAL('clicked()'), self.procWord)


    def getPara(self):
        paras = OrderedDict()
        paras['projectName'] = '工程名称'
        paras['checkRecordNo'] = '检验记录编号'
        paras['checkNo'] = '检验编号'
        paras['checkProject'] = '检验项目'
        paras['checkAccord'] = '检验依据'
        paras['checkCount'] = '检验数量'
        paras['piezometerNo'] = '压力表表号'
        paras['dialgaugeNo'] = '百分表表号'
        paras['load'] = '原始载荷'
        paras['month'] = '月份'
        paras['day'] = '日期'
        paras['hour'] = '小时'
        paras['minute'] = '分钟'
        paras['f'] = 'f值'
        paras['g'] = 'g值'
        paras['l'] = 'l值'
        paras['m'] = 'm值'
        paras['r'] = 'r值'
        paras['s'] = 's值'
        paras['t'] = 't值'
        paras['u'] = 'u值'
        paras['v'] = 'v值'
        paras['w'] = 'w值'
        paras['x'] = 'x值'
        paras['y'] = 'y值'
        paras['a'] = 'a值'
        paras['b'] = 'b值'
        
        if not self.loopCheckPara(paras):
            return False
        float_paras = ['f', 'g', 'l', 'm', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'a', 'b']
        try:
            self.floattoround2(*float_paras)
            self.load = int(self.load)
            self.hour = int(self.hour)
            self.minute = int(self.minute)
            self.setfixedvalue(*float_paras)
        except:
            self.displayMessage('请检查要求是数字的值')
            return False

    def procWord(self):
        if self.getPara() == False:
            return False
        doc = self.getDoc()
        if not doc:
            return False
        table1 = doc.Tables[0]
        self.setTableHead1(table1)

        self.A = 0
        self.B = 0
        self.C = 0
        self.D = 0
        self.E = 0
        self.F = 0
        fixed_step = 0
        incr_step = 0
        row_step = 7
        load_step = 7
        minu_step = 5

        self.B = self.fixedvalue[fixed_step]
        self.D = self.fixedvalue[fixed_step+1]
        self.genIncr(self.B, self.D, 1)

        for n in range(6):
            table1.Cell(load_step, 7).Range.Text = str(int(self.load) * (2*(n+1)*2))
            if n != 0:
                self.updateTime(minu_step)
                table1.Cell(load_step, 3).Range.Text = str(self.hour)
                table1.Cell(load_step, 4).Range.Text = str(self.minute)
            incr_step = 0;
            for i in range(0, 3):
                self.A = self.incr[incr_step]
                self.C = self.incr[incr_step+1]
                if i == 0:
                    self.B = self.fixedvalue[fixed_step]
                    self.D = self.fixedvalue[fixed_step+1]
                else:
                    self.B = round(self.B + self.A, 2)
                    self.D = round(self.D + self.C, 2)
                self.E = round(((float(self.B) + float(self.D)) / 2), 3)
                self.F = round(((float(self.A) + float(self.C)) / 2), 3)

                table1.Cell(i+row_step, 8).Range.Text = str(round(self.B, 2))
                table1.Cell(i+row_step, 9).Range.Text = str(round(self.A, 2))
                table1.Cell(i+row_step, 10).Range.Text = str(round(self.D, 2))
                table1.Cell(i+row_step, 11).Range.Text = str(round(self.C, 2))
                table1.Cell(i+row_step, 16).Range.Text = str(self.F)
                table1.Cell(i+row_step, 17).Range.Text = str(self.E)
                incr_step += 2

            
            fixed_step += 2
            self.resultB = self.fixedvalue[fixed_step] - self.B
            self.resultD = self.fixedvalue[fixed_step] - self.D
            if n == 5:
                self.genIncr(self.resultB, self.resultD, 2)
            else:
                self.genIncr(self.resultB, self.resultD, 1)
            load_step += 3
            row_step += 3

        table2 = doc.Tables[1]
        self.setTableHead2(table2)
        self.updateTime(minu_step)
        table2.Cell(4, 3).Range.Text = str(self.hour)
        table2.Cell(4, 4).Range.Text = str(self.minute)
        table2.Cell(4, 7).Range.Text = str(int(self.load))
        incr_step = 0;
        for i in range(0, 3):
            self.A = self.incr[incr_step]
            self.C = self.incr[incr_step+1]
            if i == 0:
                self.B = self.fixedvalue[fixed_step]
                self.D = self.fixedvalue[fixed_step+1]
            else:
                self.B = round(self.B + self.A, 2)
                self.D = round(self.D + self.C, 2)
            self.E = round(((float(self.B) + float(self.D)) / 2), 3)
            self.F = round(((float(self.A) + float(self.C)) / 2), 3)

            table2.Cell(i+4, 8).Range.Text = str(round(self.B, 2))
            table2.Cell(i+4, 9).Range.Text = str(round(self.A, 2))
            table2.Cell(i+4, 10).Range.Text = str(round(self.D, 2))
            table2.Cell(i+4, 11).Range.Text = str(round(self.C, 2))
            table2.Cell(i+4, 16).Range.Text = str(self.F)
            table2.Cell(i+4, 17).Range.Text = str(self.E)
            incr_step += 2

        self.updateTime(minu_step)
        table2.Cell(6, 3).Range.Text = str(self.hour)
        table2.Cell(6, 4).Range.Text = str(self.minute)

        self.displayMessage('执行成功')

    def getDoc(self):
        if not os.path.exists(self.templateFile):
            self.displayMessage('未选择模板文件或者文件不存在')
            return False
        if not os.path.exists(self.saveDir):
            self.displayMessage('未选择存储目录或者目录不存在')
            return False
        word = win32.Dispatch('Word.Application')
        word.Visible = 0
        self.saveFileName = self.saveDir + '/' + str(int(time.time())) + '.doc'
        shutil.copy(self.templateFile, self.saveFileName)
        word.Documents.Open(self.saveFileName)
        doc = word.ActiveDocument
        return doc

    def chooseTemplateFile(self):
        fd = QFileDialog(self)
        if fd.exec() == QDialog.Accepted:
            self.templateFile = fd.selectedFiles()[0]
            self.ui.chooseTemplateFileLE.setText(self.templateFile)
    
    def chooseSaveDir(self):
        self.saveDir = QFileDialog.getExistingDirectory()
        self.ui.chooseSaveDirLE.setText(self.saveDir)
        
    def genIncr(self, B, D, typ):
        self.incr[0] = B
        self.incr[1] = D
        n = 2
        for i in range(2, 6):
            self.incr[i] = round(random.uniform(0, n), 2)
            n -= self.incr[i]
        if typ == 2:
            for i in range(2, 6):
                if self.incr[i] != 0:
                    self.incr[i] = -self.incr[i]

    def displayMessage(self, message):
        QMessageBox.warning(self, '警告', message)

    def checkPara(self, name, message):
        cmd = 'self.' + name + '=' + 'self.ui.' + name + 'LE.text()'
        exec(cmd)
        if not eval('self.'+name):
            self.displayMessage('请输入'+message)
            return False
        return True

    def loopCheckPara(self, kargs):
        for k in kargs:
            if not self.checkPara(k, kargs[k]):
                return False
        return True
        
    def floattoround2(self, *args):
        for arg in args:
            cmd = 'self.' +  arg + '=' + 'round(float(self.%s), 2)' % arg
            exec(cmd)
            
    def setfixedvalue(self, *args):
        for arg in args:
            cmd = 'self.fixedvalue.append(self.%s)' % arg
            exec(cmd)

    def setTableHead1(self, table):
        table.Cell(1, 2).Range.Text = self.projectName
        table.Cell(1, 4).Range.Text = self.checkRecordNo
        table.Cell(1, 6).Range.Text = self.checkNo
        table.Cell(2, 2).Range.Text = self.checkProject
        table.Cell(2, 4).Range.Text = self.checkAccord
        table.Cell(2, 6).Range.Text = self.checkCount
        table.Cell(3, 2).Range.Text = '压力表表号：' + self.piezometerNo + '\n' + '百分表表号：' + self.dialgaugeNo
        table.Cell(6, 1).Range.Text = self.month
        table.Cell(6, 2).Range.Text = self.day
        table.Cell(6, 3).Range.Text = self.hour
        table.Cell(6, 4).Range.Text = self.minute
        table.Cell(6, 7).Range.Text = self.load

    def setTableHead2(self, table):
        table.Cell(1, 2).Range.Text = self.projectName
        table.Cell(1, 4).Range.Text = self.checkRecordNo
        table.Cell(1, 6).Range.Text = self.checkNo

    def updateTime(self, step):
        self.minute += step
        if self.minute >= 60:
            self.hour  += 1
            self.minute %= 60
        

app = QApplication(sys.argv)
record = Record()
record.show()
app.exec_()
