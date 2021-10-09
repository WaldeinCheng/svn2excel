# -*- coding: UTF-8 -*-
"""
@Author  ：waldeincheng
@Date    ：2021/9/7 10:09
"""
import re
import sys
import xlwt
import xlrd
import os
import getopt
from xlutils.copy import copy


class svnlogToExcel():
    def __init__(self, argv):
        self.key = True
        self.argv = argv
        # 节点范围，版本号，提交类型，是否全部
        self.revisionRange = ""
        self.version = ""
        self.adaptType = ""
        # 是否全部 如果为1，代表发布说明分类里所有的都显示，如果为0，代表只显示发布说明：是
        self.allOrNot = ""
        # svn路径
        self.rootDir = ""
        # 提交人
        self.author = ""
        # 提交时间
        self.time = ""
        # 提交内容
        self.message = ""
        # 命令行
        self.command = ""
        # 提交类型
        self.submitType = ""
        # 问题ID
        self.bugId = ""
        # 问题描述
        self.desc = ""
        # 问题原因
        self.reason = ""
        # 修改方案
        self.modifyPlan = ""
        # 自测过程
        self.testProcess = ""
        # 是否必现
        self.isPresent = ""
        # 发布说明
        self.releaseShow = ""

    def setReleaseShow(self, releaseShow):
        self.releaseShow = releaseShow

    def getReleaseShow(self):
        return self.releaseShow

    def setIsPresent(self, isPresent):
        self.isPresent = isPresent

    def getIsPresent(self):
        return self.isPresent

    def setTestProcess(self, testProcess):
        self.testProcess = testProcess

    def getTestProcess(self):
        return self.testProcess

    def setModifyPlan(self, modifyPlan):
        self.modifyPlan = modifyPlan

    def getModifyPlan(self):
        return self.modifyPlan

    def setSubmitType(self, submitType):
        self.submitType = submitType

    def getSubmitType(self):
        return self.submitType

    def setDesc(self, desc):
        self.desc = desc

    def getDesc(self):
        return self.desc

    def setReason(self, reason):
        self.reason = reason

    def getReason(self):
        return self.reason

    def setBugId(self, bugID):
        self.bugId = bugID

    def setRevision(self, revision):
        self.revision = revision

    def setAuthor(self, author):
        self.author = author

    def setTime(self, time):
        self.time = time

    def setMessage(self, msg):
        self.message = msg

    def getBugId(self):
        return self.bugId

    def getRevision(self):
        return self.revision

    def getAuthor(self):
        return self.author

    def getTime(self):
        return self.time

    def getMessage(self):
        return self.message

    # 提取命令行参数
    def solvePatterns(self):
        try:
            opts, args = getopt.getopt(self.argv, 'hr:v:t:a:d:',
                                       ['revisionRange=', 'version=', 'adaptType=', 'all=', 'dir='])
        except getopt.GetoptError:
            print('test.py -r <revision> -v <version> -t <adaptType> -a <allornot> -d <dir>')
            sys.exit(2)
        for opt, arg in opts:
            if opt == '-h':
                print('test.py -r <revisionRange> -v <version> -t <adaptType> -a <allornot> -d <dir>')
                sys.exit()
            elif opt in ('-r', '--revisionRange'):
                self.revisionRange = arg
            elif opt in ('-v', '--version'):
                self.version = arg
            elif opt in ('-t', '--adaptType'):
                self.adaptType = arg
            elif opt in ('-a', '--allornot'):
                self.allOrNot = arg
            elif opt in ('-d', '--dir'):
                self.rootDir = arg

    def readsvnlog(self):
        # rootDir = "../"  # svn路径

        self.command = 'svn log ' '-r ' + self.revisionRange + ' ' + self.rootDir
        # print(self.command)
        rootLogList = os.popen(self.command)
        res = rootLogList.read()
        rootLogList.close()

        if self.allOrNot == '1':
            self.publish_flag = True
        else:
            self.publish_flag = False
        i = 0
        message = ""
        result = []
        for line in res.splitlines():
            if line is None or (len(line) < 1):
                continue
            if '--------' in line:
                continue
            i = i + 1
            if line.count("|") >= 3:
                self.svnlog = svnlogToExcel(sys.argv[1:])
                submitType = re.findall(r'【提交类型】:(.*?)【对应版本】', message, flags=0)
                adaptType = re.findall(r'【对应版本】:(.*?)【问题单号】', message, flags=0)
                problemList = re.findall(r'【问题单号】:(.*?)【问题描述】', message, flags=0)
                if problemList:
                    bugid = re.findall(r'Bug #\d+', problemList[0],re.I)
                    taskid = re.findall(r'Task #\d+', problemList[0],re.I)
                    if bugid:
                        problemList = bugid
                    if taskid:
                        problemList = taskid
                desc = re.findall(r'【问题描述】:(.*?)【问题原因】', message, flags=0)
                reason = re.findall(r'【问题原因】:(.*?)【修改方案】', message, flags=0)
                modifyPlan = re.findall(r'【修改方案】:(.*?)【自测过程】', message, flags=0)
                testProcess = re.findall(r'【自测过程】:(.*?)【是否必现】', message, flags=0)
                isPresent = re.findall(r'【是否必现】:(.*?)【发布说明】', message, flags=0)
                releaseShow = re.findall(r'【发布说明】:(.)', message, flags=0)
                if adaptType and self.publish_flag:
                    if self.adaptType in adaptType[0].strip() or 'ALL' in adaptType[0].strip():
                        if submitType:
                            self.svnlog.setSubmitType(submitType[0])
                        else:
                            self.svnlog.setSubmitType("")
                        if problemList:
                            self.svnlog.setBugId(problemList[0])
                        else:
                            self.svnlog.setBugId("")
                        if desc:
                            self.svnlog.setDesc(desc[0])
                        else:
                            self.svnlog.setDesc("")
                        if reason:
                            self.svnlog.setReason(reason[0])
                        else:
                            self.svnlog.setReason("")
                        if modifyPlan:
                            self.svnlog.setModifyPlan(modifyPlan[0])
                        else:
                            self.svnlog.setModifyPlan("")
                        if testProcess:
                            self.svnlog.setTestProcess(testProcess[0])
                        else:
                            self.svnlog.setTestProcess("")
                        if isPresent:
                            self.svnlog.setIsPresent(isPresent[0])
                        else:
                            self.svnlog.setIsPresent("")
                        # if releaseShow:
                        #     self.svnlog.setReleaseShow(releaseShow[0])
                        # else:
                        #     self.svnlog.setReleaseShow("")
                        self.svnlog.setAuthor(self.temp.getAuthor())
                        self.svnlog.setTime(self.temp.getTime())
                        self.svnlog.setMessage(message)
                elif adaptType and not self.publish_flag and releaseShow:
                    if (self.adaptType in adaptType[0].strip() or 'ALL' in adaptType[0].strip()) and releaseShow[0] == '是':
                        if submitType:
                            self.svnlog.setSubmitType(submitType[0])
                        else:
                            self.svnlog.setSubmitType("")
                        if problemList:
                            self.svnlog.setBugId(problemList[0])
                        else:
                            self.svnlog.setBugId("")
                        if desc:
                            self.svnlog.setDesc(desc[0])
                        else:
                            self.svnlog.setDesc("")
                        if reason:
                            self.svnlog.setReason(reason[0])
                        else:
                            self.svnlog.setReason("")
                        if modifyPlan:
                            self.svnlog.setModifyPlan(modifyPlan[0])
                        else:
                            self.svnlog.setModifyPlan("")
                        if testProcess:
                            self.svnlog.setTestProcess(testProcess[0])
                        else:
                            self.svnlog.setTestProcess("")
                        if isPresent:
                            self.svnlog.setIsPresent(isPresent[0])
                        else:
                            self.svnlog.setIsPresent("")
                        # if releaseShow:
                        #     self.svnlog.setReleaseShow(releaseShow[0])
                        # else:
                        #     self.svnlog.setReleaseShow("")
                        self.svnlog.setAuthor(self.temp.getAuthor())
                        self.svnlog.setTime(self.temp.getTime())
                        self.svnlog.setMessage(message)
                message = ""
                if len(self.svnlog.getMessage()) > 0:
                    result.append(self.svnlog)
                self.temp = svnlogToExcel(sys.argv[1:])
                tmpList = line.split("|")
                self.temp.setRevision(tmpList[0])
                self.temp.setAuthor(tmpList[1])
                self.temp.setTime(tmpList[2])
            else:
                message = message + line
        # TODO 最后一个日志，需要单独处理
        submitType = re.findall(r'【提交类型】:(.*?)【对应版本】', message, flags=0)
        adaptType = re.findall(r'【对应版本】:(.*?)【问题单号】', message, flags=0)
        problemList = re.findall(r'【问题单号】:(.*?)【问题描述】', message, flags=0)
        if problemList:
            bugid = re.findall(r'Bug #\d+', problemList[0], re.I)
            taskid = re.findall(r'Task #\d+', problemList[0], re.I)
            if bugid:
                problemList = bugid
            if taskid:
                problemList = taskid
        desc = re.findall(r'【问题描述】:(.*?)【问题原因】', message, flags=0)
        reason = re.findall(r'【问题原因】:(.*?)【修改方案】', message, flags=0)
        modifyPlan = re.findall(r'【修改方案】:(.*?)【自测过程】', message, flags=0)
        testProcess = re.findall(r'【自测过程】:(.*?)【是否必现】', message, flags=0)
        isPresent = re.findall(r'【是否必现】:(.*?)【发布说明】', message, flags=0)
        releaseShow = re.findall(r'【发布说明】:(.)', message, flags=0)
        if adaptType and self.publish_flag:
            if self.adaptType in adaptType[0].strip() or 'ALL' in adaptType[0].strip():
                if submitType:
                    self.temp.setSubmitType(submitType[0])
                else:
                    self.temp.setSubmitType("")
                if problemList:
                    self.temp.setBugId(problemList[0])
                else:
                    self.temp.setBugId("")
                if desc:
                    self.temp.setDesc(desc[0])
                else:
                    self.temp.setDesc("")
                if reason:
                    self.temp.setReason(reason[0])
                else:
                    self.temp.setReason("")
                if modifyPlan:
                    self.temp.setModifyPlan(modifyPlan[0])
                else:
                    self.temp.setModifyPlan("")
                if testProcess:
                    self.temp.setTestProcess(testProcess[0])
                else:
                    self.temp.setTestProcess("")
                if isPresent:
                    self.temp.setIsPresent(isPresent[0])
                else:
                    self.temp.setIsPresent("")
                # if releaseShow:
                #     self.temp.setReleaseShow(releaseShow[0])
                # else:
                #     self.temp.setReleaseShow("")
                self.temp.setAuthor(self.temp.getAuthor())
                self.temp.setTime(self.temp.getTime())
        elif adaptType and not self.publish_flag and releaseShow:
            if releaseShow[0] =='否':
                self.key = False
            if (self.adaptType in adaptType[0].strip() or 'ALL' in adaptType[0].strip()) and releaseShow[0] == '是':
                if submitType:
                    self.temp.setSubmitType(submitType[0])
                else:
                    self.temp.setSubmitType("")
                if problemList:
                    self.temp.setBugId(problemList[0])
                else:
                    self.temp.setBugId("")
                if desc:
                    self.temp.setDesc(desc[0])
                else:
                    self.temp.setDesc("")
                if reason:
                    self.temp.setReason(reason[0])
                else:
                    self.temp.setReason("")
                if modifyPlan:
                    self.temp.setModifyPlan(modifyPlan[0])
                else:
                    self.temp.setModifyPlan("")
                if testProcess:
                    self.temp.setTestProcess(testProcess[0])
                else:
                    self.temp.setTestProcess("")
                if isPresent:
                    self.temp.setIsPresent(isPresent[0])
                else:
                    self.temp.setIsPresent("")
                # if releaseShow:
                #     self.temp.setReleaseShow(releaseShow[0])
                # else:
                #     self.temp.setReleaseShow("")
                self.temp.setAuthor(self.temp.getAuthor())
                self.temp.setTime(self.temp.getTime())
        self.temp.setMessage(message)
        if self.key:
            result.append(self.temp)
        # TODO 表格处理
        # 读取目录下模板
        try:
            data = xlrd.open_workbook('./8910_Release_Notes.xls)')
            book = copy(wb=data)
        except Exception:
            book = xlwt.Workbook(encoding='utf-8')
        # book = xlwt.Workbook(encoding='utf-8')
        sheet = book.add_sheet(self.version)
        # 初始化
        style = xlwt.XFStyle()
        style2 = xlwt.XFStyle()
        style3 = xlwt.XFStyle()
        style.font = self.getFont()
        style.alignment = self.getAlignment()
        style.borders = self.getBorders()
        style3.borders = self.getBorders()

        # 背景颜色
        pattern1 = xlwt.Pattern()
        pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern1.pattern_fore_colour = 44
        pattern2 = xlwt.Pattern()
        pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern2.pattern_fore_colour = 43
        style.pattern = pattern1
        style2.pattern = pattern2
        style2.borders = self.getBorders()
        sheet.write_merge(0, 1, 0, 8, 'Release Notes', style)
        # sheet.write(0,0,'Release Notes')
        # sheet.write(1,0,'Version')
        sheet.write(2, 0, 'Version', style2)
        sheet.write_merge(2, 2, 1, 8, self.version, style3)
        sheet.write(3, 0, 'Revision', style2)
        sheet.write_merge(3, 3, 1, 8, self.revisionRange, style3)
        sheet.write(5, 0, '问题单号', style2)
        sheet.write(5, 1, '提交人', style2)
        sheet.write(5, 2, '日期', style2)
        sheet.write(5, 3, '提交类型', style2)
        sheet.write(5, 4, '问题描述', style2)
        sheet.write(5, 5, '问题原因', style2)
        sheet.write(5, 6, '修改方案', style2)
        sheet.write(5, 7, '自测过程', style2)
        sheet.write(5, 8, '是否必现', style2)
        # sheet.write(5, 9, '发布说明', style2)
        j = 6
        # 特殊列宽一些
        sheet.col(0).width = 256 * 15
        sheet.col(0).width = 256 * 15
        sheet.col(4).width = 256 * 30
        sheet.col(5).width = 256 * 30
        sheet.col(6).width = 256 * 30
        for i in result:
            if 'Task' in i.getBugId():
                taskLink = re.search(r'(\d+)', i.getBugId(), re.I)
                click = "http://36.7.87.100:90/pro/task-view-"+taskLink.group(0)+".html"
                sheet.write(j, 0, xlwt.Formula(
                    'HYPERLINK("%s";"%s")' % (click, i.getBugId())))
            elif 'Bug' in i.getBugId():
                bugLink = re.search(r'(\d+)', i.getBugId(), re.I)
                click = "http://36.7.87.100:90/pro/bug-view-" + bugLink.group(0) + ".html"
                sheet.write(j, 0, xlwt.Formula(
                    'HYPERLINK("%s";"%s")' % (click, i.getBugId())))
            # sheet.write(j, 0, i.getBugId(),style3)
            sheet.write(j, 1, i.getAuthor(), style3)
            sheet.write(j, 2, i.getTime(), style3)
            sheet.write(j, 3, i.getSubmitType(), style3)
            sheet.write(j, 4, i.getDesc(), style3)
            sheet.write(j, 5, i.getReason(), style3)
            sheet.write(j, 6, i.getModifyPlan(), style3)
            sheet.write(j, 7, i.getTestProcess(), style3)
            sheet.write(j, 8, i.getIsPresent(), style3)
            # sheet.write(j, 9, i.getReleaseShow(), style3)
            j = j + 1
        book.save('./版本日志文档/'+self.version+'.xls')

    # 字体设置
    def getFont(self):
        font = xlwt.Font()
        # 加粗
        font.bold = True
        # 字体大小
        font.height = 20 * 11

        return font

    # 表格设置
    def getAlignment(self):
        alignment = xlwt.Alignment()
        # 水平居中
        alignment.horz = 0x02
        # 上下居中
        alignment.vert = 0x01

        return alignment

    # 边框
    def getBorders(self):
        border = xlwt.Borders()
        border.left = 1
        border.right = 1
        border.top = 1
        border.bottom = 1

        return border


if __name__ == '__main__':
    if not os.path.exists("版本日志文档"):
        os.mkdir("版本日志文档")
    tt = svnlogToExcel(sys.argv[1:])
    tt.solvePatterns()
    tt.readsvnlog()
