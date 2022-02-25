# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import QMessageBox, QLineEdit, QApplication, QMainWindow, QAbstractItemView, QTableWidgetItem, \
    QFileDialog
import _thread
from PyQt5 import QtCore
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
import poplib
from email.parser import Parser
from email.header import decode_header, Header
from email.utils import parseaddr
import webbrowser
import time
import idenfy
import pop3
import smtp
import trans
from addr_book import Contact
from pypinyin import pinyin, Style    # pip3 install pypinyin

app = QApplication(sys.argv)

global user
user = {
    'pop3': None,
    'user': None,
    'pass': None
}

global is_Idenfy
is_Idenfy = False

# email类
class Email(object):
    def __init__(self):
        self.send_name = None
        self.send_addr = None
        self.receive_name = None
        self.receive_addr = None
        self.email_time = None
        self.attach_tag = 0
        self.charset = None
        self.content = None
        self.title = None
        self.index = -1  # 标记邮件是否为未被初始化的邮件
        self.attachment_files = {}
        # self.is_trans_email = False
        # self.is_trans_attach = False

    # 判断邮件是否被初始化
    def is_email_initi(self):
        if self.index == -1:
            return False
        else:
            return True

    # 解码函数
    def decode_str(self, s):
        value, charset = decode_header(s)[0]
        # self.charset = charset
        if charset:
            value = value.decode(charset)
        return value

    # 获取附件信息
    def get_attach_data(self, part, file_name):
        h = Header(file_name)
        # 对附件名称进行解码
        dh = decode_header(h)
        filename = dh[0][0]
        if dh[0][1]:
            # 将附件名称可读化
            filename = self.decode_str(str(filename, dh[0][1]))

        data = part.get_payload(decode=True)
        # self.attachment_data = data
        self.attachment_files[filename] = data
        # print( self.attachment_files)

    # 下载附件
    def download_attach(self, filename, path):
        with open(filename, 'wb') as f:
            f.write(self.attachment_files[filename])
            f.close()

    # 解析邮件信息
    def parser_email_from_POP3(self, msg, index):
        # 标记邮件在邮箱中的索引
        self.index = index
        # 提取头部信息
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                # 标题
                if header == 'Subject':
                    value = self.decode_str(value)
                    self.title = value
                # 发送方
                elif header == 'From':
                    hdr, addr = parseaddr(value)
                    name = self.decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
                    self.send_name = name
                    self.send_addr = addr
                # 接收方
                else:
                    hdr, addr = parseaddr(value)
                    name = self.decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
                    self.receive_addr = addr
                    self.receive_name = name
            print('%s: %s' % (header, value))

            # 时间
        date = decode_header(msg['Received'])
        self.email_time = str(date).split(";")[-1].split("'")[0]
        if self.email_time[0] == "\\":
            self.email_time = ' ' + self.email_time.split("\\t")[1]

        # 获取邮件主体信息
        attachment_files = []
        for part in msg.walk():
            # 获取附件名称类型
            file_name = part.get_filename()
            # 获取数据类型
            contentType = part.get_content_type()
            # 获取编码格式
            mycode = part.get_content_charset()
            # 如果是附件，测调用附件处理函数
            if file_name:
                self.attach_tag = 1
                self.get_attach_data(part, file_name)
            # 正文部分处理
            elif contentType == 'text/plain':
                data = part.get_payload(decode=True)
                content = data.decode(mycode)
                print('正文：', content)
                self.content = content

            elif contentType == 'text/html':
                data = part.get_payload(decode=True)
                html_content = data.decode(mycode)
                #GET_HTML = "demo_" + str(time.time()) + ".html"  # 命名生成的html
                #f = open(GET_HTML, 'w')
                message = html_content
                # f.write(message)
                # f.close()
                # webbrowser.open(GET_HTML, new = 1)
                self.content = message
                #self.content_html = GET_HTML

        print('附件名列表：', attachment_files)


# receive_Server 类
class Receive_server():
    def __init__(self, user_dic):
        self.user_mail = user_dic['user']
        self.password = user_dic['pass']
        self.pop_server = user_dic['pop3']  # pop.163.com
        self.Connect_Server()
        self.attach = ''
        self.email_list = []

    # 连接服务
    def Connect_Server(self):
        self.server = poplib.POP3(self.pop_server)
        self.server.user(self.user_mail)
        self.server.pass_(self.password)

    # 关闭服务器资源
    def _close_(self):
        self.server.close()

    # 获取邮件数目
    def Get_Email_Count(self):
        email_num, email_size = self.server.stat()
        return email_num


# send_server类
class Send_server(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.send_ui = smtp.Ui_MainWindow()
        self.send_ui.setupUi(self)
        self.send_ui.pushButton_2.clicked.connect(self.close)
        self.send_ui.pushButton.clicked.connect(self.get_email_from_UI)
        self.send_ui.pushButton_3.clicked.connect(self.attach_click)
        self.send_ui.pushButton_9.clicked.connect(self.import_addr_book)
        self.send_ui.pushButton_10.clicked.connect(self.export_addr_book)
        self.send_ui.pushButton_7.clicked.connect(self.Update_addr_book)
        self.send_ui.pushButton_6.clicked.connect(self.save_to_addr_book)
        self.send_ui.pushButton_5.clicked.connect(self.dele_in_addr_book)
        self.send_ui.pushButton_8.clicked.connect(self.modify_in_addr_book)
        self.send_ui.pushButton_4.clicked.connect(self.search_in_addr_book)
        self.send_ui.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.send_ui.tableWidget.itemClicked.connect(self.addr_book_click)
        self.send_attach_path = ''

    # 载入附件路径信息
    def attach_click(self):
        global pop3_Server
        filepath, _ = QFileDialog.getOpenFileName()  # 显示文件夹 返回filepath是一个元组
        self.send_attach_path = filepath
        self.display_attach()

    # 将附件名显示到文本框中
    def display_attach(self):
        filename = self.send_attach_path.split("/")[-1]
        self.send_ui.lineEdit_3.setText(QtCore.QCoreApplication.translate("MainWindow", filename))

    # 发送端从UI获取信息
    def get_email_from_UI(self):
        global pop3_Server
        try:
            smtp_server = self.send_ui.lineEdit.text()
            send_name = self.send_ui.lineEdit_4.text()
            recv_mail = self.send_ui.lineEdit_5.text()
            send_title = self.send_ui.lineEdit_6.text()
            send_text = self.send_ui.textEdit.toPlainText()
            send_attach_path = self.send_attach_path
            smtp_tail = user['pop3']
            smpt_tail_list = smtp_tail.split('.')
            tail = '@' + smpt_tail_list[1] + '.' + smpt_tail_list[2]
            fromEmailAddr = pop3_Server.user_mail + tail  # 邮件发送方邮箱地址
            print(fromEmailAddr)
            password = pop3_Server.password  # 密码(部分邮箱为授权码)
            try:
                toEmailAddrs = recv_mail.split(';')
                toEmailAddrs = [x.strip() for x in toEmailAddrs]
            except:
                toEmailAddrs = [recv_mail]  # 邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发

        except:
            QMessageBox.about(self, "Message", "填写完整信息")
        _thread.start_new_thread(self.Send_Email,
                                 (smtp_server, send_name, recv_mail, send_title,
                                  send_text, send_attach_path, fromEmailAddr, password, toEmailAddrs,''))


    # 将UI获取的信息封装为邮件并发送
    def Send_Email(self, smtp_server, send_name, recv_mail, send_title, send_text, send_attach, fromEmailAddr, password,
                   toEmailAddrs, tran_attach_files):

        try:
            # 设置email信息
            # ---------------------------发送带附件邮件-----------------------------
            # 邮件内容设置
            message = MIMEMultipart()
            # 邮件主题
            message['Subject'] = send_title
            # 发送方信息
            message['From'] = formataddr([send_name, fromEmailAddr])
            # 接受方信息
            message['To'] = recv_mail
            # 邮件正文内容
            message.attach(MIMEText(send_text, 'plain', 'utf-8'))

            # 构造附件
            if tran_attach_files:
                att1 = MIMEText(list(tran_attach_files.values())[0], 'base64', 'utf-8')
                att1['Content-type'] = 'application/octet-stream'
                att1['Content-Disposition'] = 'attachment; filename=' + '"' + list(tran_attach_files.keys())[0] + '"'
                message.attach(att1)
            elif send_attach:
                att1 = MIMEText(open(send_attach, 'rb').read(), 'base64', 'utf-8')
                att1['Content-type'] = 'application/octet-stream'
                att1['Content-Disposition'] = 'attachment; filename=' + '"' + str(str(send_attach).split("/")[-1]) + '"'
                message.attach(att1)
            # ---------------------------------------------------------------------

            # 登录并发送邮件
            try:
                server = smtplib.SMTP(smtp_server)  # 163邮箱服务器地址，端口默认为25
                server.login(fromEmailAddr, password)
                server.sendmail(fromEmailAddr, recv_mail.split(';'), message.as_string())
                print('success')
                #QMessageBox.information(self, "Message", "发送成功！", QMessageBox.Yes | QMessageBox.No)
                self.close()
                server.quit()
            except smtplib.SMTPException as e:
                print("error:", e)
        except:
            QMessageBox.critical(self, "失败", "pop验证失败", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

    # 导入通讯录
    def import_addr_book(self):
        path, _ = QFileDialog.getOpenFileName()
        addr_book = Contact()
        addr_book.read(path)   # 从指定路径读入通讯录
        addr_book.write('')    # 存入本地通讯录文件(当前路径下‘contact.txt’)
        self.Update_addr_book()
        if path:
            QMessageBox.information(self, "Message", "导入通讯录成功！", QMessageBox.Yes | QMessageBox.No)

    # 导出通讯录
    def export_addr_book(self):
        path = QFileDialog.getExistingDirectory()
        addr_book = Contact()
        addr_book.read('')  # 从本地通讯录文件(当前路径下‘contact.txt’)读入
        addr_book.write(path)  # 存入指定路径
        if path:
            QMessageBox.information(self, "Message", "导出通讯录成功！", QMessageBox.Yes | QMessageBox.No)

    # 读取本地通讯录并显示在table中
    def Update_addr_book(self):
        # 清空列表
        allrownum = self.send_ui.tableWidget.rowCount()
        for i in range(allrownum):
            self.send_ui.tableWidget.removeRow(0)
        # 读取本地通讯录
        addr_book = Contact()
        addr_book.read('')
        # 显示
        for name,email in addr_book.contacts_dict.items():
            rrow = self.send_ui.tableWidget.rowCount()
            self.send_ui.tableWidget.insertRow(rrow)
            self.send_ui.tableWidget.setItem(rrow, 0, QTableWidgetItem(str(name)))
            self.send_ui.tableWidget.setItem(rrow, 1, QTableWidgetItem(str(email)))

    # 点击通讯录一栏 则自动填充收件人邮箱
    def addr_book_click(self):
        now_current_addr = self.send_ui.tableWidget.currentIndex().row()
        recv_name = self.send_ui.tableWidget.item(now_current_addr, 0).text()
        recv_email = self.send_ui.tableWidget.item(now_current_addr, 1).text()
        self.send_ui.lineEdit_7.setText(recv_name)
        self.send_ui.lineEdit_5.setText(recv_email)

    # 将当前收件人邮箱保存至本地通讯录，用户自行命名昵称
    def save_to_addr_book(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            name = self.send_ui.lineEdit_7.text()
            email = self.send_ui.lineEdit_5.text()
            addr_book = Contact()
            addr_book.read('')
            addr_book.add_contact(name, email)
            addr_book.write('')
            self.Update_addr_book()
        # if result:
        #     QMessageBox.information(self, "Message", "保存该联系人成功！", QMessageBox.Yes | QMessageBox.No)
        # else:
        #     QMessageBox.information(self, "Message", "该联系人已存在！", QMessageBox.Yes | QMessageBox.No)

    # 删除本地通讯录中该收件人信息
    def dele_in_addr_book(self):
        dele_name = self.send_ui.lineEdit_7.text()
        addr_book = Contact()
        addr_book.read('')
        result = addr_book.delete_contact(dele_name)
        addr_book.write('')
        self.Update_addr_book()
        if result:
            QMessageBox.information(self, "Message", "删除该联系人成功！", QMessageBox.Yes | QMessageBox.No)
        else:
            QMessageBox.information(self, "Message", "该联系人不存在！", QMessageBox.Yes | QMessageBox.No)

    # 修改本地通讯录中该收件人信息
    def modify_in_addr_book(self):
        name = self.send_ui.lineEdit_7.text()
        email = self.send_ui.lineEdit_5.text()
        addr_book = Contact()
        addr_book.read('')
        addr_book.modify_contact(name, email)
        addr_book.write('')
        self.Update_addr_book()

    # 按昵称在本地通讯录中寻找，将邮箱显示在“收件人邮箱”处
    def search_in_addr_book(self):
        name = self.send_ui.lineEdit_7.text()
        addr_book = Contact()
        addr_book.read('')
        email = addr_book.search_contact(name)
        self.send_ui.lineEdit_5.setText(email)


class Transmit_server(QMainWindow):
    def __init__(self,mainWin):
        QMainWindow.__init__(self)
        self.send_ui = trans.Ui_MainWindow()
        self.send_ui.setupUi(self)
        self.send_ui.pushButton_2.clicked.connect(self.close)
        self.send_ui.pushButton.clicked.connect(mainWin.transfer_email)

sendUi = None

# 收件箱ui类
class mainWin(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.main_ui = pop3.Ui_MainWindow()
        self.main_ui.setupUi(self)
        self.main_ui.pushButton_2.clicked.connect(self.Send)
        self.main_ui.pushButton_4.clicked.connect(self.Update)
        self.main_ui.tableWidget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.main_ui.tableWidget.itemClicked.connect(self.itemclick)
        self.main_ui.textEdit.document().setMaximumBlockCount(100)
        self.main_ui.pushButton_5.clicked.connect(self.load_attach_path)          #下载附件选择路径跳转函数
        self.main_ui.pushButton_6.clicked.connect(self.download_attach)           #下载附件跳转函数
        self.main_ui.pushButton_7.clicked.connect(self.delete_email)

        self.main_ui.pushButton_19.clicked.connect(self.save_to_addr_book)

        self.main_ui.comboBox.currentIndexChanged.connect(
            lambda: self.ChoseSearchObject(self.main_ui.comboBox.currentIndex()))

        self.main_ui.comboBox_2.currentIndexChanged.connect(
            lambda: self.ChoseSortObject(self.main_ui.comboBox_2.currentIndex()))

    def ChoseSearchObject(self,tag):
        if tag == 1:
            self.search_by_title()
        elif tag == 2:
            self.search_by_name()
            #print('似乎发生了什么事')
        elif tag == 3:
            self.search_by_time()
            #print('似乎又发生了什么事')
        elif tag == 4:
            self.search_by_addr()

    def ChoseSortObject(self,tag):
        if tag == 1:
            self.Update()
        elif tag == 2:
            self.sort_by_time_reverse()
        elif tag == 3:
            self.sort_by_title()
        elif tag == 4:
            self.sort_by_title_reverse()
        elif tag == 5:
            self.sort_by_sender()
        elif tag == 6:
            self.sort_by_sender_reverse()

    # 点击邮件列表发生事件函数
    def itemclick(self):
        global pop3_Server
        now_current_row = self.main_ui.tableWidget.currentIndex().row()
        index = self.main_ui.tableWidget.item(now_current_row, 6).text()
        for An_email in pop3_Server.email_list:
            # print('right now An_email.index:',An_email.index)
            # print('wanted index:',index)
            if str(An_email.index) == str(index):
                break
            else:
                continue
        try:
            rowtitle = An_email.title
            rowsender = An_email.send_name
            etime = An_email.email_time
            rowaddr = An_email.send_addr
            contents = An_email.content
            if An_email.attachment_files:
                attach_name = list(An_email.attachment_files.keys())[0]
            else:
                attach_name = ''
        except:
            raise TypeError

        self.main_ui.textEdit.clear()
        self.main_ui.lineEdit.clear()
        self.main_ui.lineEdit_2.clear()
        self.main_ui.lineEdit_3.clear()
        self.main_ui.lineEdit_4.clear()
        _thread.start_new_thread(self.Dis_mail_data, (rowtitle, rowsender, rowaddr, contents, etime, attach_name, index))

    # 显示邮件详细信息
    def Dis_mail_data(self, title, sender, addr, cont, etime, attach_name, index):
        self.main_ui.lineEdit.setText(title)
        self.main_ui.lineEdit_2.setText(sender)
        self.main_ui.lineEdit_3.setText(addr)
        self.main_ui.lineEdit_4.setText(etime)
        self.main_ui.lineEdit_5.setText(attach_name)
        self.main_ui.lineEdit_6.setText(index)

        # print(len(cont))
        if len(cont) > 5000:
            self.main_ui.textEdit.append(" ")
        else:
            self.main_ui.textEdit.append(cont)

    # 刷新判断函数
    def Update(self):
        global pop3_Server, is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            _thread.start_new_thread(self.Upthread, ())

    # 刷新操作函数
    def Upthread(self):
        global pop3_Server
        # 清空列表
        allrownum = self.main_ui.tableWidget.rowCount()
        for i in range(allrownum):
            self.main_ui.tableWidget.removeRow(0)

        # 清空reciever_Server中的邮件列表
        pop3_Server.email_list = []

        # 获取邮箱邮件数
        mail_count = pop3_Server.Get_Email_Count()

        # 获取最新20条邮件
        if mail_count > 20:
            for i in range(mail_count, mail_count - 20,-1):
                resp, lines, octets = pop3_Server.server.retr(i)
                msg_content = b'\r\n'.join(lines).decode('utf-8')
                # 解析邮件
                msg = Parser().parsestr(msg_content)
                iEmail = Email()
                iEmail.parser_email_from_POP3(msg, i)
                pop3_Server.email_list.append(iEmail)
                print(iEmail.index)
                if iEmail.attachment_files:
                    _thread.start_new_thread(self.Display, (
                        iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                        list(iEmail.attachment_files.keys())[0],iEmail.index))
                else:
                    _thread.start_new_thread(self.Display, (
                        iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time, '',iEmail.index))
                # print('display success!')
        else:
            for i in range(mail_count, 0, -1):
                resp, lines, octets = pop3_Server.server.retr(i)
                msg_content = b'\r\n'.join(lines).decode('utf-8')
                # 解析邮件
                msg = Parser().parsestr(msg_content)
                iEmail = Email()
                iEmail.parser_email_from_POP3(msg, i)
                pop3_Server.email_list.append(iEmail)
                print(iEmail.index)
                if iEmail.attachment_files:
                    _thread.start_new_thread(self.Display, (
                        iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                        list(iEmail.attachment_files.keys())[0],iEmail.index))
                else:
                    _thread.start_new_thread(self.Display, (
                        iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time, '',iEmail.index))

    # 显示邮件table
    def Display(self, title, name, addr, content, etime, attach, index):
        rrow = self.main_ui.tableWidget.rowCount()
        self.main_ui.tableWidget.insertRow(rrow)
        self.main_ui.tableWidget.setItem(rrow, 0, QTableWidgetItem(str(title)))
        self.main_ui.tableWidget.setItem(rrow, 1, QTableWidgetItem(str(name)))
        self.main_ui.tableWidget.setItem(rrow, 2, QTableWidgetItem(str(etime)))
        self.main_ui.tableWidget.setItem(rrow, 3, QTableWidgetItem(addr))
        self.main_ui.tableWidget.setItem(rrow, 4, QTableWidgetItem(str(content)))
        self.main_ui.tableWidget.setItem(rrow, 5, QTableWidgetItem(str(attach)))
        self.main_ui.tableWidget.setItem(rrow, 6, QTableWidgetItem(str(index)))
        # print('Display success!')

    #调用发件箱界面及创建Send_server对象
    def Send(self):
        global is_Idenfy, user, sendUi
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            sendUi = Send_server()
            sendUi.send_ui.label_8.setText(user['user'])
            sendUi.show()

    #浏览附件下载路径
    def load_attach_path(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            filepath = QFileDialog.getExistingDirectory()  # 选择文件夹 返回filepath是一个元组
            self.main_ui.lineEdit_5.setText(QtCore.QCoreApplication.translate("MainWindow", filepath))  #显示文件夹

    #下载附件
    def download_attach(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            global pop3_Server
            path = self.main_ui.lineEdit_5.text()
            index = self.main_ui.lineEdit_6.text()
            for An_email in pop3_Server.email_list:
                if str(An_email.index) == str(index):
                    break
                else:
                    continue
            count = len(An_email.attachment_files)
            for i in range(count):
                filename = list(An_email.attachment_files.keys())[i]
                data = list(An_email.attachment_files.values())[i]
                if str(path) != str(filename):
                    with open(path + '/' + filename, 'wb') as f:
                        # 保存附件
                        f.write(data)
                    f.close()
                else:
                    with open(filename, 'wb') as f:
                        # 保存附件
                        f.write(data)
                    f.close()
            print('Download success at :', path)
            QMessageBox.information(self, "Message", "附件下载成功！", QMessageBox.Yes | QMessageBox.No)

    #将当前邮件的发件人保存至通讯录内
    def save_to_addr_book(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            name = self.main_ui.lineEdit_2.text()
            email = self.main_ui.lineEdit_3.text()
            addr_book = Contact()
            addr_book.read('')
            result = addr_book.add_contact(name,email)
            addr_book.write('')
            if result:
                QMessageBox.information(self, "Message", "保存该联系人成功！", QMessageBox.Yes | QMessageBox.No)
            else:
                QMessageBox.information(self, "Message", "该联系人已存在！", QMessageBox.Yes | QMessageBox.No)

    #删除邮件并刷新
    def delete_email(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            global pop3_Server
            index = self.main_ui.lineEdit_6.text()
            pop3_Server.server.dele(index)
            pop3_Server.server.quit()
            pop3_Server.Connect_Server()
            self.Update()
            QMessageBox.information(self, "Message", "邮件删除成功！", QMessageBox.Yes | QMessageBox.No)

    #按发件人昵称查找
    def search_by_name(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            global pop3_Server
            #创建两个发送者和index下标对应的list用于查找
            name_list = []
            index_list = []
            for iEmail in pop3_Server.email_list:
                name_list.append(iEmail.send_name)
                index_list.append(iEmail.index)
            search_name = self.main_ui.lineEdit_10.text()
            #查找过程
            result_list = []
            for index,name in enumerate(name_list):
               if name.find(search_name) != -1:
                   # print('search_name:',search_name)
                   # print('name_in_list:',name)
                   result_list.append(index)
            result_list = [index_list[x] for x in result_list]
            print(result_list)

            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                    else:
                        continue

    # 按发件人地址查找
    def search_by_addr(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            global pop3_Server
            # 创建两个发送者和index下标对应的list用于查找
            addr_list = []
            index_list = []
            for iEmail in pop3_Server.email_list:
                addr_list.append(iEmail.send_addr)
                index_list.append(iEmail.index)
            search_name = self.main_ui.lineEdit_10.text()
            # 查找过程
            result_list = []
            for index, name in enumerate(addr_list):
                if name.find(search_name) != -1:
                    # print('search_name:',search_name)
                    # print('name_in_list:',name)
                    result_list.append(index)
            result_list = [index_list[x] for x in result_list]
            print(result_list)

            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.send_addr, 'is being displayed!')
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.send_addr, 'is being displayed!')
                    else:
                        continue

    # 按发信时间查找
    def search_by_time(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            global pop3_Server
            # 创建两个发送者和index下标对应的list用于查找
            time_list = []
            index_list = []
            for iEmail in pop3_Server.email_list:
                time_list.append(iEmail.email_time)
                index_list.append(iEmail.index)
            search_name = self.main_ui.lineEdit_10.text()
            # 查找过程
            result_list = []
            for index, name in enumerate(time_list):
                if name.find(search_name) != -1:
                    # print('search_name:',search_name)
                    # print('name_in_list:',name)
                    result_list.append(index)
            result_list = [index_list[x] for x in result_list]
            print(result_list)

            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.email_time, 'is being displayed!')
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.email_time, 'is being displayed!')
                    else:
                        continue

    # 按邮件标题查找
    def search_by_title(self):
        global pop3_Server, is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        # 创建两个发送者和index下标对应的list用于查找
        else:
            title_list = []
            index_list = []
            for iEmail in pop3_Server.email_list:
                title_list.append(iEmail.title)
                index_list.append(iEmail.index)
            search_name = self.main_ui.lineEdit_10.text()
            # 查找过程
            result_list = []
            for index, name in enumerate(title_list):
                if name.find(search_name) != -1:
                    # print('search_name:',search_name)
                    # print('name_in_list:',name)
                    result_list.append(index)
            result_list = [index_list[x] for x in result_list]
            print(result_list)

            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                    else:
                        continue

    # 按邮件标题排序(降序)
    def sort_by_title(self):
        # 排序过程:中英文分别排序
        global  is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            contain_chinese_dic = {}
            not_contain_dic = {}
            for iEmail in pop3_Server.email_list:
                if '\u4e00' <= iEmail.title[0] <= '\u9fa5':
                    contain_chinese_dic[iEmail.index] = str(iEmail.title)
                else:
                    not_contain_dic[iEmail.index] = str(iEmail.title)
            contain_chinese_list = sorted(contain_chinese_dic.items(),
                                          key=lambda kv: [pinyin(kv[1], style=Style.TONE3), kv[0]])
            not_contain_list = sorted(not_contain_dic.items(), key=lambda kv: [kv[1], kv[0]])
            result_list = not_contain_list + contain_chinese_list
            print(result_list)

            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index, title in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                            time.sleep(0.1)
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time, '',
                                iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                            time.sleep(0.1)
                        break

    # 按发信人排序(降序)
    def sort_by_sender(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        # 排序过程:中英文分别排序
        else:
            contain_chinese_dic = {}
            not_contain_dic = {}
            for iEmail in pop3_Server.email_list:
                if '\u4e00' <= iEmail.send_name[0] <= '\u9fa5':
                    contain_chinese_dic[iEmail.index] = str(iEmail.send_name)
                else:
                    not_contain_dic[iEmail.index] = str(iEmail.send_name)
            contain_chinese_list = sorted(contain_chinese_dic.items(),
                                          key=lambda kv: [pinyin(kv[1], style=Style.TONE3), kv[0]])
            not_contain_list = sorted(not_contain_dic.items(), key=lambda kv: [kv[1], kv[0]])
            result_list = not_contain_list + contain_chinese_list
            print(result_list)
            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for items in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(items[0]):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        break

    # 按发信时间排序(升序)
    def sort_by_time_reverse(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        else:
            unsort_dic = {}
            for iEmail in pop3_Server.email_list:
                unsort_dic[iEmail.index] = str(iEmail.index)
            result_list = sorted(unsort_dic.items(), key=lambda kv: [kv[1], kv[0]])
            print(result_list)
            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for items in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(items[0]):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        break

    # 按邮件标题排序(升序）
    def sort_by_title_reverse(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        # 排序过程:中英文分别排序
        else:
            contain_chinese_dic = {}
            not_contain_dic = {}
            for iEmail in pop3_Server.email_list:
                if '\u4e00' <= iEmail.title[0] <= '\u9fa5':
                    contain_chinese_dic[iEmail.index] = str(iEmail.title)
                else:
                    not_contain_dic[iEmail.index] = str(iEmail.title)
            contain_chinese_list = sorted(contain_chinese_dic.items(),
                                          key=lambda kv: [pinyin(kv[1], style=Style.TONE3), kv[0]])
            not_contain_list = sorted(not_contain_dic.items(), key=lambda kv: [kv[1], kv[0]])
            result_list = not_contain_list + contain_chinese_list
            result_list.reverse()
            print(result_list)
            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for index, title in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(index):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                            time.sleep(0.1)
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time, '',
                                iEmail.index))
                            print(iEmail.index, iEmail.title, 'is being displayed!')
                            time.sleep(0.1)
                        break

    # 按邮件发件人排序(升序）
    def sort_by_sender_reverse(self):
        global is_Idenfy
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
        # 排序过程:中英文分别排序
        else:
            contain_chinese_dic = {}
            not_contain_dic = {}
            for iEmail in pop3_Server.email_list:
                if '\u4e00' <= iEmail.send_name[0] <= '\u9fa5':
                    contain_chinese_dic[iEmail.index] = str(iEmail.send_name)
                else:
                    not_contain_dic[iEmail.index] = str(iEmail.send_name)
            contain_chinese_list = sorted(contain_chinese_dic.items(),
                                          key=lambda kv: [pinyin(kv[1], style=Style.TONE3), kv[0]])
            not_contain_list = sorted(not_contain_dic.items(), key=lambda kv: [kv[1], kv[0]])
            result_list = not_contain_list + contain_chinese_list
            result_list.reverse()
            print(result_list)
            # 清空显示列表
            allrownum = self.main_ui.tableWidget.rowCount()
            for i in range(allrownum):
                self.main_ui.tableWidget.removeRow(0)
            # 显示到table中
            for items in result_list:
                for iEmail in pop3_Server.email_list:
                    if str(iEmail.index) == str(items[0]):
                        if iEmail.attachment_files:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                list(iEmail.attachment_files.keys())[0], iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        else:
                            _thread.start_new_thread(self.Display, (
                                iEmail.title, iEmail.send_name, iEmail.send_addr, iEmail.content, iEmail.email_time,
                                '', iEmail.index))
                            print(iEmail.index, iEmail.send_name, 'is being displayed!')
                            time.sleep(0.1)
                        break

    #获取转发信息
    def transfer_email(self):
        global pop3_Server,is_Idenfy, user,  transUi
        #sendUi = Send_server()
        #获取当前index对应的email对象
        if not is_Idenfy:
            QMessageBox.about(self, "Message", "请先登录")
            transUi.close()
        else:
                index = self.main_ui.lineEdit_6.text()
                for An_email in pop3_Server.email_list:
                    if str(An_email.index) == str(index):
                        break
                    else:
                        continue



                else:
                    write_tag = False
                    try:
                        smtp_server = 'smtp.163.com'
                        send_name = transUi.send_ui.lineEdit.text()
                        recv_mail = transUi.send_ui.lineEdit_2.text()
                        smtp_tail = user['pop3']
                        smpt_tail_list = smtp_tail.split('.')
                        tail = '@' + smpt_tail_list[1] + '.' + smpt_tail_list[2]
                        fromEmailAddr = pop3_Server.user_mail + tail  # 邮件发送方邮箱地址
                        print(fromEmailAddr)
                        password = pop3_Server.password  # 密码(部分邮箱为授权码)
                        toEmailAddrs = [recv_mail]  # 邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发
                        write_tag = True
                        trans_content_header = '-------- 转发邮件信息 --------' + '\n' \
                                               + '发件人：' + send_name + '<'+ str(An_email.send_addr) +'>' + '\n'\
                                               + '发送日期：' + str(An_email.email_time) +'\n'\
                                               + '收件人：' + str(fromEmailAddr) + '\n'\
                                               + '主题：' + str(An_email.title) + '\n'
                        trans_content = trans_content_header + An_email.content
                        attachment_files = An_email.attachment_files
                        title =  An_email.title
                    except:
                        QMessageBox.about(self, "Message", "填写完整信息")

                    if write_tag:
                        self.Trans_Email(smtp_server, send_name, recv_mail, title,
                                                  trans_content, '', fromEmailAddr, password, toEmailAddrs,
                                                  attachment_files)

    #连接SMTP服务器并转发
    def Trans_Email(self, smtp_server, send_name, recv_mail, send_title, send_text, send_attach, fromEmailAddr,
                    password, toEmailAddrs, tran_attach_files):
        try:
            # 设置email信息
            # ---------------------------发送带附件邮件-----------------------------
            # 邮件内容设置
            message = MIMEMultipart()
            # 邮件主题
            message['Subject'] = send_title
            # 发送方信息
            message['From'] = formataddr([send_name, fromEmailAddr])
            # 接受方信息
            message['To'] = toEmailAddrs[0]
            # 邮件正文内容
            message.attach(MIMEText(send_text, 'plain', 'utf-8'))

            # 构造附件
            if tran_attach_files:
                for i in range(len(tran_attach_files)):
                    att1 = MIMEText(list(tran_attach_files.values())[i], 'base64', 'utf-8')
                    att1['Content-type'] = 'application/octet-stream'
                    att1['Content-Disposition'] = 'attachment; filename=' + '"' + list(tran_attach_files.keys())[
                        0] + '"'
                    message.attach(att1)
            elif send_attach:
                att1 = MIMEText(open(send_attach, 'rb').read(), 'base64', 'utf-8')
                att1['Content-type'] = 'application/octet-stream'
                att1['Content-Disposition'] = 'attachment; filename=' + '"' + str(str(send_attach).split("/")[-1]) + '"'
                message.attach(att1)
            # ---------------------------------------------------------------------

            # 登录并发送邮件
            try:
                server = smtplib.SMTP(smtp_server)  # 163邮箱服务器地址，端口默认为25
                server.login(fromEmailAddr, password)
                server.sendmail(fromEmailAddr, toEmailAddrs, message.as_string())
                print('success')
                QMessageBox.information(self, "Message", "发送成功！", QMessageBox.Yes | QMessageBox.No)
                # self.close()
                server.quit()
            except smtplib.SMTPException as e:
                print("error:", e)
        except:
            QMessageBox.critical(self, "失败", "pop验证失败", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

# 登录界面ui类
class loginWin(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.login_ui = idenfy.Ui_MainWindow()
        self.login_ui.setupUi(self)
        self.login_ui.pushButton_2.clicked.connect(self.close)
        self.login_ui.pushButton.clicked.connect(self.Login)
        self.login_ui.lineEdit_2.setEchoMode(QLineEdit.Password)

    def Login(self):
        global is_Idenfy, pop3_Server, user
        print("login....")
        user['user'] = str(self.login_ui.lineEdit.text())
        user['pass'] = str(self.login_ui.lineEdit_2.text())
        user['pop3'] = str(self.login_ui.lineEdit_3.text())
        # print(user)
        try:
            pop3_Server = Receive_server(user)
            print("连接成功.......")
            QMessageBox.about(self, "连接成功", "验证成功！")
            self.close()
            is_Idenfy = True
        except:
            print("pop验证失败！")
            QMessageBox.critical(self, "失败", "pop验证失败", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)


if __name__ == '__main__':
    ui = mainWin()
    loginUi = loginWin()
    global transUi
    transUi = Transmit_server(ui)
    ui.main_ui.pushButton.clicked.connect(loginUi.show)
    ui.main_ui.pushButton_12.clicked.connect(transUi.show)
    ui.main_ui.pushButton_3.clicked.connect(QApplication.quit)
    ui.show()

    sys.exit(app.exec_())
