import poplib
import time
import email
import os
import random
import pymongo
import re
import sys
import xlwt, xlrd
import operator
from email.parser import Parser
from email.utils import parseaddr
from email.header import decode_header
from xlutils.copy import copy


class Logger(object):
    def __init__(self, filename='all.log', stream=sys.stdout):
        self.terminal = stream
        self.filename = filename

    def write(self, message):
        self.terminal.write(message)
        with open(self.filename, 'a', encoding='utf8') as f:
            f.write(message)

    def flush(self):
        pass


class GetMailFiles():
    def __init__(self):
        self.email_name = 'zhangym1@chinatelecom.cn'
        self.password = 'phoenixnash13'
        self.pop3_server = 'pop.chinatelecom.cn'
        self.root_path =os.getcwd().replace("\\", '/')

        # 初始化mongo数据库连接
        try:
            host = '127.0.0.1'
            port = 27017
            db_name = 'chinatelecom_mail'
            sheet_name = 'chinatelecom_mail'
            mongo_client = pymongo.MongoClient(host=host, port=port)
            mongo_db = mongo_client[db_name]
            self.mongo_object = mongo_db[sheet_name]
        except Exception as e:
            print('因为 %s，数据库连接初始化失败！' % e)

    # subject传输的时候进行了编码，所以要进行解码操作，才能正常显示
    def decode_str(self, input_encode_string):
        value, charset = decode_header(input_encode_string)[0]
        try:
            if charset:
                value = value.decode(charset)
        except Exception as e:
            print('因为 %s ，解码失败！' % e)
            value = ''
        return value

    def get_mail_file_data(self, mail_content, path_name):
        return_data = []
        for file in mail_content.walk():
            file_name = file.get_filename()
            try:
                if file_name:
                    # 将str类型的文件名转换为Header类型的数据
                    h = email.header.Header(file_name)
                    # 对附件名称进行解码
                    dh = email.header.decode_header(h)
                    # dh[0][0]是附件名本身，dh[0][1]是编码的规则类似于us-ascii
                    filename = dh[0][0]
                    if dh[0][1]:
                        filename = self.decode_str(str(filename, dh[0][1]))
                    file_data = file.get_payload(decode=True)

                    # 判断文件夹是否存在
                    if not os.path.exists(path_name):
                        os.makedirs(path_name)
                    return_data.append(file_data)
                    write_file_name = re.sub('[\/:*?"<>→]', '-', str(filename))
                    with open(path_name + '\\' + str(write_file_name), 'wb') as f:
                        f.write(file_data)
                    print('附件-%s-保存成功！' % filename)
            except Exception as e:
                print('附件保存失败！因为%s' %(e))
        return return_data

    def save_to_excel(self, data_to_excel):
        head = ['邮件编号', '邮件标题', '发送人', '收件人', '邮件日期']
        try:
            if os.path.isfile('邮箱数据.xls'):
                old_file = xlrd.open_workbook('邮箱数据.xls')
                try:
                    sheet_name = old_file.sheet_names()[0]
                except Exception as e:
                    sheet_name = ''
                if sheet_name == '邮件信息':
                    # 获取表格
                    sheet = old_file.sheet_by_name('邮件信息')
                    # 获取表格已有的有效行数
                    i = sheet.nrows
                    # 判断表头是否一致 pass
                    if operator.eq(sheet.row_values(0), head):
                        # 在原有的基础上直接添加数据
                        workbook = copy(old_file)
                        sheet = workbook.get_sheet(0)
                    else:
                        # 这种情况表示已存在的xls文件的表头不对，进行修改
                        workbook = copy(old_file)
                        sheet = workbook.get_sheet(0)
                        for h in range(len(head)):
                            sheet.write(0, h, head[h])
                else:
                    # 没有邮件信息这个表，在邮箱数据.xls这个表基础上新建邮件信息这个sheet
                    workbook = copy(old_file)
                    sheet = workbook.add_sheet('邮件信息')
                    # 添加excel头
                    for h in range(len(head)):
                        sheet.write(0, h, head[h])
                    i = 1
            else:
                # 不存在这个文件，直接创建邮箱数据.xls这个excel
                workbook = xlwt.Workbook(encoding='utf-8')
                sheet = workbook.add_sheet('邮件信息')
                # 添加excel头
                for h in range(len(head)):
                    sheet.write(0, h, head[h])
                i = 1

            # 写入数据到excel文件
            for each in data_to_excel:
                sheet.write(i, 0, each['mail_number'])
                sheet.write(i, 1, each['mail_subject'])
                sheet.write(i, 2, each['mail_from_addr'])
                sheet.write(i, 3, each['mail_to_addr'])
                sheet.write(i, 4, each['mail_date_format'])
                i += 1
            workbook.save('邮箱数据.xls')
            print('邮箱数据写入excel成功！')
        except Exception as e:
            print('因为 %s ，邮箱数据写入excel失败！' % e)

    def mail_main(self):
        server = poplib.POP3_SSL(self.pop3_server, 995)
        server.set_debuglevel(1)

        # 身份认证
        server.user(self.email_name)
        server.pass_(self.password)
        print('Messages: %s. Size: %s' % server.stat())
        resp, mails, octets = server.list()

        # 保存需要写入excel的数据，爬取完毕之后一次性写入excel
        data_to_excel = []
        # 遍历邮件
        for i in range(len(mails), 0, -1):
            # lines存储邮件的原始文件
            resp, lines, octets = server.retr(i)
            mail_content = b'\r\n'.join(lines).decode('utf8')
            mail_content = Parser().parsestr(mail_content)
            hdr1, mail_from_addr = parseaddr(mail_content.get('From'))
            mail_subject = self.decode_str(mail_content.get('Subject'))

            if mail_subject == '':
                print('第 %s 封邮件解码失败！跳过!' % i)
                continue
            if mail_from_addr == 'chendj@spdi.com.cn' and '光缆工程' in mail_subject:

                # 去掉主题中一些特殊字符，含有这些特殊字符会无法创建文件夹
                mail_subject_str = re.sub('[\/:*?"<>→]', '-', mail_subject)
                path_name = ''
                if '室分' in mail_subject or '室内覆盖' in mail_subject:
                    path_name = self.root_path + '/室分/' + mail_subject_str + '/'
                elif '主干' in mail_subject:
                    path_name = self.root_path + '/主干/' + mail_subject_str + '/'
                else:
                    path_name = self.root_path + '/其他/' + mail_subject_str + '/'

                # 判断邮件附件是否已经保存了
                if os.path.exists(path_name):
                    print('%s-邮件已爬取！跳过！' % mail_subject)
                    continue
                print('正在获取%s的附件文件！' % mail_subject)

                # 获取邮件的附件数据，并保存到本地，同时返回数据，方便写入数据库
                mail_file_data = self.get_mail_file_data(mail_content, path_name)
                mail_item = {}
                try:
                    # 保存邮件数据到mongo数据库，判断数据库中是否已经存在该记录了
                    if not self.mongo_object.find({'mail_subject': mail_subject}).count():
                        mail_date = time.strptime(mail_content.get('Date')[0:24], '%a, %d %b %Y %H:%M:%S')
                        mail_date_format = time.strftime('%Y%m%d %H:%M:%S', mail_date)

                        # 因为存在多个收件人的情况，所以这里比获取发件人复杂
                        mail_to_addrs = mail_content.get('To').split(',')
                        mail_to_addrs_list = list(map(self.decode_str, mail_to_addrs))

                        # 记录邮件序号
                        mail_item['mail_number'] = i
                        for i in range(len(mail_to_addrs)):
                            mail_to_addrs_list[i] = mail_to_addrs_list[i] + mail_to_addrs[i].split(' ')[-1]

                        mail_item['mail_subject'] = mail_subject
                        mail_item['mail_from_addr'] = mail_from_addr
                        mail_item['mail_to_addr'] = mail_to_addrs_list
                        mail_item['mail_date_format'] = mail_date_format
                        mail_item['mail_file_data'] = mail_file_data

                        # 保存到mongo数据库
                        mongo_data_dict = dict(mail_item)
                        self.mongo_object.insert(mongo_data_dict)

                        # 构建保存在excel的数据
                        del mail_item['mail_file_data']
                        data_to_excel.append(mail_item)
                    else:
                        print('数据库中已存在-%s-的邮件数据库了！' % mail_subject)
                except Exception as e:
                    print('因为 %s，保存邮件数据到数据库出错！' % e)
                time.sleep(random.randint(1, 5) + random.randint(4, 8) / 10)
        print('邮件附件下载完毕！')

        # 保存数据到excel文件
        self.save_to_excel(data_to_excel)
        server.quit()


if __name__ == '__main__':
    sys.stdout = Logger('all.log', sys.stdout)
    GetMailFiles = GetMailFiles()
    GetMailFiles.mail_main()
