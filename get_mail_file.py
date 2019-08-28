import poplib
import time
import email
import os
import random
import pymysql
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


# 传入的两个参数，一个是data_to_excel写入excel的数据，是一个list，list里面包含字典
# 字典就是要写入excel的数据
# 第二个参数也是一个字典，data_for_xls包含head（xls的表头），sheet（xls表的名称）
# xls_file_name（xls文件的文件名）
def save_to_excel(data_to_excel, data_for_xls):
    try:
        if os.path.isfile(data_for_xls['xls_file_name']):
            old_file = xlrd.open_workbook(data_for_xls['xls_file_name'])
            try:
                sheet_name = old_file.sheet_names()[0]
            except Exception as e:
                sheet_name = ''
            if sheet_name == data_for_xls['sheet']:
                # 获取表格
                sheet = old_file.sheet_by_name(data_for_xls['sheet'])
                # 获取表格已有的有效行数
                i = sheet.nrows
                # 判断表头是否一致 pass
                if operator.eq(sheet.row_values(0), data_for_xls['head']):
                    # 在原有的基础上直接添加数据
                    workbook = copy(old_file)
                    sheet = workbook.get_sheet(0)
                else:
                    # 这种情况表示已存在的xls文件的表头不对，进行修改 pass
                    workbook = copy(old_file)
                    sheet = workbook.get_sheet(0)
                    for h in range(len(data_for_xls['head'])):
                        sheet.write(0, h, data_for_xls['head'][h])
            else:
                # 没有邮件信息这个表，在邮箱数据.xls这个表基础上新建邮件信息这个sheet
                workbook = copy(old_file)
                sheet = workbook.add_sheet(data_for_xls['sheet'])
                # 添加excel头
                for h in range(len(data_for_xls['head'])):
                    sheet.write(0, h, data_for_xls['head'][h])
                i = 1
        else:
            # 不存在这个文件，直接创建邮箱数据.xls这个excel
            workbook = xlwt.Workbook(encoding='utf-8')
            sheet = workbook.add_sheet(data_for_xls['sheet'])
            # 添加excel头
            for h in range(len(data_for_xls['head'])):
                sheet.write(0, h, data_for_xls['head'][h])
            i = 1

        # 写入数据到excel文件
        for each in data_to_excel:
            each_sort = sorted(each.items(), key=lambda d:d[0])
            for j in range(len(data_for_xls['head'])):
                sheet.write(i, j, each_sort[j][-1])
            i += 1
        workbook.save(data_for_xls['xls_file_name'])
        print('邮箱数据写入excel成功！')
    except Exception as e:
        print('因为 %s ，邮箱数据写入excel失败！' % e)


class GetMailFiles():
    def __init__(self):
        self.email_name = 'zhangym1@chinatelecom.cn'
        self.password = 'phoenixnash13'
        self.pop3_server = 'pop.chinatelecom.cn'
        self.root_path =os.getcwd().replace("\\", '/')

        # 初始化mysql数据库连接
        try:
            mysql_host = '127.0.0.1'
            mysql_user = 'zhang'
            mysql_pass = '19940327'
            mysql_db = 'chinatelecom_mail'
            self.conn = pymysql.connect(host=mysql_host, user=mysql_user, passwd=mysql_pass, db=mysql_db)
            self.cursor = self.conn.cursor()
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
                    write_file_name = re.sub('[\/:*?"<>→]', '-', str(filename))
                    with open(path_name + '\\' + str(write_file_name), 'wb') as f:
                        f.write(file_data)
                    print('附件-%s-保存成功！' % filename)
                    return_data.append(path_name + str(write_file_name))
            except Exception as e:
                print('附件保存失败！因为%s' %(e))
        return return_data

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
        for i in range(len(mails), 832, -1):
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
                mail_file_path = self.get_mail_file_data(mail_content, path_name)
                mail_item = {}
                try:
                    sql_select = "select mail_id from chinatelecom_mail where mail_id = %d" % i
                    self.cursor.execute(sql_select)
                    if self.cursor.rowcount == 0:
                    # 保存邮件数据到mongo数据库，判断数据库中是否已经存在该记录了
                    # if not self.mongo_object.find({'mail_subject': mail_subject}).count():
                        mail_date = time.strptime(mail_content.get('Date')[0:24], '%a, %d %b %Y %H:%M:%S')
                        mail_date_format = time.strftime('%Y%m%d %H:%M:%S', mail_date)

                        # 因为存在多个收件人的情况，所以这里比获取发件人复杂
                        mail_to_addrs = mail_content.get('To').split(',')
                        mail_to_addrs_list = list(map(self.decode_str, mail_to_addrs))

                        # 记录邮件序号
                        mail_item['a_mail_number'] = i
                        for i in range(len(mail_to_addrs)):
                            mail_to_addrs_list[i] = mail_to_addrs_list[i] + mail_to_addrs[i].split(' ')[-1]

                        mail_item['b_mail_subject'] = mail_subject
                        mail_item['c_mail_from_addr'] = mail_from_addr
                        mail_item['d_mail_to_addr'] = ';'.join(mail_to_addrs_list)
                        mail_item['e_mail_date_format'] = mail_date_format
                        mail_item['f_mail_file_path'] = ';'.join(mail_file_path)

                        self.cursor.execute(r'insert ignore into chinatelecom_mail values(%s,%s,%s,%s,%s,%s)',
                                            [mail_item['a_mail_number'], mail_item['b_mail_subject'],
                                             mail_item['c_mail_from_addr'], mail_item['d_mail_to_addr'],
                                             mail_item['e_mail_date_format'], mail_item['f_mail_file_path']])
                        self.conn.commit()
                        print('邮件 %s 的相关数据已经保存到了数据库了' % mail_subject)
                        data_to_excel.append(mail_item)
                    else:
                        print('数据库中已存在-%s-的邮件数据库了！' % mail_subject)
                except Exception as e:
                    print('因为 %s，保存邮件数据到数据库出错！' % e)
                time.sleep(random.randint(1, 5) + random.randint(4, 8) / 10)
        print('邮件附件下载完毕！')

        # 保存数据到excel文件
        data_for_xls = {}
        data_for_xls['head'] = ['邮件编号', '邮件标题', '发送人', '收件人', '邮件日期', '邮件附件位置']
        data_for_xls['sheet'] = '邮件信息'
        data_for_xls['xls_file_name'] = '邮箱数据.xls'
        save_to_excel(data_to_excel, data_for_xls)
        server.quit()


if __name__ == '__main__':
    sys.stdout = Logger('all.log', sys.stdout)
    GetMailFiles = GetMailFiles()
    GetMailFiles.mail_main()