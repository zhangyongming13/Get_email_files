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
import zipfile, rarfile
import shutil
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


def un_zip_rar(path_names):
    # 即使传进来的文件路径有多个，但是预算文件不存在两个的情况，所以预算文件的路径就一个
    budget_file_path = ''
    for path_name in path_names.split(';'):
        # 将windows下路径分隔符反斜杠全部替换成斜杠，避免兼容性问题
        path_name = path_name.replace('\\', '/')
        un_file_path = '/'.join(path_name.split('/')[0:-1]) + '/'

        if path_name.split('.')[-1].lower() == 'zip':
            try:
                with zipfile.ZipFile(path_name, 'r') as f:
                    for file_name in f.namelist():
                        # zipfile这个包默认编码是cp437，所以先转换为unicode再进行gbk编码，这样中文就不会乱码
                        # split是去除压缩包里面的文件夹，达到只解压文件的目的
                        right_file_name = file_name.encode('cp437').decode('gbk').split('/')[-1]

                        # 判断这个name名字是文件类型还是文件夹路径
                        if re.match(r'.+\..+', right_file_name):
                            right_path_file_name = un_file_path + right_file_name

                            if '预算' in right_path_file_name:
                                budget_file_path = right_path_file_name
                            with open(right_path_file_name, 'wb') as file:
                                with f.open(file_name, 'r') as origin_file:
                                    shutil.copyfileobj(origin_file, file)
                print('%s 解压完毕！' % ''.join(path_name.split('/')[-1]))
            except Exception as e:
                print('%s 解压失败，原因：%s' % (''.join(path_name.split('/')[-1]), e))
        elif path_name.split('.')[-1].lower() == 'rar':
            try:
                with rarfile.RarFile(path_name) as rar_file:
                    for name in rar_file.namelist():

                        # 判断这个name名字是文件类型还是文件夹路径
                        if re.match(r'.+\..+', name):
                            right_path_file_name = un_file_path + name.split('/')[-1]

                            if '预算' in right_path_file_name:
                                budget_file_path = right_path_file_name
                            with open(right_path_file_name, 'wb') as file:
                                with rar_file.open(name, 'r') as origin_file:
                                    shutil.copyfileobj(origin_file, file)
                print('%s 解压完毕！' % ''.join(path_name.split('/')[-1]))
            except Exception as e:
                print('rar文件解压失败，原因：%s' % e)
        else:
            print('错误的压缩包文件类型!')
    return budget_file_path


def save_end_to_settings_file(settings_file_path, end):
    with open(settings_file_path, 'r', encoding='utf8') as f:
        old_text = f.readlines()
    with open(settings_file_path, 'w', encoding='utf8') as f:
        for line in old_text:
            if 'end=' in line:
                line = 'end=' + str(end) + '\n'
            f.write(line)
    print('写入成功！')


def get_budget_from_excel(budget_file_path):
    budget_dict = {}
    if budget_file_path == '':
        budget_dict['tax_deduction_price'] = 0
        budget_dict['value_added_tax'] = 0
        budget_dict['tax_included_price'] = 0
        print('预算文件未找到，请手动记录预算情况！')
    else:
        try:
            budget_workbook = xlrd.open_workbook(budget_file_path)
            budget_sheet = budget_workbook.sheet_by_name('表一')
            # excel表里面有效的行数和列数
            # print(budget_sheet.nrows)
            # print(budget_sheet.ncols)

            # round进行四舍五入的操作
            for row in range(budget_sheet.nrows-1, 0, -1):
                total_row = budget_sheet.row_values(row)
                if '总计' in total_row:
                    tax_deduction_price = round(total_row[-4], 2)  # 除税价
                    value_added_tax = round(total_row[-3], 2)  # 增值税
                    tax_included_price = round(total_row[-2], 2)  # 含税价
                    budget_dict['tax_deduction_price'] = tax_deduction_price
                    budget_dict['value_added_tax'] = value_added_tax
                    budget_dict['tax_included_price'] = tax_included_price
                    break

        except Exception as e:
            print('因为 %s 的原因，预算读取失败！' % e)
    return budget_dict


class GetMailFiles():
    def __init__(self):
        self.settings_file_path = 'settings.txt'
        settings_data = self.get_settings_from_txt(self.settings_file_path)

        # 邮件服务器参数初始化
        self.email_name = settings_data['email_name']
        self.password = settings_data['password']
        self.pop3_server = settings_data['pop3_server']

        self.root_path =os.getcwd().replace("\\", '/')

        self.data_for_xls = {}
        self.data_for_xls['head'] = ['邮件编号', '邮件标题', '发送人', '收件人', '邮件日期', '邮件附件位置', '除税价', '增值税', '含税价']
        self.data_for_xls['sheet'] = '邮件信息'
        self.data_for_xls['xls_file_name'] = '邮箱数据.xls'
        # 保存需要写入excel的数据，爬取完毕之后一次性写入excel
        self.data_to_excel = []

        # 初始化mysql数据库连接
        try:
            self.conn = pymysql.connect(host=settings_data['mysql_host'], user=settings_data['mysql_user'],
                                        passwd=settings_data['mysql_pass'], db=settings_data['mysql_db'])
            self.cursor = self.conn.cursor()
        except Exception as e:
            print('因为 %s，数据库连接初始化失败！' % e)

    # 将存储于chinatelecom_mail的邮箱数据写入到chinatelecom_mails_files
    # 前期mail_id不是自增的，而是邮件在服务器的序号，这样的话会出现问题，服务器删除之后，新邮件数据写入的时候
    # 可能会造成mail_id（主键重复）
    def move_database_data(self):
        sql_select = "select * from chinatelecom_mails_files"
        self.cursor.execute(sql_select)
        result = self.cursor.fetchall()
        for row in result:
            print(row)
            self.cursor.execute(
                r'insert ignore into chinatelecom_mail_files values(%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                [0, row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]])
        self.conn.commit()
        print('转移数据成功！')

    def get_settings_from_txt(self, settings_file_path):
        return_data = {}
        try:
            with open(settings_file_path, 'r', encoding='utf8') as f:
                settings = f.readlines()
            for i in settings:
                if 'email_name' in i:
                    return_data['email_name'] = i.split('\'')[1]
                    continue
                if 'password' in i:
                    return_data['password'] = i.split('\'')[1]
                    continue
                if 'pop3_server' in i:
                    return_data['pop3_server'] = i.split('\'')[1]
                    continue
                if 'end=' in i:
                    return_data['end'] = int(i.split('=')[1])
                    continue
                if 'mysql_host' in i:
                    return_data['mysql_host'] = i.split('\'')[1]
                    continue
                if 'mysql_user' in i:
                    return_data['mysql_user'] = i.split('\'')[1]
                    continue
                if 'mysql_pass' in i:
                    return_data['mysql_pass'] = i.split('\'')[1]
                    continue
                if 'mysql_db' in i:
                    return_data['mysql_db'] = i.split('\'')[1]
                    continue
            if len(return_data) < 8:
                print('请检查配置文件是否配置完毕！')
                return_data.clear()
        except Exception as e:
            print('因为%s配置数据读取错误！' % e)
        return return_data

    # 保存到爬取到的邮件的地方，下次爬取就是从最新到end这里，而不是0
    def save_end_to_settings_file(self, settings_file_path, end):
        with open(settings_file_path, 'r', encoding='utf8') as f:
            old_text = f.readlines()
        with open(settings_file_path, 'w', encoding='utf8') as f:
            for line in old_text:
                if 'end=' in line:
                    line = 'end=' + str(end) + '\n'
                f.write(line)
        print('配置文件中end写入成功！')

    # subject传输的时候进行了编码，所以要进行解码操作，才能正常显示
    def decode_str(self, input_encode_string):
        value, charset = decode_header(input_encode_string)[0]
        try:
            if charset:
                # 邮件中可能会有非法字符，所以添加ingore忽略掉
                value = value.decode(charset, 'ignore')
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
        try:
            server = poplib.POP3_SSL(self.pop3_server, 995)
            server.set_debuglevel(1)

            # 身份认证
            server.user(self.email_name)
            server.pass_(self.password)
            print('Messages: %s. Size: %s' % server.stat())
            resp, mails, octets = server.list()

            # 记录又多少封邮件是符合条件的
            flag = 0
            # 遍历邮件
            for i in range(len(mails), 0, -1):

                # lines存储邮件的原始文件
                resp, lines, octets = server.retr(i)
                try:
                    mail_content = b'\r\n'.join(lines).decode('ANSI')
                except:
                    print('邮件编码不是ansi，跳过！')
                    continue
                mail_content = Parser().parsestr(mail_content)
                hdr1, mail_from_addr = parseaddr(mail_content.get('From'))
                mail_subject = self.decode_str(mail_content.get('Subject'))

                if mail_subject == '':
                    print('第 %s 封邮件解码失败！跳过!' % i)
                    continue
                if mail_from_addr == 'chendj@spdi.com.cn' and '光缆工' in mail_subject:

                    # 去掉主题中一些特殊字符，含有这些特殊字符会无法创建文件夹
                    mail_subject_str = re.sub('[\/:*?"<>→]', '-', mail_subject)
                    path_name = ''
                    if '室分' in mail_subject or '室内覆盖' in mail_subject:
                        path_name = self.root_path + '/室分/' + mail_subject_str + '/'
                    elif '主干' in mail_subject:
                        path_name = self.root_path + '/主干/' + mail_subject_str + '/'
                    else:
                        path_name = self.root_path + '/其他/' + mail_subject_str + '/'

                    # 判断邮件附件是否已经保存了,由于都是顺序爬取的，遇到有保存到本地的话，后天的视为都已经爬取过了
                    if os.path.exists(path_name):
                        print('%s-邮件已爬取！跳过！' % mail_subject)
                        break
                    print('正在获取%s的附件文件！' % mail_subject)

                    # 获取邮件的附件数据，并保存到本地，同时返回数据，方便写入数据库
                    mail_file_path = self.get_mail_file_data(mail_content, path_name)
                    mail_item = {}
                    try:
                        # 获取邮件中相关数据
                        mail_date = time.strptime(mail_content.get('Date')[0:24], '%a, %d %b %Y %H:%M:%S')
                        mail_date_format = time.strftime('%Y%m%d %H:%M:%S', mail_date)
                        # 因为存在多个收件人的情况，所以这里比获取发件人复杂
                        mail_to_addrs = mail_content.get('To').split(',')
                        mail_to_addrs_list = list(map(self.decode_str, mail_to_addrs))
                        # 记录邮件序号，并构建邮件收件人数据
                        mail_item['a_mail_number'] = i
                        for index in range(len(mail_to_addrs)):
                            mail_to_addrs_list[index] = mail_to_addrs_list[index] + mail_to_addrs[index].split(' ')[-1]

                        mail_item['b_mail_subject'] = mail_subject
                        mail_item['c_mail_from_addr'] = mail_from_addr
                        mail_item['d_mail_to_addr'] = ';'.join(mail_to_addrs_list)
                        mail_item['e_mail_date_format'] = mail_date_format
                        mail_item['f_mail_file_path'] = ';'.join(mail_file_path)

                        # 进行文件压缩包解压以及预算数据的获取
                        budget_path_file = un_zip_rar(mail_item['f_mail_file_path'])
                        budget_dict = get_budget_from_excel(budget_path_file)

                        # 分别是除税价、增值税和含税价
                        mail_item['g_tax_deduction_price'] = budget_dict['tax_deduction_price']
                        mail_item['h_value_added_tax'] = budget_dict['value_added_tax']
                        mail_item['i_tax_included_price'] = budget_dict['tax_included_price']

                        # 尝试从数据库中获取有没有当前mail_subject的邮件数据，replace_or_insert用来记录
                        # replace into是replace还是insert
                        sql_select = "select mail_id from chinatelecom_mail_files where mail_subject = " + '\'' + mail_subject + '\''
                        replace_or_insert = self.cursor.execute(sql_select)
                        if replace_or_insert == 1:
                            print('数据库原有 %s 的数据，现更新数据库内的数据！' % mail_subject)
                        elif replace_or_insert == 0:
                            print('邮件 %s 的相关数据正在保存到数据库！' % mail_subject)
                        # 使用replace语句代替insert ignore语句，这样的话如果要插入的数据存在主键或者唯一索引相同的情况，
                        # 这里设置了mail_id主键，自增不用管，mail_subject唯一索引，如果插入的数据mail_subject和原来有相同的情况
                        # 这样的话，就先删除原有数据，然后再新添加一条，确保数据的实效性，如果用insert ignore这样数据可能一直都是旧数据
                        # 不存在唯一索引相同的情况的话，replace into 和 insert into等价
                        self.cursor.execute(
                            r'replace into chinatelecom_mail_files values(%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                            [0, mail_item['b_mail_subject'],
                             mail_item['c_mail_from_addr'], mail_item['d_mail_to_addr'],
                             mail_item['e_mail_date_format'], mail_item['f_mail_file_path'],
                             mail_item['g_tax_deduction_price'], mail_item['h_value_added_tax'],
                             mail_item['i_tax_included_price']])
                        self.conn.commit()

                        # 保存相关的数据到data_to_excel这个列表中
                        self.data_to_excel.append(mail_item)
                        flag += 1

                        # 每8个数据保存数据到excel文件
                        if flag % 8 == 0 and self.data_to_excel:
                            save_to_excel(self.data_to_excel, self.data_for_xls)
                            print('%s至%s，共%s条符合条件的数据已经保存到了excel中！' % (len(mails), i, flag))
                            # 清空已经保存的数据
                            self.data_to_excel.clear()
                    except Exception as e:
                        print('因为 %s，保存邮件数据到数据库出错！' % e)
                    time.sleep(random.randint(1, 5) + random.randint(4, 8) / 10)

            # 记录已经爬取的邮件的序号，后面就不会重复爬取了
            # if i == self.end + 1:
            #     self.save_end_to_settings_file(self.settings_file_path, len(mails))
            print('邮件附件下载完毕！')

        except Exception as e:
            print('错误：%s' %e)
        finally:
            # 保存数据到excel文件
            if self.data_to_excel:
                save_to_excel(self.data_to_excel, self.data_for_xls)
            server.quit()
            self.conn.close()


if __name__ == '__main__':
    sys.stdout = Logger('all.log', sys.stdout)
    GetMailFiles = GetMailFiles()
    # GetMailFiles.test()
    # GetMailFiles.move_database_data()
    GetMailFiles.mail_main()
