# -*- coding: utf-8 -*-
import pickle  # Python的标准模块，可以将纯Python对象存储在文件中

class Contact:
    """ 联系人类，保存联系人的信息:包括了姓名、电话、邮件、地址和生日"""
    total_amount = 0  # 类变量，用于记录当前通讯录的总人数
    contacts_dict = {}  # 创建空的联系人字典,字典以键值对的形式存储信息

    @classmethod
    def add_contact(cls,name,email):
        """添加联系人"""
        # name = input("请输入添加的联系人姓名：")  # input函数可以接受一个字符串作为参数，等待用户输入后将返回用户输入的文本
        if name in Contact.contacts_dict:  # 运算符in来检查某一键值对是否存在于字典中
            print("该联系人已经存在")
            return True
        else:
            Contact.contacts_dict[name] = email
            Contact.total_amount += 1
            print("添加成功，当前已有联系人{}人".format(Contact.total_amount))
            return False

    @classmethod
    def delete_contact(cls,name):
        """根据姓名删除联系人"""
        # name = input("请输入要删除的联系人姓名：")
        if name in Contact.contacts_dict:
            del Contact.contacts_dict[name]  # del语句用来删除某一键值对
            print(Contact.contacts_dict)
            Contact.total_amount -= 1
            print("删除成功，当前已有联系人{}人".format(Contact.total_amount))
            return True
        else:
            print("{}人不在通讯录中".format(name))
            return False

    @classmethod
    def search_contact(cls,name):
        """根据姓名搜索联系人"""
        # name = input("请输入要搜索的联系人姓名：")
        if name in Contact.contacts_dict:
            print(Contact.contacts_dict[name])
            # TODO：查到则显示在table中
            return Contact.contacts_dict[name]
        else:
            print("{}人不在通讯录中".format(name))
            return ''

    @classmethod
    def modify_contact(cls,name,modify_email):
        """修改联系人信息，注意，由于name是字典的键，所以name的值不可以修改"""
        # name = input("请输入要修改的联系人姓名：")
        if name in Contact.contacts_dict:
            print("修改前：")
            print(Contact.contacts_dict[name])
            Contact.contacts_dict[name] = modify_email
            print("修改后：")
            print(Contact.contacts_dict[name])
        else:
            print("{}人不在通讯录中".format(name))

    @classmethod
    def write(cls,path):
        """将通讯录写入文件"""
        # 写入的方式打开文件
        if path == '':
            f = open('contact.txt', 'wb')
        else:
            # ------------
            # text = "------Address Book------\n"
            # for name, email in Contact.contacts_dict:
            #     text = text + (str(name) + ':' + str(email) + '\n')
            # with open(path + '/' + 'contact.txt', 'wb') as f:
            #     # 保存附件
            #     f.write(text)
            # f.close()
            #-----------
            f = open(path + '/' + 'contact.txt', 'wb')
        pickle.dump(Contact.contacts_dict, f)
        # 关闭文件
        f.close()

    @classmethod
    def read(cls,path):
        file = 'contact.txt'
        try:
            # 读的方式打开文件
            if path == '':
                f = open('contact.txt', 'rb')
            else:
                f = open(path, 'rb')
            # 拆封
            Contact.contacts_dict = pickle.load(f)
            print("已经保存联系人{}位".format(len(Contact.contacts_dict)))
            Contact.total_amount += len(Contact.contacts_dict)
            print(Contact.contacts_dict)
            # 关闭文件
            f.close()
        except EOFError:
            print("尚未保存联系人")
        except FileNotFoundError:
            f.close()


# def contact_menu():
#     print("欢迎来到Leaf通讯录，系统提供以下功能："
#           "1：添加"
#           "2：删除"
#           "3：修改"
#           "4：搜索"
#           "5：退出")
#
#
# contact_person = Contact()
# contact_person.read()
# while True:
#     try:
#         contact_menu()
#         choice = int(input("请选择功能：输入对应的数字"))
#         if choice == 1:  # 添加
#             contact_person.add_contact()
#         elif choice == 2:  # 删除
#             contact_person.delete_contact()
#         elif choice == 3:  # 修改
#             contact_person.modify_contact()
#         elif choice == 4:  # 搜索
#             contact_person.search_contact()
#         elif choice == 5:  # 退出
#             contact_person.write()  # 字典中数据重写写到文件中
#             break
#         else:
#             print("输入不合法，请重新输入")
#     except ValueError:
#         print("请输入相应的数字")
