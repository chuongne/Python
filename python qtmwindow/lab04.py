from curses import use_default_colors
from doctest import DocFileCase
from msilib.schema import InstallExecuteSequence
import os
from pprint import isreadable

from hehe import taoProfile, xoauser

Username =""
password =""
domain =""


def taoUser(username, password, ou, domain):
    command ="dsadd user "+ chr(34)+"cn="+username+",ou="+ou+","+domain+chr(34)+" -pwd "+ password
    print(command)
    os.system(command)

def doiPass(username, password, ou, domain):
    command="dsmod user"+ chr(34)+"cn="+username+",ou="+ou+","+domain+chr(34)+" -pwd "+ password
    print(command)
    os.system(command)



def menu():
    while True:
        print(" ")
        print("Nhan 1 de tao user")
        print("Nhan 2 de doi mat khau user")
        print("Nhan 3 de tao profile cho user")
        print("Nhan 4 cai service")
        print("Nhan 5 de xoa user")
        print("Nhan 6 de tao user tu csv file")
        print("Nhan 7 remove Service")
        print("Nhan 8 de thiet lap user co quyen remote desktop")
        chon = int(input("Nhap su lua chon cua ban: "))
        if chon == 1:
            username = input("Nhap ten user: ")
            password = input("nhap mat khau: ")
            taoUser(username, password, "demo", "dc=demo", "dc=com")
        elif chon == 2:
            username + input("Nhap username: ")
            newpassword = input("Nhap mat khau moi: ")
            ou = input("Nhap OU: ")
            domain = "dc=demo,dc=com"
            doiPass(username, ou, domain)
        elif chon == 3:
            username = input("Nhap username: ")
            password = input("Nhap mat khau: ")
            ou = input("Nhap OU: ")
            domain = "dc=demo,dc=com"
            taoProfile(username, ou, domain)
        elif chon == 4:
            username = input("Nhap ten user")
            xoauser(username, "data", "dc=demo,dc=com")

        elif chon == 5:
            username=input("Nhap ten user: ")
            xoauser(username,"data","dc=hehe,dc=com")

        elif chon == 6:
            DocFileCase("DATASERVER.csv")
        elif chon == 7:
            removeService()
        elif chon == 8:
            username = input("Nhap ten username : ")
            addUsertoRemoteDesktop(username)

menu()
            
            
