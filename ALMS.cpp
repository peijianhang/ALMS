#include"ALMS.h"
#include "libxl.h"
#include"Linkman.h"
#include "Excel_op.h"
#include <iostream>
#include<string.h>
#pragma comment(lib,"libxl.lib")
using namespace libxl;

ALMS::ALMS()
{
	Linkman linkman[100];
	Excel_op excel_op;
	excel_op.init();
}

void ALMS::showMenu()//界面
{
	cout << "*********************************" << endl;
	cout << "***   欢迎使用通讯录管理系统  *****" << endl;
	cout << "*****   1、增加通讯录信息   *******" << endl;
	cout << "*****   2、显示通讯录信息   *******" << endl;
	cout << "*****   3、删除通讯录信息   *******" << endl;
	cout << "*****   4、修改通讯录信息   *******" << endl;
	cout << "*****   5、查找通讯录信息   *******" << endl;
	cout << "*****   6、按照姓名通讯录排序  *******" << endl;
	cout << "*****   7、清空所有通讯录   *******" << endl;
	cout << "*****   8、退出管理程序   *******" << endl;
	cout << "*********************************" << endl;
}

//0
void ALMS::exitProgram()
{
	int opt;
	cout << "确认退出？" << endl;
	cout << "1、确认" << endl;
	cout << "2、返回" << endl;
	cin >> opt;
	if (opt == 1)
	{
		exit(1);
	}
	else {
		system("pause");
		system("cls");
		showMenu();
	}
}
//1
void ALMS::Add(Linkman linkman)
{
	Excel_op excel_op;
	cout << "请输入她的姓名" << endl;
	cin >> linkman.name;
	cout << "请输入她的手机号码" << endl;
	cin >> linkman.cellphone;
	cout << "请输入她的家庭电话" << endl;
	cin >> linkman.phone_home;
	cout << "请输入她的工作电话" << endl;
	cin >> linkman.phone_work;
	cout << "请输入她的邮件地址" << endl;
	cin >> linkman.email_address;
	cout << "请输入她的家庭地址 " << endl;
	cin >> linkman.address;
	excel_op.add(linkman);
	//Sleep(3 * 1000);
	//system("cls");
}
//2
void ALMS::Show()
{
	Excel_op excel_op;
	excel_op.show();
	//Sleep(5 * 1000);
	//system("cls");
}
//3
void ALMS::Delete()
{
	Linkman linkman;
	Excel_op excel_op;
	cout << "请输入她的姓名" << endl;
	cin >> linkman.name;
	excel_op.Delete(linkman.name);
	excel_op.show();
	//Sleep(4 * 1000);
	//system("cls");
}
//4
void ALMS::Modify()
{
	Linkman linkman1;
	Linkman linkman2;
	Excel_op excel_op;
	cout << "请输入她的姓名" << endl;
	cin >> linkman1.name;
	cout << "请输入她现在的姓名" << endl;
	cin >> linkman2.name;
	cout << "请输入她现在的手机号码" << endl;
	cin >> linkman2.cellphone;
	cout << "请输入她现在的家庭电话" << endl;
	cin >> linkman2.phone_home;
	cout << "请输入她现在的工作电话" << endl;
	cin >> linkman2.phone_work;
	cout << "请输入她现在的邮件地址" << endl;
	cin >> linkman2.email_address;
	cout << "请输入她现在的家庭地址 " << endl;
	cin >> linkman2.address;
	excel_op.modify(linkman1.name, linkman2.name, linkman2.cellphone, linkman2.phone_home, linkman2.phone_work, linkman2.email_address, linkman2.address);
	//Sleep(4 * 1000);
	//system("cls");
}
//5
void ALMS::Query()
{
	Linkman linkman;
	Excel_op excel_op;
	cout << "请输入她的姓名" << endl;
	cin >> linkman.name;
	excel_op.query(linkman.name);
	//Sleep(4 * 1000);
	//system("cls");
}
//6
void ALMS::Sort_name()
{

	Excel_op excel_op;
	excel_op.sort_name();
	//Sleep(4 * 1000);
	//system("cls");
}
//7
void ALMS::Delete_all()
{
	int opt;
	cout << "确认清空？" << endl;
	cout << "1、确认" << endl;
	cout << "2、返回" << endl;
	cin >> opt;
	if (opt == 1)
	{
		Excel_op excel_op;
		excel_op.Alldelete();

	}
	else {
		system("cls");
		return;
	}

	//Sleep(3 * 1000);
	//system("cls");

}


