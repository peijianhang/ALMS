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

void ALMS::showMenu()//����
{
	cout << "*********************************" << endl;
	cout << "***   ��ӭʹ��ͨѶ¼����ϵͳ  *****" << endl;
	cout << "*****   1������ͨѶ¼��Ϣ   *******" << endl;
	cout << "*****   2����ʾͨѶ¼��Ϣ   *******" << endl;
	cout << "*****   3��ɾ��ͨѶ¼��Ϣ   *******" << endl;
	cout << "*****   4���޸�ͨѶ¼��Ϣ   *******" << endl;
	cout << "*****   5������ͨѶ¼��Ϣ   *******" << endl;
	cout << "*****   6����������ͨѶ¼����  *******" << endl;
	cout << "*****   7���������ͨѶ¼   *******" << endl;
	cout << "*****   8���˳��������   *******" << endl;
	cout << "*********************************" << endl;
}

//0
void ALMS::exitProgram()
{
	int opt;
	cout << "ȷ���˳���" << endl;
	cout << "1��ȷ��" << endl;
	cout << "2������" << endl;
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
	cout << "��������������" << endl;
	cin >> linkman.name;
	cout << "�����������ֻ�����" << endl;
	cin >> linkman.cellphone;
	cout << "���������ļ�ͥ�绰" << endl;
	cin >> linkman.phone_home;
	cout << "���������Ĺ����绰" << endl;
	cin >> linkman.phone_work;
	cout << "�����������ʼ���ַ" << endl;
	cin >> linkman.email_address;
	cout << "���������ļ�ͥ��ַ " << endl;
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
	cout << "��������������" << endl;
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
	cout << "��������������" << endl;
	cin >> linkman1.name;
	cout << "�����������ڵ�����" << endl;
	cin >> linkman2.name;
	cout << "�����������ڵ��ֻ�����" << endl;
	cin >> linkman2.cellphone;
	cout << "�����������ڵļ�ͥ�绰" << endl;
	cin >> linkman2.phone_home;
	cout << "�����������ڵĹ����绰" << endl;
	cin >> linkman2.phone_work;
	cout << "�����������ڵ��ʼ���ַ" << endl;
	cin >> linkman2.email_address;
	cout << "�����������ڵļ�ͥ��ַ " << endl;
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
	cout << "��������������" << endl;
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
	cout << "ȷ����գ�" << endl;
	cout << "1��ȷ��" << endl;
	cout << "2������" << endl;
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


