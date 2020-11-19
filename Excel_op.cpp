#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <conio.h>
#include <iomanip>
#include <windows.h>
#include <cstring>
#include "Excel_op.h"
#include "Linkman.h"
#include "libxl.h"
#include "ALMS.h"
#pragma comment(lib,"libxl.lib")
using namespace libxl;

bool Excel_op::init()//表格的初始化
{

	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		Sheet* sheet = book->addSheet("Mysheet1");
		if (!sheet)
		{
			return false;
		}

		//设置标题字体
		Font* titleFont = book->addFont();
		titleFont->setName("Arial Black");
		titleFont->setColor(COLOR_RED);//设置颜色为红色
		titleFont->setSize(20);//设置字体大小

		Format* titleFormat = book->addFormat();
		titleFormat->setAlignH(ALIGNH_CENTER);//设置中心对齐
		titleFormat->setFont(titleFont);

		//设置属性字体和格式
		Font* attributeFont = book->addFont();
		attributeFont->setName("宋体");
		attributeFont->setBold(TRUE);

		Format* attributeFormat = book->addFormat();
		attributeFormat->setAlignH(ALIGNH_CENTER);
		attributeFormat->setFont(attributeFont);

		//设置行的宽度
		sheet->setCol(1, 1, 25);

		//设置列高
		sheet->setRow(1, 15);

		//写入字符串数据

		sheet->writeStr(1, 0, "编号", attributeFormat);
		sheet->writeStr(1, 1, "姓名", attributeFormat);
		sheet->writeStr(1, 2, "手机号码", attributeFormat);
		sheet->writeStr(1, 3, "家庭电话", attributeFormat);
		sheet->writeStr(1, 4, "工作电话 ", attributeFormat);
		sheet->writeStr(1, 5, "邮件地址 ", attributeFormat);
		sheet->writeStr(1, 6, "家庭地址 ", attributeFormat);


	}
	else {
		Sheet* sheet = book->getSheet(0);
		if (!sheet)
		{
			return false;
		}


	}
	book->save("ALMS.xls");
	book->release();
	return true;
}

bool Excel_op::show()//显示通讯录信息
{

	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();
	int colfirst = sheet->firstCol();
	int collast = sheet->lastCol();
	int linkman_numbers = 0;
	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		const char* name0 = sheet->readStr(i, 1) ? sheet->readStr(i, 1) : "null";
		if ((strcmp(name0, "") != 0) && (strcmp(name0, "null") != 0))
		{
			linkman_numbers++;
		}
	}
	//cout << rowfirst << rowlast<< endl;

	cout << "当前通讯录中有 " << linkman_numbers << "个联系人" << endl;

	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出
	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		cout << setw(4) << "编号：" << setw(4) << sheet->readNum(i, 0) << endl;
		cout << setw(4) << "姓名：" << setw(10) << (sheet->readStr(i, 1) ? sheet->readStr(i, 1) : "null") << endl;
		cout << setw(6) << "手机号码：" << setw(20) << sheet->readNum(i, 2) << endl;
		cout << setw(6) << "家庭号码：" << setw(20) << sheet->readNum(i, 3) << endl;
		cout << setw(6) << "家庭号码：" << setw(20) << sheet->readNum(i, 4) << endl;
		cout << setw(6) << "邮件地址：" << setw(20) << (sheet->readStr(i, 5) ? sheet->readStr(i, 5) : "null") << endl;
		cout << setw(6) << "地址：" << setw(20) << (sheet->readStr(i, 6) ? sheet->readStr(i, 6) : "null") << endl;
		cout << endl;
	}

	if (book->save("ALMS.xls"))
	{
		::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
		return true;
	}
}

bool Excel_op::add(Linkman linkman)//增加用户信息
{


	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}

	int id = 0; //联系人编号
	double cellphone = 0;//手机号码 
	double phone_home = 0;//家庭电话
	double phone_work = 0;//工作电话
	const CHAR* name = "null";//姓名
	const CHAR* email_address = "null";//邮件地址 
	const CHAR* address = "null";//家庭地址 

	cellphone = linkman.cellphone;
	phone_home = linkman.phone_home;
	phone_work = linkman.phone_work;
	name = linkman.name;
	email_address = linkman.email_address;
	address = linkman.address;

	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();
	int i = rowfirst;




	while (i++ < rowlast - 1)
	{
		const char* name0 = sheet->readStr(i, 1) ? sheet->readStr(i, 1) : "null";
		if (strcmp(name0, "") == 0 || (strcmp(name0, "null") == 0))
		{
			id = i - 1;
			sheet->writeNum(i, 0, id);
			sheet->writeStr(i, 1, name);
			sheet->writeNum(i, 2, cellphone);
			sheet->writeNum(i, 3, phone_home);
			sheet->writeNum(i, 4, phone_work);
			sheet->writeStr(i, 5, email_address);
			sheet->writeStr(i, 6, address);



			if (book->save("ALMS.xls"))
			{
				::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
				show();
				return true;
			}
		}
	}
	if (i == rowlast)
	{
		sheet->insertRow(rowlast, rowlast);
		sheet->writeNum(rowlast, 0, rowlast - 1);
		sheet->writeStr(rowlast, 1, name);
		sheet->writeNum(rowlast, 2, cellphone);
		sheet->writeNum(rowlast, 3, phone_home);
		sheet->writeNum(rowlast, 4, phone_work);
		sheet->writeStr(rowlast, 5, email_address);
		sheet->writeStr(rowlast, 6, address);


		if (book->save("ALMS.xls"))
		{
			::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
			show();
			return true;
		}
	}

}

bool Excel_op::query(const TCHAR* Name)//根据姓名查找通讯录信息
{
	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();

	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出

	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		const CHAR* name = sheet->readStr(i, 1);

		// int id = sheet->readNum(i, 0);
		if (strcmp(name, Name) == 0)
		{
			cout << setw(4) << "编号：" << setw(4) << sheet->readNum(i, 0) << endl;
			cout << setw(4) << "姓名：" << setw(10) << (sheet->readStr(i, 1) ? sheet->readStr(i, 1) : "null") << endl;
			cout << setw(6) << "手机号码：" << setw(20) << sheet->readNum(i, 2) << endl;
			cout << setw(6) << "家庭号码：" << setw(20) << sheet->readNum(i, 3) << endl;
			cout << setw(6) << "家庭号码：" << setw(20) << sheet->readNum(i, 4) << endl;
			cout << setw(6) << "邮件地址：" << setw(20) << (sheet->readStr(i, 5) ? sheet->readStr(i, 5) : "null") << endl;
			cout << setw(6) << "地址：" << setw(20) << (sheet->readStr(i, 6) ? sheet->readStr(i, 6) : "null") << endl;
			cout << endl;
			return true;
		}
		if (i == rowlast - 1) {
			cout << "查无此人" << endl;
		}
	}
}

bool Excel_op::Delete(const CHAR* Name)//根据姓名删除通讯录信息
{

	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();


	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出
	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		const CHAR* name = sheet->readStr(i, 1);
		if (strcmp(name, Name) == 0)
		{
			sheet->writeNum(i, 0, 0);
			sheet->writeStr(i, 1, "null");
			sheet->writeNum(i, 2, 0);
			sheet->writeNum(i, 3, 0);
			sheet->writeNum(i, 4, 0);
			sheet->writeStr(i, 5, "null");
			sheet->writeStr(i, 6, "null");

			break;
		}
		if (i == rowlast - 1) {
			cout << "查无此人" << endl;
		}
	}
	if (book->save("ALMS.xls"))
	{
		::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
		return true;
	}
}

bool Excel_op::Alldelete()//清空所有通讯录
{
	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();

	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出
	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		sheet->writeNum(i, 0, 0);
		sheet->writeStr(i, 1, "null");
		sheet->writeNum(i, 2, 0);
		sheet->writeNum(i, 3, 0);
		sheet->writeNum(i, 4, 0);
		sheet->writeStr(i, 5, "null");
		sheet->writeStr(i, 6, "null");
	}
	if (book->save("ALMS.xls"))
	{
		::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
		show();


		return true;
	}
}

bool Excel_op::modify(const TCHAR* Name1, const TCHAR* Name2, double cellphone, double phone_home, double phone_work, const TCHAR* email_address, const TCHAR* address)//根据姓名修改通讯录信息
{
	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();
	int colfirst = sheet->firstCol();
	int collast = sheet->lastCol();

	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出

	for (int i = rowfirst + 1; i < rowlast; i++)
	{
		const TCHAR* name = sheet->readStr(i, 1);

		// int id = sheet->readNum(i, 0);
		if (strcmp(name, Name1) == 0)
		{

			sheet->writeStr(i, 1, Name2);
			sheet->writeNum(i, 2, cellphone);
			sheet->writeNum(i, 3, phone_home);
			sheet->writeNum(i, 4, phone_work);
			sheet->writeStr(i, 5, email_address);
			sheet->writeStr(i, 6, address);
			break;
		}
		if (i == rowlast - 1) {
			cout << "查无此人,修改无效" << endl;
		}
	}
	if (book->save("ALMS.xls"))
	{
		::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
		return true;
	}
	show();
}

bool Excel_op::sort_name()
{
	Book* book = xlCreateBook();
	if (!book)
	{
		return false;
	}

	if (!book->load("ALMS.xls"))
	{
		return false;
	}

	Sheet* sheet = book->getSheet(0);
	if (!sheet)
	{
		return false;
	}
	int rowfirst = sheet->firstRow();
	int rowlast = sheet->lastRow();
	int colfirst = sheet->firstCol();
	int collast = sheet->lastCol();

	wcout.imbue(locale(locale(), "", LC_CTYPE));//unicode 宽字节中文输出

	for (int i = rowfirst + 1; i < rowlast-1; i++)
	{

		Linkman linkman1;
		const CHAR* name2 = sheet->readStr(i, 1);
		name2 = name2 ? name2 : "null";
		std::strcpy(linkman1.name, name2);
		Linkman linkman2;
		const CHAR* name4 = sheet->readStr(i + 1, 1);
		name4 = name4 ? name4 : "null";
		std::strcpy(linkman2.name, name4);
		if (strcmp(linkman1.name, linkman2.name) >= 0)
		{

			linkman1.cellphone = sheet->readNum(i, 2);
			linkman1.phone_home = sheet->readNum(i, 3);
			linkman1.phone_work = sheet->readNum(i, 4);
			const CHAR* email_address2 = sheet->readStr(i, 5);
			email_address2 = email_address2 ? email_address2 : "null";
			std::strcpy(linkman1.email_address, email_address2);
			const CHAR* address2 = sheet->readStr(i, 6);
			address2 = address2 ? address2 : "null";
			std::strcpy(linkman1.address, address2);

			Linkman linkman3;
			const CHAR* name3 = sheet->readStr(i+1, 1);
			name3= name3 ? name3 : "null";
			std::strcpy(linkman3.name, name3);

			linkman3.cellphone = sheet->readNum(i + 1, 2);
			linkman3.phone_home = sheet->readNum(i + 1, 3);
			linkman3.phone_work = sheet->readNum(i + 1, 4);
			const CHAR* email_address3 = sheet->readStr(i + 1, 5);
			email_address3 = email_address3 ? email_address3 : "null";
			std::strcpy(linkman3.email_address, email_address3);
			const CHAR* address3 = sheet->readStr(i + 1, 6);
			address3 = address3 ? address3 : "null";
			std::strcpy(linkman3.address, address3);

			Delete(linkman2.name);

			sheet->writeStr(i+1, 1, linkman1.name);
			sheet->writeNum(i+1, 2, linkman1.cellphone);
			sheet->writeNum(i+1, 3, linkman1.phone_home);
			sheet->writeNum(i+1, 4, linkman1.phone_work);
			sheet->writeStr(i+1, 5, linkman1.email_address);
			sheet->writeStr(i+1, 6, linkman1.address);

			Delete(linkman1.name);

			sheet->writeStr(i, 1, linkman3.name);
			sheet->writeNum(i, 2, linkman3.cellphone);
			sheet->writeNum(i, 3, linkman3.phone_home);
			sheet->writeNum(i, 4, linkman3.phone_work);
			sheet->writeStr(i, 5, linkman3.email_address);
			sheet->writeStr(i, 6, linkman3.address);
			break;

		}
	}
	if (book->save("ALMS.xls"))
	{
		::ShellExecute(NULL, "open", "ALMS.xls", NULL, NULL, SW_SHOW);
		show();
		return true;
	}
}












	//sheet->insertRow(2, 2);
	//for (int i = rowfirst + 1; i < rowlast; i++)
	//{
	//	Linkman linkman1;
	//	Linkman linkman2;

	//	char* name1 = new char[100];
	//	const CHAR* name2 = sheet->readStr(i, 1);
	//	name2 = name2 ? name2 : "null";
	//	std::strcpy(name1, name2);
	//	linkman1.name = name1;

	//	char* name3 = new char[100];
	//	const CHAR* name4 = sheet->readStr(i + 1, 1);
	//	name4 = name4 ? name4 : "null";
	//	std::strcpy(name3, name4);
	//	linkman2.name = name3;

	//	cout << linkman1.name << linkman2.name << endl;
	//	

	//	if (strcmp(linkman1.name, linkman2.name) > 0) {//if 1>3 交换  -> name1:小的 i，name3:大的 i+1
	//		char* name5 = new char[100];
	//		std::strcpy(name5, linkman1.name);
	//		std::strcpy(linkman1.name, linkman2.name);
	//		std::strcpy(linkman2.name, name5);
	//		cout << linkman1.name << linkman2.name << endl;


	//		double cellphone1 = sheet->readNum(i + 1, 2);
	//		double cellphone3 = sheet->readNum(i, 2);
	//		double phone_home1 = sheet->readNum(i + 1, 3);
	//		double phone_home3 = sheet->readNum(i, 3);
	//		double phone_work1 = sheet->readNum(i + 1, 4);
	//		double phone_work3 = sheet->readNum(i, 4);


	//		char* email_address1 = new char[100];
	//		const CHAR* email_address2 = sheet->readStr(i, 5);
	//		email_address2 = email_address2 ? email_address2 : "null";
	//		std::strcpy(email_address1, email_address2);
	//		char* email_address3 = new char[100];
	//		const CHAR* email_address4 = sheet->readStr(i + 1, 5);
	//		email_address4 = email_address4 ? email_address4 : "null";
	//		std::strcpy(email_address3, email_address4);

	//		char* email_address5 = new char[100];
	//		std::strcpy(email_address5, email_address1);
	//		std::strcpy(email_address1, email_address3);
	//		std::strcpy(email_address3, email_address5);
	//		//cout << email_address1 << email_address3 << endl;

	//		char* address1 = new char[100];
	//		const CHAR* address2 = sheet->readStr(i, 6);
	//		address2 = address2 ? address2 : "null";
	//		std::strcpy(address1, address2);
	//		char* address3 = new char[100];
	//		const CHAR* address4 = sheet->readStr(i + 1, 6);
	//		address4 = address4 ? address4 : "null";
	//		std::strcpy(address3, address4);

	//		char* address5 = new char[100];
	//		std::strcpy(address5, address1);
	//		std::strcpy(address1, address3);
	//		std::strcpy(address3, address5);

	//		/*cout << name1 << name3
	//			<< endl
	//			<< cellphone1 << cellphone3
	//			<< endl
	//			<< phone_home1 << phone_home3
	//			<<endl
	//			<< phone_work1 << phone_work3
	//			<<endl
	//			<< email_address1 << email_address3
	//			<<endl
	//			<< address1 << address3
	//			<< endl;*/


	//		sheet->writeStr(i, 1, name1);
	//		sheet->writeStr(i + 1, 1, name3);
	//		sheet->writeNum(i, 2, cellphone1);
	//		sheet->writeNum(i + 1, 2, cellphone3);
	//		sheet->writeNum(i, 3, phone_home1);
	//		sheet->writeNum(i + 1, 3, phone_home3);
	//		sheet->writeNum(i, 4, phone_work1);
	//		sheet->writeNum(i + 1, 4, phone_work3);
	//		sheet->writeStr(i, 1, email_address3);
	//		sheet->writeStr(i + 1, 1, email_address1);
	//		sheet->writeStr(i, 1, address3);
	//		sheet->writeStr(i + 1, 1, address1);


	//		delete address1;

	//		delete address3;

	//		delete email_address1;

	//		delete email_address3;

	//	}

	//	delete name1;

	//	delete name3;}


