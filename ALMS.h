#pragma once
#include"Linkman.h"
#include<iostream>
using namespace std;

class ALMS
{
public:
	ALMS();  //构造函数判断文本状态并初始化
	void showMenu();//显示菜单
	void exitProgram();//退出管理程序
	void Add(Linkman linkman);//增加通讯录
	void Show();//显示通讯录
	void Delete();//删除通讯录记录
	void Modify();//修改通讯录
	void Query();//查找通讯录信息 
	void Delete_all();//清空通讯录
	void Sort_name();//通讯录排序 
	

};