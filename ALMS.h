#pragma once
#include"Linkman.h"
#include<iostream>
using namespace std;

class ALMS
{
public:
	ALMS();  //���캯���ж��ı�״̬����ʼ��
	void showMenu();//��ʾ�˵�
	void exitProgram();//�˳��������
	void Add(Linkman linkman);//����ͨѶ¼
	void Show();//��ʾͨѶ¼
	void Delete();//ɾ��ͨѶ¼��¼
	void Modify();//�޸�ͨѶ¼
	void Query();//����ͨѶ¼��Ϣ 
	void Delete_all();//���ͨѶ¼
	void Sort_name();//ͨѶ¼���� 
	

};