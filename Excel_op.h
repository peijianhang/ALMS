#pragma once
#include <windows.h>
#include"Linkman.h"
#include "ALMS.h"

class Excel_op
{

public:
	bool init();
	bool show();
	bool add(Linkman linkman);
	bool query(const CHAR* Name);
	bool Delete(const CHAR* Name);
	bool Alldelete();
	bool modify(const CHAR* Name1, const CHAR* Name2, double cellphone, double phone_home, double phone_work, const CHAR* email_address, const CHAR* address);//根据姓名修改通讯录信息
	bool sort_name();
	
};
