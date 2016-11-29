#pragma once
#include "BaseRWExcel.h"
class CServerLanguage :
	public BaseRWExcel
{
public:
	CServerLanguage();
	~CServerLanguage();
	//得到更新文件
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void setFileHead(Excel::RangePtr range);
};

