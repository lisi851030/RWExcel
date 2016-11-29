#pragma once
#include "BaseRWExcel.h"
class CClientLanguage :
	public BaseRWExcel
{
public:
	CClientLanguage();
	~CClientLanguage();
	//得到更新文件
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	//合并文件
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void CClientLanguage::setFileHead(Excel::RangePtr range);
};

