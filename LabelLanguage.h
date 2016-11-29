#pragma once
#include "BaseRWExcel.h"
class CLabelLanguage :
	public BaseRWExcel
{
public:
	CLabelLanguage();
	~CLabelLanguage();
	void StringChangeExcel(Excel::_ApplicationPtr excelApp);
	//得到更新文件
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	//合并文件
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void CLabelLanguage::setFileHead(Excel::RangePtr range);
};

