#pragma once
#include "BaseRWExcel.h"
class CLabelLanguage :
	public BaseRWExcel
{
public:
	CLabelLanguage();
	~CLabelLanguage();
	void StringChangeExcel(Excel::_ApplicationPtr excelApp);
	//�õ������ļ�
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	//�ϲ��ļ�
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void CLabelLanguage::setFileHead(Excel::RangePtr range);
};

