#pragma once
#include "ClientLanguage.h"
class CDevelopLanguage :
	public CClientLanguage
{
public:
	CDevelopLanguage();
	~CDevelopLanguage();
	void StringChangeExcel(Excel::_ApplicationPtr excelApp);
	//�ϲ��ļ�
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void CDevelopLanguage::setFileHead(Excel::RangePtr range);
};

