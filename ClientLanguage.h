#pragma once
#include "BaseRWExcel.h"
class CClientLanguage :
	public BaseRWExcel
{
public:
	CClientLanguage();
	~CClientLanguage();
	//�õ������ļ�
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	//�ϲ��ļ�
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
protected:
	void CClientLanguage::setFileHead(Excel::RangePtr range);
};

