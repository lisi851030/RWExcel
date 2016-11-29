#pragma once
#include "BaseRWExcel.h"
class CSkillLanguage :
	public BaseRWExcel
{
public:
	CSkillLanguage();
	~CSkillLanguage();
	void CSkillLanguage::StringChangeExcel(Excel::_ApplicationPtr excelApp);
};

