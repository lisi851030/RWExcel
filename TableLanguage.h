#pragma once
#include "BaseRWExcel.h"
#include <map>
#include <list>
class CTableLanguage :
	public BaseRWExcel
{
public:
	CTableLanguage();
	~CTableLanguage();
	void GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType);
	void MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType);
	//∑÷¿Î”Ô—‘
	void LeaveLanguage(Excel::_ApplicationPtr excelApp, bool create_all, string vec_file);
	void BuildLanguageFile(Excel::_ApplicationPtr excelApp, string strPlatform);
protected:
	_bstr_t connectKey(_bstr_t value1, _bstr_t value2, _bstr_t connect);
	void setFileHead(Excel::RangePtr range);
	bool ReadBaseLanguageFile(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, string strPath, string strFileName, vector<string> &vecFiles);
	void saveOneLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, _bstr_t tar, string strFileName);
	void saveBaseLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang);
	void readOneExcel(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, StructChange &stucChange);
	string strClientName = "tb_table_away_client";
	string strServerName = "tb_table_away_server";
	string getLanguagePath();
	map<_bstr_t, _bstr_t> GetBaseTargetMap(Excel::_ApplicationPtr excelApp);
private:
	void readChange(Excel::_WorksheetPtr &sheet, list<StructChange> &lstChange, vector<string> &vecFiles);
};

