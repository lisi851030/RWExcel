#pragma once
#import "C:\\Program Files (x86)\\Common Files\\microsoft shared\\OFFICE14\\MSO.DLL" \
	rename("RGB", "MsoRGB") \
	rename("SearchPath", "MsoSearchPath")
using namespace Office;
#import "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA\\VBA6\\VBE6EXT.OLB"
using namespace VBIDE;

#import "D:\\Program Files (x86)\\Microsoft Office\\Office14\\EXCEL.EXE" \
	rename("DialogBox", "ExcelDialogBox") \
	rename("RGB", "ExcelRGB") \
	rename("CopyFile", "ExcelCopyFile") \
	rename("ReplaceText", "ExcelReplaceText") \
	exclude("IFont", "IPicture") \
	no_auto_exclude
using namespace Excel;

#include <iostream>
#include <string>
#include <vector>
using namespace std;


struct StructChange
{
	_bstr_t strKey1;
	_bstr_t strKey2;
	_bstr_t strKey3;
	_bstr_t strTableName;
	int startID;
};
struct StructID
{
	_bstr_t strKey;
	int oldID;
	int newID;
};
struct StructLanguageData
{
	_bstr_t strData;
	_bstr_t strTableName;
	_bstr_t strTarget;
};

struct StructMapData
{
	int seted;
	_bstr_t data;
};

struct StructIDToLanguage
{
	int id;
	_bstr_t desc;
	_bstr_t language;
};
class BaseRWExcel
{
private:
	string strSourcePath;
	string strTransPath;
	string strBaseFileName;
	string strTransFileName;
public:
	BaseRWExcel();
	~BaseRWExcel();
	void setStrSourcePath(string value){
		this->strSourcePath = value;
	}
	string getStrSourcePath()
	{
		return strSourcePath;
	}
	void setStrTransPath(string value){
		this->strTransPath = value;
	}
	string getTransPath()
	{
		return strTransPath;
	}
	void setStrBaseFileName(string value)
	{
		this->strBaseFileName = value;
	}
	string getStrBaseFileName()
	{
		return this->strBaseFileName;
	}
	void setStrTransFileName(string value)
	{
		this->strTransFileName = value;
	}
	string getStrTransFileName()
	{
		return this->strTransFileName;
	}
	//得到更新文件
	void virtual GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType){};
	//合并文件
	void virtual MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType){};
	//分离语言，抽离用
	//virtual void LeaveLanguage();
	//字符串转换成excel，开发配置和静态用
	//virtual void StringChangeExcel() = 0;
protected:
	bool GetSheet(const string strFileName, const string strPath, const Excel::_ApplicationPtr excelApp, Excel::_WorksheetPtr &sheet);
	bool FindValue(vector<string> &vecFiles, string &strFile);
	string GBKToUTF8(const std::string& strGBK);
	std::vector<std::string> Split(const  std::string& s, const std::string& delim);
	_bstr_t strConnect = "_________________";
	_bstr_t strHaveTrans = _bstr_t("1");
	virtual void setFileHead(Excel::RangePtr range) {};
	int GetColumnID(Excel::_WorksheetPtr &sheet, _bstr_t name);
	_bstr_t getTranslateData(_variant_t oldData);
};

