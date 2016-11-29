#include "RWExcel.h"
#include <iostream>
#include <string>
#include "tchar.h"

#include  <direct.h>  
#include  <stdio.h>  
#include <string>
#include <list>
#include <map>
#include <iostream>
#include <string>
#include <cstring>
#include <vector>
#include <unordered_map>  
#include <fstream>
#include "ClientLanguage.h"
#include "ServerLanguage.h"
#include "TableLanguage.h"
#include "LabelLanguage.h"
#include "DevelopLanguage.h"
#include "SkillLanguage.h"
#include "shlwapi.h"
#include <Exception>
#pragma comment(lib,"shlwapi.lib")
using namespace std;
using namespace Excel;
CTableLanguage* getTableLanguage()
{
	CTableLanguage *tableLang = new CTableLanguage();
	tableLang->setStrBaseFileName("away_language");
	tableLang->setStrTransFileName("language_change");
	return tableLang;
}
CLabelLanguage* getLabelLanguage()
{
	CLabelLanguage *lang = new CLabelLanguage();
	lang->setStrBaseFileName("tb_table_labellang");
	lang->setStrTransFileName("labellang_change");
	return lang;
}
CDevelopLanguage* getDevelopLanguage()
{
	CDevelopLanguage *lang = new CDevelopLanguage();
	lang->setStrBaseFileName("tb_table_developlang");
	lang->setStrTransFileName("developlang_change");
	return lang;
}
CServerLanguage* getServerLanguage()
{
	CServerLanguage *lang = new CServerLanguage();
	lang->setStrBaseFileName("tb_table_serverlanguage");
	lang->setStrTransFileName("serverlang_change");
	return lang;
}
CClientLanguage* getClientLanguage()
{
	CClientLanguage *lang = new CClientLanguage();
	lang->setStrBaseFileName("tb_table_clientlanguage");
	lang->setStrTransFileName("clientlang_change");
	return lang;
}
bool createExcel(Excel::_ApplicationPtr &excel_app)
{
	try
	{
		//创建Excel
		HRESULT hr = excel_app.CreateInstance(L"Excel.Application");
		excel_app->PutVisible(0x22e, true);

		/*判断当前Excel的版本*/
		_bstr_t strExcelVersion = excel_app->Version;
		std::cout << "当前Excel版本 " << strExcelVersion << std::endl;
		return true;
	}
	catch (exception e)
	{
		std::cout << e.what() << endl;
		return false;
	}
}
void leaveLanguage(string strPath)
{
	int select;
	std::cout << "1抽离所有表 \n2抽离指定表 \n请输入：" << endl;
	std::cin >> select;
	string str_file;
	bool create_all = true;
	if (select == 1)
	{
		create_all = true;
	}else if (select == 2)
	{
		std::cout << "输入表明，以逗号隔开" << endl;
		std::cin >> str_file;
		create_all = false;
	}
	else
	{
		leaveLanguage(strPath);
		return;
	}
	CTableLanguage *tableLang;
	Excel::_ApplicationPtr excel_app;
	if (!createExcel(excel_app))
		return;
	tableLang = getTableLanguage();
	tableLang->setStrSourcePath(strPath + "data\\");
	tableLang->LeaveLanguage(excel_app, create_all, str_file);
	delete(tableLang);
	excel_app->Quit();
}
void getOneTransFile(BaseRWExcel *tar, Excel::_ApplicationPtr excelApp, string strPath, string lang_type)
{
	tar->setStrSourcePath(strPath + "data\\");
	tar->setStrTransPath(strPath + lang_type + "\\");
	tar->GetUpdateFile(excelApp, lang_type);
	delete(tar);
}
void getTransFiles(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	string lang_type;
	std::cout << "请输入语言版本的后缀如英文：en \n请输入：" << endl;
	std::cin >> lang_type;
	if (!PathIsDirectory((strPath + lang_type).c_str()))
	{
		CreateDirectory((strPath + lang_type).c_str(), NULL);
	}
	if (!createExcel(excel_app))
		return;
	//得到抽离配置更新文件
	getOneTransFile(getTableLanguage(), excel_app, strPath, lang_type);
	//得到静态文本更新文件
	getOneTransFile(getLabelLanguage(), excel_app, strPath, lang_type);
	//得到开发配置翻译更新文件
	getOneTransFile(getDevelopLanguage(), excel_app, strPath, lang_type);
	//得到客户端翻译更新文件
	getOneTransFile(getClientLanguage(), excel_app, strPath, lang_type);
	//得到服务器翻译更新文件
	getOneTransFile(getServerLanguage(), excel_app, strPath, lang_type);
	excel_app->Quit(); 
}
void getTableExcel(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	string lang_type;
	std::cout << "请输入语言版本的后缀如英文：en \n请输入：" << endl;
	std::cin >> lang_type;
	if (!PathIsDirectory((strPath + lang_type).c_str()))
	{
		CreateDirectory((strPath + lang_type).c_str(), NULL);
	}
	if (!createExcel(excel_app))
		return;
	//得到抽离配置更新文件
	CTableLanguage *table = getTableLanguage();
	table->setStrSourcePath(strPath + "data\\");
	table->setStrTransPath(strPath + lang_type + "\\");
	table->BuildLanguageFile(excel_app, lang_type);
	excel_app->Quit();
}
void mergeOneFile(BaseRWExcel *tar, Excel::_ApplicationPtr excelApp, string strPath, string lang_type)
{
	tar->setStrSourcePath(strPath + "data\\");
	tar->setStrTransPath(strPath + lang_type + "\\");
	tar->MergeLanguageFile(excelApp, lang_type);
	delete(tar);
}
void mergeFiles(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	if (!createExcel(excel_app))
		return;
	string lang_type;
	std::cout << "请输入语言版本的后缀如英文：en \n请输入：" << endl;
	std::cin >> lang_type;
	if (!PathIsDirectory((strPath + lang_type).c_str()))
	{
		cout << "需要合并的目录不存在" << endl;
		return;
	}
	//得到抽离配置更新文件
	mergeOneFile(getTableLanguage(), excel_app, strPath, lang_type);
	//得到静态文本更新文件
	mergeOneFile(getLabelLanguage(), excel_app, strPath, lang_type);
	//得到开发配置翻译更新文件
	mergeOneFile(getDevelopLanguage(), excel_app, strPath, lang_type);
	//得到服务器翻译更新文件
	mergeOneFile(getServerLanguage(), excel_app, strPath, lang_type);
	//得到客户端翻译更新文件
	mergeOneFile(getClientLanguage(), excel_app, strPath, lang_type);
	excel_app->Quit();
}
void getLabelExcel(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	if (!createExcel(excel_app))
		return;
	CLabelLanguage *lbl_lang = getLabelLanguage();
	lbl_lang->setStrSourcePath(strPath + "data\\");
	lbl_lang->StringChangeExcel(excel_app);
	delete(lbl_lang);
	excel_app->Quit();
}
void getDevelopExcel(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	if (!createExcel(excel_app))
		return;
	CDevelopLanguage *dev_lang = getDevelopLanguage();
	dev_lang->setStrSourcePath(strPath + "data\\");
	dev_lang->StringChangeExcel(excel_app);
	delete(dev_lang);
	excel_app->Quit();
}
void getSkillExcel(string strPath)
{
	Excel::_ApplicationPtr excel_app;
	if (!createExcel(excel_app))
		return;
	CSkillLanguage *lang = new CSkillLanguage();
	lang->setStrBaseFileName("skill_desc");
	lang->setStrTransFileName("labellang_change");
	lang->setStrSourcePath(strPath + "data\\");
	lang->StringChangeExcel(excel_app);
	delete(lang);
	excel_app->Quit();
}
int main()
{
	char local_url[MAX_PATH];
	_getcwd(local_url, MAX_PATH);
	string str_path = string(local_url) + "\\";
	CoInitialize(NULL);

	std::cout << "注意：请把此工具跟策划配置文件的data放在同一个目录下" << endl;
	std::cout << "1抽离中文 \n2得到更新翻译文件 \n3合并翻译 \n4生成静态文本的excel \n5生成客户端开发填翻译的excel \n6得到某个语言版本的抽离翻译 \n7得到技能翻译 \n请输入：" << endl;
	int select;
	std::cin >> select;
	if (select == 1)
	{
		leaveLanguage(str_path);
	}
	else if(select == 2){
		getTransFiles(str_path);
	}
	else if (select == 3)
	{
		mergeFiles(str_path);
	}
	else if (select == 4)
	{
		getLabelExcel(str_path);
	}
	else if (select == 5)
	{
		getDevelopExcel(str_path);
	}
	else if (select == 6)
	{
		getTableExcel(str_path);
	}
	else if (select == 7)
	{
		getSkillExcel(str_path);
	}
	CoUninitialize();
	system("pause");
	return 0;
}
