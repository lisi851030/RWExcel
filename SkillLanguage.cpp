#include "SkillLanguage.h"
#include "LabelLanguage.h"
#include  <direct.h>  
#include <fstream>
#include <list>
#include <map>

CSkillLanguage::CSkillLanguage()
{
}


CSkillLanguage::~CSkillLanguage()
{
}
void CSkillLanguage::StringChangeExcel(Excel::_ApplicationPtr excelApp)
{
	ifstream in;
	in.open("skill.lua");
	string data;
	char ch;
	while (!in.eof())

	{
		in.read(&ch, 1);
		data += ch;
	}
	string test = GBKToUTF8(data);

	list<vector<string>> lst_data;
	vector<string> vec_datas;
	vector<string> vec_one_data;
	vec_datas = Split(data, "@");
	vector<string>::iterator vec_itor = vec_datas.begin();
	for (int i = 1; i < vec_datas.size(); i++)
	{
		if (vec_datas[i] != "")
		{
			cout << vec_datas[i] << "   " << i << endl;
			vec_one_data = Split(vec_datas[i], "$");
			if (vec_one_data.size() == 2)
			{
				lst_data.push_back(vec_one_data);
			}
		}
	}

	Excel::_WorksheetPtr sheet;

	//保存变更文件
	Excel::_WorkbookPtr book = excelApp->Workbooks->Add();
	sheet = book->Sheets->Add();
	sheet->Name = L"skil_desc";
	Excel::RangePtr range = sheet->Cells;
	//写入第一列头4位
	range->PutItem(1, 1, "skill_id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "中文所在UI");
	//写入第三列头4位
	range->PutItem(1, 2, "skill_des");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "client");
	range->PutItem(4, 2, "文字内容");
	range->PutColumnWidth(80);
	int now_row = 5;
	for (list<vector<string>>::iterator itor = lst_data.begin(); itor != lst_data.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor)[0].c_str());
		range->PutItem(now_row, 2, (*itor)[1].c_str());
		now_row++;
	}
	sheet->SaveAs((getStrSourcePath() + "skil_desc").c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}

