#include "DevelopLanguage.h"
#include <map>
#include  <direct.h>  
#include <fstream>
#include <map>
#include <list>


CDevelopLanguage::CDevelopLanguage()
{
}


CDevelopLanguage::~CDevelopLanguage()
{
}
void CDevelopLanguage::setFileHead(Excel::RangePtr range)
{
	//创建4个头部
	//写入第一列头4位
	range->PutItem(1, 1, "key_str");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "key值");
	//写入第二列头4位
	range->PutItem(1, 2, "desc");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "中文描述");
	//写入第三列头4位
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "mix");
	range->PutItem(3, 3, "client");
	range->PutItem(4, 3, "文字内容");
	//写入第四列头4位
	range->PutItem(1, 4, "state");
	range->PutItem(2, 4, "int");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "翻译状态");
	range->PutColumnWidth(80);
}
void CDevelopLanguage::StringChangeExcel(Excel::_ApplicationPtr excelApp)
{
	ifstream in;
	in.open("language.lua");
	string data;
	char ch;
	while (!in.eof())

	{
		in.read(&ch, 1);
		data += ch;
	}
	string test = GBKToUTF8(data);

	map<string, string> map_data;
	vector<string> vec_datas;
	vector<string> vec_one_data;
	vec_datas = Split(data, "§");
	vector<string>::iterator vec_itor = vec_datas.begin();
	for (int i = 1; i < vec_datas.size(); i++)
	{
		if (vec_datas[i] != "")
		{
			cout << vec_datas[i] << "   " << i << endl;
			vec_one_data = Split(vec_datas[i], "№");
			if (vec_one_data.size() == 3)
			{
				if (vec_one_data[2] == "table")
				{
					map_data.insert(map<string, string>::value_type(vec_one_data[0], vec_one_data[1]));
				}
				else{
					map_data.insert(map<string, string>::value_type(vec_one_data[0], vec_one_data[1]));
				}
			}
		}
	}
	Excel::_WorksheetPtr sheet;

	//保存变更文件
	Excel::_WorkbookPtr book = excelApp->Workbooks->Add();
	sheet = book->Sheets->Add();
	sheet->Name = L"tb_table_developlang";
	Excel::RangePtr range = sheet->Cells;
	//写入第一列头4位
	range->PutItem(1, 1, "key_str");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "中文的key值");
	//写入第二列头4位
	range->PutItem(1, 2, "language");
	range->PutItem(2, 2, "mix");
	range->PutItem(3, 2, "client");
	range->PutItem(4, 2, "文字内容");
	range->PutColumnWidth(80);
	int now_row = 5;
	for (map<string, string>::iterator itor = map_data.begin(); itor != map_data.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).first.c_str());
		range->PutItem(now_row, 2, (*itor).second.c_str());
		now_row++;
	}
	sheet->SaveAs((getStrSourcePath() + "tb_table_developlang").c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}
void CDevelopLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	list<StructIDToLanguage> lst_data;
	map<_bstr_t, _bstr_t> map_trans_data;
	//读取翻译配置
	string str_save_name = getStrTransFileName() + "_" + strLangType;
	if (GetSheet(str_save_name, getTransPath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			map_trans_data.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)), this->getTranslateData(range->GetItem(i, 3))));
		}
	}
	else{
		cout << "翻译文件不存在不存在文件:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	_bstr_t key;
	//读取该语言的配置
	if (GetSheet(str_table_file, getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			key = _bstr_t(range->GetItem(i, 1));
			if (map_trans_data[key] != _bstr_t(""))
			{
				range->PutItem(i, 3, map_trans_data[key]);
				//设置已有翻译状态
				range->PutItem(i, 4, 1);
			}
			else if (_bstr_t(range->GetItem(i, 3)) == _bstr_t("")){
				//如果沒有翻譯，設置中文進去
				range->PutItem(i, 3, (_bstr_t)range->GetItem(i, 2));
			}
		}
	}
	else{
		cout << "不存在文件:" << str_table_file << endl;
		return;
	}
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}
