#include "ServerLanguage.h"
#include <map>
#include <list>


struct StructServerData
{
	int lang_type;
	int lang_id;
	_bstr_t lang_string;
	_bstr_t lang_trans;
	_bstr_t state;
};
CServerLanguage::CServerLanguage()
{
}


CServerLanguage::~CServerLanguage()
{
}
_bstr_t connectKey(_bstr_t value1, _bstr_t value2, _bstr_t value3, _bstr_t connect)
{
	return value1 + connect + value2 + connect + value3;
}

void CServerLanguage::setFileHead(Excel::RangePtr range)
{
	//创建4个头部
	//写入第一列头4位
	range->PutItem(1, 1, "lang_type");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "all");
	range->PutItem(4, 1, "文字类型");
	//写入第二列头4位
	range->PutItem(1, 2, "lang_id");
	range->PutItem(2, 2, "int");
	range->PutItem(3, 2, "all");
	range->PutItem(4, 2, "文字ID");
	//写入第三列头4位
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "中文描述");
	//写入第四列头4位
	range->PutItem(1, 4, "lang_string");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "all");
	range->PutItem(4, 4, "文字内容");
	//写入第五列头4位
	range->PutItem(1, 5, "state");
	range->PutItem(2, 5, "int");
	range->PutItem(3, 5, "none");
	range->PutItem(4, 5, "翻译状态");
	range->PutColumnWidth(80);
}

void CServerLanguage::GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	map<_bstr_t, StructServerData> vec_old_data;
	map<_bstr_t, StructServerData> vec_new_data;
	_bstr_t key;
	//读取新表的数据
	if (GetSheet(getStrBaseFileName(), getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			StructServerData stuc;
			if (_bstr_t(range->GetItem(i, 3)) != _bstr_t(""))
			{
				stuc.lang_type = range->GetItem(i, 1);
				stuc.lang_id = range->GetItem(i, 2);
				stuc.lang_string = range->GetItem(i, 3);
				//将一二个字段连接起来作为KEY值
				vec_new_data.insert(map<_bstr_t, StructServerData>::value_type(connectKey(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2)), _bstr_t(""), strConnect), stuc));
			}
		}
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//根据语言类型得到旧的配置
	if (GetSheet(str_table_file, getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			key = connectKey(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2)), _bstr_t(""), strConnect);
			StructServerData stuc;
			stuc.lang_type = range->GetItem(i, 1);
			stuc.lang_id = range->GetItem(i, 2);
			stuc.lang_string = range->GetItem(i, 3);
			stuc.lang_trans = range->GetItem(i, 4);
			stuc.state = range->GetItem(i, 5);
			vec_old_data.insert(map<_bstr_t, StructServerData>::value_type(key, stuc));
		}
		range->Clear();
	}
	else{
		//生成新表，并且保存
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getStrSourcePath() + str_table_file).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
	}
	sheet->Name = str_table_file.c_str();
	//填入头部4行
	setFileHead(range);
	map<_bstr_t, StructServerData>::iterator itor = vec_new_data.begin();
	list<StructServerData> lst_change;
	int now_row = 5;
	while (itor != vec_new_data.end())
	{
		range->PutItem(now_row, 1, (*itor).second.lang_type);
		range->PutItem(now_row, 2, (*itor).second.lang_id);
		range->PutItem(now_row, 3, (*itor).second.lang_string);
		range->PutItem(now_row, 4, vec_old_data[(*itor).first].lang_trans);
		//判断两个值是否相同，相同则不用处理，不同则放入变更列表
		if (vec_old_data[(*itor).first].lang_string != (*itor).second.lang_string || vec_old_data[(*itor).first].lang_trans == _bstr_t("") || vec_old_data[(*itor).first].state != strHaveTrans)
		{
			lst_change.push_back(list<StructServerData>::value_type((*itor).second));
			range->PutItem(now_row, 5, "");
		}
		else {
			//如果内容还相同，则将原来表里面的翻译写进去
			range->PutItem(now_row, 4, vec_old_data[(*itor).first].lang_trans);
			range->PutItem(now_row, 5, 1);
		}
		now_row++;
		itor++;
	}
	excelApp->ActiveWorkbook->Save();
	excelApp->Workbooks->Close();
	string str_save_name = getStrTransFileName() + "_" + strLangType;
	if (!GetSheet(str_save_name, getTransPath(), excelApp, sheet))
	{
		//生成新表，并且保存
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getTransPath() + str_save_name).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	}
	else{
		//清空原来的表
		sheet->Cells->Clear();
	}
	sheet->Name = str_save_name.c_str();
	range = sheet->Cells;
	//写入第一列头4位
	range->PutItem(1, 1, "lang_type");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "all");
	range->PutItem(4, 1, "文字类型");
	//写入第二列头4位
	range->PutItem(1, 2, "lang_id");
	range->PutItem(2, 2, "int");
	range->PutItem(3, 2, "all");
	range->PutItem(4, 2, "文字ID");
	//写入第三列头4位
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "中文描述");
	//写入第四列头4位
	range->PutItem(1, 4, "language");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "翻译内容");
	range->PutColumnWidth(80);
	now_row = 5;
	for (list<StructServerData>::iterator itor = lst_change.begin(); itor != lst_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).lang_type);
		range->PutItem(now_row, 2, (*itor).lang_id);
		range->PutItem(now_row, 3, (*itor).lang_string);
		now_row++;
	}
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}
void CServerLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	list<StructIDToLanguage> lst_data;
	map<_bstr_t, _bstr_t> map_trans_data;
	_bstr_t key;
	//读取翻译配置
	string str_save_name = getStrTransFileName() + "_" + strLangType;
	if (GetSheet(str_save_name, getTransPath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			key = connectKey(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2)), _bstr_t(range->GetItem(i, 3)), strConnect);
			map_trans_data.insert(map<_bstr_t, _bstr_t>::value_type(key, this->getTranslateData(range->GetItem(i, 4))));
		}
	}
	else{
		cout << "翻译文件不存在不存在文件:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//读取该语言的配置
	if (GetSheet(str_table_file, getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			key = connectKey(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2)), _bstr_t(range->GetItem(i, 3)), strConnect);
			if (map_trans_data[key] != _bstr_t(""))
			{
				range->PutItem(i, 4, map_trans_data[key]);
				range->PutItem(i, 5, 1);
			}
			else if (_bstr_t(range->GetItem(i, 4)) == _bstr_t("")){
				//如果沒有翻譯，設置中文進去
				range->PutItem(i, 4, (_bstr_t)range->GetItem(i, 3));
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