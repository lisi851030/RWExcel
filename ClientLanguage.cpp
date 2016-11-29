#include "ClientLanguage.h"
#include <map>
#include <list>


CClientLanguage::CClientLanguage()
{
}


CClientLanguage::~CClientLanguage()
{
}
void CClientLanguage::setFileHead(Excel::RangePtr range)
{
	//创建4个头部
	//写入第一列头4位
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "id");
	//写入第二列头4位
	range->PutItem(1, 2, "desc");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "中文描述");
	//写入第三列头4位
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "client");
	range->PutItem(4, 3, "文字内容");
	//写入第四列头4位
	range->PutItem(1, 4, "state");
	range->PutItem(2, 4, "int");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "翻译状态");
	range->PutColumnWidth(80);
}
void CClientLanguage::GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	map<_bstr_t, _bstr_t> vec_old_data;
	map<_bstr_t, _bstr_t> vec_new_data;
	map<_bstr_t, _bstr_t> vec_trans_data;
	map<_bstr_t, _bstr_t> vec_trans_state;
	//读取新表的数据
	if (GetSheet(getStrBaseFileName(), getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			if (_bstr_t(range->GetItem(i, 1)) != _bstr_t(""))
			{
				vec_new_data.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2))));
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
			vec_old_data.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2))));
			vec_trans_data.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 3))));
			vec_trans_state.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 4))));
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
	map<_bstr_t, _bstr_t>::iterator itor = vec_new_data.begin();
	map<_bstr_t, _bstr_t> vec_change;
	int now_row = 5;
	while (itor != vec_new_data.end())
	{
		range->PutItem(now_row, 1, (*itor).first);
		range->PutItem(now_row, 2, (*itor).second);
		range->PutItem(now_row, 3, vec_trans_data[(*itor).first]);
		//判断两个值是否相同，相同则不用处理，不同则放入变更列表
		if (vec_old_data[(*itor).first] != (*itor).second || vec_trans_state[(*itor).first] != strHaveTrans)
		{
			vec_change.insert(map<_bstr_t, _bstr_t>::value_type((*itor).first, (*itor).second));
			range->PutItem(now_row, 4, "");
		}
		else {
			//如果内容还相同，则将原来表里面的翻译写进去
			range->PutItem(now_row, 3, vec_trans_data[(*itor).first]);
			range->PutItem(now_row, 4, 1);
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
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "文字的ID");
	//写入第二列头4位
	range->PutItem(1, 2, "desc");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "中文描述");
	//写入第三列头4位
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "翻译内容");
	range->PutColumnWidth(80);
	now_row = 5;
	for (map<_bstr_t, _bstr_t>::iterator itor = vec_change.begin(); itor != vec_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).first);
		range->PutItem(now_row, 2, (*itor).second);
		now_row++;
	}
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}
void CClientLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
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
			map_trans_data.insert(map<_bstr_t, _bstr_t>::value_type(_bstr_t(range->GetItem(i, 1)) + strConnect + _bstr_t(range->GetItem(i, 2)), this->getTranslateData(range->GetItem(i, 3))));
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
			key = _bstr_t(range->GetItem(i, 1)) + strConnect + _bstr_t(range->GetItem(i, 2));
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