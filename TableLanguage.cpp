#include "TableLanguage.h"
#include  <direct.h>  
struct StructTableData
{
	int id;
	_bstr_t language;
	_bstr_t table_name;
	_bstr_t desc;
	_bstr_t state;
};

CTableLanguage::CTableLanguage()
{
}


CTableLanguage::~CTableLanguage()
{
}
_bstr_t CTableLanguage::connectKey(_bstr_t value1, _bstr_t value2, _bstr_t connect)
{
	return value1 + connect + value2;
}
void CTableLanguage::setFileHead(Excel::RangePtr range)
{
	//创建4个头部
	//写入第一列头4位
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "id");
	//写入第三列头4位
	range->PutItem(1, 2, "table_name");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "文字内容");
	//写入第二列头4位
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "中文描述");
	//写入第四列头4位
	range->PutItem(1, 4, "language");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "client");
	range->PutItem(4, 4, "文字内容");
	//写入第五列头4位
	range->PutItem(1, 5, "state");
	range->PutItem(2, 5, "int");
	range->PutItem(3, 5, "none");
	range->PutItem(4, 5, "翻译状态");
	range->PutColumnWidth(80);
}
void CTableLanguage::readChange(Excel::_WorksheetPtr &sheet, list<StructChange> &lstChange, vector<string> &vecFiles)
{
	Excel::RangePtr range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;

	std::list<int> lstID;
	bool create;
	bool point_file = vecFiles.size() > 0;

	for (int i = 2; i <= row_count; i++)
	{
		_variant_t file_name = range->GetItem(i, 1);
		//如果是指定文件，但是在文件名字列表找不到，则跳过
		if (point_file && !FindValue(vecFiles, string(_bstr_t(file_name))))
		{
			continue;
		}
		create = false;
		StructChange stuc;
		//保存表名字
		stuc.strTableName = range->GetItem(i, 1);
		//读取第一个中文字段
		std::cout << "更新表： " << _bstr_t(file_name) << endl;

		if (_bstr_t(range->GetItem(i, 2)) != _bstr_t(""))
		{
			create = true;
			stuc.strKey1 = range->GetItem(i, 2);
		}
		//读取第二个中文字段
		if (_bstr_t(range->GetItem(i, 3)) != _bstr_t(""))
		{
			create = true;
			stuc.strKey2 = range->GetItem(i, 3);
		}
		//读取第二个中文字段
		if (_bstr_t(range->GetItem(i, 4)) != _bstr_t(""))
		{
			create = true;
			stuc.strKey3 = range->GetItem(i, 4);
		}
		if (create)
		{
			stuc.startID = range->GetItem(i, 5);
			if (stuc.startID != 0)
			{
				lstChange.push_back(stuc);
			}
		}
	}
}
void CTableLanguage::readOneExcel(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, StructChange &stucChange)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	bool need_save = false;
	if (GetSheet(string(stucChange.strTableName), getStrSourcePath(), excelApp, sheet))
	{
		book = excelApp->GetActiveWorkbook();
		range = sheet->Cells;
	}
	else{
		std::cout << stucChange.strTableName << "表， 不存在，无法打开" << endl;
		return;
	}
	int row_count = sheet->GetUsedRange()->Rows->Count;
	int column_count = sheet->GetUsedRange()->Columns->Count;
	int max_column = 0;
	int set_column;
	list<StructID> lst_id;
	map<_bstr_t, StructID> map_have_column;
	map<_bstr_t, StructID>::iterator it;
	_bstr_t key;
	_bstr_t old_tar;
	for (int i = 1; i <= column_count; i++)
	{
		//每一列第一个值不为空，才写进去
		if (_bstr_t(range->GetItem(1, i)) != _bstr_t(""))
		{
			StructID have;
			have.strKey = _bstr_t(range->GetItem(1, i));
			have.oldID = i;
			have.newID = i;
			map_have_column.insert(map<_bstr_t, StructID>::value_type(have.strKey, have));
			max_column++;
		}
	}
	for (int i = 1; i <= max_column; i++)
	{
		key = "";
		if (stucChange.strKey1 == _bstr_t(range->GetItem(1, i)))
		{
			key = stucChange.strKey1;
		}
		else if (stucChange.strKey2 == _bstr_t(range->GetItem(1, i)))
		{
			key = stucChange.strKey2;
		}
		else if (stucChange.strKey3 == _bstr_t(range->GetItem(1, i)))
		{
			key = stucChange.strKey3;
		}
		if (key != _bstr_t(""))
		{
			StructID stuc_id;
			it = map_have_column.find(key + "_id");
			if (map_have_column.end() != it)
			{
				stuc_id.oldID = map_have_column[key].oldID;
				stuc_id.newID = map_have_column[key + "_id"].newID;
			}
			else{
				old_tar = range->GetItem(3, i);
				//添加一个新的列
				range->PutItem(1, ++max_column, key + "_id");
				range->PutItem(2, max_column, "int");
				//如果新的发布对象跟旧的发布对象不一致，要保存
				if ((_bstr_t)range->GetItem(3, max_column) != old_tar)
					need_save = true;
				//保留原来导出设置，因为可能会导出一个服务器需要的单独的语言表
				range->PutItem(3, max_column, old_tar);
				range->PutItem(4, max_column, key + "_id");
				stuc_id.oldID = i;
				stuc_id.newID = max_column;
				//该位置放入列表
				//改成不发布到客户端和服务器
				range->PutItem(3, i, "none");
			}
			lst_id.push_back(stuc_id);
		}
	}
	int nowid = stucChange.startID;
	_bstr_t str_data;
	list<StructID>::iterator itor = lst_id.begin();
	map<_bstr_t, int> map_data_to_key;
	map<_bstr_t, int>::iterator map_itor;
	for (int i = 5; i <= row_count; i++)
	{
		for (list<StructID>::iterator itor = lst_id.begin(); itor != lst_id.end(); itor++)
		{
			//不为空才写进去
			str_data = _bstr_t(range->GetItem(i, (*itor).oldID));
			if (str_data != _bstr_t(""))
			{
				map_itor = map_data_to_key.find(str_data);
				//判断这个表里面是否已经有这样一句文字了
				if (map_itor == map_data_to_key.end())
				{
					//没有这个文字，需要新增
					StructLanguageData data;
					data.strData = str_data;
					data.strTableName = stucChange.strTableName;
					data.strTarget = range->GetItem(3, (*itor).newID);
					//将ID对应的中文放入map
					mapLang.insert(map<int, StructLanguageData>::value_type(nowid, data));
					map_data_to_key.insert(map<_bstr_t, int>::value_type(str_data, nowid));
					//如果这个位置的ID变更了也要保存
					if ((_bstr_t)range->GetItem(i, (*itor).newID) == _bstr_t("") || (int)range->GetItem(i, (*itor).newID) != nowid)
						need_save = true;
					range->PutItem(i, (*itor).newID, nowid);
					nowid++;
				}
				else{
					//已经有这个文字了。则取文字的ID即可
					range->PutItem(i, (*itor).newID, (*map_itor).second);
				}
			}
		}
	}
	book->Close(need_save);
}
//保存基础语言表
void CTableLanguage::saveBaseLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	int row_count = 0;
	//文件不存在，则创建
	if (!GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet))
	{
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		range = sheet->Cells;
		sheet->Name = getStrBaseFileName().c_str();
		sheet->SaveAs((getLanguagePath() + getStrBaseFileName()).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	}
	else{
		book = excelApp->GetActiveWorkbook();
		range = sheet->Cells;
		row_count = sheet->GetUsedRange()->Rows->Count;
		printf("row_count = %d", row_count);
	}
	//写入第一列头4位
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "文字ID");
	//写入第三列头4位（注释表列）
	range->PutItem(1, 2, "table_name");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "文字所在表名");
	//写入第三列头4位
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "文字内容");
	//写入第四列头4位
	range->PutItem(1, 4, "target");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "发布位置");
	range->PutColumnWidth(80);
	int now_row = 5;
	bool need_save = false;
	for (map<int, StructLanguageData>::iterator itor = mapLang.begin(); itor != mapLang.end(); itor++)
	{
		if ((int)range->GetItem(now_row, 1) != (*itor).first || (_bstr_t)range->GetItem(now_row, 2) != (*itor).second.strTableName ||
			(_bstr_t)range->GetItem(now_row, 3) != (*itor).second.strData || (_bstr_t)range->GetItem(now_row, 4) != (*itor).second.strTarget)
		{
			need_save = true;
		}
		range->PutItem(now_row, 1, (*itor).first);
		range->PutItem(now_row, 2, (*itor).second.strTableName);
		range->PutItem(now_row, 3, (*itor).second.strData);
		range->PutItem(now_row, 4, (*itor).second.strTarget);
		now_row++;
	}
	for (int i = now_row; i <= row_count; i++)
	{
		need_save = true;
		range->PutItem(i, 1, "");
		range->PutItem(i, 2, "");
		range->PutItem(i, 3, "");
		range->PutItem(i, 4, "");
	}
	book->Close(need_save);
}
//保存一个Table的语言，分开服务器客户端
void CTableLanguage::saveOneLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, _bstr_t tar, string strFileName)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	bool need_save = false;
	//文件不存在，则创建
	if (!GetSheet(strFileName, getStrSourcePath(), excelApp, sheet))
	{
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->Name = strFileName.c_str();
		sheet->SaveAs((getStrSourcePath() + strFileName).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
		//写入第一列头4位
		range->PutItem(1, 1, "id");
		range->PutItem(2, 1, "int");
		range->PutItem(3, 1, tar);
		range->PutItem(4, 1, "文字ID");
		//写入第三列头4位（注释表列）
		range->PutItem(1, 2, "language");
		range->PutItem(2, 2, "string");
		range->PutItem(3, 2, tar);
		range->PutItem(4, 2, "文字内容");
		range->PutColumnWidth(80);
		need_save = true;
	}else{
		book = excelApp->GetActiveWorkbook();
		range = sheet->Cells;
	}
	int now_row = 5;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	for (map<int, StructLanguageData>::iterator itor = mapLang.begin(); itor != mapLang.end(); itor++)
	{
		if ((*itor).second.strTarget == _bstr_t("all") || (*itor).second.strTarget == tar)
		{
			if ((int)range->GetItem(now_row, 1) != (*itor).first || (_bstr_t)range->GetItem(now_row, 2) != (*itor).second.strData)
			{
				need_save = true;
			}
			range->PutItem(now_row, 1, (*itor).first);
			range->PutItem(now_row, 2, (*itor).second.strData);
			now_row++;
		}
	}
	if (now_row != row_count + 1)
		need_save = true;
	//如果当前数量小于之前数量，表示删除了数据
	if (now_row < row_count + 1)
	{
		for (int i = now_row; i <= row_count; i++)
		{
			range->PutItem(i, 1, _bstr_t(""));
			range->PutItem(i, 2, _bstr_t(""));
		}
	}
	book->Close(need_save);
}
//读取中文基础表
bool CTableLanguage::ReadBaseLanguageFile(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, string strPath, string strFileName, vector<string> &vecFiles)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	//文件不存在，就不用执行下面了
	if (!GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet))
	{
		return false;
	}
	bool all_file = vecFiles.size() == 0;
	range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	int id_index = GetColumnID(sheet, "id");
	int table_name_index = GetColumnID(sheet, "table_name");
	int language_index = GetColumnID(sheet, "language");
	int target_index = GetColumnID(sheet, "target");
	for (int i = 5; i <= row_count; i++)
	{
		//如果是找全部文件，或者在更新列表中没有找到这个表的话。就使用旧数据
		if (all_file || !FindValue(vecFiles, string(_bstr_t(range->GetItem(i, 2)))))
		{
			StructLanguageData data;
			data.strTableName = range->GetItem(i, table_name_index);
			data.strData = range->GetItem(i, language_index);
			data.strTarget = range->GetItem(i, target_index);
			//将ID对应的中文放入map
			mapLang.insert(map<int, StructLanguageData>::value_type(range->GetItem(i, id_index), data));
		}
	}
	book = excelApp->GetActiveWorkbook();
	book->Close(false);
	return true;
}
//读取中文基础表中各个翻译内容发布的对象（客户端还是服务端）
map<_bstr_t, _bstr_t> CTableLanguage::GetBaseTargetMap(Excel::_ApplicationPtr excelApp)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	map<_bstr_t, _bstr_t> map_target;
	//文件不存在，就不用执行下面了
	if (!GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet))
	{
		return map_target;
	}
	range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	int id_index = GetColumnID(sheet, "id");
	int table_name_index = GetColumnID(sheet, "table_name");
	int language_index = GetColumnID(sheet, "language");
	int target_index = GetColumnID(sheet, "target");
	std::cout << row_count << endl << id_index << table_name_index << language_index << target_index << endl;
	for (int i = 5; i <= row_count; i++)
	{
		//生成一个对象的map
		map_target.insert(map<_bstr_t, _bstr_t>::value_type(connectKey((_bstr_t)range->GetItem(i, language_index), (_bstr_t)range->GetItem(i, table_name_index), strConnect), (_bstr_t)range->GetItem(i, target_index)));
	}
	book = excelApp->GetActiveWorkbook();
	book->Close(false);
	return map_target;
}
void CTableLanguage::LeaveLanguage(Excel::_ApplicationPtr excelApp, bool create_all, string strFiles)
{
	std::vector<string> vec_file;
	if (!create_all)
	{
		vec_file = Split(strFiles, ",");
	}

	Excel::_WorksheetPtr sheet;
	map<int, StructLanguageData> map_lang;
	list<StructChange> lstChange;
	//如果有旧的文件，而且不是创建全部，也读取旧的文件做合并
	if (!create_all)
	{
		ReadBaseLanguageFile(excelApp, map_lang, getLanguagePath(), getStrBaseFileName(), vec_file);
	}
	//读取lang_change的配置
	GetSheet("lang_config", getStrSourcePath(), excelApp, sheet);
	//将各个表需要修改的字段保存下来
	readChange(sheet, lstChange, vec_file);
	//遍历需要变更的表
	list<StructChange>::iterator itor = lstChange.begin();
	while (itor != lstChange.end())
	{
		readOneExcel(excelApp, map_lang, *itor);
		itor++;
	}
	saveBaseLanguage(excelApp, map_lang);
	//保存客户端文件
	saveOneLanguage(excelApp, map_lang, "client", strClientName);
	saveOneLanguage(excelApp, map_lang, "server", strServerName);
}

void CTableLanguage::BuildLanguageFile(Excel::_ApplicationPtr excelApp, string strPlatform)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	vector<string> vec;
	map<int, StructLanguageData> map_lang;
	string str_base_file = "";
	string str_end_name = "";
	if (strPlatform == "ch" || strPlatform == "")
	{
		str_base_file = getStrBaseFileName();
		str_end_name = "";
	}
	else{
		str_base_file = getStrBaseFileName() + "_" + strPlatform;
		str_end_name = "_" + strPlatform;
	}
	bool find_file = false;
	//如果不为空的话，表示更新非中文，非中文要从基础语言表中得到发布对象
	if (str_end_name != "")
	{
		//得到基础语言表中各句话的发布对象
		map<_bstr_t, _bstr_t> map_target = GetBaseTargetMap(excelApp);
		//读取翻译来源的文件
		if (!GetSheet(str_base_file, getLanguagePath(), excelApp, sheet))
		{
			std::cout << "源翻译文件不存在！" << endl;
			return;
		}
		int row_count = sheet->GetUsedRange()->Rows->Count;
		int id_index = GetColumnID(sheet, "id");
		int table_name_index = GetColumnID(sheet, "table_name");
		int language_index = GetColumnID(sheet, "language");
		int desc_index = GetColumnID(sheet, "desc");
		range = sheet->Cells;
		for (int i = 5; i <= row_count; i++)
		{
			if (map_target[connectKey((_bstr_t)range->GetItem(i, desc_index), (_bstr_t)range->GetItem(i, table_name_index), strConnect)] != _bstr_t(""))
			{
				StructLanguageData data;
				data.strTableName = range->GetItem(i, table_name_index);
				data.strData = range->GetItem(i, language_index);
				data.strTarget = map_target[connectKey((_bstr_t)range->GetItem(i, desc_index), (_bstr_t)range->GetItem(i, table_name_index), strConnect)];
				//翻译如果不是简体和繁体，错误码的表，如果描述跟翻译一样的话，表示没翻译，没翻译就不显示了，游戏会自动显示错误码ID
				if (!(strPlatform != "ch" && strPlatform != "tw" && data.strTableName == _bstr_t("tb_table_errcode") && range->GetItem(i, language_index) == range->GetItem(i, desc_index)))
				{
					//将ID对应的中文放入map
					map_lang.insert(map<int, StructLanguageData>::value_type(range->GetItem(i, id_index), data));
				}
			}
		}
		find_file = true;
		book = excelApp->GetActiveWorkbook();
		book->Close(false);
	}
	else{
		find_file = ReadBaseLanguageFile(excelApp, map_lang, getLanguagePath(), getStrBaseFileName(), vec);
	}
	if (find_file)
	{
		//保存客户端文件
		saveOneLanguage(excelApp, map_lang, "client", strClientName + str_end_name);
		saveOneLanguage(excelApp, map_lang, "server", strServerName + str_end_name);
	}
	else{
		std::cout << "源翻译文件不存在！" << endl;
	}
}


//---------------------------------------------------------------提取更新文件相关begin--------------------------------------------------------//
///比较一个表的变更
void compareOneFile(Excel::_WorksheetPtr &sheet, list<_bstr_t> &lstComData, StructChange &stucChange, list<StructLanguageData> &lstChange)
{
	Excel::RangePtr range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	int column_count = sheet->GetUsedRange()->Columns->Count;
	list<int> lst_id;
	map<_bstr_t, StructID> map_have_column;
	map<_bstr_t, StructID>::iterator it;
	//得到有中文的字段ID
	for (int i = 1; i <= column_count; i++)
	{
		//每一列第一个值不为空，才写进去
		if ((stucChange.strKey1 != _bstr_t("") && _bstr_t(range->GetItem(1, i)) == stucChange.strKey1) ||
			(stucChange.strKey2 != _bstr_t("") && _bstr_t(range->GetItem(1, i)) == stucChange.strKey2))
		{
			lst_id.push_back(i);
		}
	}

	_bstr_t str_data;
	list<_bstr_t>::iterator com_itor;
	map<_bstr_t, int> map_data_to_key;
	map<_bstr_t, int>::iterator map_itor;
	//遍历表中的中文，如果中文在原来表中已经有，不处理，如果没有则放入变更表里
	for (int i = 5; i <= row_count; i++)
	{
		for (list<int>::iterator itor = lst_id.begin(); itor != lst_id.end(); itor++)
		{
			//不为空才写进去
			str_data = _bstr_t(range->GetItem(i, *itor));
			com_itor = std::find(lstComData.begin(), lstComData.end(), str_data);
			//原来的表找不到需要的数据，则写入更新列表
			if (com_itor == lstComData.end())
			{
				//如果显示内容不为空，则将放入新增列表里面
				if (str_data != _bstr_t(""))
				{
					map_itor = map_data_to_key.find(str_data);
					//判断这个表里面是否已经有这样一句文字了
					if (map_itor == map_data_to_key.end())
					{
						//没有这个文字，需要新增
						StructLanguageData data;
						data.strData = str_data;
						data.strTableName = stucChange.strTableName;
						lstChange.push_back(data);
						map_data_to_key.insert(map<_bstr_t, int>::value_type(str_data, 1));
					}
				}
			}
		}
	}
}
///得到翻译更新文件
void CTableLanguage::GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	string str_save_name = getStrTransFileName() + "_" + strLangType;
	Excel::_WorkbookPtr book;
	Excel::_WorksheetPtr sheet;
	map<_bstr_t, list<_bstr_t>> map_change;
	list<StructChange> lstChange;
	list<StructLanguageData> lst_refresh;
	Excel::RangePtr range;
	//如果没有原来的翻译表，不能比较
	if (!GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet))
	{
		std::cout << "还没有away_language表，不能比较！" << endl;
		excelApp->Workbooks->Close();
		return;
	}
	//读取源表的数据
	GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet);
	range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	list<StructTableData> lst_source_data;
	for (int i = 5; i <= row_count; i++)
	{
		StructTableData stuc;
		stuc.id = range->GetItem(i, 1);
		stuc.table_name = range->GetItem(i, 2);
		stuc.desc = range->GetItem(i, 3);
		stuc.language = range->GetItem(i, 3);//源翻译的desc和language一样
		lst_source_data.push_back(stuc);
	}
	map<_bstr_t, StructTableData> map_now_trans;
	//读取该语言表的数据
	if (GetSheet(str_table_file, getLanguagePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			StructTableData stuc;
			stuc.id = range->GetItem(i, 1);
			stuc.table_name = range->GetItem(i, 2);
			stuc.desc = range->GetItem(i, 3);
			stuc.language = range->GetItem(i, 4);
			stuc.state = range->GetItem(i, 5);
			//将翻译内容+加所在表名连接起来作为KEY值
			map_now_trans.insert(map<_bstr_t, StructTableData>::value_type(connectKey(stuc.desc, stuc.table_name, strConnect), stuc));
		}
		range->Clear();
	}
	else{
		//生成新表，并且保存
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getLanguagePath() + str_table_file).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
	}
	sheet->Name = str_table_file.c_str();
	//填入头部4行
	setFileHead(range);
	int now_row = 5;
	list<StructTableData> lst_change;
	_bstr_t key;
	//遍历新的源表，得到差异
	for (list<StructTableData>::iterator itor = lst_source_data.begin(); itor != lst_source_data.end(); itor++)
	{
		//更新基本信息
		range->PutItem(now_row, 1, (*itor).id);
		range->PutItem(now_row, 2, (*itor).table_name);
		range->PutItem(now_row, 3, (*itor).desc);
		key = connectKey((*itor).desc, (*itor).table_name, strConnect);
		//如果旧map中有数据，则没有更新
		if (map_now_trans.find(key) != map_now_trans.end() && map_now_trans[key].language != _bstr_t("") && map_now_trans[key].state == strHaveTrans)
		{
			//将旧翻译数据写入
			range->PutItem(now_row, 4, map_now_trans[connectKey((*itor).desc, (*itor).table_name, strConnect)].language);
			range->PutItem(now_row, 5, 1);
		}
		else{
			//将中文写入
			range->PutItem(now_row, 4, (*itor).desc);
			range->PutItem(now_row, 5, "");
			//放入更新列表
			lst_change.push_back((*itor));
		}
		now_row++;
	}
	excelApp->ActiveWorkbook->Save();
	excelApp->Workbooks->Close();
	//得到更新文件
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
	range->PutItem(1, 1, "table_name");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "文字在的表名");
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
	for (list<StructTableData>::iterator itor = lst_change.begin(); itor != lst_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).table_name);
		range->PutItem(now_row, 2, (*itor).desc);
		now_row++;
	}
	//保存文件
	excelApp->ActiveWorkbook->Save();
	//关闭文件
	excelApp->Workbooks->Close();
}
//---------------------------------------------------------------提取更新文件相关end--------------------------------------------------------//
void CTableLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
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
			key = connectKey(_bstr_t(range->GetItem(i, 2)), _bstr_t(range->GetItem(i, 1)), strConnect);
			map_trans_data.insert(map<_bstr_t, _bstr_t>::value_type(key, this->getTranslateData(range->GetItem(i, 3))));
		}
	}
	else{
		cout << "翻译文件不存在不存在文件:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//读取该语言的配置
	if (GetSheet(str_table_file, getLanguagePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			key = connectKey(_bstr_t(range->GetItem(i, 3)), _bstr_t(range->GetItem(i, 2)), strConnect);
			_bstr_t test = _bstr_t(range->GetItem(i, 4));
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
string CTableLanguage::getLanguagePath()
{
	return getStrSourcePath() + "language\\";
}