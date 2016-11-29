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
	//����4��ͷ��
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "id");
	//д�������ͷ4λ
	range->PutItem(1, 2, "table_name");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "��������");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 4, "language");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "client");
	range->PutItem(4, 4, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 5, "state");
	range->PutItem(2, 5, "int");
	range->PutItem(3, 5, "none");
	range->PutItem(4, 5, "����״̬");
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
		//�����ָ���ļ����������ļ������б��Ҳ�����������
		if (point_file && !FindValue(vecFiles, string(_bstr_t(file_name))))
		{
			continue;
		}
		create = false;
		StructChange stuc;
		//���������
		stuc.strTableName = range->GetItem(i, 1);
		//��ȡ��һ�������ֶ�
		std::cout << "���±� " << _bstr_t(file_name) << endl;

		if (_bstr_t(range->GetItem(i, 2)) != _bstr_t(""))
		{
			create = true;
			stuc.strKey1 = range->GetItem(i, 2);
		}
		//��ȡ�ڶ��������ֶ�
		if (_bstr_t(range->GetItem(i, 3)) != _bstr_t(""))
		{
			create = true;
			stuc.strKey2 = range->GetItem(i, 3);
		}
		//��ȡ�ڶ��������ֶ�
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
		std::cout << stucChange.strTableName << "�� �����ڣ��޷���" << endl;
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
		//ÿһ�е�һ��ֵ��Ϊ�գ���д��ȥ
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
				//���һ���µ���
				range->PutItem(1, ++max_column, key + "_id");
				range->PutItem(2, max_column, "int");
				//����µķ���������ɵķ�������һ�£�Ҫ����
				if ((_bstr_t)range->GetItem(3, max_column) != old_tar)
					need_save = true;
				//����ԭ���������ã���Ϊ���ܻᵼ��һ����������Ҫ�ĵ��������Ա�
				range->PutItem(3, max_column, old_tar);
				range->PutItem(4, max_column, key + "_id");
				stuc_id.oldID = i;
				stuc_id.newID = max_column;
				//��λ�÷����б�
				//�ĳɲ��������ͻ��˺ͷ�����
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
			//��Ϊ�ղ�д��ȥ
			str_data = _bstr_t(range->GetItem(i, (*itor).oldID));
			if (str_data != _bstr_t(""))
			{
				map_itor = map_data_to_key.find(str_data);
				//�ж�����������Ƿ��Ѿ�������һ��������
				if (map_itor == map_data_to_key.end())
				{
					//û��������֣���Ҫ����
					StructLanguageData data;
					data.strData = str_data;
					data.strTableName = stucChange.strTableName;
					data.strTarget = range->GetItem(3, (*itor).newID);
					//��ID��Ӧ�����ķ���map
					mapLang.insert(map<int, StructLanguageData>::value_type(nowid, data));
					map_data_to_key.insert(map<_bstr_t, int>::value_type(str_data, nowid));
					//������λ�õ�ID�����ҲҪ����
					if ((_bstr_t)range->GetItem(i, (*itor).newID) == _bstr_t("") || (int)range->GetItem(i, (*itor).newID) != nowid)
						need_save = true;
					range->PutItem(i, (*itor).newID, nowid);
					nowid++;
				}
				else{
					//�Ѿ�����������ˡ���ȡ���ֵ�ID����
					range->PutItem(i, (*itor).newID, (*map_itor).second);
				}
			}
		}
	}
	book->Close(need_save);
}
//����������Ա�
void CTableLanguage::saveBaseLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	int row_count = 0;
	//�ļ������ڣ��򴴽�
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
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "����ID");
	//д�������ͷ4λ��ע�ͱ��У�
	range->PutItem(1, 2, "table_name");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "�������ڱ���");
	//д�������ͷ4λ
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 4, "target");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "����λ��");
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
//����һ��Table�����ԣ��ֿ��������ͻ���
void CTableLanguage::saveOneLanguage(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, _bstr_t tar, string strFileName)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	bool need_save = false;
	//�ļ������ڣ��򴴽�
	if (!GetSheet(strFileName, getStrSourcePath(), excelApp, sheet))
	{
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->Name = strFileName.c_str();
		sheet->SaveAs((getStrSourcePath() + strFileName).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
		//д���һ��ͷ4λ
		range->PutItem(1, 1, "id");
		range->PutItem(2, 1, "int");
		range->PutItem(3, 1, tar);
		range->PutItem(4, 1, "����ID");
		//д�������ͷ4λ��ע�ͱ��У�
		range->PutItem(1, 2, "language");
		range->PutItem(2, 2, "string");
		range->PutItem(3, 2, tar);
		range->PutItem(4, 2, "��������");
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
	//�����ǰ����С��֮ǰ��������ʾɾ��������
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
//��ȡ���Ļ�����
bool CTableLanguage::ReadBaseLanguageFile(Excel::_ApplicationPtr excelApp, map<int, StructLanguageData> &mapLang, string strPath, string strFileName, vector<string> &vecFiles)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	//�ļ������ڣ��Ͳ���ִ��������
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
		//�������ȫ���ļ��������ڸ����б���û���ҵ������Ļ�����ʹ�þ�����
		if (all_file || !FindValue(vecFiles, string(_bstr_t(range->GetItem(i, 2)))))
		{
			StructLanguageData data;
			data.strTableName = range->GetItem(i, table_name_index);
			data.strData = range->GetItem(i, language_index);
			data.strTarget = range->GetItem(i, target_index);
			//��ID��Ӧ�����ķ���map
			mapLang.insert(map<int, StructLanguageData>::value_type(range->GetItem(i, id_index), data));
		}
	}
	book = excelApp->GetActiveWorkbook();
	book->Close(false);
	return true;
}
//��ȡ���Ļ������и����������ݷ����Ķ��󣨿ͻ��˻��Ƿ���ˣ�
map<_bstr_t, _bstr_t> CTableLanguage::GetBaseTargetMap(Excel::_ApplicationPtr excelApp)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	map<_bstr_t, _bstr_t> map_target;
	//�ļ������ڣ��Ͳ���ִ��������
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
		//����һ�������map
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
	//����оɵ��ļ������Ҳ��Ǵ���ȫ����Ҳ��ȡ�ɵ��ļ����ϲ�
	if (!create_all)
	{
		ReadBaseLanguageFile(excelApp, map_lang, getLanguagePath(), getStrBaseFileName(), vec_file);
	}
	//��ȡlang_change������
	GetSheet("lang_config", getStrSourcePath(), excelApp, sheet);
	//����������Ҫ�޸ĵ��ֶα�������
	readChange(sheet, lstChange, vec_file);
	//������Ҫ����ı�
	list<StructChange>::iterator itor = lstChange.begin();
	while (itor != lstChange.end())
	{
		readOneExcel(excelApp, map_lang, *itor);
		itor++;
	}
	saveBaseLanguage(excelApp, map_lang);
	//����ͻ����ļ�
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
	//�����Ϊ�յĻ�����ʾ���·����ģ�������Ҫ�ӻ������Ա��еõ���������
	if (str_end_name != "")
	{
		//�õ��������Ա��и��仰�ķ�������
		map<_bstr_t, _bstr_t> map_target = GetBaseTargetMap(excelApp);
		//��ȡ������Դ���ļ�
		if (!GetSheet(str_base_file, getLanguagePath(), excelApp, sheet))
		{
			std::cout << "Դ�����ļ������ڣ�" << endl;
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
				//����������Ǽ���ͷ��壬������ı��������������һ���Ļ�����ʾû���룬û����Ͳ���ʾ�ˣ���Ϸ���Զ���ʾ������ID
				if (!(strPlatform != "ch" && strPlatform != "tw" && data.strTableName == _bstr_t("tb_table_errcode") && range->GetItem(i, language_index) == range->GetItem(i, desc_index)))
				{
					//��ID��Ӧ�����ķ���map
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
		//����ͻ����ļ�
		saveOneLanguage(excelApp, map_lang, "client", strClientName + str_end_name);
		saveOneLanguage(excelApp, map_lang, "server", strServerName + str_end_name);
	}
	else{
		std::cout << "Դ�����ļ������ڣ�" << endl;
	}
}


//---------------------------------------------------------------��ȡ�����ļ����begin--------------------------------------------------------//
///�Ƚ�һ����ı��
void compareOneFile(Excel::_WorksheetPtr &sheet, list<_bstr_t> &lstComData, StructChange &stucChange, list<StructLanguageData> &lstChange)
{
	Excel::RangePtr range = sheet->Cells;
	int row_count = sheet->GetUsedRange()->Rows->Count;
	int column_count = sheet->GetUsedRange()->Columns->Count;
	list<int> lst_id;
	map<_bstr_t, StructID> map_have_column;
	map<_bstr_t, StructID>::iterator it;
	//�õ������ĵ��ֶ�ID
	for (int i = 1; i <= column_count; i++)
	{
		//ÿһ�е�һ��ֵ��Ϊ�գ���д��ȥ
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
	//�������е����ģ����������ԭ�������Ѿ��У����������û�������������
	for (int i = 5; i <= row_count; i++)
	{
		for (list<int>::iterator itor = lst_id.begin(); itor != lst_id.end(); itor++)
		{
			//��Ϊ�ղ�д��ȥ
			str_data = _bstr_t(range->GetItem(i, *itor));
			com_itor = std::find(lstComData.begin(), lstComData.end(), str_data);
			//ԭ���ı��Ҳ�����Ҫ�����ݣ���д������б�
			if (com_itor == lstComData.end())
			{
				//�����ʾ���ݲ�Ϊ�գ��򽫷��������б�����
				if (str_data != _bstr_t(""))
				{
					map_itor = map_data_to_key.find(str_data);
					//�ж�����������Ƿ��Ѿ�������һ��������
					if (map_itor == map_data_to_key.end())
					{
						//û��������֣���Ҫ����
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
///�õ���������ļ�
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
	//���û��ԭ���ķ�������ܱȽ�
	if (!GetSheet(getStrBaseFileName(), getLanguagePath(), excelApp, sheet))
	{
		std::cout << "��û��away_language�����ܱȽϣ�" << endl;
		excelApp->Workbooks->Close();
		return;
	}
	//��ȡԴ�������
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
		stuc.language = range->GetItem(i, 3);//Դ�����desc��languageһ��
		lst_source_data.push_back(stuc);
	}
	map<_bstr_t, StructTableData> map_now_trans;
	//��ȡ�����Ա������
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
			//����������+�����ڱ�������������ΪKEYֵ
			map_now_trans.insert(map<_bstr_t, StructTableData>::value_type(connectKey(stuc.desc, stuc.table_name, strConnect), stuc));
		}
		range->Clear();
	}
	else{
		//�����±����ұ���
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getLanguagePath() + str_table_file).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
	}
	sheet->Name = str_table_file.c_str();
	//����ͷ��4��
	setFileHead(range);
	int now_row = 5;
	list<StructTableData> lst_change;
	_bstr_t key;
	//�����µ�Դ���õ�����
	for (list<StructTableData>::iterator itor = lst_source_data.begin(); itor != lst_source_data.end(); itor++)
	{
		//���»�����Ϣ
		range->PutItem(now_row, 1, (*itor).id);
		range->PutItem(now_row, 2, (*itor).table_name);
		range->PutItem(now_row, 3, (*itor).desc);
		key = connectKey((*itor).desc, (*itor).table_name, strConnect);
		//�����map�������ݣ���û�и���
		if (map_now_trans.find(key) != map_now_trans.end() && map_now_trans[key].language != _bstr_t("") && map_now_trans[key].state == strHaveTrans)
		{
			//���ɷ�������д��
			range->PutItem(now_row, 4, map_now_trans[connectKey((*itor).desc, (*itor).table_name, strConnect)].language);
			range->PutItem(now_row, 5, 1);
		}
		else{
			//������д��
			range->PutItem(now_row, 4, (*itor).desc);
			range->PutItem(now_row, 5, "");
			//��������б�
			lst_change.push_back((*itor));
		}
		now_row++;
	}
	excelApp->ActiveWorkbook->Save();
	excelApp->Workbooks->Close();
	//�õ������ļ�
	if (!GetSheet(str_save_name, getTransPath(), excelApp, sheet))
	{
		//�����±����ұ���
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getTransPath() + str_save_name).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	}
	else{
		//���ԭ���ı�
		sheet->Cells->Clear();
	}
	sheet->Name = str_save_name.c_str();
	range = sheet->Cells;
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "table_name");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "�����ڵı���");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "desc");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "��������");
	range->PutColumnWidth(80);
	now_row = 5;
	for (list<StructTableData>::iterator itor = lst_change.begin(); itor != lst_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).table_name);
		range->PutItem(now_row, 2, (*itor).desc);
		now_row++;
	}
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
	excelApp->Workbooks->Close();
}
//---------------------------------------------------------------��ȡ�����ļ����end--------------------------------------------------------//
void CTableLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	list<StructIDToLanguage> lst_data;
	map<_bstr_t, _bstr_t> map_trans_data;
	_bstr_t key;
	//��ȡ��������
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
		cout << "�����ļ������ڲ������ļ�:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//��ȡ�����Ե�����
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
				//����]�з��g���O�������Mȥ
				range->PutItem(i, 4, (_bstr_t)range->GetItem(i, 3));
			}
		}
	}
	else{
		cout << "�������ļ�:" << str_table_file << endl;
		return;
	}
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
	excelApp->Workbooks->Close();
}
string CTableLanguage::getLanguagePath()
{
	return getStrSourcePath() + "language\\";
}