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
	//����4��ͷ��
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "lang_type");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "all");
	range->PutItem(4, 1, "��������");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "lang_id");
	range->PutItem(2, 2, "int");
	range->PutItem(3, 2, "all");
	range->PutItem(4, 2, "����ID");
	//д�������ͷ4λ
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 4, "lang_string");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "all");
	range->PutItem(4, 4, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 5, "state");
	range->PutItem(2, 5, "int");
	range->PutItem(3, 5, "none");
	range->PutItem(4, 5, "����״̬");
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
	//��ȡ�±������
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
				//��һ�����ֶ�����������ΪKEYֵ
				vec_new_data.insert(map<_bstr_t, StructServerData>::value_type(connectKey(_bstr_t(range->GetItem(i, 1)), _bstr_t(range->GetItem(i, 2)), _bstr_t(""), strConnect), stuc));
			}
		}
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//�����������͵õ��ɵ�����
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
		//�����±����ұ���
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getStrSourcePath() + str_table_file).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
	}
	sheet->Name = str_table_file.c_str();
	//����ͷ��4��
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
		//�ж�����ֵ�Ƿ���ͬ����ͬ���ô�����ͬ��������б�
		if (vec_old_data[(*itor).first].lang_string != (*itor).second.lang_string || vec_old_data[(*itor).first].lang_trans == _bstr_t("") || vec_old_data[(*itor).first].state != strHaveTrans)
		{
			lst_change.push_back(list<StructServerData>::value_type((*itor).second));
			range->PutItem(now_row, 5, "");
		}
		else {
			//������ݻ���ͬ����ԭ��������ķ���д��ȥ
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
	range->PutItem(1, 1, "lang_type");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "all");
	range->PutItem(4, 1, "��������");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "lang_id");
	range->PutItem(2, 2, "int");
	range->PutItem(3, 2, "all");
	range->PutItem(4, 2, "����ID");
	//д�������ͷ4λ
	range->PutItem(1, 3, "desc");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "none");
	range->PutItem(4, 3, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 4, "language");
	range->PutItem(2, 4, "string");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "��������");
	range->PutColumnWidth(80);
	now_row = 5;
	for (list<StructServerData>::iterator itor = lst_change.begin(); itor != lst_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).lang_type);
		range->PutItem(now_row, 2, (*itor).lang_id);
		range->PutItem(now_row, 3, (*itor).lang_string);
		now_row++;
	}
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
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
	//��ȡ��������
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
		cout << "�����ļ������ڲ������ļ�:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	//��ȡ�����Ե�����
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