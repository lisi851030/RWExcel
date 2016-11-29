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
	//����4��ͷ��
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "int");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "id");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "desc");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "client");
	range->PutItem(4, 3, "��������");
	//д�������ͷ4λ
	range->PutItem(1, 4, "state");
	range->PutItem(2, 4, "int");
	range->PutItem(3, 4, "none");
	range->PutItem(4, 4, "����״̬");
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
	//��ȡ�±������
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
	//�����������͵õ��ɵ�����
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
		//�����±����ұ���
		book = excelApp->Workbooks->Add();
		sheet = book->Sheets->Add();
		sheet->SaveAs((getStrSourcePath() + str_table_file).c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
		range = sheet->Cells;
	}
	sheet->Name = str_table_file.c_str();
	//����ͷ��4��
	setFileHead(range);
	map<_bstr_t, _bstr_t>::iterator itor = vec_new_data.begin();
	map<_bstr_t, _bstr_t> vec_change;
	int now_row = 5;
	while (itor != vec_new_data.end())
	{
		range->PutItem(now_row, 1, (*itor).first);
		range->PutItem(now_row, 2, (*itor).second);
		range->PutItem(now_row, 3, vec_trans_data[(*itor).first]);
		//�ж�����ֵ�Ƿ���ͬ����ͬ���ô�����ͬ��������б�
		if (vec_old_data[(*itor).first] != (*itor).second || vec_trans_state[(*itor).first] != strHaveTrans)
		{
			vec_change.insert(map<_bstr_t, _bstr_t>::value_type((*itor).first, (*itor).second));
			range->PutItem(now_row, 4, "");
		}
		else {
			//������ݻ���ͬ����ԭ��������ķ���д��ȥ
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
	range->PutItem(1, 1, "id");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "���ֵ�ID");
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
	for (map<_bstr_t, _bstr_t>::iterator itor = vec_change.begin(); itor != vec_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).first);
		range->PutItem(now_row, 2, (*itor).second);
		now_row++;
	}
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
	excelApp->Workbooks->Close();
}
void CClientLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	list<StructIDToLanguage> lst_data;
	map<_bstr_t, _bstr_t> map_trans_data;
	//��ȡ��������
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
		cout << "�����ļ������ڲ������ļ�:" << str_save_name << endl;
		return;
	}
	string str_table_file = getStrBaseFileName() + "_" + strLangType;
	_bstr_t key;
	//��ȡ�����Ե�����
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
				//�������з���״̬
				range->PutItem(i, 4, 1);
			}
			else if (_bstr_t(range->GetItem(i, 3)) == _bstr_t("")){
				//����]�з��g���O�������Mȥ
				range->PutItem(i, 3, (_bstr_t)range->GetItem(i, 2));
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