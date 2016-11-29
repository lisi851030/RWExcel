#include "LabelLanguage.h"
#include  <direct.h>  
#include <fstream>
#include <list>
#include <map>


struct StructLabelData
{
	_bstr_t key_str;
	_bstr_t language;
	_bstr_t ui_name;
	_bstr_t desc;
	_bstr_t state;
};
CLabelLanguage::CLabelLanguage()
{
}


CLabelLanguage::~CLabelLanguage()
{
}
void CLabelLanguage::setFileHead(Excel::RangePtr range)
{
	//����4��ͷ��
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "uiName");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "��������UI");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "key_str");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "client");
	range->PutItem(4, 2, "keyֵ");
	//д�������ͷ4λ
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
void CLabelLanguage::GetUpdateFile(Excel::_ApplicationPtr excelApp, string strLangType)
{
	Excel::_WorksheetPtr sheet;
	Excel::_WorkbookPtr book;
	Excel::RangePtr range;
	map<_bstr_t, StructLabelData> map_old_data;
	list<StructLabelData> lst_new_data;
	//��ȡ�±������
	if (GetSheet(getStrBaseFileName(), getStrSourcePath(), excelApp, sheet))
	{
		range = sheet->Cells;
		int row_count = sheet->GetUsedRange()->Rows->Count;
		for (int i = 5; i <= row_count; i++)
		{
			if (_bstr_t(range->GetItem(i, 3)) != _bstr_t(""))
			{
				StructLabelData stuc;
				stuc.desc = range->GetItem(i, 3);
				stuc.language = stuc.desc;
				stuc.key_str = range->GetItem(i, 2);
				stuc.ui_name = range->GetItem(i, 1);
				lst_new_data.push_back(stuc);
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
			StructLabelData stuc;
			stuc.ui_name = range->GetItem(i, 1);
			stuc.key_str = range->GetItem(i, 2);
			stuc.desc = range->GetItem(i, 3);
			stuc.language = range->GetItem(i, 4);
			stuc.state = range->GetItem(i, 5);
			map_old_data.insert(map<_bstr_t, StructLabelData>::value_type(stuc.ui_name + strConnect + stuc.key_str, stuc));
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
	list<StructLabelData>::iterator itor = lst_new_data.begin();
	list<StructLabelData> lst_change;
	int now_row = 5;
	StructLabelData old_data;
	while (itor != lst_new_data.end())
	{
		range->PutItem(now_row, 1, (*itor).ui_name);
		range->PutItem(now_row, 2, (*itor).key_str);
		range->PutItem(now_row, 3, (*itor).desc);
		//����ɵķ���Ϊ�գ���д������������Ĭ��д��ɵķ���
		if (old_data.language == _bstr_t(""))
		{
			range->PutItem(now_row, 4, (*itor).desc);
		}
		else{
			range->PutItem(now_row, 4, old_data.language);
		}
		old_data = map_old_data[(*itor).ui_name + strConnect + (*itor).key_str];
		//�ж�����ֵ�Ƿ���ͬ����ͬ���ô�����ͬ��������б�
		if (old_data.language == _bstr_t("") || old_data.desc != (*itor).desc || old_data.state != strHaveTrans)
		{
			lst_change.push_back((*itor));
			range->PutItem(now_row, 5, "");
		}
		else {
			//������ݻ���ͬ����ԭ��������ķ���д��ȥ
			range->PutItem(now_row, 4, old_data.language);
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
	range->PutItem(1, 1, "uiName");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "none");
	range->PutItem(4, 1, "���ֵ�ID");
	//д���һ��ͷ4λ
	range->PutItem(1, 2, "key_str");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "none");
	range->PutItem(4, 2, "���ֵ�ID");
	//д��ڶ���ͷ4λ
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
	for (list<StructLabelData>::iterator itor = lst_change.begin(); itor != lst_change.end(); itor++)
	{
		range->PutItem(now_row, 1, (*itor).ui_name);
		range->PutItem(now_row, 2, (*itor).key_str);
		range->PutItem(now_row, 3, (*itor).desc);
		now_row++;
	}
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
	excelApp->Workbooks->Close();
}
void CLabelLanguage::MergeLanguageFile(Excel::_ApplicationPtr excelApp, string strLangType)
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
			key = _bstr_t(range->GetItem(i, 1)) + strConnect + _bstr_t(range->GetItem(i, 2)) + strConnect + _bstr_t(range->GetItem(i, 3));
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
			key = _bstr_t(range->GetItem(i, 1)) + strConnect + _bstr_t(range->GetItem(i, 2)) + strConnect + _bstr_t(range->GetItem(i, 3));
			if (map_trans_data[key] != _bstr_t(""))
			{
				range->PutItem(i, 4, map_trans_data[key]);
				range->PutItem(i, 5, strHaveTrans);
			}
			else if (_bstr_t(range->GetItem(i, 4)) == _bstr_t("")){
				//����]�з��g���O�������Mȥ ���������ַ�
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
void CLabelLanguage::StringChangeExcel(Excel::_ApplicationPtr excelApp)
{
	ifstream in;
	in.open("static_language.lua");
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
			if (vec_one_data.size() == 3)
			{
				lst_data.push_back(vec_one_data);
			}
		}
	}

	Excel::_WorksheetPtr sheet;

	//�������ļ�
	Excel::_WorkbookPtr book = excelApp->Workbooks->Add();
	sheet = book->Sheets->Add();
	sheet->Name = L"tb_table_labellang";
	Excel::RangePtr range = sheet->Cells;
	//д���һ��ͷ4λ
	range->PutItem(1, 1, "uiName");
	range->PutItem(2, 1, "string");
	range->PutItem(3, 1, "client");
	range->PutItem(4, 1, "��������UI");
	//д��ڶ���ͷ4λ
	range->PutItem(1, 2, "key_str");
	range->PutItem(2, 2, "string");
	range->PutItem(3, 2, "client");
	range->PutItem(4, 2, "���ĵ�keyֵ");
	//д�������ͷ4λ
	range->PutItem(1, 3, "language");
	range->PutItem(2, 3, "string");
	range->PutItem(3, 3, "client");
	range->PutItem(4, 3, "��������");
	range->PutColumnWidth(80);
	int now_row = 5;
	for (list<vector<string>>::iterator itor = lst_data.begin(); itor != lst_data.end(); itor++)
	{
		if (string((*itor)[1].c_str()) != "")
		{
			range->PutItem(now_row, 1, (*itor)[2].c_str());
			range->PutItem(now_row, 2, (*itor)[0].c_str());
			range->PutItem(now_row, 3, (*itor)[1].c_str());
			now_row++;
		}
	}
	sheet->SaveAs((getStrSourcePath() + "tb_table_labellang").c_str(), vtMissing, vtMissing, vtMissing, vtMissing, true);
	//�����ļ�
	excelApp->ActiveWorkbook->Save();
	//�ر��ļ�
	excelApp->Workbooks->Close();
}
