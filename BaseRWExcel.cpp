#include "BaseRWExcel.h"


BaseRWExcel::BaseRWExcel()
{
}


BaseRWExcel::~BaseRWExcel()
{
}

bool BaseRWExcel::GetSheet(const string strFileName, const string strPath, const Excel::_ApplicationPtr excelApp, Excel::_WorksheetPtr &sheet)
{
	Excel::_WorkbookPtr pWorkbook;
	string name = strPath + strFileName;
	try
	{
		/*��һ��������*/
		pWorkbook = excelApp->Workbooks->Open(_bstr_t(name.c_str()));  // open excel file
	}
	catch (...)
	{
		std::cout << strFileName << "�ļ������ڣ�" << endl;
		return false;
	}
	//�õ�WorkSheets
	Excel::SheetsPtr sheets = pWorkbook->GetSheets();
	//���ļ�����Ӧ��Sheet
	try
	{
		sheet = pWorkbook->GetActiveSheet();
	}
	catch (...)
	{
		std::cout << strFileName << " Sheet�����ڣ�" << endl;
		return false;
	}
	return true;
}
bool BaseRWExcel::FindValue(vector<string> &vecFiles, string &strFile)
{
	vector<string>::iterator itor = vecFiles.begin();
	while (itor != vecFiles.end())
	{
		if ((*itor) == strFile)
		{
			return true;
		}
		itor++;
	}
	return false;
}
std::vector<std::string> BaseRWExcel::Split(const  std::string& s, const std::string& delim)
{
	std::vector<std::string> elems;
	size_t pos = 0;
	size_t len = s.length();
	size_t delim_len = delim.length();
	if (delim_len == 0) return elems;
	while (pos < len)
	{
		int find_pos = s.find(delim, pos);
		if (find_pos < 0)
		{
			elems.push_back(s.substr(pos, len - pos));
			break;
		}
		elems.push_back(s.substr(pos, find_pos - pos));
		pos = find_pos + delim_len;
	}
	return elems;
}
string BaseRWExcel::GBKToUTF8(const std::string& strGBK)
{
	string strOutUTF8 = "";
	WCHAR * str1;
	int n = MultiByteToWideChar(CP_ACP, 0, strGBK.c_str(), -1, NULL, 0);
	str1 = new WCHAR[n];
	MultiByteToWideChar(CP_ACP, 0, strGBK.c_str(), -1, str1, n);
	n = WideCharToMultiByte(CP_UTF8, 0, str1, -1, NULL, 0, NULL, NULL);
	char * str2 = new char[n];
	WideCharToMultiByte(CP_UTF8, 0, str1, -1, str2, n, NULL, NULL);
	strOutUTF8 = str2;
	delete[]str1;
	str1 = NULL;
	delete[]str2;
	str2 = NULL;
	return strOutUTF8;
}
int BaseRWExcel::GetColumnID(Excel::_WorksheetPtr &sheet, _bstr_t name)
{
	Excel::RangePtr range = sheet->Cells;
	int column_count = sheet->GetUsedRange()->Columns->Count;
	for (int i = 1; i <= column_count; i++)
	{
		if ((_bstr_t)range->GetItem(1, i) == name)
		{
			return i;
		}
	}
	return 0;
}
_bstr_t BaseRWExcel::getTranslateData(_variant_t oldData)
{
	string str_data = (LPCSTR)_bstr_t(oldData);
	int len = str_data.length();
	for (int i = 0; i < len; i++)
	{
		//��ʱ�����жϣ���������ж����Ҫ�滻�ģ���������map���ж�
		if (str_data[i] == '��' || str_data[i] == '��')
		{
			str_data[i] = '"';
		}
	}
	return str_data.c_str();
}