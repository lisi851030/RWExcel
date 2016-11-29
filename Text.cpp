#include "Text.h"
using namespace std;
Text::Text(string path) :m_index(0)
{
   filebuf *pbuf;
   ifstream filestr;
   // ���ö����ƴ� 
   filestr.open(path.c_str(), ios::binary);
   if (!filestr)
   {
     return;
   }
   // ��ȡfilestr��Ӧbuffer�����ָ�� 
   pbuf = filestr.rdbuf();
   // ����buffer���󷽷���ȡ�ļ���С
   m_length = (int)pbuf->pubseekoff(0, ios::end, ios::in);
   pbuf->pubseekpos(0, ios::in);
   // �����ڴ�ռ�
   m_binaryStr = new char[m_length + 1];
   // ��ȡ�ļ�����
   pbuf->sgetn(m_binaryStr, m_length);
   //�ر��ļ�
   filestr.close();
}
void Text::SetIndex(size_t index)
{
	m_index = index;
}
size_t Text::Size()
{
    return m_length;
}
Text::~Text()
{
    delete[] m_binaryStr;
}