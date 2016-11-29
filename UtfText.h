#ifndef UTFTEXT_H
#define UTFTEXT_H
#include <iostream>
#include <string>
#include "Text.h"
using namespace std;
class UtfText :public Text
{
	public:
	    UtfText(string path);
		~UtfText(void);
		bool ReadOneChar(string & oneChar);
	private:
		size_t get_utf8_char_len(const char & byte);
};
#endif