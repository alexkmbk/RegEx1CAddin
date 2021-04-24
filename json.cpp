/*#include <sstream>
#include <iomanip>

std::string escape_json(const std::string &s) {
	std::ostringstream o;
	for (auto c = s.cbegin(); c != s.cend(); c++) {
		switch (*c) {
		case '"': o << "\\\""; break;
		case '\\': o << "\\\\"; break;
		case '\b': o << "\\b"; break;
		case '\f': o << "\\f"; break;
		case '\n': o << "\\n"; break;
		case '\r': o << "\\r"; break;
		case '\t': o << "\\t"; break;
		default:
			if ('\x00' <= *c && *c <= '\x1f') {
				o << "\\u"
					<< std::hex << std::setw(4) << std::setfill('0') << (int)*c;
			}
			else {
				o << *c;
			}
		}
	}
	return o.str();
}

/*std::wstring escape_json(const std::wstring &s) {
	std::wostringstream o;
	for (auto c = s.cbegin(); c != s.cend(); c++) {
		switch (*c) {
		case L'"': o << L"\\\""; break;
		case L'\\': o << L"\\\\"; break;
		case L'\b': o << L"\\b"; break;
		case L'\f': o << L"\\f"; break;
		case L'\n': o << L"\\n"; break;
		case L'\r': o << L"\\r"; break;
		case L'\t': o << L"\\t"; break;
		default:
			if (L'\x00' <= *c && *c <= L'\x1f') {
				o << L"\\u"
					<< std::hex << std::setw(4) << std::setfill(L'0') << (int)*c;
			}
			else {
				o << *c;
			}
		}
	}
	return o.str();
}*/