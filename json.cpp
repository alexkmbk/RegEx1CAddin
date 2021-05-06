#include <cstring>
#include <string>

void escape(char16_t x, char16_t* res) {
#define TO_HEX(i) (i <= 9 ? '0' + i : 'A' - 10 + i)

	if (x <= 0xFFFF)
	{
		res[2] = TO_HEX(((x & 0xF000) >> 12));
		res[3] = TO_HEX(((x & 0x0F00) >> 8));
		res[4] = TO_HEX(((x & 0x00F0) >> 4));
		res[5] = TO_HEX((x & 0x000F));
		//res[6] = '\0';
	}

}

void append_escaped_json(std::basic_string<char16_t> &out, const char16_t* in, size_t begin, size_t end) {
	for (size_t i = begin; i <= end; i++) {
		char16_t c = in[i];
/*		if (i == begin && (c == u'\n' || c == u'\r'))
			continue;*/
		switch (c) {
		case u'"': out.append(u"\\\"", 2); break;
		case u'\\': out.append(u"\\\\", (sizeof(u"\\\\") / sizeof(char16_t)) - 1); break;
		case u'\b': out.append(u"\\b", (sizeof(u"\\b") / sizeof(char16_t)) - 1); break;
		case u'\f': out.append(u"\\f", (sizeof(u"\\f") / sizeof(char16_t)) - 1); break;
		case u'\n': out.append(u"\\n", (sizeof(u"\\n") / sizeof(char16_t)) - 1); break;
		case u'\r': out.append(u"\\r", (sizeof(u"\\r") / sizeof(char16_t)) - 1); break;
		case u'\t': out.append(u"\\t", (sizeof(u"\\t") / sizeof(char16_t)) - 1); break;
		default:
			if (u'\x00' <= c && c <= u'\x1f') {
				char16_t escaped[7] = u"\\u0000";
				escape(c, escaped);
				out.append(escaped, 6);
			}
			else {
				out +=c;
			}
		}
	}
}

size_t append_escaped_json_raw(char16_t* out, const char16_t* in, size_t begin, size_t end) {
	size_t offset = 0;
	for (size_t i = begin; i <= end; i++) {
		char16_t c = in[i];
		switch (c) {
		case u'"': memcpy(out + offset, u"\\\"", 4); offset += 2;  break;
		case u'\\': memcpy(out + offset, u"\\\\", 4); offset += 2;  break;
		case u'\b': memcpy(out + offset, u"\\b", 4); offset += 2;  break;
		case u'\f': memcpy(out + offset, u"\\f", 4); offset += 2;  break;
		case u'\n': memcpy(out + offset, u"\\n", 4); offset += 2;  break;
		case u'\r': memcpy(out + offset, u"\\r", 4); offset += 2;  break;
		case u'\t': memcpy(out + offset, u"\\t", 4); offset += 2;  break;
		default:
			if (u'\x00' <= c && c <= u'\x1f') {
				char16_t escaped[7] = u"\\u0000";
				escape(c, escaped);
				memcpy(out + offset, escaped, 6 * sizeof(char16_t));
				offset += 6;
			}
			else {
				out[offset++] = c;
			}
		}
	}
	return offset;
}

/*std::string escape_json(const std::string &s) {
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

std::wstring escape_json(const std::wstring &s) {
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