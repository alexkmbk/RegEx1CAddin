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