#include "StrConv.h"

#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)

inline int is_high_surrogate(char16_t uc) { return (uc & 0xfffffc00) == 0xd800; }
inline int is_low_surrogate(char16_t uc) { return (uc & 0xfffffc00) == 0xdc00; }

inline char32_t surrogate_to_utf32(char16_t high, char16_t low) {
	return (high << 10) + low - 0x35fdc00;
}

// The algorithm is based on this answer:
//https://stackoverflow.com/a/23920015/2134488
//
void convertUTF16ToUTF32(char16_t *input,
	const size_t input_size,
	std::basic_string<wchar_t> &output)
{
	int i = 0;
	const char16_t * const end = input + input_size;
	while (input < end) {
		const char16_t uc = *input++;
		if (!((uc - 0xd800u) < 2048u)) {
			output[i++] = uc;
		}
		else {
			if (is_high_surrogate(uc) && input < end && is_low_surrogate(*input)) {
				output[i++] = (wchar_t)surrogate_to_utf32(uc, *input++);
			}
			else {
				output[i++] = 0; //ERROR
			}
		}
	}
}

// The algorithm is based on this answer:
//https://stackoverflow.com/questions/955484/is-it-possible-to-convert-utf32-text-to-utf16-using-only-windows-api
//
unsigned int convertUTF32ToUTF16(const wchar_t *input, size_t input_size, char16_t *output)
{
	char16_t *start = output;
	const wchar_t * const end = input + input_size;
	while (input < end) {
		const wchar_t cUTF32 = *input++;
		if (cUTF32 < 0x10000)
		{
			*output++ = cUTF32;
		}
		else {
			unsigned int t = cUTF32 - 0x10000;
			wchar_t h = (((t << 12) >> 22) + 0xD800);
			wchar_t l = (((t << 22) >> 22) + 0xDC00);
			*output++ = h;
			*output++ = (l & 0x0000FFFF);
		}
	}
	return (output - start) * 2; // size in bytes
}

#endif

inline void tolowerPtr(char16_t*);

// Приведение к нижнему регистру только для Кириллицы и Латиницы
//
void tolowerStr(std::basic_string<char16_t> & s)
{
	char16_t* c = const_cast<char16_t*>(s.c_str());
	size_t l = s.size();
	for (char16_t* c2 = c; c2 < c + l; c2++)tolowerPtr(c2);
};

inline void tolowerPtr(char16_t *p)
{
	if (((*p >= u'А') && (*p <= u'Я')) || ((*p >= u'A') && (*p <= u'Z')))
		*p = *p + 32;

	else if (*p == u'Ё')
		*p = u'ё';
}

