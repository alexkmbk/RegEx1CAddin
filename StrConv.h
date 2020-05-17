#pragma once

#include <string>

#if defined( __linux__ ) || defined(__APPLE__) || defined(__ANDROID__)
void convertUTF16ToUTF32(char16_t *input, const size_t input_size, std::basic_string<wchar_t> &output);
unsigned int convertUTF32ToUTF16(const wchar_t *input, size_t input_size, char16_t *output);
#endif

// Приведение к нижнему регистру только для Кириллицы и Латиницы
//
void tolowerStr(std::basic_string<char16_t> & s);

