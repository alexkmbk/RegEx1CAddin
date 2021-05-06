#pragma once

//#include <sstream>

//std::string escape_json(const std::string &s);

size_t append_escaped_json_raw(char16_t* out, const char16_t* in, size_t begin, size_t end);
void append_escaped_json(std::basic_string<char16_t> &out, const char16_t* in, size_t begin, size_t end);


