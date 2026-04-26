#ifndef VERSION_H
#define VERSION_H

#define VERSION_MAJOR               1
#define VERSION_MINOR               0

#define STRINGIZE2(s) #s
#define STRINGIZE(s) STRINGIZE2(s)

// Отображаемая строка версии (только Major.Minor)
#define VERSION_STRING \
    STRINGIZE(VERSION_MAJOR) "." \
    STRINGIZE(VERSION_MINOR)

#define VERSION_WSTRING2(s) L##s
#define VERSION_WSTRING(s) VERSION_WSTRING2(s)

#define VERSION_STRING_W VERSION_WSTRING(VERSION_STRING)

// Для ресурсов Windows (VERSIONINFO) всегда нужно 4 числа
#define VERSION_COMMA \
    VERSION_MAJOR, VERSION_MINOR, 0, 0

#endif // VERSION_H
