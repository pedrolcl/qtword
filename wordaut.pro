QT += widgets axcontainer
CONFIG += c++11 cmdline
DEFINES += QT_DEPRECATED_WARNINGS

DUMPCPP=$$absolute_path("dumpcpp.exe", $$dirname(QMAKE_QMAKE))
TYPELIBS = $$system($$DUMPCPP -getfile {00020905-0000-0000-C000-000000000046})

isEmpty(TYPELIBS) {
    message("Microsoft Word type library not found!")
    REQUIRES += StackOverflow Sucks
} else {
    SOURCES  = main.cpp
}

target.path = C:/tmp/$${TARGET}
!isEmpty(target.path): INSTALLS += target
