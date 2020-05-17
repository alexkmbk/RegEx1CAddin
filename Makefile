ifeq ($(x32),1)
CPU = 32
else
CPU = 64
endif

SOURCES=AddInNative.cpp StrConv.cpp 

ifeq ($(macos),1)
TARGETDIR = "$(CURDIR)/macos/"
TARGET = $(TARGETDIR)RegExMac64.so
LIBPATHS = -Llib/macos/
LIBS= boost_regex
else
TARGETDIR = $(CURDIR)/linux/x$(CPU)/
TARGET = $(TARGETDIR)RegEx$(CPU).so
LIBPATHS = -Llib/linux/x$(CPU)/
LIBS= boost_regex
endif

OBJECTS=$(SOURCES:.cpp=.o)
INCLUDES=-Iinclude
CXXLAGS=$(CXXFLAGS) $(INCLUDES) -m$(CPU) -finput-charset=UTF-8 -fPIC -std=c++14 -O1

all: $(TARGET)

-include $(OBJECTS:.o=.d)

%.o: %.cpp
	g++ -c  $(CXXLAGS) $*.cpp -o $*.o
	g++ -MM $(CXXLAGS) $*.cpp >  $*.d
	@mv -f $*.d $*.d.tmp
	@sed -e 's|.*:|$*.o:|' < $*.d.tmp > $*.d
	@sed -e 's/.*://' -e 's/\\$$//' < $*.d.tmp | fmt -1 | \
	  sed -e 's/^ *//' -e 's/$$/:/' >> $*.d
	@rm -f $*.d.tmp

$(TARGET): $(OBJECTS) Makefile
	g++ $(CXXLAGS) $(LIBPATHS) -shared $(OBJECTS) -o $(TARGET) $(addprefix -l, $(LIBS))

clean:
	rm -f $(CURDIR)/*.d $(CURDIR)/*.o 