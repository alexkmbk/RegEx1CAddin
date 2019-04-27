# Cmake script for project AddInNative
# autor: 1C

CMAKE_MINIMUM_REQUIRED(VERSION 2.6 FATAL_ERROR)

message("CMAKE_SIZEOF_VOID_P:" CMAKE_SIZEOF_VOID_P)

SET(CPU "32" CACHE STRING "Some user-specified option")

#if(CMAKE_SIZEOF_VOID_P EQUAL 8)
#    set(CPU "64")
#else()
#    set(CPU "32")
#endif()

message("CPU:" ${CPU} " SYSTEM: " ${CMAKE_GENERATOR})

PROJECT(AddInNative)

SET(AddInNative_SRC 
    AddInNative.cpp
    AddInNative.h
    dllmain.cpp
    stdafx.cpp
    stdafx.h
    )

include_directories(${CMAKE_SOURCE_DIR}/include)

SET (CMAKE_PREFIX_PATH ${CMAKE_PREFIX_PATH} ${CMAKE_SOURCE_DIR})

add_library(${PROJECT_NAME} SHARED ${AddInNative_SRC} ${AddInDef_SRC})

if (UNIX)
#    set_source_files_properties(${AddInNative_SRC}
#                              PROPERTIES COMPILE_FLAGS
#                              "-fPIC ")
set (CMAKE_CXX_FLAGS " -std=c++11 -O1 -fPIC ")
							  
target_link_libraries(${PROJECT_NAME} -L${CMAKE_SOURCE_DIR}/lib/linux/x${CPU}/)
target_link_libraries(${PROJECT_NAME} -lboost_regex)
							  
else()
    target_link_libraries(${PROJECT_NAME} -L${CMAKE_SOURCE_DIR}/lib/win/x${CPU}/)
	
	if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    target_link_libraries(${PROJECT_NAME} -lboost_regex-vc141-mt-s-x64-1_69)
    elseif(CMAKE_SIZEOF_VOID_P EQUAL 4)
    target_link_libraries(${PROJECT_NAME} -lboost_regex-vc141-mt-s-x32-1_69)
    endif()

    SET(AddInDef_SRC AddInNative.def)
    SET(CMAKE_LINK_DEF_FILE_FLAG AddInNative.def)
    add_definitions(-DUNICODE -DWIN32 )
endif()

set_target_properties( ${PROJECT_NAME} PROPERTIES
   CLEAN_DIRECT_OUTPUT 1
   # avoid conflicts between library and binary target names
   OUTPUT_NAME ${PROJECT_NAME}${MySuffix1}${MySuffix2} 
   #WINDOWS_EXPORT_ALL_SYMBOLS TRUE
   )

