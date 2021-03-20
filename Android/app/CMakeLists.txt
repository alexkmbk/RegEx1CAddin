# For more information about using CMake with Android Studio, read the
# documentation: https://d.android.com/studio/projects/add-native-code.html

# Sets the minimum version of CMake required to build the native library.

cmake_minimum_required(VERSION 3.4.1)

include_directories(${CMAKE_CURRENT_SOURCE_DIR}/src/include)
include_directories(${CMAKE_CURRENT_SOURCE_DIR}/src/main/cpp)
#include_directories(${PROJECT_SOURCE_DIR}../../../)

add_compile_options(-O1)
set(CMAKE_ANDROID_STL_TYPE "c++_static")

# Creates and names a library, sets it as either STATIC
# or SHARED, and provides the relative paths to its source code.
# You can define multiple libraries, and CMake builds them for you.
# Gradle automatically packages shared libraries with your APK.

add_library( # Sets the name of the library.
             AddInNative

             # Sets the library as a shared library.
             SHARED

             # Provides a relative path to your source file(s).
             src/main/cpp/jnienv.cpp ${PROJECT_SOURCE_DIR}../../../StrConv.cpp ${PROJECT_SOURCE_DIR}../../../AddInNative.cpp)

# Searches for a specified prebuilt library and stores the path as a
# variable. Because CMake includes system libraries in the search path by
# default, you only need to specify the name of the public NDK library
# you want to add. CMake verifies that the library exists before
# completing its build.

find_library( # Sets the name of the path variable.
              log-lib

              # Specifies the name of the NDK library that
              # you want CMake to locate.
              log )

#${CMAKE_CURRENT_SOURCE_DIR} - the path to the CMakeLists.txt file

target_link_libraries(AddInNative -L${CMAKE_CURRENT_SOURCE_DIR}/libs/android/)
#target_link_libraries(AddInNative -lboost_regex_armeabi-v7a)
target_link_libraries(AddInNative -lboost_regex_x86)

# Specifies libraries CMake should link to your target library. You
# can link multiple libraries, such as libraries you define in this
# build script, prebuilt third-party libraries, or system libraries.

target_link_libraries( # Specifies the target library.
                       AddInNative

                       # Links the target library to the log library
                       # included in the NDK.
                       ${log-lib} )