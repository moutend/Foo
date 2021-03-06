cmake_minimum_required(VERSION 3.12)
project(Foo)

if(CMAKE_SOURCE_DIR STREQUAL CMAKE_BINARY_DIR OR EXISTS "${CMAKE_BINARY_DIR}/CMakeLists.txt")
  message(FATAL_ERROR "In-source builds are not allowed.")
endif()

if(NOT WIN32)
  message(FATAL "This project targets Windows")
endif()

# windows.winmd search path
set( CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -AI\"C:/Program Files (x86)/Windows Kits/10/UnionMetadata/10.0.17763.0\"" )
set( CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -AI\"C:/Program Files (x86)/Microsoft Visual Studio/2017/Community/Common7/IDE/VC/vcpackages\"" )

find_package(cpprestsdk CONFIG REQUIRED)

add_subdirectory(lib/IFoo)
add_subdirectory(lib/cpplogger)

set(CMAKE_CXX_STANDARD 17)
add_compile_options("/ZW")

file(GLOB SOURCES "./src/*")
add_library(Foo SHARED ${SOURCES})
target_link_libraries(Foo IFoo cpplogger cpprestsdk::cpprest cpprestsdk::cpprestsdk_zlib_internal cpprestsdk::cpprestsdk_boost_internal cpprestsdk::cpprestsdk_brotli_internal shlwapi.lib)
