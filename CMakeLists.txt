cmake_minimum_required(VERSION 3.10)
project(docx2latex)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(PkgConfig REQUIRED)
pkg_check_modules(LIBZIP REQUIRED libzip)
pkg_check_modules(TINYXML2 REQUIRED tinyxml2)

add_executable(docx2latex doc2latex.cpp)

target_include_directories(docx2latex PRIVATE
    ${LIBZIP_INCLUDE_DIRS}
    ${TINYXML2_INCLUDE_DIRS}
)

target_link_libraries(docx2latex
    ${LIBZIP_LIBRARIES}
    ${TINYXML2_LIBRARIES}
)
