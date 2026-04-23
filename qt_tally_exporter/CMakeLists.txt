cmake_minimum_required(VERSION 3.21)
project(TallyQtExporter LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTORCC ON)

find_package(Qt6 REQUIRED COMPONENTS Core Gui Widgets Network Xml)

qt_add_executable(TallyQtExporter
    src/main.cpp
    src/MainWindow.cpp
    src/TallyService.cpp
    resources/resources.qrc
    resources/app.rc
)

target_include_directories(TallyQtExporter PRIVATE src)

target_link_libraries(TallyQtExporter PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
    Qt6::Network
    Qt6::Xml
)

set_target_properties(TallyQtExporter PROPERTIES
    WIN32_EXECUTABLE TRUE
)

install(TARGETS TallyQtExporter RUNTIME DESTINATION .)
