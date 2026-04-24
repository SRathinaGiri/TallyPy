#include "MainWindow.h"

#include <QApplication>
#include <QIcon>

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);
    app.setApplicationName("Tally Qt Exporter");
    app.setApplicationVersion(APP_VERSION);
    app.setOrganizationName("TallyXML");
    app.setWindowIcon(QIcon(":/app_icon.ico"));
    MainWindow window;
    window.setWindowIcon(QIcon(":/app_icon.ico"));
    window.show();
    return app.exec();
}
