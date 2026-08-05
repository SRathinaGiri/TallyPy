#include "BannerDialog.h"
#include "MainWindow.h"

#include <QApplication>
#include <QIcon>

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);
    app.setApplicationName("Tally Qt Exporter");
    app.setApplicationVersion(APP_VERSION);
    app.setOrganizationName("TallyXML");
    app.setWindowIcon(QIcon(":/assets/app_icon.ico"));

    BannerDialog banner;
    banner.setWindowIcon(QIcon(":/assets/app_icon.ico"));
    banner.exec();

    MainWindow window;
    window.setWindowIcon(QIcon(":/assets/app_icon.ico"));
    window.show();
    return app.exec();
}
