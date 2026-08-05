#include "BannerDialog.h"

#include <QDialogButtonBox>
#include <QFont>
#include <QLabel>
#include <QPixmap>
#include <QPushButton>
#include <QVBoxLayout>

BannerDialog::BannerDialog(QWidget *parent)
    : QDialog(parent) {
    setWindowTitle("About Tally QT Exporter");
    setModal(true);
    setMinimumWidth(760);

    auto *layout = new QVBoxLayout(this);
    layout->setContentsMargins(18, 18, 18, 18);
    layout->setSpacing(10);

    auto *banner = new QLabel(this);
    banner->setAlignment(Qt::AlignCenter);
    banner->setPixmap(QPixmap(":/assets/tally_qt_banner.png").scaledToWidth(720, Qt::SmoothTransformation));

    auto *author = new QLabel("by S. Rathinagiri", this);
    QFont authorFont = author->font();
    authorFont.setPointSize(10);
    authorFont.setItalic(true);
    author->setFont(authorFont);
    author->setAlignment(Qt::AlignCenter);

    auto *link = new QLabel(
        "<a href=\"https://github.com/SRathinaGiri/TallyPy\">github.com/SRathinaGiri/TallyPy</a>",
        this);
    link->setTextFormat(Qt::RichText);
    link->setTextInteractionFlags(Qt::TextBrowserInteraction);
    link->setOpenExternalLinks(true);
    link->setAlignment(Qt::AlignCenter);

    auto *buttons = new QDialogButtonBox(QDialogButtonBox::Ok, this);
    buttons->button(QDialogButtonBox::Ok)->setText("Continue");
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);

    layout->addWidget(banner);
    layout->addWidget(author);
    layout->addWidget(link);
    layout->addSpacing(4);
    layout->addWidget(buttons);
}
