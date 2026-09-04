#include "MainWindow.h"

#include "BannerDialog.h"

#include <QAbstractTableModel>
#include <QApplication>
#include <QDateTime>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QGuiApplication>
#include <QGridLayout>
#include <QGroupBox>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QLocale>
#include <QMessageBox>
#include <QMetaObject>
#include <QPushButton>
#include <QRect>
#include <QScreen>
#include <QSettings>
#include <QSplitter>
#include <QStandardPaths>
#include <QStringConverter>
#include <QTextStream>
#include <QVBoxLayout>
#include <thread>

namespace {
QRect initialWindowGeometry() {
    constexpr int preferredWidth = 1320;
    constexpr int preferredHeight = 840;
    constexpr int desktopMargin = 32;

    const QScreen *screen = QGuiApplication::primaryScreen();
    if (!screen) {
        return QRect(80, 80, preferredWidth, preferredHeight);
    }

    const QRect available = screen->availableGeometry();
    const int width = qMin(preferredWidth, qMax(available.width() - desktopMargin, available.width() * 9 / 10));
    const int height = qMin(preferredHeight, qMax(available.height() - desktopMargin, available.height() * 9 / 10));
    const int x = available.x() + qMax(0, (available.width() - width) / 2);
    const int y = available.y() + qMax(0, (available.height() - height) / 2);
    return QRect(x, y, width, height);
}

QStringList parseCsvLine(const QString &line) {
    QStringList values;
    QString value;
    bool quoted = false;
    for (int i = 0; i < line.size(); ++i) {
        const QChar ch = line.at(i);
        if (quoted) {
            if (ch == '"') {
                if (i + 1 < line.size() && line.at(i + 1) == '"') {
                    value.append('"');
                    ++i;
                } else {
                    quoted = false;
                }
            } else {
                value.append(ch);
            }
        } else if (ch == '"') {
            quoted = true;
        } else if (ch == ',') {
            values.append(value);
            value.clear();
        } else {
            value.append(ch);
        }
    }
    values.append(value);
    return values;
}

QString trimCsvLine(QString line) {
    if (line.endsWith('\n')) {
        line.chop(1);
    }
    if (line.endsWith('\r')) {
        line.chop(1);
    }
    return line;
}

double csvNumber(const QString &value) {
    QString normalized = value.trimmed();
    normalized.remove(',');
    bool ok = false;
    const double amount = normalized.toDouble(&ok);
    return ok ? amount : 0.0;
}

struct AmountSummary {
    double amount = 0.0;
    double debitAmount = 0.0;
    double creditAmount = 0.0;
};

AmountSummary summarizeAmountColumns(const TallyTable &table) {
    AmountSummary summary;
    if (table.csvPath.isEmpty()) {
        return summary;
    }

    QFile file(table.csvPath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return summary;
    }

    const QStringList headers = parseCsvLine(trimCsvLine(QString::fromUtf8(file.readLine())));
    const int amountIndex = headers.indexOf("Amount");
    const int debitIndex = headers.indexOf("DebitAmount");
    const int creditIndex = headers.indexOf("CreditAmount");
    if (amountIndex < 0 && debitIndex < 0 && creditIndex < 0) {
        return summary;
    }

    while (!file.atEnd()) {
        const QStringList row = parseCsvLine(trimCsvLine(QString::fromUtf8(file.readLine())));
        if (amountIndex >= 0 && amountIndex < row.size()) {
            summary.amount += csvNumber(row.at(amountIndex));
        }
        if (debitIndex >= 0 && debitIndex < row.size()) {
            summary.debitAmount += csvNumber(row.at(debitIndex));
        }
        if (creditIndex >= 0 && creditIndex < row.size()) {
            summary.creditAmount += csvNumber(row.at(creditIndex));
        }
    }
    return summary;
}

QString formatCount(int value) {
    return QLocale(QLocale::English, QLocale::India).toString(value);
}

QString formatAmount(double value) {
    if (qAbs(value) < 0.0000001) {
        value = 0.0;
    }
    return QLocale(QLocale::English, QLocale::India).toString(value, 'f', 2);
}

class CsvTableModel final : public QAbstractTableModel {
public:
    explicit CsvTableModel(const TallyTable &table, QObject *parent = nullptr)
        : QAbstractTableModel(parent), table_(table), file_(table.csvPath) {
        buildIndex();
    }

    int rowCount(const QModelIndex &parent = QModelIndex()) const override {
        return parent.isValid() ? 0 : offsets_.size();
    }

    int columnCount(const QModelIndex &parent = QModelIndex()) const override {
        return parent.isValid() ? 0 : table_.columns.size();
    }

    QVariant data(const QModelIndex &index, int role = Qt::DisplayRole) const override {
        if (!index.isValid() || role != Qt::DisplayRole || index.column() < 0 || index.column() >= table_.columns.size()) {
            return {};
        }
        const QStringList row = rowValues(index.row());
        return index.column() < row.size() ? row.at(index.column()) : QString();
    }

    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override {
        if (role != Qt::DisplayRole) {
            return {};
        }
        if (orientation == Qt::Horizontal) {
            return section >= 0 && section < table_.columns.size() ? table_.columns.at(section) : QVariant();
        }
        return section + 1;
    }

private:
    void buildIndex() {
        QFile file(table_.csvPath);
        if (!file.open(QIODevice::ReadOnly)) {
            return;
        }
        file.readLine();
        while (!file.atEnd()) {
            const qint64 offset = file.pos();
            const QByteArray line = file.readLine();
            if (!line.isEmpty()) {
                offsets_.append(offset);
            }
        }
    }

    QStringList rowValues(int row) const {
        if (row < 0 || row >= offsets_.size()) {
            return {};
        }
        if (row == cachedRow_) {
            return cachedValues_;
        }
        if (!file_.isOpen() && !file_.open(QIODevice::ReadOnly)) {
            return {};
        }
        if (!file_.seek(offsets_.at(row))) {
            return {};
        }
        cachedRow_ = row;
        cachedValues_ = parseCsvLine(trimCsvLine(QString::fromUtf8(file_.readLine())));
        return cachedValues_;
    }

    TallyTable table_;
    QVector<qint64> offsets_;
    mutable QFile file_;
    mutable int cachedRow_ = -1;
    mutable QStringList cachedValues_;
};
}

MainWindow::MainWindow() {
    setWindowTitle("Tally Qt Exporter");
    buildUi();
    setGeometry(initialWindowGeometry());
    logFilePath_ = createLogFilePath();
    TallyService::setLogCallback([this](const QString &message) {
        QMetaObject::invokeMethod(this, [this, message]() { logMessage(message); }, Qt::QueuedConnection);
    });
    logMessage("Log file: " + logFilePath_);
    setStatus("Ready");
}

void MainWindow::buildUi() {
    auto *central = new QWidget(this);
    auto *mainLayout = new QVBoxLayout(central);

    auto *topLayout = new QHBoxLayout();
    auto *connectionBox = new QGroupBox("Connection", central);
    auto *connectionLayout = new QGridLayout(connectionBox);

    hostEdit_ = new QLineEdit("localhost", connectionBox);
    portEdit_ = new QLineEdit("9000", connectionBox);
    companyEdit_ = new QLineEdit(connectionBox);
    fromDateEdit_ = new QLineEdit(connectionBox);
    toDateEdit_ = new QLineEdit(connectionBox);

    connectionLayout->addWidget(new QLabel("Host", connectionBox), 0, 0);
    connectionLayout->addWidget(hostEdit_, 0, 1);
    connectionLayout->addWidget(new QLabel("Port", connectionBox), 1, 0);
    connectionLayout->addWidget(portEdit_, 1, 1);
    connectionLayout->addWidget(new QLabel("Company", connectionBox), 2, 0);
    connectionLayout->addWidget(companyEdit_, 2, 1);
    connectionLayout->addWidget(new QLabel("From Date", connectionBox), 3, 0);
    connectionLayout->addWidget(fromDateEdit_, 3, 1);
    connectionLayout->addWidget(new QLabel("To Date", connectionBox), 4, 0);
    connectionLayout->addWidget(toDateEdit_, 4, 1);

    auto *buttonsLayout = new QHBoxLayout();
    connectButton_ = new QPushButton("Connect", connectionBox);
    loadButton_ = new QPushButton("Load Tables", connectionBox);
    cancelButton_ = new QPushButton("Cancel", connectionBox);
    cancelButton_->setEnabled(false);
    buttonsLayout->addWidget(connectButton_);
    buttonsLayout->addWidget(loadButton_);
    buttonsLayout->addWidget(cancelButton_);
    buttonsLayout->addStretch(1);
    progressBar_ = new QProgressBar(connectionBox);
    progressBar_->setRange(0, 0);
    progressBar_->setVisible(false);
    progressBar_->setFixedWidth(160);
    buttonsLayout->addWidget(progressBar_);
    connectionLayout->addLayout(buttonsLayout, 5, 0, 1, 2);

    auto *maintenanceLayout = new QHBoxLayout();
    clearCacheButton_ = new QPushButton("Clear Cache", connectionBox);
    clearLogsButton_ = new QPushButton("Clear Logs", connectionBox);
    aboutButton_ = new QPushButton("About", connectionBox);
    maintenanceLayout->addWidget(clearCacheButton_);
    maintenanceLayout->addWidget(clearLogsButton_);
    maintenanceLayout->addWidget(aboutButton_);
    maintenanceLayout->addStretch(1);
    connectionLayout->addLayout(maintenanceLayout, 6, 0, 1, 2);

    auto *detailsBox = new QGroupBox("Detected", central);
    auto *detailsLayout = new QGridLayout(detailsBox);
    detectedCompanyEdit_ = new QLineEdit(detailsBox);
    detectedFromEdit_ = new QLineEdit(detailsBox);
    detectedToEdit_ = new QLineEdit(detailsBox);
    detectedCompanyEdit_->setReadOnly(true);
    detectedFromEdit_->setReadOnly(true);
    detectedToEdit_->setReadOnly(true);
    statusLabel_ = new QLabel("Ready", detailsBox);
    exportDirEdit_ = new QLineEdit(detailsBox);
    browseExportDirButton_ = new QPushButton("Browse", detailsBox);
    overwriteCheckBox_ = new QCheckBox("Overwrite existing files", detailsBox);

    detailsLayout->addWidget(new QLabel("Company", detailsBox), 0, 0);
    detailsLayout->addWidget(detectedCompanyEdit_, 0, 1);
    detailsLayout->addWidget(new QLabel("From", detailsBox), 1, 0);
    detailsLayout->addWidget(detectedFromEdit_, 1, 1);
    detailsLayout->addWidget(new QLabel("To", detailsBox), 2, 0);
    detailsLayout->addWidget(detectedToEdit_, 2, 1);
    detailsLayout->addWidget(new QLabel("Status", detailsBox), 3, 0);
    detailsLayout->addWidget(statusLabel_, 3, 1);
    detailsLayout->addWidget(new QLabel("Export Folder", detailsBox), 4, 0);
    auto *exportDirLayout = new QHBoxLayout();
    exportDirLayout->addWidget(exportDirEdit_);
    exportDirLayout->addWidget(browseExportDirButton_);
    detailsLayout->addLayout(exportDirLayout, 4, 1);
    detailsLayout->addWidget(overwriteCheckBox_, 5, 1);

    auto *exportButtonsLayout = new QHBoxLayout();
    auto *exportAllButton = new QPushButton("Export All CSVs", detailsBox);
    auto *exportVouchersButton = new QPushButton("Export Vouchers", detailsBox);
    auto *exportAllVouchersButton = new QPushButton("Export All Vouchers", detailsBox);
    auto *exportLedgersButton = new QPushButton("Export Ledgers", detailsBox);
    auto *exportStockItemsButton = new QPushButton("Export Stock Items", detailsBox);
    auto *exportStockVouchersButton = new QPushButton("Export Stock Vouchers", detailsBox);
    exportButtonsLayout->addWidget(exportAllButton);
    exportButtonsLayout->addWidget(exportVouchersButton);
    exportButtonsLayout->addWidget(exportAllVouchersButton);
    exportButtonsLayout->addWidget(exportLedgersButton);
    exportButtonsLayout->addWidget(exportStockItemsButton);
    exportButtonsLayout->addWidget(exportStockVouchersButton);
    exportButtonsLayout->addStretch(1);
    detailsLayout->addLayout(exportButtonsLayout, 6, 0, 1, 2);

    topLayout->addWidget(connectionBox, 1);
    topLayout->addWidget(detailsBox, 1);
    mainLayout->addLayout(topLayout);

    auto *splitter = new QSplitter(Qt::Vertical, central);
    tabWidget_ = new QTabWidget(splitter);
    logEdit_ = new QPlainTextEdit(splitter);
    logEdit_->setReadOnly(true);
    logEdit_->setPlaceholderText("Run connection or export actions to see logs.");
    splitter->setStretchFactor(0, 5);
    splitter->setStretchFactor(1, 2);

    auto *summaryPage = new QWidget(tabWidget_);
    auto *summaryLayout = new QGridLayout(summaryPage);
    summaryLayout->setColumnStretch(0, 0);
    summaryLayout->setColumnStretch(1, 1);
    summaryLayout->setColumnStretch(2, 1);
    summaryLayout->setColumnStretch(3, 1);
    summaryLayout->setColumnStretch(4, 1);

    auto *countsBox = new QGroupBox("Counts", summaryPage);
    auto *countsLayout = new QGridLayout(countsBox);
    countsLayout->setHorizontalSpacing(24);
    countsLayout->setVerticalSpacing(14);
    const QList<QPair<QString, QString>> countRows = {
        {"voucher_count", "Vouchers"},
        {"all_voucher_count", "All Vouchers"},
        {"ledger_count", "Ledgers"},
        {"stock_item_count", "Stock Items"},
        {"inventory_count", "Stock Vouchers"},
    };
    QFont countLabelFont;
    countLabelFont.setPointSize(11);
    QFont countValueFont;
    countValueFont.setPointSize(18);
    countValueFont.setBold(true);
    for (int row = 0; row < countRows.size(); ++row) {
        auto *label = new QLabel(countRows.at(row).second, countsBox);
        auto *value = new QLabel("0", countsBox);
        label->setFont(countLabelFont);
        value->setFont(countValueFont);
        value->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        countsLayout->addWidget(label, row, 0);
        countsLayout->addWidget(value, row, 1);
        summaryLabels_.insert(countRows.at(row).first, value);
    }

    auto *amountsBox = new QGroupBox("Accounting Totals", summaryPage);
    auto *amountsLayout = new QGridLayout(amountsBox);
    amountsLayout->setHorizontalSpacing(32);
    amountsLayout->setVerticalSpacing(18);
    const QStringList amountHeaders = {"Table", "Amount", "Debit Amount", "Credit Amount", "Debit - Credit"};
    QFont amountHeaderFont;
    amountHeaderFont.setPointSize(11);
    amountHeaderFont.setBold(true);
    for (int column = 0; column < amountHeaders.size(); ++column) {
        auto *header = new QLabel(amountHeaders.at(column), amountsBox);
        header->setFont(amountHeaderFont);
        header->setAlignment(column == 0 ? Qt::AlignLeft : Qt::AlignRight);
        amountsLayout->addWidget(header, 0, column);
    }
    const QList<QPair<QString, QString>> amountRows = {
        {"voucher", "Vouchers CSV"},
        {"all_voucher", "All Vouchers CSV"},
    };
    QFont amountRowFont;
    amountRowFont.setPointSize(12);
    QFont amountValueFont;
    amountValueFont.setPointSize(18);
    amountValueFont.setBold(true);
    for (int row = 0; row < amountRows.size(); ++row) {
        const int layoutRow = row + 1;
        auto *rowLabel = new QLabel(amountRows.at(row).second, amountsBox);
        rowLabel->setFont(amountRowFont);
        amountsLayout->addWidget(rowLabel, layoutRow, 0);
        const QString prefix = amountRows.at(row).first;
        const QStringList keys = {prefix + "_amount", prefix + "_debit", prefix + "_credit", prefix + "_balance"};
        for (int column = 0; column < keys.size(); ++column) {
            auto *value = new QLabel("0.00", amountsBox);
            value->setFont(amountValueFont);
            value->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
            amountsLayout->addWidget(value, layoutRow, column + 1);
            summaryLabels_.insert(keys.at(column), value);
        }
    }
    for (int column = 1; column < amountHeaders.size(); ++column) {
        amountsLayout->setColumnMinimumWidth(column, 140);
    }

    summaryLayout->addWidget(countsBox, 0, 0);
    summaryLayout->addWidget(amountsBox, 0, 1, 1, 4);
    summaryLayout->setRowStretch(1, 1);
    tabWidget_->addTab(summaryPage, "Summary");
    resetSummary();

    const QList<TallyTable> tableDefs = {
        {"voucher_df", "Vouchers", "vouchers.csv", QStringList()},
        {"all_voucher_df", "All Vouchers", "allvouchers.csv", QStringList()},
        {"ledger_df", "Ledgers", "ledgers.csv", QStringList()},
        {"stock_item_df", "Stock Items", "stock_items.csv", QStringList()},
        {"inventory_df", "Stock Vouchers", "stock_vouchers.csv", QStringList()},
    };

    for (const auto &table : tableDefs) {
        auto *page = new QWidget(tabWidget_);
        auto *layout = new QVBoxLayout(page);
        auto *tableView = new QTableView(page);
        tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
        tableView->setSelectionBehavior(QAbstractItemView::SelectRows);
        tableView->setAlternatingRowColors(true);
        tableView->setSortingEnabled(false);
        tableView->setWordWrap(false);
        tableView->setVerticalScrollMode(QAbstractItemView::ScrollPerPixel);
        tableView->horizontalHeader()->setStretchLastSection(false);
        tableView->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        tableView->verticalHeader()->setDefaultSectionSize(22);
        layout->addWidget(tableView);
        tabWidget_->addTab(page, table.title);
        tableViews_.insert(table.id, tableView);
        tables_.insert(table.id, table);
    }

    mainLayout->addWidget(splitter, 1);
    setCentralWidget(central);

    connect(connectButton_, &QPushButton::clicked, this, [this]() { connectToTally(); });
    connect(loadButton_, &QPushButton::clicked, this, [this]() { loadTables(); });
    connect(cancelButton_, &QPushButton::clicked, this, [this]() { cancelCurrentOperation(); });
    connect(clearCacheButton_, &QPushButton::clicked, this, [this]() { clearCache(); });
    connect(clearLogsButton_, &QPushButton::clicked, this, [this]() { clearLogs(); });
    connect(aboutButton_, &QPushButton::clicked, this, [this]() { showAbout(); });
    connect(browseExportDirButton_, &QPushButton::clicked, this, [this]() { chooseExportDirectory(); });
    connect(exportDirEdit_, &QLineEdit::editingFinished, this, [this]() { saveExportSettings(); });
    connect(overwriteCheckBox_, &QCheckBox::toggled, this, [this]() { saveExportSettings(); });
    connect(exportAllButton, &QPushButton::clicked, this, [this]() { exportAllTables(); });
    connect(exportVouchersButton, &QPushButton::clicked, this, [this]() { exportTable("voucher_df"); });
    connect(exportAllVouchersButton, &QPushButton::clicked, this, [this]() { exportTable("all_voucher_df"); });
    connect(exportLedgersButton, &QPushButton::clicked, this, [this]() { exportTable("ledger_df"); });
    connect(exportStockItemsButton, &QPushButton::clicked, this, [this]() { exportTable("stock_item_df"); });
    connect(exportStockVouchersButton, &QPushButton::clicked, this, [this]() { exportTable("inventory_df"); });
    loadExportSettings();
}

void MainWindow::connectToTally() {
    if (operationRunning_.load()) {
        return;
    }
    setBusy(true);
    setStatus("Connecting to Tally...");
    operationRunning_.store(true);
    TallyService::resetCancellation();
    const QString host = hostEdit_->text().trimmed();
    const QString port = portEdit_->text().trimmed();

    std::thread([this, host, port]() {
        try {
            const CompanyInfo info = TallyService::getCompanyInfo(host, port);
            if (info.name.isEmpty()) {
                throw std::runtime_error("Active company could not be detected.");
            }
            QMetaObject::invokeMethod(this, [this, info]() {
                applyCompanyInfo(info);
                setStatus("Connected to " + info.name);
                setBusy(false);
            }, Qt::QueuedConnection);
        } catch (const std::exception &ex) {
            const QString message = ex.what();
            QMetaObject::invokeMethod(this, [this, message]() {
                setStatus("Connect failed");
                QMessageBox::critical(this, "Tally Qt Exporter", message);
                logMessage(QString("Connect failed: %1").arg(message));
                setBusy(false);
            }, Qt::QueuedConnection);
        }
    }).detach();
}

void MainWindow::loadTables() {
    if (operationRunning_.load()) {
        return;
    }
    setBusy(true);
    setStatus("Loading tables...");
    operationRunning_.store(true);
    TallyService::resetCancellation();
    const QString host = hostEdit_->text().trimmed();
    const QString port = portEdit_->text().trimmed();
    const QString company = companyEdit_->text().trimmed();
    const QString fromDate = fromDateEdit_->text().trimmed();
    const QString toDate = toDateEdit_->text().trimmed();

    std::thread([this, host, port, company, fromDate, toDate]() {
        try {
            const TallyDataBundle bundle = TallyService::loadAllData(host, port, company, fromDate, toDate);
            QMetaObject::invokeMethod(this, [this, bundle]() {
                applyLoadedData(bundle);
                setStatus("Loaded data for " + bundle.companyName);
                setBusy(false);
            }, Qt::QueuedConnection);
        } catch (const std::exception &ex) {
            const QString message = ex.what();
            QMetaObject::invokeMethod(this, [this, message]() {
                setStatus(message == "Operation cancelled by user." ? message : "Load failed");
                if (message != "Operation cancelled by user.") {
                    QMessageBox::critical(this, "Tally Qt Exporter", message);
                }
                logMessage(QString("Load stopped: %1").arg(message));
                setBusy(false);
            }, Qt::QueuedConnection);
        }
    }).detach();
}

void MainWindow::applyCompanyInfo(const CompanyInfo &info) {
    detectedCompanyEdit_->setText(info.name);
    detectedFromEdit_->setText(formatRawDate(info.startDateRaw));
    detectedToEdit_->setText(formatRawDate(info.endDateRaw));

    if (companyEdit_->text().trimmed().isEmpty()) {
        companyEdit_->setText(info.name);
    }
    if (fromDateEdit_->text().trimmed().isEmpty()) {
        fromDateEdit_->setText(info.startDateRaw);
    }
    if (toDateEdit_->text().trimmed().isEmpty()) {
        toDateEdit_->setText(info.endDateRaw);
    }
}

void MainWindow::cancelCurrentOperation() {
    if (!operationRunning_.load()) {
        return;
    }
    logMessage("Cancel requested. Aborting active Tally request after the current network call stops.");
    setStatus("Cancelling...");
    TallyService::cancelCurrentOperation();
}

void MainWindow::clearCache() {
    QString cacheRoot = QStandardPaths::writableLocation(QStandardPaths::CacheLocation);
    if (cacheRoot.isEmpty()) {
        cacheRoot = QDir::currentPath() + "/.tally_cache";
    } else {
        cacheRoot += "/tally_xml";
    }
    int removed = 0;
    QDir dir(cacheRoot);
    if (dir.exists()) {
        const QFileInfoList files = dir.entryInfoList(QDir::Files | QDir::NoDotAndDotDot);
        removed = files.size();
        dir.removeRecursively();
    }
    logMessage(QString("Cache cleared. Removed %1 cached XML file(s).").arg(removed));
    QMessageBox::information(this, "Cache Cleared", QString("Removed %1 cached XML file(s).").arg(removed));
}

void MainWindow::clearLogs() {
    int removed = 0;
    const QFileInfo currentLog(logFilePath_);
    QDir dir(currentLog.absolutePath());
    if (dir.exists()) {
        const QFileInfoList files = dir.entryInfoList({"*.log"}, QDir::Files | QDir::NoDotAndDotDot);
        for (const QFileInfo &file : files) {
            if (QFile::remove(file.absoluteFilePath())) {
                ++removed;
            }
        }
    }
    logEdit_->clear();
    logFilePath_ = createLogFilePath();
    logMessage(QString("Logs cleared. Removed %1 log file(s). New log file: %2").arg(removed).arg(logFilePath_));
    QMessageBox::information(this, "Logs Cleared", QString("Removed %1 log file(s).").arg(removed));
}

void MainWindow::showAbout() {
    BannerDialog dialog(this);
    dialog.setWindowIcon(windowIcon());
    dialog.exec();
}

void MainWindow::chooseExportDirectory() {
    const QString startDir = defaultExportDirectory();
    const QString folder = QFileDialog::getExistingDirectory(this, "Select Default Export Folder", startDir);
    if (folder.isEmpty()) {
        return;
    }
    exportDirEdit_->setText(QDir::toNativeSeparators(folder));
    saveExportSettings();
}

void MainWindow::loadExportSettings() {
    QSettings settings;
    exportDirEdit_->setText(QDir::toNativeSeparators(settings.value("export/defaultDirectory").toString()));
    overwriteCheckBox_->setChecked(settings.value("export/overwriteExisting", false).toBool());
}

void MainWindow::saveExportSettings() const {
    QSettings settings;
    settings.setValue("export/defaultDirectory", QDir::fromNativeSeparators(exportDirEdit_->text().trimmed()));
    settings.setValue("export/overwriteExisting", overwriteCheckBox_->isChecked());
}

void MainWindow::applyLoadedData(const TallyDataBundle &bundle) {
    detectedCompanyEdit_->setText(bundle.companyName);
    detectedFromEdit_->setText(formatRawDate(bundle.fromDateRaw));
    detectedToEdit_->setText(formatRawDate(bundle.toDateRaw));

    for (auto it = bundle.tables.constBegin(); it != bundle.tables.constEnd(); ++it) {
        tables_[it.key()] = it.value();
        populateTableView(tableViews_.value(it.key()), it.value());
    }
    updateSummary(bundle);

    logMessage(QString("Loaded company=%1, from=%2, to=%3")
                   .arg(bundle.companyName, formatRawDate(bundle.fromDateRaw), formatRawDate(bundle.toDateRaw)));
}

void MainWindow::exportTable(const QString &tableId) {
    const TallyTable table = tables_.value(tableId);
    if (table.rowCount == 0 || table.csvPath.isEmpty()) {
        QMessageBox::warning(this, "Tally Qt Exporter", "This table is empty. Load data first.");
        return;
    }

    QString path;
    const QString exportDir = defaultExportDirectory();
    if (!exportDir.isEmpty()) {
        QDir dir(exportDir);
        if (!dir.exists() && !dir.mkpath(".")) {
            QMessageBox::critical(this, "Tally Qt Exporter", "Unable to create export folder: " + exportDir);
            return;
        }
        path = dir.filePath(table.defaultFileName);
    } else {
        path = QFileDialog::getSaveFileName(this, "Save CSV", table.defaultFileName, "CSV Files (*.csv)");
        if (path.isEmpty()) {
            return;
        }
    }

    QString errorMessage;
    if (!writeCsvFile(path, table, &errorMessage)) {
        QMessageBox::critical(this, "Tally Qt Exporter", errorMessage);
        return;
    }

    setStatus("Saved " + QFileInfo(path).fileName());
    logMessage("Exported " + path);
}

void MainWindow::exportAllTables() {
    bool hasAnyData = false;
    for (auto it = tables_.constBegin(); it != tables_.constEnd(); ++it) {
        if (it.value().rowCount > 0 && !it.value().csvPath.isEmpty()) {
            hasAnyData = true;
            break;
        }
    }
    if (!hasAnyData) {
        QMessageBox::warning(this, "Tally Qt Exporter", "Load data first.");
        return;
    }

    QString folder = defaultExportDirectory();
    if (folder.isEmpty()) {
        folder = QFileDialog::getExistingDirectory(this, "Select Export Folder");
    }
    if (folder.isEmpty()) {
        return;
    }
    QDir dir(folder);
    if (!dir.exists() && !dir.mkpath(".")) {
        QMessageBox::critical(this, "Tally Qt Exporter", "Unable to create export folder: " + folder);
        return;
    }

    for (auto it = tables_.constBegin(); it != tables_.constEnd(); ++it) {
        if (it.value().rowCount == 0 || it.value().csvPath.isEmpty()) {
            continue;
        }
        const QString path = dir.filePath(it.value().defaultFileName);
        QString errorMessage;
        if (!writeCsvFile(path, it.value(), &errorMessage)) {
            QMessageBox::critical(this, "Tally Qt Exporter", errorMessage);
            return;
        }
    }

    setStatus("Exported all CSVs");
    logMessage("Exported all CSVs to " + QDir::toNativeSeparators(dir.absolutePath()));
    QMessageBox::information(this, "Tally Qt Exporter", "All CSV files were exported successfully.");
}

void MainWindow::populateTableView(QTableView *tableView, const TallyTable &table) {
    if (!tableView) {
        return;
    }

    if (QAbstractItemModel *oldModel = tableModels_.take(table.id)) {
        oldModel->deleteLater();
    }
    auto *model = new CsvTableModel(table, tableView);
    tableModels_.insert(table.id, model);
    tableView->setModel(model);
    tableView->horizontalHeader()->setDefaultSectionSize(140);
}

void MainWindow::updateSummary(const TallyDataBundle &bundle) {
    const auto setLabel = [this](const QString &key, const QString &value) {
        if (QLabel *label = summaryLabels_.value(key, nullptr)) {
            label->setText(value);
        }
    };

    const TallyTable voucherTable = bundle.tables.value("voucher_df");
    const TallyTable allVoucherTable = bundle.tables.value("all_voucher_df");
    const TallyTable ledgerTable = bundle.tables.value("ledger_df");
    const TallyTable stockItemTable = bundle.tables.value("stock_item_df");
    const TallyTable inventoryTable = bundle.tables.value("inventory_df");

    setLabel("voucher_count", formatCount(voucherTable.rowCount));
    setLabel("all_voucher_count", formatCount(allVoucherTable.rowCount));
    setLabel("ledger_count", formatCount(ledgerTable.rowCount));
    setLabel("stock_item_count", formatCount(stockItemTable.rowCount));
    setLabel("inventory_count", formatCount(inventoryTable.rowCount));

    const AmountSummary voucherSummary = summarizeAmountColumns(voucherTable);
    const AmountSummary allVoucherSummary = summarizeAmountColumns(allVoucherTable);
    const auto setAmountSummary = [&](const QString &prefix, const AmountSummary &summary) {
        setLabel(prefix + "_amount", formatAmount(summary.amount));
        setLabel(prefix + "_debit", formatAmount(summary.debitAmount));
        setLabel(prefix + "_credit", formatAmount(summary.creditAmount));
        setLabel(prefix + "_balance", formatAmount(summary.debitAmount - summary.creditAmount));
    };
    setAmountSummary("voucher", voucherSummary);
    setAmountSummary("all_voucher", allVoucherSummary);
}

void MainWindow::resetSummary() {
    for (auto it = summaryLabels_.begin(); it != summaryLabels_.end(); ++it) {
        it.value()->setText(it.key().endsWith("_count") ? "0" : "0.00");
    }
}

bool MainWindow::writeCsvFile(const QString &path, const TallyTable &table, QString *errorMessage) {
    if (table.csvPath.isEmpty() || !QFileInfo::exists(table.csvPath)) {
        if (errorMessage) {
            *errorMessage = "The generated table file is missing. Load data again.";
        }
        return false;
    }
    if (QFileInfo(table.csvPath).absoluteFilePath() == QFileInfo(path).absoluteFilePath()) {
        return true;
    }
    if (QFileInfo::exists(path)) {
        if (!overwriteCheckBox_->isChecked()) {
            if (errorMessage) {
                *errorMessage = "File already exists and overwrite is disabled: " + path;
            }
            return false;
        }
        if (!QFile::remove(path)) {
            if (errorMessage) {
                *errorMessage = "Unable to overwrite file: " + path;
            }
            return false;
        }
    }
    if (!QFile::copy(table.csvPath, path)) {
        if (errorMessage) {
            *errorMessage = "Unable to save file: " + path;
        }
        return false;
    }
    return true;
}

QString MainWindow::defaultExportDirectory() const {
    return QDir::fromNativeSeparators(exportDirEdit_->text().trimmed());
}

void MainWindow::setBusy(bool busy) {
    operationRunning_.store(busy);
    hostEdit_->setDisabled(busy);
    portEdit_->setDisabled(busy);
    companyEdit_->setDisabled(busy);
    fromDateEdit_->setDisabled(busy);
    toDateEdit_->setDisabled(busy);
    connectButton_->setDisabled(busy);
    loadButton_->setDisabled(busy);
    clearCacheButton_->setDisabled(busy);
    clearLogsButton_->setDisabled(busy);
    aboutButton_->setDisabled(busy);
    exportDirEdit_->setDisabled(busy);
    browseExportDirButton_->setDisabled(busy);
    overwriteCheckBox_->setDisabled(busy);
    cancelButton_->setEnabled(busy);
    for (auto it = tableViews_.begin(); it != tableViews_.end(); ++it) {
        it.value()->setDisabled(busy);
    }
    progressBar_->setVisible(busy);
}

void MainWindow::setStatus(const QString &message) {
    statusLabel_->setText(message);
    logMessage(message);
}

void MainWindow::logMessage(const QString &message) {
    const QString stamp = QDateTime::currentDateTime().toString("yyyy-MM-dd HH:mm:ss");
    const QString line = QString("[%1] %2").arg(stamp, message);
    logEdit_->appendPlainText(line);
    if (!logFilePath_.isEmpty()) {
        QFile file(logFilePath_);
        if (file.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Append)) {
            QTextStream out(&file);
            out.setEncoding(QStringConverter::Utf8);
            out << line << "\n";
        }
    }
}

QString MainWindow::createLogFilePath() const {
    const QString logDirPath = QDir::currentPath() + "/.tally_logs";
    QDir().mkpath(logDirPath);
    const QString stamp = QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss");
    return QDir(logDirPath).filePath(QString("tally_qt_exporter_%1.log").arg(stamp));
}

QString MainWindow::formatRawDate(const QString &value) const {
    if (value.size() == 8) {
        return QString("%1-%2-%3").arg(value.mid(0, 4), value.mid(4, 2), value.mid(6, 2));
    }
    return value;
}
