#include "MainWindow.h"

#include <QApplication>
#include <QDateTime>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QGridLayout>
#include <QGroupBox>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QMessageBox>
#include <QPushButton>
#include <QSplitter>
#include <QTextStream>
#include <QVBoxLayout>

namespace {
QString csvEscape(const QString &value) {
    QString escaped = value;
    escaped.replace('"', "\"\"");
    if (escaped.contains(',') || escaped.contains('"') || escaped.contains('\n') || escaped.contains('\r')) {
        return "\"" + escaped + "\"";
    }
    return escaped;
}
}

MainWindow::MainWindow() {
    setWindowTitle("Tally Qt Exporter");
    resize(1320, 840);
    buildUi();
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
    auto *connectButton = new QPushButton("Connect", connectionBox);
    auto *loadButton = new QPushButton("Load Tables", connectionBox);
    buttonsLayout->addWidget(connectButton);
    buttonsLayout->addWidget(loadButton);
    buttonsLayout->addStretch(1);
    progressBar_ = new QProgressBar(connectionBox);
    progressBar_->setRange(0, 0);
    progressBar_->setVisible(false);
    progressBar_->setFixedWidth(160);
    buttonsLayout->addWidget(progressBar_);
    connectionLayout->addLayout(buttonsLayout, 5, 0, 1, 2);

    auto *detailsBox = new QGroupBox("Detected", central);
    auto *detailsLayout = new QGridLayout(detailsBox);
    detectedCompanyEdit_ = new QLineEdit(detailsBox);
    detectedFromEdit_ = new QLineEdit(detailsBox);
    detectedToEdit_ = new QLineEdit(detailsBox);
    detectedCompanyEdit_->setReadOnly(true);
    detectedFromEdit_->setReadOnly(true);
    detectedToEdit_->setReadOnly(true);
    statusLabel_ = new QLabel("Ready", detailsBox);
    statsLabel_ = new QLabel("Vouchers: 0 | Ledgers: 0 | Stock Items: 0 | Stock Vouchers: 0", detailsBox);

    detailsLayout->addWidget(new QLabel("Company", detailsBox), 0, 0);
    detailsLayout->addWidget(detectedCompanyEdit_, 0, 1);
    detailsLayout->addWidget(new QLabel("From", detailsBox), 1, 0);
    detailsLayout->addWidget(detectedFromEdit_, 1, 1);
    detailsLayout->addWidget(new QLabel("To", detailsBox), 2, 0);
    detailsLayout->addWidget(detectedToEdit_, 2, 1);
    detailsLayout->addWidget(new QLabel("Status", detailsBox), 3, 0);
    detailsLayout->addWidget(statusLabel_, 3, 1);
    detailsLayout->addWidget(statsLabel_, 4, 0, 1, 2);

    auto *exportButtonsLayout = new QHBoxLayout();
    auto *exportAllButton = new QPushButton("Export All CSVs", detailsBox);
    auto *exportVouchersButton = new QPushButton("Export Vouchers", detailsBox);
    auto *exportLedgersButton = new QPushButton("Export Ledgers", detailsBox);
    auto *exportStockItemsButton = new QPushButton("Export Stock Items", detailsBox);
    auto *exportStockVouchersButton = new QPushButton("Export Stock Vouchers", detailsBox);
    exportButtonsLayout->addWidget(exportAllButton);
    exportButtonsLayout->addWidget(exportVouchersButton);
    exportButtonsLayout->addWidget(exportLedgersButton);
    exportButtonsLayout->addWidget(exportStockItemsButton);
    exportButtonsLayout->addWidget(exportStockVouchersButton);
    exportButtonsLayout->addStretch(1);
    detailsLayout->addLayout(exportButtonsLayout, 5, 0, 1, 2);

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

    const QList<TallyTable> tableDefs = {
        {"voucher_df", "Vouchers", "vouchers.csv", QStringList()},
        {"ledger_df", "Ledgers", "ledgers.csv", QStringList()},
        {"stock_item_df", "Stock Items", "stock_items.csv", QStringList()},
        {"inventory_df", "Stock Vouchers", "stock_vouchers.csv", QStringList()},
    };

    for (const auto &table : tableDefs) {
        auto *page = new QWidget(tabWidget_);
        auto *layout = new QVBoxLayout(page);
        auto *tableWidget = new QTableWidget(page);
        tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);
        tableWidget->setSelectionBehavior(QAbstractItemView::SelectRows);
        tableWidget->setAlternatingRowColors(true);
        tableWidget->horizontalHeader()->setStretchLastSection(false);
        tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);
        layout->addWidget(tableWidget);
        tabWidget_->addTab(page, table.title);
        tableWidgets_.insert(table.id, tableWidget);
        tables_.insert(table.id, table);
    }

    mainLayout->addWidget(splitter, 1);
    setCentralWidget(central);

    connect(connectButton, &QPushButton::clicked, this, [this]() { connectToTally(); });
    connect(loadButton, &QPushButton::clicked, this, [this]() { loadTables(); });
    connect(exportAllButton, &QPushButton::clicked, this, [this]() { exportAllTables(); });
    connect(exportVouchersButton, &QPushButton::clicked, this, [this]() { exportTable("voucher_df"); });
    connect(exportLedgersButton, &QPushButton::clicked, this, [this]() { exportTable("ledger_df"); });
    connect(exportStockItemsButton, &QPushButton::clicked, this, [this]() { exportTable("stock_item_df"); });
    connect(exportStockVouchersButton, &QPushButton::clicked, this, [this]() { exportTable("inventory_df"); });
}

void MainWindow::connectToTally() {
    setBusy(true);
    setStatus("Connecting to Tally...");
    QApplication::processEvents();

    try {
        const CompanyInfo info = TallyService::getCompanyInfo(hostEdit_->text().trimmed(), portEdit_->text().trimmed());
        if (info.name.isEmpty()) {
            throw std::runtime_error("Active company could not be detected.");
        }
        applyCompanyInfo(info);
        setStatus("Connected to " + info.name);
    } catch (const std::exception &ex) {
        setStatus("Connect failed");
        QMessageBox::critical(this, "Tally Qt Exporter", ex.what());
        logMessage(QString("Connect failed: %1").arg(ex.what()));
    }

    setBusy(false);
}

void MainWindow::loadTables() {
    setBusy(true);
    setStatus("Loading tables...");
    QApplication::processEvents();

    try {
        const TallyDataBundle bundle = TallyService::loadAllData(
            hostEdit_->text().trimmed(),
            portEdit_->text().trimmed(),
            companyEdit_->text().trimmed(),
            fromDateEdit_->text().trimmed(),
            toDateEdit_->text().trimmed()
        );
        applyLoadedData(bundle);
        setStatus("Loaded data for " + bundle.companyName);
    } catch (const std::exception &ex) {
        setStatus("Load failed");
        QMessageBox::critical(this, "Tally Qt Exporter", ex.what());
        logMessage(QString("Load failed: %1").arg(ex.what()));
    }

    setBusy(false);
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

void MainWindow::applyLoadedData(const TallyDataBundle &bundle) {
    detectedCompanyEdit_->setText(bundle.companyName);
    detectedFromEdit_->setText(formatRawDate(bundle.fromDateRaw));
    detectedToEdit_->setText(formatRawDate(bundle.toDateRaw));

    int voucherCount = 0;
    int ledgerCount = 0;
    int stockItemCount = 0;
    int inventoryCount = 0;

    for (auto it = bundle.tables.constBegin(); it != bundle.tables.constEnd(); ++it) {
        tables_[it.key()] = it.value();
        populateTableWidget(tableWidgets_.value(it.key()), it.value());
        if (it.key() == "voucher_df") voucherCount = it.value().rows.size();
        if (it.key() == "ledger_df") ledgerCount = it.value().rows.size();
        if (it.key() == "stock_item_df") stockItemCount = it.value().rows.size();
        if (it.key() == "inventory_df") inventoryCount = it.value().rows.size();
    }

    statsLabel_->setText(QString("Vouchers: %1 | Ledgers: %2 | Stock Items: %3 | Stock Vouchers: %4")
                             .arg(voucherCount)
                             .arg(ledgerCount)
                             .arg(stockItemCount)
                             .arg(inventoryCount));

    logMessage(QString("Loaded company=%1, from=%2, to=%3")
                   .arg(bundle.companyName, formatRawDate(bundle.fromDateRaw), formatRawDate(bundle.toDateRaw)));
}

void MainWindow::exportTable(const QString &tableId) {
    const TallyTable table = tables_.value(tableId);
    if (table.rows.isEmpty()) {
        QMessageBox::warning(this, "Tally Qt Exporter", "This table is empty. Load data first.");
        return;
    }

    const QString path = QFileDialog::getSaveFileName(this, "Save CSV", table.defaultFileName, "CSV Files (*.csv)");
    if (path.isEmpty()) {
        return;
    }

    QString errorMessage;
    if (!writeCsvFile(path, table, &errorMessage)) {
        QMessageBox::critical(this, "Tally Qt Exporter", errorMessage);
        return;
    }

    setStatus("Saved " + QFileInfo(path).fileName());
}

void MainWindow::exportAllTables() {
    bool hasAnyData = false;
    for (auto it = tables_.constBegin(); it != tables_.constEnd(); ++it) {
        if (!it.value().rows.isEmpty()) {
            hasAnyData = true;
            break;
        }
    }
    if (!hasAnyData) {
        QMessageBox::warning(this, "Tally Qt Exporter", "Load data first.");
        return;
    }

    const QString folder = QFileDialog::getExistingDirectory(this, "Select Export Folder");
    if (folder.isEmpty()) {
        return;
    }

    for (auto it = tables_.constBegin(); it != tables_.constEnd(); ++it) {
        const QString path = folder + "/" + it.value().defaultFileName;
        QString errorMessage;
        if (!writeCsvFile(path, it.value(), &errorMessage)) {
            QMessageBox::critical(this, "Tally Qt Exporter", errorMessage);
            return;
        }
    }

    setStatus("Exported all CSVs");
    QMessageBox::information(this, "Tally Qt Exporter", "All CSV files were exported successfully.");
}

void MainWindow::populateTableWidget(QTableWidget *tableWidget, const TallyTable &table) {
    if (!tableWidget) {
        return;
    }

    tableWidget->clear();
    tableWidget->setColumnCount(table.columns.size());
    tableWidget->setHorizontalHeaderLabels(table.columns);
    tableWidget->setRowCount(table.rows.size());

    for (int rowIndex = 0; rowIndex < table.rows.size(); ++rowIndex) {
        const QVariantMap &row = table.rows[rowIndex];
        for (int colIndex = 0; colIndex < table.columns.size(); ++colIndex) {
            const QString &column = table.columns[colIndex];
            auto *item = new QTableWidgetItem(row.value(column).toString());
            tableWidget->setItem(rowIndex, colIndex, item);
        }
    }
}

bool MainWindow::writeCsvFile(const QString &path, const TallyTable &table, QString *errorMessage) {
    QFile file(path);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Truncate)) {
        if (errorMessage) {
            *errorMessage = "Unable to open file for writing: " + path;
        }
        return false;
    }

    QTextStream out(&file);
    out.setEncoding(QStringConverter::Utf8);
    out << QChar(0xFEFF);
    out << table.columns.join(',') << "\n";
    for (const QVariantMap &row : table.rows) {
        QStringList values;
        for (const QString &column : table.columns) {
            values << csvEscape(row.value(column).toString());
        }
        out << values.join(',') << "\n";
    }
    return true;
}

void MainWindow::setBusy(bool busy) {
    hostEdit_->setDisabled(busy);
    portEdit_->setDisabled(busy);
    companyEdit_->setDisabled(busy);
    fromDateEdit_->setDisabled(busy);
    toDateEdit_->setDisabled(busy);
    for (auto it = tableWidgets_.begin(); it != tableWidgets_.end(); ++it) {
        it.value()->setDisabled(busy);
    }
    progressBar_->setVisible(busy);
}

void MainWindow::setStatus(const QString &message) {
    statusLabel_->setText(message);
    logMessage(message);
}

void MainWindow::logMessage(const QString &message) {
    const QString stamp = QDateTime::currentDateTime().toString("HH:mm:ss");
    logEdit_->appendPlainText(QString("[%1] %2").arg(stamp, message));
}

QString MainWindow::formatRawDate(const QString &value) const {
    if (value.size() == 8) {
        return QString("%1-%2-%3").arg(value.mid(0, 4), value.mid(4, 2), value.mid(6, 2));
    }
    return value;
}
