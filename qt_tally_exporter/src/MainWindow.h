#pragma once

#include "TallyService.h"

#include <QAbstractItemModel>
#include <QLabel>
#include <QCheckBox>
#include <QLineEdit>
#include <QMainWindow>
#include <QMap>
#include <QPlainTextEdit>
#include <QProgressBar>
#include <QPushButton>
#include <QTableView>
#include <QTabWidget>
#include <atomic>

class MainWindow : public QMainWindow {
public:
    MainWindow();

private:
    void buildUi();
    void connectToTally();
    void loadTables();
    void cancelCurrentOperation();
    void clearCache();
    void clearLogs();
    void showAbout();
    void chooseExportDirectory();
    void loadExportSettings();
    void saveExportSettings() const;
    void applyCompanyInfo(const CompanyInfo &info);
    void applyLoadedData(const TallyDataBundle &bundle);
    void exportTable(const QString &tableId);
    void exportAllTables();
    void populateTableView(QTableView *tableView, const TallyTable &table);
    bool writeCsvFile(const QString &path, const TallyTable &table, QString *errorMessage = nullptr);
    void setBusy(bool busy);
    void setStatus(const QString &message);
    void logMessage(const QString &message);
    QString createLogFilePath() const;
    QString defaultExportDirectory() const;
    QString formatRawDate(const QString &value) const;

    QLineEdit *hostEdit_;
    QLineEdit *portEdit_;
    QLineEdit *companyEdit_;
    QLineEdit *fromDateEdit_;
    QLineEdit *toDateEdit_;
    QLineEdit *detectedCompanyEdit_;
    QLineEdit *detectedFromEdit_;
    QLineEdit *detectedToEdit_;
    QLineEdit *exportDirEdit_;
    QCheckBox *overwriteCheckBox_;
    QLabel *statusLabel_;
    QLabel *statsLabel_;
    QProgressBar *progressBar_;
    QPushButton *connectButton_;
    QPushButton *loadButton_;
    QPushButton *cancelButton_;
    QPushButton *clearCacheButton_;
    QPushButton *clearLogsButton_;
    QPushButton *aboutButton_;
    QPushButton *browseExportDirButton_;
    QPlainTextEdit *logEdit_;
    QTabWidget *tabWidget_;
    QString logFilePath_;
    std::atomic_bool operationRunning_{false};

    QMap<QString, TallyTable> tables_;
    QMap<QString, QTableView *> tableViews_;
    QMap<QString, QAbstractItemModel *> tableModels_;
};
