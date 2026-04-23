#pragma once

#include "TallyService.h"

#include <QLabel>
#include <QLineEdit>
#include <QMainWindow>
#include <QMap>
#include <QPlainTextEdit>
#include <QProgressBar>
#include <QTableWidget>
#include <QTabWidget>

class MainWindow : public QMainWindow {
public:
    MainWindow();

private:
    void buildUi();
    void connectToTally();
    void loadTables();
    void applyCompanyInfo(const CompanyInfo &info);
    void applyLoadedData(const TallyDataBundle &bundle);
    void exportTable(const QString &tableId);
    void exportAllTables();
    void populateTableWidget(QTableWidget *tableWidget, const TallyTable &table);
    bool writeCsvFile(const QString &path, const TallyTable &table, QString *errorMessage = nullptr);
    void setBusy(bool busy);
    void setStatus(const QString &message);
    void logMessage(const QString &message);
    QString formatRawDate(const QString &value) const;

    QLineEdit *hostEdit_;
    QLineEdit *portEdit_;
    QLineEdit *companyEdit_;
    QLineEdit *fromDateEdit_;
    QLineEdit *toDateEdit_;
    QLineEdit *detectedCompanyEdit_;
    QLineEdit *detectedFromEdit_;
    QLineEdit *detectedToEdit_;
    QLabel *statusLabel_;
    QLabel *statsLabel_;
    QProgressBar *progressBar_;
    QPlainTextEdit *logEdit_;
    QTabWidget *tabWidget_;

    QMap<QString, TallyTable> tables_;
    QMap<QString, QTableWidget *> tableWidgets_;
};
