#pragma once

#include <QMap>
#include <QString>
#include <QStringList>
#include <QVariantMap>
#include <QVector>

struct TallyTable {
    QString id;
    QString title;
    QString defaultFileName;
    QStringList columns;
    QVector<QVariantMap> rows;
};

struct CompanyInfo {
    QString name;
    QString startDateRaw;
    QString endDateRaw;
};

struct TallyDataBundle {
    QString companyName;
    QString fromDateRaw;
    QString toDateRaw;
    QMap<QString, TallyTable> tables;
};

class TallyService {
public:
    static CompanyInfo getCompanyInfo(const QString &host, const QString &port);
    static TallyDataBundle loadAllData(const QString &host, const QString &port, const QString &company,
                                       const QString &fromDate, const QString &toDate);
};
