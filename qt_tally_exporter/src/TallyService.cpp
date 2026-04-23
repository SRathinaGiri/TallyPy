#include "TallyService.h"

#include <QByteArray>
#include <QEventLoop>
#include <QMap>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QRegularExpression>
#include <QSet>
#include <QStringConverter>
#include <QTimer>
#include <QUrl>
#include <QDomDocument>
#include <algorithm>
#include <cmath>
#include <stdexcept>

namespace {
const QSet<QString> kAccountingVoucherTypes = {
    "Sales",
    "Purchase",
    "Journal",
    "Receipt",
    "Payment",
    "Debit Note",
    "Credit Note",
    "Contra",
};

const QSet<QString> kBsPrimaryGroups = {
    "Capital Account", "Reserves & Surplus",
    "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans",
    "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors",
    "Fixed Assets",
    "Investments",
    "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors",
    "Misc. Expenses (ASSET)",
    "Suspense Account",
    "Branch / Divisions",
};

const QSet<QString> kPlPrimaryGroups = {
    "Sales Accounts",
    "Purchase Accounts",
    "Direct Incomes",
    "Indirect Incomes",
    "Direct Expenses",
    "Indirect Expenses",
};

const QSet<QString> kPrimaryGroups = []() {
    QSet<QString> groups = kBsPrimaryGroups;
    groups.unite(kPlPrimaryGroups);
    return groups;
}();

const QMap<QString, QString> kCurrencyFallbacks = {
    {"INR", QString::fromUtf8("₹")},
    {"INDIAN RUPEE", QString::fromUtf8("₹")},
    {"RUPEE", QString::fromUtf8("₹")},
    {"RUPEES", QString::fromUtf8("₹")},
    {"RS", QString::fromUtf8("₹")},
    {"RS.", QString::fromUtf8("₹")},
    {"USD", "$"},
    {"US DOLLAR", "$"},
    {"DOLLAR", "$"},
    {"EUR", QString::fromUtf8("€")},
    {"EURO", QString::fromUtf8("€")},
    {"GBP", QString::fromUtf8("£")},
    {"POUND", QString::fromUtf8("£")},
    {"POUND STERLING", QString::fromUtf8("£")},
    {"AED", QString::fromUtf8("د.إ")},
    {"DIRHAM", QString::fromUtf8("د.إ")},
    {"", ""},
};

const QStringList kVoucherColumns = {
    "Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName",
    "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger",
    "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName",
    "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName",
    "FromDate", "ToDate"
};

const QStringList kLedgerColumns = {
    "MasterID", "Name", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN",
    "StartingFrom", "CurrencyName", "StateName", "Parent", "PartyGSTIN",
    "OpeningBalance", "ClosingBalance", "CompanyName", "FromDate", "ToDate"
};

const QStringList kStockItemColumns = {
    "Name", "Parent", "Category", "LedgerName", "OpeningBalance", "OpeningValue",
    "BasicValue", "BasicQty", "OpeningRate", "ClosingBalance", "ClosingValue",
    "ClosingRate", "CompanyName", "FromDate", "ToDate"
};

const QStringList kStockVoucherColumns = {
    "Date", "VoucherTypeName", "VoucherNumber", "StockItemName", "BilledQty",
    "Rate", "Amount", "GodownName", "BatchName", "VoucherNarration", "CompanyName",
    "FromDate", "ToDate"
};

struct GroupInfo {
    QString parent;
    QString nature;
    QString primaryGroup;
};

QString cleanText(const QString &text) {
    QString value = text;
    static const QRegularExpression controlChars("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    value.remove(controlChars);
    return value.trimmed();
}

QString stripNs(const QString &tag) {
    const int idx = tag.indexOf('}');
    return idx >= 0 ? tag.mid(idx + 1) : tag;
}

QString xmlCleanup(QString xmlText) {
    static const QRegularExpression charRefs(R"(&#(x[0-9A-Fa-f]+|\d+);)");
    QString cleaned;
    cleaned.reserve(xmlText.size());
    int last = 0;
    auto it = charRefs.globalMatch(xmlText);
    while (it.hasNext()) {
        auto match = it.next();
        cleaned += xmlText.mid(last, match.capturedStart() - last);
        const QString value = match.captured(1);
        bool ok = false;
        uint codepoint = 0;
        if (value.startsWith('x', Qt::CaseInsensitive)) {
            codepoint = value.mid(1).toUInt(&ok, 16);
        } else {
            codepoint = value.toUInt(&ok, 10);
        }
        if (ok && (codepoint == 9 || codepoint == 10 || codepoint == 13 ||
                   (codepoint >= 32 && codepoint <= 55295) ||
                   (codepoint >= 57344 && codepoint <= 65533) ||
                   (codepoint >= 65536 && codepoint <= 1114111))) {
            cleaned += match.captured(0);
        }
        last = match.capturedEnd();
    }
    cleaned += xmlText.mid(last);
    xmlText = cleaned;

    static const QRegularExpression controlChars("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    xmlText.remove(controlChars);
    xmlText.replace(QRegularExpression(R"(&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;))"), "&amp;");
    xmlText.replace(QRegularExpression(R"(<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*))"), "<\\1\\2");
    xmlText.replace(QRegularExpression(R"(\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*")"), "");
    xmlText.replace(QRegularExpression(R"(\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*'[^']*')"), "");
    return xmlText;
}

QDomDocument parseXmlRoot(const QString &xmlText) {
    QDomDocument doc;
    QString errorMsg;
    int errorLine = 0;
    int errorColumn = 0;
    if (!doc.setContent(xmlCleanup(xmlText), &errorMsg, &errorLine, &errorColumn)) {
        throw std::runtime_error(QString("XML parse error at line %1, column %2: %3")
                                     .arg(errorLine)
                                     .arg(errorColumn)
                                     .arg(errorMsg)
                                     .toStdString());
    }
    return doc;
}

QList<QDomElement> directChildren(const QDomElement &elem, const QString &localName) {
    QList<QDomElement> children;
    QDomNode node = elem.firstChild();
    const QString wanted = localName.toUpper();
    while (!node.isNull()) {
        if (node.isElement()) {
            QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                children.append(child);
            }
        }
        node = node.nextSibling();
    }
    return children;
}

QString directChildText(const QDomElement &elem, const QString &localName) {
    QDomNode node = elem.firstChild();
    const QString wanted = localName.toUpper();
    while (!node.isNull()) {
        if (node.isElement()) {
            QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                return cleanText(child.text());
            }
        }
        node = node.nextSibling();
    }
    return {};
}

QString firstDescendantText(const QDomElement &elem, const QString &localName) {
    const QString wanted = localName.toUpper();
    QDomNode node = elem.firstChild();
    while (!node.isNull()) {
        if (node.isElement()) {
            const QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                const QString value = cleanText(child.text());
                if (!value.isEmpty()) {
                    return value;
                }
            }
            const QString nested = firstDescendantText(child, localName);
            if (!nested.isEmpty()) {
                return nested;
            }
        }
        node = node.nextSibling();
    }
    return {};
}

QString firstNonEmptyText(const QDomElement &elem, const QStringList &names) {
    for (const QString &name : names) {
        const QString value = directChildText(elem, name);
        if (!value.isEmpty()) {
            return value;
        }
    }
    return {};
}

QString formatTallyDate(const QString &value) {
    const QString text = cleanText(value);
    static const QRegularExpression eightDigits(R"(^\d{8}$)");
    if (eightDigits.match(text).hasMatch()) {
        return QString("%1-%2-%3").arg(text.mid(0, 4), text.mid(4, 2), text.mid(6, 2));
    }
    return text;
}

QString escapeXml(const QString &value) {
    QString escaped = value.toHtmlEscaped();
    escaped.replace('\'', "&apos;");
    return escaped;
}

QString normalizeAmountText(const QString &value) {
    QString text = cleanText(value);
    text.remove(',');
    if (text.isEmpty()) {
        return {};
    }
    static const QRegularExpression numberPattern(R"([-+]?\d+(?:\.\d+)?)");
    auto matches = numberPattern.globalMatch(text);
    QString token;
    while (matches.hasNext()) {
        token = matches.next().captured(0);
    }
    return token.isEmpty() ? text : token;
}

double toDoubleValue(const QString &value, double defaultValue = 0.0) {
    const QString normalized = normalizeAmountText(value);
    if (normalized.isEmpty()) {
        return defaultValue;
    }
    bool ok = false;
    const double result = normalized.toDouble(&ok);
    return ok ? result : defaultValue;
}

QString numberToString(double value) {
    if (std::abs(value) < 0.0000001) {
        value = 0.0;
    }
    return QString::number(value, 'f', 2);
}

QPair<QString, QString> natureFromPrimaryGroup(const QString &primaryGroup) {
    const QString pg = cleanText(primaryGroup).toLower();
    if (QStringList({
            "current assets", "fixed assets", "investments", "misc. expenses (asset)",
            "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)",
            "stock-in-hand", "sundry debtors"
        }).contains(pg)) {
        return {"BS", "Assets"};
    }
    if (QStringList({
            "capital account", "current liabilities", "loans (liability)", "suspense account",
            "branch / divisions", "bank od a/c", "duties & taxes", "provisions",
            "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"
        }).contains(pg)) {
        return {"BS", "Liabilities"};
    }
    if (QStringList({"direct incomes", "indirect incomes", "sales accounts"}).contains(pg)) {
        return {"PL", "Income"};
    }
    if (QStringList({"direct expenses", "indirect expenses", "purchase accounts"}).contains(pg)) {
        return {"PL", "Expenses"};
    }
    return {"Unknown", "Unknown"};
}

QString ledgerPrimaryGroup(const QString &ledgerName, const QMap<QString, QVariantMap> &ledgerMeta) {
    QSet<QString> seen;
    QString current = cleanText(ledgerName);
    while (!current.isEmpty() && !seen.contains(current)) {
        seen.insert(current);
        const QVariantMap meta = ledgerMeta.value(current);
        const QString parent = cleanText(meta.value("Parent").toString());
        if (parent.isEmpty()) {
            return {};
        }
        if (kPrimaryGroups.contains(parent)) {
            return parent;
        }
        current = parent;
    }
    return {};
}

QString postToTally(const QString &url, const QString &xmlText, int timeoutMs = 120000) {
    QNetworkAccessManager manager;
    QNetworkRequest request{QUrl(url)};
    request.setHeader(QNetworkRequest::ContentTypeHeader, "text/xml; charset=utf-8");

    QEventLoop loop;
    QNetworkReply *reply = manager.post(request, xmlText.toUtf8());
    QObject::connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);

    QTimer timer;
    timer.setSingleShot(true);
    QObject::connect(&timer, &QTimer::timeout, &loop, &QEventLoop::quit);
    timer.start(timeoutMs);
    loop.exec();

    if (timer.isActive()) {
        timer.stop();
    } else {
        reply->abort();
        reply->deleteLater();
        throw std::runtime_error("Request to Tally timed out.");
    }

    const auto networkError = reply->error();
    const QString errorText = reply->errorString();
    const QByteArray body = reply->readAll();
    reply->deleteLater();

    if (networkError != QNetworkReply::NoError) {
        throw std::runtime_error(QString("Tally request failed: %1").arg(errorText).toStdString());
    }
    return QString::fromUtf8(body);
}

QString buildLedgerRequestXml(const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE>"
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildVoucherRequestXml(const QString &company, const QString &fromDate, const QString &toDate) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    staticVars << QString("<SVFROMDATE TYPE='Date'>%1</SVFROMDATE>").arg(escapeXml(fromDate));
    staticVars << QString("<SVTODATE TYPE='Date'>%1</SVTODATE>").arg(escapeXml(toDate));
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVouchers</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<SYSTEM TYPE='Formulae' NAME='IsAccountingVoucher'>"
        "($VoucherTypeName = \"Sales\") OR ($VoucherTypeName = \"Purchase\") OR "
        "($VoucherTypeName = \"Journal\") OR ($VoucherTypeName = \"Receipt\") OR "
        "($VoucherTypeName = \"Payment\") OR ($VoucherTypeName = \"Debit Note\") OR "
        "($VoucherTypeName = \"Credit Note\")"
        "</SYSTEM>"
        "<OBJECT NAME=\"All Ledger Entries\">"
        "<COMPUTE>EntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>EntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        "</OBJECT>"
        "<COLLECTION NAME=\"MyVouchers\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, "
        "PartyGSTIN, IsOptional, AllLedgerEntries.LedgerName, AllLedgerEntries.Amount, "
        "AllLedgerEntries.IsDeemedPositive, AllLedgerEntries.EntryLedgerMasterID, "
        "AllLedgerEntries.EntryParentLedger, AllLedgerEntries.EntryPrimaryGroup, "
        "AllLedgerEntries.EntryLedgerGSTIN</FETCH>"
        "<FILTER>IsAccountingVoucher</FILTER>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildStockItemRequestXml(const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyStockItems</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyStockItems\"><TYPE>StockItem</TYPE>"
        "<FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH>"
        "<COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE>"
        "<COMPUTE>ClosingValue:$_ClosingValue</COMPUTE>"
        "<COMPUTE>ClosingRate:$_ClosingRate</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildInventoryEntriesRequestXml(const QString &company, const QString &fromDate, const QString &toDate) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    staticVars << QString("<SVFROMDATE TYPE='Date'>%1</SVFROMDATE>").arg(escapeXml(fromDate));
    staticVars << QString("<SVTODATE TYPE='Date'>%1</SVTODATE>").arg(escapeXml(toDate));
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyInventoryVouchers</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyInventoryVouchers\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, "
        "InventoryEntries.*, AllInventoryEntries.*, InventoryEntriesIn.*, InventoryEntriesOut.*</FETCH>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildCompanyListRequestXml() {
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>List of Companies</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "</DESC></BODY></ENVELOPE>"
    );
}

QPair<QMap<QString, QString>, QMap<QString, GroupInfo>> fetchTallyMetadata(const QString &url, const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }

    const QString vtypeXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllVoucherTypes</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllVoucherTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    const QString groupXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllGroups</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    QMap<QString, QString> vtypeMap;
    QMap<QString, GroupInfo> groupMap;

    try {
        const QDomDocument vtypeDoc = parseXmlRoot(postToTally(url, vtypeXml));
        QDomNodeList vtypes = vtypeDoc.elementsByTagName("VOUCHERTYPE");
        for (int i = 0; i < vtypes.size(); ++i) {
            const QDomElement elem = vtypes.at(i).toElement();
            const QString name = directChildText(elem, "NAME");
            const QString parent = directChildText(elem, "PARENT");
            if (!name.isEmpty()) {
                vtypeMap.insert(name, parent.isEmpty() ? name : parent);
            }
        }

        const QDomDocument groupDoc = parseXmlRoot(postToTally(url, groupXml));
        QDomNodeList groups = groupDoc.elementsByTagName("GROUP");
        for (int i = 0; i < groups.size(); ++i) {
            const QDomElement elem = groups.at(i).toElement();
            const QString name = directChildText(elem, "NAME");
            if (!name.isEmpty()) {
                groupMap.insert(name, {
                    directChildText(elem, "PARENT"),
                    directChildText(elem, "NATURE"),
                    directChildText(elem, "_PRIMARYGROUP")
                });
            }
        }

        const QSet<QString> baseTypes = {
            "Sales", "Purchase", "Journal", "Receipt", "Payment", "Debit Note", "Credit Note", "Contra", "Stock Journal"
        };
        for (int pass = 0; pass < 5; ++pass) {
            for (auto it = vtypeMap.begin(); it != vtypeMap.end(); ++it) {
                const QString parentName = it.value();
                if (!parentName.isEmpty() && !baseTypes.contains(parentName) && vtypeMap.contains(parentName)) {
                    it.value() = vtypeMap.value(parentName);
                }
            }
        }

        for (int pass = 0; pass < 5; ++pass) {
            for (auto it = groupMap.begin(); it != groupMap.end(); ++it) {
                GroupInfo &info = it.value();
                if (!info.parent.isEmpty() && groupMap.contains(info.parent)) {
                    if (info.nature.isEmpty()) {
                        info.nature = groupMap.value(info.parent).nature;
                    }
                    if (info.primaryGroup.isEmpty()) {
                        info.primaryGroup = groupMap.value(info.parent).primaryGroup;
                    }
                }
            }
        }
    } catch (...) {
    }

    return {vtypeMap, groupMap};
}

QVector<QVariantMap> parseLedgers(const QDomDocument &doc, const QMap<QString, GroupInfo> &groupMap) {
    QVector<QVariantMap> rows;
    QMap<QString, QVariantMap> ledgerLookup;
    QDomNodeList nodes = doc.elementsByTagName("LEDGER");

    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement elem = nodes.at(i).toElement();
        if (stripNs(elem.tagName()).toUpper() != "LEDGER") {
            continue;
        }

        const QString name = cleanText(elem.attribute("NAME").isEmpty() ? directChildText(elem, "NAME") : elem.attribute("NAME"));
        if (name.isEmpty()) {
            continue;
        }

        const QString parent = directChildText(elem, "PARENT");
        const GroupInfo gInfo = groupMap.value(parent);
        const QString primaryGroup = !gInfo.primaryGroup.isEmpty() ? gInfo.primaryGroup :
                                     firstNonEmptyText(elem, {"PRIMARYGROUP"}).isEmpty() ? firstDescendantText(elem, "PRIMARYGROUP") :
                                                                                           firstNonEmptyText(elem, {"PRIMARYGROUP"});

        QVariantMap row;
        row.insert("MasterID", cleanText(elem.attribute("MASTERID").isEmpty() ? directChildText(elem, "MASTERID") : elem.attribute("MASTERID")));
        row.insert("Name", name);
        row.insert("PrimaryGroup", primaryGroup);
        row.insert("Nature", "");
        row.insert("NatureOfGroup", gInfo.nature);
        row.insert("PAN", firstNonEmptyText(elem, {"INCOMETAXNUMBER", "PAN"}).isEmpty() ? firstDescendantText(elem, "INCOMETAXNUMBER") : firstNonEmptyText(elem, {"INCOMETAXNUMBER", "PAN"}));
        row.insert("StartingFrom", firstNonEmptyText(elem, {"STARTINGFROM"}).isEmpty() ? firstDescendantText(elem, "STARTINGFROM") : firstNonEmptyText(elem, {"STARTINGFROM"}));
        row.insert("CurrencyNameRaw", firstNonEmptyText(elem, {"CURRENCYNAME"}).isEmpty() ? firstDescendantText(elem, "CURRENCYNAME") : firstNonEmptyText(elem, {"CURRENCYNAME"}));
        row.insert("CurrencyFormalNameRaw", firstNonEmptyText(elem, {"CURRENCYFORMALNAME"}).isEmpty() ? firstDescendantText(elem, "CURRENCYFORMALNAME") : firstNonEmptyText(elem, {"CURRENCYFORMALNAME"}));
        row.insert("CurrencySymbolRaw", firstNonEmptyText(elem, {"CURRENCYSYMBOL"}).isEmpty() ? firstDescendantText(elem, "CURRENCYSYMBOL") : firstNonEmptyText(elem, {"CURRENCYSYMBOL"}));
        row.insert("CurrencyOriginalSymbolRaw", firstNonEmptyText(elem, {"CURRENCYORIGINALSYMBOL"}).isEmpty() ? firstDescendantText(elem, "CURRENCYORIGINALSYMBOL") : firstNonEmptyText(elem, {"CURRENCYORIGINALSYMBOL"}));
        row.insert("StateName", firstNonEmptyText(elem, {"STATENAME"}).isEmpty() ? firstDescendantText(elem, "STATENAME") : firstNonEmptyText(elem, {"STATENAME"}));
        row.insert("Parent", parent);
        row.insert("PartyGSTIN", firstNonEmptyText(elem, {"PARTYGSTIN", "GSTIN"}).isEmpty() ? firstDescendantText(elem, "PARTYGSTIN") : firstNonEmptyText(elem, {"PARTYGSTIN", "GSTIN"}));
        row.insert("OpeningBalance", numberToString(toDoubleValue(firstNonEmptyText(elem, {"OPENINGBALANCE"}).isEmpty() ? firstDescendantText(elem, "OPENINGBALANCE") : firstNonEmptyText(elem, {"OPENINGBALANCE"}))));
        row.insert("ClosingBalance", numberToString(toDoubleValue(firstNonEmptyText(elem, {"CLOSINGBALANCE"}).isEmpty() ? firstDescendantText(elem, "CLOSINGBALANCE") : firstNonEmptyText(elem, {"CLOSINGBALANCE"}))));
        rows.append(row);
        ledgerLookup.insert(name, row);
    }

    for (QVariantMap &row : rows) {
        QString primaryGroup = row.value("PrimaryGroup").toString();
        if (primaryGroup.isEmpty()) {
            primaryGroup = ledgerPrimaryGroup(row.value("Name").toString(), ledgerLookup);
            row.insert("PrimaryGroup", primaryGroup);
        }

        if (row.value("NatureOfGroup").toString().isEmpty() && groupMap.contains(primaryGroup)) {
            row.insert("NatureOfGroup", groupMap.value(primaryGroup).nature);
        }

        QString nature = row.value("Nature").toString();
        QString natureOfGroup = row.value("NatureOfGroup").toString();
        if (!natureOfGroup.isEmpty()) {
            const QString lowered = natureOfGroup.toLower();
            if (lowered == "assets" || lowered == "liabilities") {
                nature = "BS";
            } else if (lowered == "income" || lowered == "expenses") {
                nature = "PL";
            }
        }
        if (nature.isEmpty() && !primaryGroup.isEmpty()) {
            const auto pair = natureFromPrimaryGroup(primaryGroup);
            nature = pair.first;
            natureOfGroup = pair.second;
        }
        row.insert("Nature", nature);
        row.insert("NatureOfGroup", natureOfGroup);

        const QString currencyKey = cleanText((row.value("CurrencyFormalNameRaw").toString().isEmpty()
                                               ? row.value("CurrencyNameRaw").toString()
                                               : row.value("CurrencyFormalNameRaw").toString())).toUpper();
        row.insert("CurrencyName", kCurrencyFallbacks.value(currencyKey,
                                                            cleanText(row.value("CurrencySymbolRaw").toString().isEmpty()
                                                                          ? row.value("CurrencyOriginalSymbolRaw").toString()
                                                                          : row.value("CurrencySymbolRaw").toString())));
    }

    std::sort(rows.begin(), rows.end(), [](const QVariantMap &a, const QVariantMap &b) {
        return a.value("MasterID").toInt() == b.value("MasterID").toInt()
                   ? a.value("Name").toString() < b.value("Name").toString()
                   : a.value("MasterID").toInt() < b.value("MasterID").toInt();
    });

    QVector<QVariantMap> output;
    for (const QVariantMap &row : rows) {
        QVariantMap ordered;
        for (const QString &column : kLedgerColumns) {
            ordered.insert(column, row.value(column).toString());
        }
        output.append(ordered);
    }
    return output;
}

QVector<QVariantMap> parseVouchers(const QDomDocument &doc, const QMap<QString, QVariantMap> &ledgerMeta,
                                   const QString &company, const QString &fromDate, const QString &toDate,
                                   const QMap<QString, QString> &vtypeMap) {
    QVector<QVariantMap> rows;
    const QString formattedFromDate = formatTallyDate(fromDate);
    const QString formattedToDate = formatTallyDate(toDate);
    QDomNodeList nodes = doc.elementsByTagName("VOUCHER");

    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement voucher = nodes.at(i).toElement();
        if (stripNs(voucher.tagName()).toUpper() != "VOUCHER") {
            continue;
        }

        const QString voucherType = directChildText(voucher, "VOUCHERTYPENAME");
        const QString baseType = vtypeMap.value(voucherType, voucherType);
        if (!kAccountingVoucherTypes.contains(baseType)) {
            continue;
        }

        const QString voucherDate = formatTallyDate(directChildText(voucher, "DATE"));
        const QString voucherNumber = directChildText(voucher, "VOUCHERNUMBER");
        const QString partyLedgerName = directChildText(voucher, "PARTYLEDGERNAME").isEmpty() ? "N/A" : directChildText(voucher, "PARTYLEDGERNAME");
        const QString voucherGstin = directChildText(voucher, "PARTYGSTIN");
        const QString voucherNarration = firstNonEmptyText(voucher, {"NARRATION", "VOUCHERNARRATION"});
        const QString isOptional = directChildText(voucher, "ISOPTIONAL").toUpper() == "YES" ? "Yes" : "No";
        const QString voucherCompany = firstNonEmptyText(voucher, {"COMPANYNAME", "SVCURRENTCOMPANY"}).isEmpty()
                                           ? company
                                           : firstNonEmptyText(voucher, {"COMPANYNAME", "SVCURRENTCOMPANY"});

        QList<QDomElement> entryNodes = directChildren(voucher, "ALLLEDGERENTRIES.LIST");
        if (entryNodes.isEmpty()) {
            entryNodes = directChildren(voucher, "LEDGERENTRIES.LIST");
        }

        for (const QDomElement &entry : entryNodes) {
            const QString ledgerName = directChildText(entry, "LEDGERNAME");
            const double amountValue = toDoubleValue(directChildText(entry, "AMOUNT"));
            const QString isDeemedPositive = directChildText(entry, "ISDEEMEDPOSITIVE").toUpper();
            if (ledgerName.isEmpty() || std::abs(amountValue) < 0.0000001) {
                continue;
            }

            const double baseAmount = std::abs(amountValue);
            const double signedAmount = isDeemedPositive == "YES" ? -baseAmount : baseAmount;
            const QString drCr = signedAmount < 0 ? "Dr" : "Cr";
            const double debitAmount = signedAmount < 0 ? baseAmount : 0.0;
            const double creditAmount = signedAmount > 0 ? baseAmount : 0.0;

            QVariantMap meta = ledgerMeta.value(ledgerName);
            QString primaryGroup = meta.value("PrimaryGroup").toString();
            QString parentLedger = meta.value("Parent").toString();
            QString ledgerGstin = meta.value("PartyGSTIN").toString();
            QString ledgerMasterId = meta.value("MasterID").toString();
            QString nature = meta.value("Nature").toString();
            QString natureOfGroup = meta.value("NatureOfGroup").toString();
            QString pan = meta.value("PAN").toString();

            const QString entryMasterId = directChildText(entry, "ENTRYLEDGERMASTERID");
            const QString entryParent = directChildText(entry, "ENTRYPARENTLEDGER");
            const QString entryPrimaryGroup = directChildText(entry, "ENTRYPRIMARYGROUP");
            const QString entryGstin = directChildText(entry, "ENTRYLEDGERGSTIN");

            if (!entryMasterId.isEmpty()) ledgerMasterId = entryMasterId;
            if (!entryParent.isEmpty()) parentLedger = entryParent;
            if (!entryPrimaryGroup.isEmpty()) primaryGroup = entryPrimaryGroup;
            if (!entryGstin.isEmpty()) ledgerGstin = entryGstin;
            if (nature.isEmpty() && !primaryGroup.isEmpty()) {
                const auto pair = natureFromPrimaryGroup(primaryGroup);
                nature = pair.first;
                natureOfGroup = pair.second;
            }

            QVariantMap row;
            row.insert("Date", voucherDate);
            row.insert("VoucherTypeName", voucherType);
            row.insert("BaseVoucherType", baseType);
            row.insert("VoucherNumber", voucherNumber);
            row.insert("LedgerName", ledgerName);
            row.insert("MasterID", ledgerMasterId);
            row.insert("Amount", numberToString(signedAmount));
            row.insert("DrCr", drCr);
            row.insert("DebitAmount", numberToString(debitAmount));
            row.insert("CreditAmount", numberToString(creditAmount));
            row.insert("ParentLedger", parentLedger);
            row.insert("PrimaryGroup", primaryGroup);
            row.insert("Nature", nature);
            row.insert("NatureOfGroup", natureOfGroup);
            row.insert("PAN", pan);
            row.insert("PartyLedgerName", partyLedgerName);
            row.insert("PartyGSTIN", voucherGstin);
            row.insert("LedgerGSTIN", ledgerGstin);
            row.insert("VoucherNarration", voucherNarration);
            row.insert("IsOptional", isOptional);
            row.insert("CompanyName", voucherCompany);
            row.insert("FromDate", formattedFromDate);
            row.insert("ToDate", formattedToDate);
            rows.append(row);
        }
    }
    return rows;
}

QVector<QVariantMap> parseStockItems(const QDomDocument &doc) {
    QVector<QVariantMap> rows;
    QDomNodeList nodes = doc.elementsByTagName("STOCKITEM");
    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement elem = nodes.at(i).toElement();
        if (stripNs(elem.tagName()).toUpper() != "STOCKITEM") {
            continue;
        }

        const QString name = cleanText(elem.attribute("NAME").isEmpty() ? directChildText(elem, "NAME") : elem.attribute("NAME"));
        if (name.isEmpty()) {
            continue;
        }

        QVariantMap row;
        row.insert("Name", name);
        row.insert("Parent", directChildText(elem, "PARENT"));
        row.insert("Category", directChildText(elem, "CATEGORY"));
        row.insert("LedgerName", directChildText(elem, "LEDGERNAME"));
        row.insert("OpeningBalance", numberToString(toDoubleValue(directChildText(elem, "OPENINGBALANCE"))));
        row.insert("OpeningValue", numberToString(toDoubleValue(directChildText(elem, "OPENINGVALUE"))));
        row.insert("BasicValue", numberToString(toDoubleValue(directChildText(elem, "BASICVALUE"))));
        row.insert("BasicQty", numberToString(toDoubleValue(directChildText(elem, "BASICQTY"))));
        row.insert("OpeningRate", numberToString(toDoubleValue(directChildText(elem, "OPENINGRATE"))));
        row.insert("ClosingBalance", numberToString(toDoubleValue(directChildText(elem, "CLOSINGBALANCE"))));
        row.insert("ClosingValue", numberToString(toDoubleValue(directChildText(elem, "CLOSINGVALUE"))));
        row.insert("ClosingRate", numberToString(toDoubleValue(directChildText(elem, "CLOSINGRATE"))));
        rows.append(row);
    }
    return rows;
}

QVector<QVariantMap> parseInventoryEntries(const QDomDocument &doc, const QString &company) {
    QVector<QVariantMap> rows;
    QDomNodeList nodes = doc.elementsByTagName("VOUCHER");
    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement voucher = nodes.at(i).toElement();
        if (stripNs(voucher.tagName()).toUpper() != "VOUCHER") {
            continue;
        }

        const QString voucherType = directChildText(voucher, "VOUCHERTYPENAME");
        if (voucherType.contains("Order", Qt::CaseInsensitive)) {
            continue;
        }

        const QString voucherDate = formatTallyDate(directChildText(voucher, "DATE"));
        const QString voucherNumber = directChildText(voucher, "VOUCHERNUMBER");
        const QString voucherNarration = firstNonEmptyText(voucher, {"NARRATION", "VOUCHERNARRATION"});
        const QString voucherCompany = firstNonEmptyText(voucher, {"COMPANYNAME", "SVCURRENTCOMPANY"}).isEmpty()
                                           ? company
                                           : firstNonEmptyText(voucher, {"COMPANYNAME", "SVCURRENTCOMPANY"});

        QDomNode childNode = voucher.firstChild();
        while (!childNode.isNull()) {
            if (childNode.isElement()) {
                const QDomElement inv = childNode.toElement();
                if (stripNs(inv.tagName()).toUpper().contains("INVENTORYENTRIES")) {
                    const QString itemName = directChildText(inv, "STOCKITEMNAME");
                    if (!itemName.isEmpty()) {
                        const bool isInward = directChildText(inv, "ISDEEMEDPOSITIVE").toUpper() == "YES";
                        const double amount = std::abs(toDoubleValue(directChildText(inv, "AMOUNT")));
                        const double qty = std::abs(toDoubleValue(directChildText(inv, "BILLEDQTY")));
                        const double rate = toDoubleValue(directChildText(inv, "RATE"));

                        QString godown;
                        QString batch;
                        const QList<QDomElement> batches = directChildren(inv, "BATCHALLOCATIONS.LIST");
                        if (!batches.isEmpty()) {
                            godown = directChildText(batches.first(), "GODOWNNAME");
                            batch = directChildText(batches.first(), "BATCHNAME");
                        }

                        QVariantMap row;
                        row.insert("Date", voucherDate);
                        row.insert("VoucherTypeName", voucherType);
                        row.insert("VoucherNumber", voucherNumber);
                        row.insert("StockItemName", itemName.trimmed());
                        row.insert("BilledQty", numberToString(isInward ? qty : -qty));
                        row.insert("Rate", numberToString(rate));
                        row.insert("Amount", numberToString(isInward ? amount : -amount));
                        row.insert("GodownName", godown);
                        row.insert("BatchName", batch);
                        row.insert("VoucherNarration", voucherNarration);
                        row.insert("CompanyName", voucherCompany);
                        rows.append(row);
                    }
                }
            }
            childNode = childNode.nextSibling();
        }
    }
    return rows;
}

TallyTable makeTable(const QString &id, const QString &title, const QString &fileName, const QStringList &columns,
                     const QVector<QVariantMap> &rows, const QString &companyName, const QString &fromDate, const QString &toDate) {
    TallyTable table;
    table.id = id;
    table.title = title;
    table.defaultFileName = fileName;
    table.columns = columns;

    for (QVariantMap row : rows) {
        row.insert("CompanyName", companyName);
        row.insert("FromDate", formatTallyDate(fromDate));
        row.insert("ToDate", formatTallyDate(toDate));
        for (const QString &column : columns) {
            if (!row.contains(column)) {
                row.insert(column, "");
            }
        }
        QVariantMap ordered;
        for (const QString &column : columns) {
            ordered.insert(column, row.value(column).toString());
        }
        table.rows.append(ordered);
    }
    return table;
}
}

CompanyInfo TallyService::getCompanyInfo(const QString &host, const QString &port) {
    const QString url = QString("http://%1:%2").arg(host, port);
    const QString activeCompanyXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt, Guid</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    const auto extractCompany = [](const QDomDocument &doc) -> CompanyInfo {
        const QDomNodeList nodes = doc.elementsByTagName("COMPANY");
        for (int i = 0; i < nodes.size(); ++i) {
            const QDomElement cmp = nodes.at(i).toElement();
            const QString name = cleanText(cmp.attribute("NAME").isEmpty() ? directChildText(cmp, "NAME") : cmp.attribute("NAME"));
            if (!name.isEmpty()) {
                return {name, directChildText(cmp, "STARTINGFROM"), directChildText(cmp, "ENDINGAT")};
            }
        }

        const QDomElement root = doc.documentElement();
        const QString currentCompany = firstDescendantText(root, "SVCURRENTCOMPANY");
        if (!currentCompany.isEmpty()) {
            return {currentCompany, firstDescendantText(root, "STARTINGFROM"), firstDescendantText(root, "ENDINGAT")};
        }

        const QString companyName = firstDescendantText(root, "COMPANYNAME");
        if (!companyName.isEmpty()) {
            return {companyName, firstDescendantText(root, "STARTINGFROM"), firstDescendantText(root, "ENDINGAT")};
        }

        return {};
    };

    const QDomDocument activeDoc = parseXmlRoot(postToTally(url, activeCompanyXml, 10000));
    CompanyInfo info = extractCompany(activeDoc);
    if (!info.name.isEmpty()) {
        return info;
    }

    const QDomDocument companyListDoc = parseXmlRoot(postToTally(url, buildCompanyListRequestXml(), 10000));
    info = extractCompany(companyListDoc);
    if (!info.name.isEmpty()) {
        return info;
    }

    return {};
}

TallyDataBundle TallyService::loadAllData(const QString &host, const QString &port, const QString &company,
                                          const QString &fromDate, const QString &toDate) {
    const QString url = QString("http://%1:%2").arg(host, port);
    QString selectedCompany = cleanText(company);
    QString selectedFrom = cleanText(fromDate);
    QString selectedTo = cleanText(toDate);

    if (selectedCompany.isEmpty() || selectedFrom.isEmpty() || selectedTo.isEmpty()) {
        const CompanyInfo info = getCompanyInfo(host, port);
        if (selectedCompany.isEmpty()) selectedCompany = info.name;
        if (selectedFrom.isEmpty()) selectedFrom = info.startDateRaw;
        if (selectedTo.isEmpty()) selectedTo = info.endDateRaw;
    }

    const auto metadata = fetchTallyMetadata(url, selectedCompany);
    const QMap<QString, QString> &vtypeMap = metadata.first;
    const QMap<QString, GroupInfo> &groupMap = metadata.second;

    const QDomDocument ledgerDoc = parseXmlRoot(postToTally(url, buildLedgerRequestXml(selectedCompany)));
    const QVector<QVariantMap> ledgerRows = parseLedgers(ledgerDoc, groupMap);
    QMap<QString, QVariantMap> ledgerMeta;
    for (const QVariantMap &row : ledgerRows) {
        ledgerMeta.insert(row.value("Name").toString(), row);
    }

    const QDomDocument voucherDoc = parseXmlRoot(postToTally(url, buildVoucherRequestXml(selectedCompany, selectedFrom, selectedTo)));
    const QString status = cleanText(firstDescendantText(voucherDoc.documentElement(), "STATUS"));
    if (status == "0") {
        const QString errorText = firstDescendantText(voucherDoc.documentElement(), "LINEERROR");
        throw std::runtime_error((errorText.isEmpty() ? QString("Tally returned STATUS=0") : errorText).toStdString());
    }
    const QVector<QVariantMap> voucherRows = parseVouchers(voucherDoc, ledgerMeta, selectedCompany, selectedFrom, selectedTo, vtypeMap);

    const QDomDocument stockDoc = parseXmlRoot(postToTally(url, buildStockItemRequestXml(selectedCompany)));
    const QVector<QVariantMap> stockRows = parseStockItems(stockDoc);

    const QDomDocument inventoryDoc = parseXmlRoot(postToTally(url, buildInventoryEntriesRequestXml(selectedCompany, selectedFrom, selectedTo)));
    const QVector<QVariantMap> inventoryRows = parseInventoryEntries(inventoryDoc, selectedCompany);

    TallyDataBundle bundle;
    bundle.companyName = selectedCompany;
    bundle.fromDateRaw = selectedFrom;
    bundle.toDateRaw = selectedTo;
    bundle.tables.insert("voucher_df", makeTable("voucher_df", "Vouchers", "vouchers.csv", kVoucherColumns, voucherRows, selectedCompany, selectedFrom, selectedTo));
    bundle.tables.insert("ledger_df", makeTable("ledger_df", "Ledgers", "ledgers.csv", kLedgerColumns, ledgerRows, selectedCompany, selectedFrom, selectedTo));
    bundle.tables.insert("stock_item_df", makeTable("stock_item_df", "Stock Items", "stock_items.csv", kStockItemColumns, stockRows, selectedCompany, selectedFrom, selectedTo));
    bundle.tables.insert("inventory_df", makeTable("inventory_df", "Stock Vouchers", "stock_vouchers.csv", kStockVoucherColumns, inventoryRows, selectedCompany, selectedFrom, selectedTo));
    return bundle;
}
