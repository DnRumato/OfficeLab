#include "mainwindow.h"
#include "ui_mainwindow.h"


QString generateRandomString(int length) {
    const QString possibleCharacters("abcdefghijklmnopqrstuvwxyz");
    const QString possibleUpper("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
    QString randomString;

    for(int i=0; i<length; ++i) {
        int index = rand() % possibleCharacters.length();
        randomString.append(possibleCharacters.at(index));
    }

    int index = rand() % possibleUpper.length();
    randomString.insert(0, possibleUpper.at(index));

    return randomString;
}

class AlignDelegate : public QStyledItemDelegate {
public:
    AlignDelegate(Qt::Alignment alignment, QObject *parent = nullptr)
        : QStyledItemDelegate(parent), m_alignment(alignment) {}
    void initStyleOption(QStyleOptionViewItem *option, const QModelIndex &index) const override {
        QStyledItemDelegate::initStyleOption(option, index);
        option->displayAlignment = m_alignment;
    }
private:
    Qt::Alignment m_alignment;
};

void sortTable(QTableWidget* table, int number, bool flag) {
    table->setSortingEnabled(true);
    for (int i = 0; i < table->rowCount(); ++i) {
        table->item(i, 0)->setData(Qt::UserRole, table->item(i, 0)->text());
        QDate date = QDate::fromString(table->item(i, 1)->text(), "dd.MM.yyyy");
        table->item(i, 1)->setData(Qt::UserRole, date);
        double number1 = table->item(i, 2)->text().toDouble();
        table->item(i, 2)->setData(Qt::UserRole, number1);
        double number2 = table->item(i, 3)->text().toDouble();
        table->item(i, 3)->setData(Qt::UserRole, number2);
    }
    if(flag)
        table->sortByColumn(number, Qt::AscendingOrder);
    else
        table->sortByColumn(number, Qt::DescendingOrder);
}





MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_clicked()
{
    int n = ui->spinBox->value();

    ui->tableWidget->setRowCount(0);
    ui->tableWidget->setRowCount(n);

    ui->tableWidget->setItemDelegateForColumn(0, new AlignDelegate(Qt::AlignLeft));
    ui->tableWidget->setItemDelegateForColumn(1, new AlignDelegate(Qt::AlignCenter));
    ui->tableWidget->setItemDelegateForColumn(2, new AlignDelegate(Qt::AlignRight));
    ui->tableWidget->setItemDelegateForColumn(3, new AlignDelegate(Qt::AlignRight));

    for (int i = 0; i < n; ++i) {
        int length = 2 + rand() % 8;
        QString name = generateRandomString(length);
        QDate date = QDate::currentDate().addDays(rand() % 365);
        QString dateString = date.toString("dd.MM.yyyy");
        double num1 = (double)(rand() % 10000) / 100;
        double num2 = (double)(rand() % 10000) / 100;

        ui->tableWidget->setItem(i, 0, new QTableWidgetItem(name));
        ui->tableWidget->setItem(i, 1, new QTableWidgetItem(dateString));
        ui->tableWidget->setItem(i, 2, new QTableWidgetItem(QString::number(num1, 'f', 2)));
        ui->tableWidget->setItem(i, 3, new QTableWidgetItem(QString::number(num2, 'f', 2)));
    }
}


void MainWindow::on_pushButton_2_clicked()
{
    sortTable(ui->tableWidget, ui->spinBox_2->value(), false);
}


void MainWindow::on_pushButton_3_clicked()
{
    sortTable(ui->tableWidget, ui->spinBox_2->value(), true);
}


void MainWindow::on_pushButton_4_clicked()
{
   /* QProcess unzipProcess;
        unzipProcess.start("unzip", QStringList() << "D:/Documents/OfficeLab/Договор на аренду помещения.docx" << "-d" << "temp");
        unzipProcess.waitForFinished();

        QFile file("temp/word/document.xml");
        file.open(QIODevice::ReadWrite | QIODevice::Text);
        //if (file.open(QIODevice::ReadWrite | QIODevice::Text)) {
            QString xml = file.readAll();
            xml.replace("{{Имя}}", ui->lineEdit->text());
            xml.replace("{{Реквизиты}}", ui->lineEdit_2->text());
            xml.replace("{{НомерДоговора}}", ui->lineEdit_3->text());
            file.resize(0);
            file.write(xml.toUtf8());
        //}
        file.close();

        QProcess zipProcess;
        zipProcess.setWorkingDirectory("temp");
        zipProcess.start("zip", QStringList() << "-r" << "D:/Documents/OfficeLab/Договор на аренду помещения.docx" << ".");
        zipProcess.waitForFinished();

        QDir("temp").removeRecursively();
      */
    QDesktopServices::openUrl(QUrl::fromLocalFile("D:/Documents/OfficeLab/Договор на аренду помещения.docx"));
}


void MainWindow::on_pushButton_5_clicked()
{
    QString filter = "Word Files (*.docx *.dotx)";
    QString fileName = QFileDialog::getOpenFileName(this, "Открыть файл", QDir::currentPath(), filter);
    if (!fileName.isEmpty()) {
        QDesktopServices::openUrl(QUrl::fromLocalFile(fileName));
    }
}

