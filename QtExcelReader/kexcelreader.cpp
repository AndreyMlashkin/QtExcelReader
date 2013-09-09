#include "kexcelreader.h"

KExcelReader::KExcelReader(QObject *parent)
    : QObject(parent)
{
    init();
}

KExcelReader::~KExcelReader()
{
    Q_ASSERT(m_excel && m_workbooks);

    if (m_workbook)
    {
        m_workbook->dynamicCall("Close()");
    }

    if (m_excel)
    {
        m_excel->dynamicCall("Quit()");
    }
}

void KExcelReader::init()
{
    m_excel = new QAxObject( "Excel.Application", this );

    Q_ASSERT(m_excel);
    if ( m_excel )
    {
        m_workbooks = m_excel->querySubObject( "Workbooks" );
    }
}

bool KExcelReader::open( const QString& xlsFile )
{
    Q_ASSERT(m_excel && m_workbooks);

    m_workbook = m_workbooks->querySubObject( "Open(const QString&)", xlsFile );
    if (m_workbook)
    {
        m_sheets = m_workbook->querySubObject( "Worksheets" );
    }

    return m_sheets != 0;
}

int KExcelReader::sheetCount()
{
    return m_sheets->dynamicCall("Count()").toInt();
}

int KExcelReader::rowCount( QAxObject* sheet )
{
    QAxObject* rows = sheet->querySubObject( "Rows" );
    int count = rows->dynamicCall( "Count()" ).toInt(); //always returns 255

    return count;
}

int KExcelReader::columnCount( QAxObject* sheet )
{
    QAxObject* columns = sheet->querySubObject( "Columns" );
    int count = columns->property("Count").toInt(); //always returns 65535

    return count;
}

QList<QVariantList> KExcelReader::values(int sheetNumber)
{
    Q_ASSERT(sheetNumber > 0);

    QList<QVariantList> data; //Data list from excel, each QVariantList is worksheet row

    //sheet pointer
    QAxObject* sheet = m_sheets->querySubObject( "Item( int )", sheetNumber );

    int rowCnt = rowCount(sheet);
    bool rowCountRecalculated = false;
    for (int row=1; row <= rowCnt; row++)
    {
        QVariantList dataRow;
        bool isEmpty = true; //When all the cells of row are empty, 
                             //it means that file is at end (of course, it maybe not right for different excel files.
                             //It's just criteria to calculate somehow row count for my file)

        for (int column=1; column <= columnCount(sheet); column++)
        {
            QAxObject* cell = sheet->querySubObject( "Cells( int, int )", row, column );
            QVariant value = cell->dynamicCall( "Value()" );

            if (!value.toString().isEmpty() && isEmpty)
            {
                isEmpty = false;
            }

            if (value.isValid())
            {
                dataRow.append(value);
            }
            else
            {
                if (!rowCountRecalculated)
                {
                    rowCountRecalculated = true;
                    rowCnt = row+1;
                }
                else
                {
                    break;
                }
            }
        }

        if (isEmpty) //criteria to get out of cycle
            break;

        data.push_back(dataRow);
    }

    return data;
}
