#include <QApplication>
#include <QStandardPaths>
#include <QDir>
#include "MSWORD.h"

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    Word::Application word;
    if (!word.isNull()) {
        word.SetVisible(false);

        Word::Documents* docs = word.Documents();
        Word::Document* newDoc = docs->Add();
        Word::Paragraph* p = newDoc->Content()->Paragraphs()->Add();
        p->Range()->SetText("Hello Word Document from Qt!");
        p->Range()->InsertParagraphAfter();
        p->Range()->SetText("That's it!");

        QDir outDir(QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation));

        QVariant fileName = outDir.absoluteFilePath("wordaut.docx");
        QVariant format = Word::wdFormatXMLDocument;
        newDoc->SaveAs2(fileName, format);

        QVariant fileName2 = outDir.absoluteFilePath("wordaut2.doc");
        QVariant format2 = Word::wdFormatDocument;
        newDoc->SaveAs2(fileName2, format2);

        newDoc->Close();
        word.Quit();
    }

    return 0;
}
