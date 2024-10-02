#include <vcl.h>
//#include "Excel_2K_SRVR.h"

class TExcelApp1 {
        public:
        Variant app ;
        Variant books ;
        Variant book ;
        Variant sheet ;
        TExcelApp1 (boolean Visible);
        TExcelApp1 (AnsiString filename, boolean Visible);
        NewBook (int NomSheets, String NameSheets[]); //Имена листов 3, { "One", "Two", "Three" };
        OpenSheet (WideString sheetname) {sheet = book.OlePropertyGet("WorkSheets", sheetname); sheet.OleProcedure("Activate");};
        AnsiString CellGet (int i, int j) {Variant cell = sheet.OlePropertyGet("Cells", i, j); return cell.OlePropertyGet("Value");};
        void CellSet (int i, int j, WideString textcell) {Variant cell = sheet.OlePropertyGet("Cells", i, j);
                                                          cell.OlePropertySet("NumberFormat","@"); cell.OlePropertySet("Value",textcell);};
        };

TExcelApp1::TExcelApp1(boolean Visible)
        {
        app = CreateOleObject("Excel.Application");
        app.OlePropertySet("Visible", Visible);
        books = app.OlePropertyGet("Workbooks");
        };

TExcelApp1::TExcelApp1(AnsiString filename, boolean Visible)
        {
        //TExcelApp1::TExcelApp1(Visible);
        app = CreateOleObject("Excel.Application");
        app.OlePropertySet("Visible", Visible);
        books = app.OlePropertyGet("Workbooks");
        books.Exec(Procedure("Open")<<filename);
        book = books.OlePropertyGet("item",1);
        };

TExcelApp1::NewBook(int NomSheets, String NameSheets[]) //Имена листов 3, { "One", "Two", "Three" };
        {
        app.OlePropertySet("SheetsInNewWorkbook", NomSheets);
        books.Exec(Procedure("Add"));
        book = books.OlePropertyGet("item",1);
        for (int i = 0; i < NomSheets; i++)
                { sheet= book.OlePropertyGet("WorkSheets",i+1);
                  sheet.OlePropertySet("Name", WideString(NameSheets[i]));
                }

        };