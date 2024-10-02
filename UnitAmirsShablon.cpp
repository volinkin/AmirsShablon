//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN //SHGetSpecialFolderPath
#include <vcl.h>
#pragma hdrstop

#include "UnitAmirsShablon.h"
#include "UnitAmirs.h"

/*
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <DB.hpp>
#include <IBDatabase.hpp>
#include <IBQuery.hpp>
#include <IBTable.hpp>
#include <ComCtrls.hpp>
#include "CSPIN.h"
#include <DBCtrls.hpp>
#include <DBGrids.hpp>
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include "Excel_2K_SRVR.h"
#include "Word_2K_SRVR.h"
#include <OleServer.hpp>
#include <IBCustomDataSet.hpp>    */
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
        {
        }



//void __fastcall (__closure *FunctionOnDataChange)(TObject* Sender, TField* Field);

class TAmirsShablon :public TAmirs
        {
        //TComboBox * ComboBox1;
        //TComboBox * ComboBox2;

        TMemo * Memo1;
        public:
        int x;
        TAmirsShablon(TForm * Form, AnsiString NameApp, TMemo * Memo, TComboBox * ComboBox1, TDateTimePicker *DateTimePicker1, TComboBox * ComboBox2, TDBNavigator * DBNavigator1);
        void Refresh();
        };

TAmirsShablon * AmirsShablon;

void __fastcall TForm1::OnDataChange(TObject* Sender, TField* Field)
        {
         //MessageBox(NULL, "7", "OndataChange", MB_OK);
         if (!AmirsShablon->IBQuery->Fields->FieldByName("NUM")->IsNull)
         Edit1->Text = AmirsShablon->IBQuery->Fields->FieldByName("NUM")->Value; else Edit1->Text ="";
         AmirsShablon->GetTextToWord();
         //Memo1->Clear();
         AmirsShablon->TextToMemo(Memo1);
         //Memo1->Perform(WM_VScroll,SB_TOP,0);
         //memo2.ScrollBars:=ssVertical;
         Memo1->SelStart = 0;         Memo1->SelLength = 0;         //memo2.ScrollBars:=ssnone;

        };


TAmirsShablon::TAmirsShablon(TForm * Form1, AnsiString NameApp, TMemo * Memo, TComboBox * ComboBox1, TDateTimePicker *DateTimePicker1, TComboBox * ComboBox2, TDBNavigator * DBNavigator1): TAmirs(Form1, NameApp)
        {
        ComboBox1->Items->Add("Шаблоны");
        ComboBox1->Items->Add("Пакеты");
        ComboBox1->Items->Add("Статистика");
        ComboBox1->Items->Add("Выборки");
        ComboBox1->ItemIndex=0;
        ComboBox2->Items->Add("Уголовные дела");
        ComboBox2->Items->Add("Гражданские дела");
        ComboBox2->Items->Add("Дела об админ. прав.");
        ComboBox2->Items->Add("Исп уголовных дел");
        ComboBox2->Items->Add("Исп гражданских дел");
        ComboBox2->Items->Add("КАС");
        ComboBox2->Items->Add("Адм дела Штраф");
        ComboBox2->Items->Add("Адм дела Обяз работы");
        ComboBox2->Items->Add("Гражданские дела.Иски");
        ComboBox2->ItemIndex=0;
        this->DataSource->DataSet=this->IBQuery;
        DBNavigator1->DataSource=this->DataSource;
        //Начальная дата
        Word year, month, day;
        DateTimePicker1->Date.DecodeDate(&year,&month,&day);
        this->Year=year;
        this->VidDela=1;
        //MessageBox(NULL, "Init1", "Amirs", MB_OK | MB_ICONERROR);
        this->Connect();
        this->GetDela();
        this->Memo1=Memo;
        this->Form1->Caption="АМИРС Шаблоны. Профиль: "+ProfileFile;
        //AmirsShablon->ErrorsInMemo(Memo);
        };

void TAmirsShablon::Refresh()
        {
        //Сюда собрать всю информацию с кантролов
        //Year=2022;
        //VidDela=1;
        Connect();
        GetDela();
        this->Form1->Caption="АМИРС Шаблоны. Профиль: "+ProfileFile;
        };


void __fastcall TForm1::FormActivate(TObject *Sender)
{
AmirsShablon = new TAmirsShablon(Form1, "AmirsShablon", Memo1, ComboBox1, DateTimePicker1, ComboBox2, DBNavigator1);
AmirsShablon->DataSource->OnDataChange=&OnDataChange;
AmirsShablon->Search(1);
AmirsShablon->ErrorsInMemo(Memo1);
}
//---------------------------------------------------------------------------


//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{
AmirsShablon->Refresh();
AmirsShablon->ErrorsInMemo(Memo1);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
Edit1->Text="";
this->Memo1->Clear();
Form1->Refresh();
AmirsShablon->ProfileSelect();
AmirsShablon->Refresh();
AmirsShablon->ErrorsInMemo(Memo1);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button3Click(TObject *Sender)
{
//AnsiString PathDocument; char Buff[MAX_PATH];
//        SHGetSpecialFolderPath(0, Buff, CSIDL_PERSONAL, 0);
//        PathDocument=Buff;
//AmirsShablon->ExcelIndex(PathDocument+"\\"+"qwerty.xlsx");

/*String **matrixUD; int cols; int rows;
String **matrixGD;
String **matrixAD;
String **matrixIspUD;
String **matrixIspGD; int colsIspGD; int rowsIspGD;
TOpenDialog *OpenDialog1 = new TOpenDialog(Form1);
OpenDialog1->InitialDir= GetSpecialFolderPath(CSIDL_PERSONAL);
if (OpenDialog1->Execute())
        {TExcelApp1 excel1(OpenDialog1->FileName, true);
        excel1.OpenSheet("Уголовные дела");
        cols=0; rows=0;
        while (!excel1.CellGet(rows+1, 1).IsEmpty()) rows++;
        while (!excel1.CellGet(1, cols+1).IsEmpty()) cols++;
        matrixUD = new String * [rows];
        for (int i = 0; i < rows; i++) matrixUD[i] = new String [cols];
        for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) matrixUD[i][j]=excel1.CellGet(i+1, j+1);

        excel1.OpenSheet("Исполнение по ГД");
        colsIspGD=0; rowsIspGD=0;
        while (!excel1.CellGet(rowsIspGD+1, 1).IsEmpty()) rowsIspGD++;
        while (!excel1.CellGet(1, colsIspGD+1).IsEmpty()) colsIspGD++;
        matrixIspGD = new String * [rowsIspGD];
        for (int i = 0; i < rowsIspGD; i++) matrixIspGD[i] = new String [colsIspGD];
        for (int i = 0; i < rowsIspGD; i++) for (int j = 0; j < colsIspGD; j++) matrixIspGD[i][j]=excel1.CellGet(i+1, j+1);
        };
        */
AmirsShablon->Proverka ();
/*
TOpenDialog *OpenDialog1 = new TOpenDialog(Form1);
OpenDialog1->InitialDir= GetSpecialFolderPath(CSIDL_PERSONAL);
if (OpenDialog1->Execute())
        {TExcelApp1 excel1(OpenDialog1->FileName, true);
        excel1.OpenSheet("Уголовные дела");
        AmirsShablon->validator[0].open(& excel1);
         excel1.OpenSheet("Гражданские дела");
        AmirsShablon->validator[1].open(& excel1);
        excel1.OpenSheet("Исполнение по ГД");
        AmirsShablon->validator[4].open(& excel1);
        excel1.Quit();
        };



 TExcelApp1 excel2(true);
 String s[]= {"Уголовные дела", "Гражданские дела", "Дела об АП", "Исполнение по УД", "Исполнение по ГД", "Проверка", "Дельта"};
 excel2.NewBook(7, s);

 //Уголовные дела
 excel2.OpenSheet("Уголовные дела");
// excel2.CellSet(1, 1, "Номер"); excel2.CellSet(2, 1, "Jgbcm");
for (int i = 0; i < AmirsShablon->validator[0].cols; i++) excel2.CellSet(1, i+1, AmirsShablon->validator[0].matrix[0][i]);
 AmirsShablon->GetDelaZapros("zaprosUD");
 AmirsShablon->IBQuery1->First();
 for (int i = 0; i < AmirsShablon->IBQuery1->RecordCount; i++)
                {excel2.CellSet(i+2, 1, AmirsShablon->GetQuery1TextField("NUM"));
                int j= AmirsShablon->validator[0].poisk(AmirsShablon->GetQuery1TextField("NUM"));
                for (int ii = 1; ii < AmirsShablon->validator[0].cols; ii++) excel2.CellSet(j+1, ii+1, AmirsShablon->validator[0].matrix[j][ii]);
                AmirsShablon->IBQuery1->Next();
                };

 //Гражданские дела
 excel2.OpenSheet("Гражданские дела");
for (int i = 0; i < AmirsShablon->validator[1].cols; i++) excel2.CellSet(1, i+1, AmirsShablon->validator[1].matrix[0][i]);
 AmirsShablon->GetDelaZapros("zaprosGDIsk");
 AmirsShablon->IBQuery1->First();
 for (int i = 0; i < AmirsShablon->IBQuery1->RecordCount; i++)
                {excel2.CellSet(i+2, 1, AmirsShablon->GetQuery1TextField("NUM"));
                int j= AmirsShablon->validator[1].poisk(AmirsShablon->GetQuery1TextField("NUM"));
                for (int ii = 1; ii < AmirsShablon->validator[1].cols; ii++) excel2.CellSet(j+1, ii+1, AmirsShablon->validator[1].matrix[j][ii]);
                AmirsShablon->IBQuery1->Next();
                };

  //Исполнение по ГД
/* excel2.OpenSheet("Исполнение по ГД");
// excel2.CellSet(1, 1, "Номер"); excel2.CellSet(2, 1, "Jgbcm");
 for (int i = 0; i < colsIspGD; i++) excel2.CellSet(1, i+1, matrixIspGD[0][i]);
 AmirsShablon->GetDelaZapros("zaprosIG");
 AmirsShablon->IBQuery1->First();
 for (int i = 0; i < AmirsShablon->IBQuery1->RecordCount; i++)
                {excel2.CellSet(i+2, 1, AmirsShablon->GetQuery1TextField("NUM"));
                int j=0; boolean poisk=true; while (j<rowsIspGD & poisk) if (matrixIspGD[j][0]==AmirsShablon->GetQuery1TextField("NUM")) poisk=false; else j++;
                if (j< rowsIspGD) for (int ii = 1; ii < colsIspGD; ii++) excel2.CellSet(j+1, ii+1, matrixIspGD[j][ii]);
                AmirsShablon->IBQuery1->Next();
                };
  */
 /*
 excel2.OpenSheet("Исполнение по ГД");
for (int i = 0; i < AmirsShablon->validator[4].cols; i++) excel2.CellSet(1, i+1, AmirsShablon->validator[4].matrix[0][i]);
 AmirsShablon->GetDelaZapros("zaprosIG");
 AmirsShablon->IBQuery1->First();
 for (int i = 0; i < AmirsShablon->IBQuery1->RecordCount; i++)
                {excel2.CellSet(i+2, 1, AmirsShablon->GetQuery1TextField("NUM"));
                int j= AmirsShablon->validator[4].poisk(AmirsShablon->GetQuery1TextField("NUM"));
                for (int ii = 1; ii < AmirsShablon->validator[4].cols; ii++) excel2.CellSet(j+1, ii+1, AmirsShablon->validator[4].matrix[j][ii]);
                AmirsShablon->IBQuery1->Next();
                };
 */
}
//---------------------------------------------------------------------------

void __fastcall TForm1::DateTimePicker1Change(TObject *Sender)
{
Edit1->Text="";
this->Memo1->Clear();
Form1->Refresh();
Word year, month, day;
DateTimePicker1->Date.DecodeDate(&year,&month,&day);
AmirsShablon->Year=year;
AmirsShablon->Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ComboBox2Change(TObject *Sender)
{
Edit1->Text="";
this->Memo1->Clear();
Form1->Refresh();
AmirsShablon->VidDela=ComboBox2->ItemIndex+1;
AmirsShablon->Refresh();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit1KeyPress(TObject *Sender, char &Key)
{
if (Key == 13)
        {AmirsShablon->Search(Edit1->Text);
        };
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
 AmirsShablon->TextToWord();
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button5Click(TObject *Sender)
{
TOpenDialog *OpenDialog1 =new TOpenDialog(Form1);
OpenDialog1->InitialDir= ExtractFilePath(Application->ExeName);
OpenDialog1->Execute();
AmirsShablon->FileNameShablon=OpenDialog1->FileName;
AmirsShablon->TextToMemo(Memo1);
//Label1->Caption=FileNameShablon;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit1Click(TObject *Sender)
{
 //MessageBox(NULL, "", "Click", MB_OK | MB_ICONERROR);
 Edit1->Text = "";
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Edit1Enter(TObject *Sender)
{
// MessageBox(NULL, "", "Enter", MB_OK | MB_ICONERROR);
 Edit1->Text = "";
}
//---------------------------------------------------------------------------

