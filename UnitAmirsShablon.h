//---------------------------------------------------------------------------

#ifndef UnitAmirsShablonH
#define UnitAmirsShablonH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <DBCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TButton *Button1;
        TDateTimePicker *DateTimePicker1;
        TDBNavigator *DBNavigator1;
        TEdit *Edit1;
        TMemo *Memo1;
        TButton *Button2;
        TComboBox *ComboBox1;
        TButton *Button4;
        TComboBox *ComboBox2;
        TButton *Button5;
        void __fastcall FormActivate(TObject *Sender);
        void __fastcall Button4Click(TObject *Sender);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall Button3Click(TObject *Sender);
        void __fastcall DateTimePicker1Change(TObject *Sender);
        void __fastcall ComboBox2Change(TObject *Sender);
        void __fastcall Edit1KeyPress(TObject *Sender, char &Key);
        void __fastcall Button2Click(TObject *Sender);
        void __fastcall Button5Click(TObject *Sender);
        void __fastcall Edit1Click(TObject *Sender);
        void __fastcall Edit1Enter(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
        void __fastcall TForm1::OnDataChange(TObject* Sender, TField* Field);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
