#pragma once

namespace DatabaseConection {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Data::OleDb;//dodanie obs³ugi bazy danych OLedb

	/// <summary>
	/// Podsumowanie informacji o MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: W tym miejscu dodaj kod konstruktora
			//
		}

	protected:
		/// <summary>
		/// Wyczyœæ wszystkie u¿ywane zasoby.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Panel^ panel1;
	protected:
	private: System::Windows::Forms::Label^ label1;
	private: System::Windows::Forms::TextBox^ txtSearch;

	private: System::Windows::Forms::Button^ btnSearch;
	private: System::Windows::Forms::Panel^ panel2;
	private: System::Windows::Forms::Label^ label7;
	private: System::Windows::Forms::Label^ label6;
	private: System::Windows::Forms::Label^ label5;
	private: System::Windows::Forms::Label^ label4;
	private: System::Windows::Forms::Label^ label3;
	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::TextBox^ txtTelephone;
	private: System::Windows::Forms::TextBox^ txtStudentID;
	private: System::Windows::Forms::TextBox^ txtAddress;
	private: System::Windows::Forms::TextBox^ txtPostcode;









	private: System::Windows::Forms::TextBox^ txtSurname;
	private: System::Windows::Forms::TextBox^ txtFirstname;




	private: System::Windows::Forms::Panel^ panel3;
	private: System::Windows::Forms::Button^ btnExit;
	private: System::Windows::Forms::Button^ btnPrint;
	private: System::Windows::Forms::Button^ btnReset;
	private: System::Windows::Forms::Button^ btnDelete;
	private: System::Windows::Forms::Button^ btnUpdate;
	private: System::Windows::Forms::Button^ btnAddData;
	private: System::Windows::Forms::Label^ label8;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Drawing::Printing::PrintDocument^ printDocument1;
	private: System::Windows::Forms::PrintPreviewDialog^ printPreviewDialog1;

	private:
		/// <summary>
		/// Wymagana zmienna projektanta.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Metoda wymagana do obs³ugi projektanta — nie nale¿y modyfikowaæ
		/// jej zawartoœci w edytorze kodu.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm::typeid));
			this->panel1 = (gcnew System::Windows::Forms::Panel());
			this->label8 = (gcnew System::Windows::Forms::Label());
			this->txtSearch = (gcnew System::Windows::Forms::TextBox());
			this->btnSearch = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->panel2 = (gcnew System::Windows::Forms::Panel());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->label7 = (gcnew System::Windows::Forms::Label());
			this->label6 = (gcnew System::Windows::Forms::Label());
			this->label5 = (gcnew System::Windows::Forms::Label());
			this->label4 = (gcnew System::Windows::Forms::Label());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->txtTelephone = (gcnew System::Windows::Forms::TextBox());
			this->txtStudentID = (gcnew System::Windows::Forms::TextBox());
			this->txtAddress = (gcnew System::Windows::Forms::TextBox());
			this->txtPostcode = (gcnew System::Windows::Forms::TextBox());
			this->txtSurname = (gcnew System::Windows::Forms::TextBox());
			this->txtFirstname = (gcnew System::Windows::Forms::TextBox());
			this->panel3 = (gcnew System::Windows::Forms::Panel());
			this->btnExit = (gcnew System::Windows::Forms::Button());
			this->btnPrint = (gcnew System::Windows::Forms::Button());
			this->btnReset = (gcnew System::Windows::Forms::Button());
			this->btnDelete = (gcnew System::Windows::Forms::Button());
			this->btnUpdate = (gcnew System::Windows::Forms::Button());
			this->btnAddData = (gcnew System::Windows::Forms::Button());
			this->printDocument1 = (gcnew System::Drawing::Printing::PrintDocument());
			this->printPreviewDialog1 = (gcnew System::Windows::Forms::PrintPreviewDialog());
			this->panel1->SuspendLayout();
			this->panel2->SuspendLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->panel3->SuspendLayout();
			this->SuspendLayout();
			// 
			// panel1
			// 
			this->panel1->BackColor = System::Drawing::SystemColors::Control;
			this->panel1->Controls->Add(this->label8);
			this->panel1->Controls->Add(this->txtSearch);
			this->panel1->Controls->Add(this->btnSearch);
			this->panel1->Location = System::Drawing::Point(44, 69);
			this->panel1->Name = L"panel1";
			this->panel1->Size = System::Drawing::Size(1041, 73);
			this->panel1->TabIndex = 0;
			// 
			// label8
			// 
			this->label8->AutoSize = true;
			this->label8->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, static_cast<System::Drawing::FontStyle>((System::Drawing::FontStyle::Bold | System::Drawing::FontStyle::Underline)),
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(238)));
			this->label8->Location = System::Drawing::Point(33, 17);
			this->label8->Name = L"label8";
			this->label8->Size = System::Drawing::Size(219, 46);
			this->label8->TabIndex = 14;
			this->label8->Text = L"DB Details";
			// 
			// txtSearch
			// 
			this->txtSearch->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtSearch->ForeColor = System::Drawing::SystemColors::ScrollBar;
			this->txtSearch->Location = System::Drawing::Point(396, 7);
			this->txtSearch->Name = L"txtSearch";
			this->txtSearch->Size = System::Drawing::Size(332, 62);
			this->txtSearch->TabIndex = 6;
			this->txtSearch->Text = L"Search";
			this->txtSearch->MouseEnter += gcnew System::EventHandler(this, &MyForm::txtSearch_MouseEnter);
			this->txtSearch->MouseLeave += gcnew System::EventHandler(this, &MyForm::txtSearch_MouseLeave);
			// 
			// btnSearch
			// 
			this->btnSearch->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnSearch->Location = System::Drawing::Point(734, 4);
			this->btnSearch->Name = L"btnSearch";
			this->btnSearch->Size = System::Drawing::Size(277, 67);
			this->btnSearch->TabIndex = 5;
			this->btnSearch->Text = L"Search";
			this->btnSearch->UseVisualStyleBackColor = true;
			this->btnSearch->Click += gcnew System::EventHandler(this, &MyForm::btnSearch_Click);
			// 
			// label1
			// 
			this->label1->BackColor = System::Drawing::SystemColors::Control;
			this->label1->BorderStyle = System::Windows::Forms::BorderStyle::FixedSingle;
			this->label1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 39.75F, System::Drawing::FontStyle::Underline, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label1->Location = System::Drawing::Point(44, 3);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(1041, 63);
			this->label1->TabIndex = 0;
			this->label1->Text = L"C++ MSAcces Database Connection";
			this->label1->TextAlign = System::Drawing::ContentAlignment::TopCenter;
			// 
			// panel2
			// 
			this->panel2->BackColor = System::Drawing::SystemColors::Control;
			this->panel2->Controls->Add(this->dataGridView1);
			this->panel2->Controls->Add(this->label7);
			this->panel2->Controls->Add(this->label6);
			this->panel2->Controls->Add(this->label5);
			this->panel2->Controls->Add(this->label4);
			this->panel2->Controls->Add(this->label3);
			this->panel2->Controls->Add(this->label2);
			this->panel2->Controls->Add(this->txtTelephone);
			this->panel2->Controls->Add(this->txtStudentID);
			this->panel2->Controls->Add(this->txtAddress);
			this->panel2->Controls->Add(this->txtPostcode);
			this->panel2->Controls->Add(this->txtSurname);
			this->panel2->Controls->Add(this->txtFirstname);
			this->panel2->Location = System::Drawing::Point(44, 148);
			this->panel2->Name = L"panel2";
			this->panel2->Size = System::Drawing::Size(648, 589);
			this->panel2->TabIndex = 1;
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Location = System::Drawing::Point(3, 395);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(642, 191);
			this->dataGridView1->TabIndex = 19;
			this->dataGridView1->CellClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MyForm::dataGridView1_CellClick);
			// 
			// label7
			// 
			this->label7->AutoSize = true;
			this->label7->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label7->Location = System::Drawing::Point(33, 326);
			this->label7->Name = L"label7";
			this->label7->Size = System::Drawing::Size(217, 46);
			this->label7->TabIndex = 18;
			this->label7->Text = L"Telephone";
			// 
			// label6
			// 
			this->label6->AutoSize = true;
			this->label6->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label6->Location = System::Drawing::Point(33, 265);
			this->label6->Name = L"label6";
			this->label6->Size = System::Drawing::Size(217, 46);
			this->label6->TabIndex = 17;
			this->label6->Text = L"Post Code";
			// 
			// label5
			// 
			this->label5->AutoSize = true;
			this->label5->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label5->Location = System::Drawing::Point(33, 209);
			this->label5->Name = L"label5";
			this->label5->Size = System::Drawing::Size(150, 46);
			this->label5->TabIndex = 16;
			this->label5->Text = L"Adress";
			// 
			// label4
			// 
			this->label4->AutoSize = true;
			this->label4->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label4->Location = System::Drawing::Point(33, 145);
			this->label4->Name = L"label4";
			this->label4->Size = System::Drawing::Size(188, 46);
			this->label4->TabIndex = 15;
			this->label4->Text = L"Surname";
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label3->Location = System::Drawing::Point(33, 84);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(212, 46);
			this->label3->TabIndex = 14;
			this->label3->Text = L"FirstName";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 30, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->label2->Location = System::Drawing::Point(33, 25);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(217, 46);
			this->label2->TabIndex = 13;
			this->label2->Text = L"Student ID";
			// 
			// txtTelephone
			// 
			this->txtTelephone->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtTelephone->Location = System::Drawing::Point(313, 329);
			this->txtTelephone->Multiline = true;
			this->txtTelephone->Name = L"txtTelephone";
			this->txtTelephone->Size = System::Drawing::Size(321, 43);
			this->txtTelephone->TabIndex = 12;
			// 
			// txtStudentID
			// 
			this->txtStudentID->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtStudentID->Location = System::Drawing::Point(313, 25);
			this->txtStudentID->Multiline = true;
			this->txtStudentID->Name = L"txtStudentID";
			this->txtStudentID->Size = System::Drawing::Size(321, 43);
			this->txtStudentID->TabIndex = 7;
			// 
			// txtAddress
			// 
			this->txtAddress->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtAddress->Location = System::Drawing::Point(313, 212);
			this->txtAddress->Multiline = true;
			this->txtAddress->Name = L"txtAddress";
			this->txtAddress->Size = System::Drawing::Size(321, 43);
			this->txtAddress->TabIndex = 10;
			// 
			// txtPostcode
			// 
			this->txtPostcode->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtPostcode->Location = System::Drawing::Point(313, 268);
			this->txtPostcode->Multiline = true;
			this->txtPostcode->Name = L"txtPostcode";
			this->txtPostcode->Size = System::Drawing::Size(321, 43);
			this->txtPostcode->TabIndex = 11;
			// 
			// txtSurname
			// 
			this->txtSurname->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtSurname->Location = System::Drawing::Point(313, 148);
			this->txtSurname->Multiline = true;
			this->txtSurname->Name = L"txtSurname";
			this->txtSurname->Size = System::Drawing::Size(321, 43);
			this->txtSurname->TabIndex = 9;
			// 
			// txtFirstname
			// 
			this->txtFirstname->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 27.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->txtFirstname->Location = System::Drawing::Point(313, 87);
			this->txtFirstname->Multiline = true;
			this->txtFirstname->Name = L"txtFirstname";
			this->txtFirstname->Size = System::Drawing::Size(321, 43);
			this->txtFirstname->TabIndex = 8;
			// 
			// panel3
			// 
			this->panel3->BackColor = System::Drawing::SystemColors::Control;
			this->panel3->Controls->Add(this->btnExit);
			this->panel3->Controls->Add(this->btnPrint);
			this->panel3->Controls->Add(this->btnReset);
			this->panel3->Controls->Add(this->btnDelete);
			this->panel3->Controls->Add(this->btnUpdate);
			this->panel3->Controls->Add(this->btnAddData);
			this->panel3->Location = System::Drawing::Point(698, 148);
			this->panel3->Name = L"panel3";
			this->panel3->Size = System::Drawing::Size(387, 589);
			this->panel3->TabIndex = 2;
			// 
			// btnExit
			// 
			this->btnExit->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnExit->Location = System::Drawing::Point(27, 459);
			this->btnExit->Name = L"btnExit";
			this->btnExit->Size = System::Drawing::Size(330, 77);
			this->btnExit->TabIndex = 9;
			this->btnExit->Text = L"Exit";
			this->btnExit->UseVisualStyleBackColor = true;
			this->btnExit->Click += gcnew System::EventHandler(this, &MyForm::btnExit_Click);
			// 
			// btnPrint
			// 
			this->btnPrint->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnPrint->Location = System::Drawing::Point(27, 376);
			this->btnPrint->Name = L"btnPrint";
			this->btnPrint->Size = System::Drawing::Size(330, 77);
			this->btnPrint->TabIndex = 8;
			this->btnPrint->Text = L"Print";
			this->btnPrint->UseVisualStyleBackColor = true;
			this->btnPrint->Click += gcnew System::EventHandler(this, &MyForm::btnPrint_Click);
			// 
			// btnReset
			// 
			this->btnReset->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnReset->Location = System::Drawing::Point(27, 293);
			this->btnReset->Name = L"btnReset";
			this->btnReset->Size = System::Drawing::Size(330, 77);
			this->btnReset->TabIndex = 7;
			this->btnReset->Text = L"Reset";
			this->btnReset->UseVisualStyleBackColor = true;
			this->btnReset->Click += gcnew System::EventHandler(this, &MyForm::btnReset_Click);
			// 
			// btnDelete
			// 
			this->btnDelete->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnDelete->Location = System::Drawing::Point(27, 210);
			this->btnDelete->Name = L"btnDelete";
			this->btnDelete->Size = System::Drawing::Size(330, 77);
			this->btnDelete->TabIndex = 6;
			this->btnDelete->Text = L"Delete";
			this->btnDelete->UseVisualStyleBackColor = true;
			this->btnDelete->Click += gcnew System::EventHandler(this, &MyForm::btnDelete_Click);
			// 
			// btnUpdate
			// 
			this->btnUpdate->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnUpdate->Location = System::Drawing::Point(27, 127);
			this->btnUpdate->Name = L"btnUpdate";
			this->btnUpdate->Size = System::Drawing::Size(330, 77);
			this->btnUpdate->TabIndex = 5;
			this->btnUpdate->Text = L"Update";
			this->btnUpdate->UseVisualStyleBackColor = true;
			this->btnUpdate->Click += gcnew System::EventHandler(this, &MyForm::btnUpdate_Click);
			// 
			// btnAddData
			// 
			this->btnAddData->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 36, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(238)));
			this->btnAddData->Location = System::Drawing::Point(27, 44);
			this->btnAddData->Name = L"btnAddData";
			this->btnAddData->Size = System::Drawing::Size(330, 77);
			this->btnAddData->TabIndex = 4;
			this->btnAddData->Text = L"Add Data";
			this->btnAddData->UseVisualStyleBackColor = true;
			this->btnAddData->Click += gcnew System::EventHandler(this, &MyForm::btnAddData_Click);
			// 
			// printDocument1
			// 
			this->printDocument1->PrintPage += gcnew System::Drawing::Printing::PrintPageEventHandler(this, &MyForm::printDocument1_PrintPage);
			// 
			// printPreviewDialog1
			// 
			this->printPreviewDialog1->AutoScrollMargin = System::Drawing::Size(0, 0);
			this->printPreviewDialog1->AutoScrollMinSize = System::Drawing::Size(0, 0);
			this->printPreviewDialog1->ClientSize = System::Drawing::Size(400, 300);
			this->printPreviewDialog1->Document = this->printDocument1;
			this->printPreviewDialog1->Enabled = true;
			this->printPreviewDialog1->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"printPreviewDialog1.Icon")));
			this->printPreviewDialog1->Name = L"printPreviewDialog1";
			this->printPreviewDialog1->Visible = false;
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::LightBlue;
			this->ClientSize = System::Drawing::Size(1134, 749);
			this->Controls->Add(this->panel3);
			this->Controls->Add(this->panel2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->panel1);
			this->Name = L"MyForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			this->panel1->ResumeLayout(false);
			this->panel1->PerformLayout();
			this->panel2->ResumeLayout(false);
			this->panel2->PerformLayout();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->panel3->ResumeLayout(false);
			this->ResumeLayout(false);

		}
#pragma endregion

		OleDbConnection^ conn = gcnew OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\portfolio\\DatabaseConection\\csave.accdb");
		//gcnew means it collects the garbage of the object created with it, important to use with CLR types, you dont have to delete things created with it
		//^ means pointer
		int checker;
		Bitmap^ bitmap;

		void ConnectionDB() {// po³¹czenie z baz¹ danych
			try
			{
				conn->Open();								//otwiera bazê danych
				OleDbCommand^ cmd = conn->CreateCommand();	//tworzy komendê która wykonuje polecenie SQL
				cmd->CommandType = CommandType::Text;
				cmd->CommandText = "SELECT * FROM csave";
				cmd->ExecuteNonQuery();						//Wykonujê t¹ komendê i wypisuje iloœc rzêdów które zosta³y zmienione

				DataTable^ dt = gcnew DataTable();
				OleDbDataAdapter^ dp = gcnew OleDbDataAdapter(cmd);	//przechowuje wyniki tej operacji
				dp->Fill(dt);
				dataGridView1->DataSource = dt;				//przypisanie danych do pola w formularzu
				conn->Close();								//zamyka bazê danych
				//MessageBox::Show("Connection is Succesful", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);

			}
			catch (Exception^ ex)
			{
				MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
			}
		}

	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
		try {
			ConnectionDB();
		}
		catch(Exception^ ex){
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void btnAddData_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			conn->Open();								
			OleDbCommand^ cmd = conn->CreateCommand();	
			cmd->CommandType = CommandType::Text;
			cmd->CommandText = "INSERT INTO csave (StudentID,Firstname,Surname,Address,Postcode,Telephone)values(?,?,?,?,?,?)";
			cmd->Parameters->AddWithValue("@StudentID", txtStudentID->Text);		//dodaje dane z pola txtStudentID w miejsce 1 zapytajnika
			cmd->Parameters->AddWithValue("@Firstname", txtFirstname->Text);
			cmd->Parameters->AddWithValue("@Surname", txtSurname->Text);
			cmd->Parameters->AddWithValue("@Address", txtAddress->Text);
			cmd->Parameters->AddWithValue("@Postcode", txtPostcode->Text);
			cmd->Parameters->AddWithValue("@Telephone", txtTelephone->Text);

			cmd->ExecuteNonQuery();						
			conn->Close();								
			MessageBox::Show("Record is Successfully Added", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);
			ConnectionDB();

		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void btnExit_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			System::Windows::Forms::DialogResult iClose;
			iClose = MessageBox::Show("Confirm if you want to exit", "C++ AccessDatabase Connector", MessageBoxButtons::YesNo, MessageBoxIcon::Question);

			if (iClose == System::Windows::Forms::DialogResult::Yes) {
				Close();
			}
		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void btnReset_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			System::Windows::Forms::DialogResult iReset;
			iReset = MessageBox::Show("Confirm if you want to reset", "C++ AccessDatabase Connector", MessageBoxButtons::YesNo, MessageBoxIcon::Question);

			if (iReset == System::Windows::Forms::DialogResult::Yes) {
				txtStudentID->Text = "";
				txtFirstname->Text = "";
				txtSurname->Text = "";
				txtAddress->Text = "";
				txtPostcode->Text = "";
				txtTelephone->Text = "";
				txtSearch->Text = "";
				MessageBox::Show("You have Successfully Reset your System", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);
				ConnectionDB();
			}
		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void btnSearch_Click(System::Object^ sender, System::EventArgs^ e) {
		checker = 0;
		try
		{
			conn->Open();
			OleDbCommand^ cmd = conn->CreateCommand();
			cmd->CommandType = CommandType::Text;
			cmd->CommandText = "SELECT * FROM csave WHERE StudentID LIKE '%' + ? + '%' OR Firstname LIKE '%' + ? + '%' OR Surname LIKE '%' + ? + '%'";
			cmd->Parameters->AddWithValue("@StudentID", txtSearch->Text);
			cmd->Parameters->AddWithValue("@Firstname", txtSearch->Text);
			cmd->Parameters->AddWithValue("@Surname", txtSearch->Text);

			cmd->ExecuteNonQuery();
			DataTable^ dt = gcnew DataTable();
			OleDbDataAdapter^ dp = gcnew OleDbDataAdapter(cmd);
			dp->Fill(dt);
			checker = Convert::ToInt32(dt->Rows->Count.ToString());
			dataGridView1->DataSource = dt;
			conn->Close();

			if (checker == 0) {
				MessageBox::Show("No record found", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);
			}
		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
			conn->Close();
		}
	}
	private: System::Void btnUpdate_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			conn->Open();
			OleDbCommand^ cmd = conn->CreateCommand();
			cmd->CommandType = CommandType::Text;
			cmd->CommandText = "UPDATE csave SET StudentID = ? WHERE Firstname = ? AND Surname = ?";
			cmd->Parameters->AddWithValue("@StudentID", txtStudentID->Text);
			cmd->Parameters->AddWithValue("@Firstname", txtFirstname->Text);
			cmd->Parameters->AddWithValue("@Surname", txtSurname->Text);
			cmd->ExecuteNonQuery();

			conn->Close();
			MessageBox::Show("Record Updated Successfully", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);
			ConnectionDB();
		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
			conn->Close();
		}
	}
	private: System::Void dataGridView1_CellClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
		try {
			txtStudentID->Text = dataGridView1->SelectedRows[0]->Cells[0]->Value->ToString();	//wybieramy 1 wartoœæ z 1 rzedu i wpisujemy j¹ do 1 pola dostêpnego w formularzu
			txtFirstname->Text = dataGridView1->SelectedRows[0]->Cells[1]->Value->ToString();
			txtSurname->Text = dataGridView1->SelectedRows[0]->Cells[2]->Value->ToString();
			txtAddress->Text = dataGridView1->SelectedRows[0]->Cells[3]->Value->ToString();
			txtPostcode->Text = dataGridView1->SelectedRows[0]->Cells[4]->Value->ToString();
			txtTelephone->Text = dataGridView1->SelectedRows[0]->Cells[5]->Value->ToString();
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
			conn->Close();
		}
	}
	private: System::Void btnDelete_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			conn->Open();
			OleDbCommand^ cmd = conn->CreateCommand();
			cmd->CommandType = CommandType::Text;
			cmd->CommandText = "DELETE * FROM csave WHERE StudentID = ?";
			cmd->Parameters->AddWithValue("@StudentID", txtStudentID->Text);		

			cmd->ExecuteNonQuery();
			conn->Close();
			MessageBox::Show("Record is Successfully Deleted", "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Information);
			ConnectionDB();
			txtStudentID->Text = "";
			txtFirstname->Text = "";
			txtSurname->Text = "";
			txtAddress->Text = "";
			txtPostcode->Text = "";
			txtTelephone->Text = "";
			txtSearch->Text = "";

		}
		catch (Exception^ ex)
		{
			MessageBox::Show(ex->Message, "C++ AccessDatabase Connector", MessageBoxButtons::OK, MessageBoxIcon::Error);
			conn->Close();
		}
	}
	private: System::Void btnPrint_Click(System::Object^ sender, System::EventArgs^ e) {
		try {
			int height = dataGridView1->Height;
			dataGridView1->Height = dataGridView1->RowCount * dataGridView1->RowTemplate->Height * 2;	//ustawia wyslokoœc tabeli na podstawie iloœci kolumn i roz³o¿enia rzêdów
			bitmap = gcnew Bitmap(dataGridView1->Width, dataGridView1->Height);							//tworzy nowy obiekt typu bitmap o wysokoœci i szerokoœci tabelki
			dataGridView1->DrawToBitmap(bitmap, System::Drawing::Rectangle(0, 0, dataGridView1->Width, dataGridView1->Height));//rysuje tabelkê w bitmapie
			printPreviewDialog1->PrintPreviewControl->Zoom = 1;
			printPreviewDialog1->ShowDialog();
			dataGridView1->Height = height;
		}
		catch (Exception^ ex) {
			MessageBox::Show(ex->Message);
		}

	}
	private: System::Void printDocument1_PrintPage(System::Object^ sender, System::Drawing::Printing::PrintPageEventArgs^ e) {

		e->Graphics->DrawImage(bitmap, 0, 0);
	}

	private: System::Void txtSearch_MouseEnter(System::Object^ sender, System::EventArgs^ e) {
		txtSearch->Text = "";
		txtSearch->Focus();
	}
	private: System::Void txtSearch_MouseLeave(System::Object^ sender, System::EventArgs^ e) {
		if (txtSearch->Text == "") {
			txtSearch->Text = "Search";
		}
	}
};
}
