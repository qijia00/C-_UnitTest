/*******************************************************************************
**
**  Copyright (c) 1999-2012 VisualSonics Inc. All Rights Reserved.
**
**  VsiTestTxFUtils.cpp
**
**	Description:
**		VsiTestTxFUtils Test
**
********************************************************************************/
#include "stdafx.h"
#include <Fcntl.h>
#include <sys/types.h>
#include <sys/stat.h>
#include <share.h>
#include <direct.h>
#include <VsiResProduct.h>
#include <VsiCommUtl.h>
#include <VsiCommUtlLib.h>

using namespace System;
using namespace System::Text;
using namespace System::Collections::Generic;
using namespace	Microsoft::VisualStudio::TestTools::UnitTesting;
using namespace System::IO;


namespace VsiCommUtlTest
{
	[TestClass]
	public ref class VsiTestTxfUtils
	{
	private:
		TestContext^ testContextInstance;
		ULONG_PTR m_cookie;

	public:
		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		property Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ TestContext
		{
			Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ get()
			{
				return testContextInstance;
			}
			System::Void set(Microsoft::VisualStudio::TestTools::UnitTesting::TestContext^ value)
			{
				testContextInstance = value;
			}
		};

		#pragma region Additional test attributes
		//
		//You can use the following additional attributes as you write your tests:
		//
		//Use ClassInitialize to run code before running the first test in the class
		//[ClassInitialize()]
		//static void MyClassInitialize(TestContext^ testContext) {};
		//
		//Use ClassCleanup to run code after all tests in a class have run
		//[ClassCleanup()]
		//static void MyClassCleanup() {};
		//
		//Use TestInitialize to run code before running each test
		[TestInitialize()]
		void MyTestInitialize()
		{
			m_cookie = 0;

			String^ strDir = testContextInstance->TestDeploymentDir;
			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strDir);

			ACTCTX actCtx;
			memset((void*)&actCtx, 0, sizeof(ACTCTX));
			actCtx.cbSize = sizeof(ACTCTX);
			actCtx.dwFlags = ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID | ACTCTX_FLAG_RESOURCE_NAME_VALID;
			actCtx.lpSource = L"VsiCommUtlTest.dll";
			actCtx.lpAssemblyDirectory = pszTestDir;
			actCtx.lpResourceName = MAKEINTRESOURCE(2);

			HANDLE hCtx = CreateActCtx(&actCtx);
			Assert::IsTrue(INVALID_HANDLE_VALUE != hCtx, "CreateActCtx failed.");

			ULONG_PTR cookie;
			BOOL bRet = ActivateActCtx(hCtx, &cookie);
			Assert::IsTrue(FALSE != bRet, "ActivateActCtx failed.");

			m_cookie = cookie;
		};
		//
		//Use TestCleanup to run code after each test has run
		[TestCleanup()]
		void MyTestCleanup()
		{
			if (0 != m_cookie)
				DeactivateActCtx(0, m_cookie);
		};
		#pragma endregion 

	public: 
		[TestMethod]
		void TestVsiCreateAndCloseTransaction()
		{
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			BOOL bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");
		}

		[TestMethod]
		void TestVsiMoveFileTransacted()
		{
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			String^ strDir = testContextInstance->TestDeploymentDir;

			// Create file
			String^ strTestFileSrc = gcnew String(strDir);
			strTestFileSrc += "\\VsiCommUtl.dll";

			String^ strTestFileSrcCopy = gcnew String(strDir);
			strTestFileSrcCopy += "\\VsiTestTransactedSource"; //create a copy to protect source

			File::Copy(strTestFileSrc, strTestFileSrcCopy);

			String^ strTestFileDest = gcnew String(strDir);
			strTestFileDest += "\\VsiTestTransactedDest";

			pin_ptr<const wchar_t> pszTestFileSrcCopy = PtrToStringChars(strTestFileSrcCopy);
			pin_ptr<const wchar_t> pszTestFileDest = PtrToStringChars(strTestFileDest);

			BOOL bRet = VsiTxF::VsiMoveFileTransacted(
				pszTestFileSrcCopy,	// Source
				pszTestFileDest,	// Destination
				NULL,	// No progress bar
				NULL,	// No progress bar
				MOVEFILE_COPY_ALLOWED,	// can also specify MOVEFILE_REPLACE_EXISTING here
				hTransaction);
			Assert::IsTrue(TRUE == bRet, "MoveFileTransacted failed"); // destination file doesn't exit, so do not need enable replace

			Assert::IsTrue(File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(!File::Exists(strTestFileDest), "MoveFileTransacted failed"); 

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(File::Exists(strTestFileDest), "MoveFileTransacted failed");
			
			bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Clean up
			File::Delete(strTestFileDest);
			
			// Now try with replace existing			
			// Create file A.txt (content: "A") and B.txt (content: "B")
			// Move file A to file B with replace, commit the move
			// Check the content of file B - should be "A"
			HANDLE hTransactionReplace = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionReplace, "Create transaction failed");

			strDir = testContextInstance->TestDeploymentDir;

			// Create file A.txt (content: "A")
			String^ strDirA = gcnew String(strDir);
			strDirA += "\\A.txt";
			FileStream^ fs = File::Create(strDirA);
			Assert::IsTrue(File::Exists(strDirA), "Create A.txt failed.");
			fs->Close();

			fs = File::OpenWrite(strDirA);
			array<Byte>^info = (gcnew UTF8Encoding( true ))->GetBytes( "A" );
			fs->Write( info, 0, info->Length );
			fs->Close();

			// Create file B.txt (content: "B")
			String^ strDirB = gcnew String(strDir);
			strDirB += "\\B.txt";
			fs = File::Create(strDirB);
			Assert::IsTrue(File::Exists(strDirB), "Create B.txt failed.");
			fs->Close();

			fs = File::OpenWrite(strDirB);
			info = (gcnew UTF8Encoding( true ))->GetBytes( "B" );
			fs->Write( info, 0, info->Length );
			fs->Close();

			// Move file A to file B with replace, commit the move
			pin_ptr<const wchar_t> pszDirA = PtrToStringChars(strDirA);
			pin_ptr<const wchar_t> pszDirB = PtrToStringChars(strDirB);

			bRet = VsiTxF::VsiMoveFileTransacted(
				pszDirA,	// Source
				pszDirB,	// Destination
				NULL,	// No progress bar
				NULL,	// No progress bar
				MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING,	// MOVEFILE_REPLACE_EXISTING,
				hTransactionReplace);
			Assert::IsTrue(TRUE == bRet, "MoveFileTransacted failed"); //if disable replace, bRet returns FALSE

			Assert::IsTrue(File::Exists(strDirA), "MoveFileTransacted failed");
			Assert::IsTrue(File::Exists(strDirB), "MoveFileTransacted failed");

			// Commit Move
			bRet = VsiTxF::VsiCommitTransaction(hTransactionReplace);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!File::Exists(strDirA), "MoveFileTransacted failed");
			Assert::IsTrue(File::Exists(strDirB), "MoveFileTransacted failed");

			// Check the content of file B - should be "A"
			fs = File::OpenRead(strDirB);
			array<Byte>^b = gcnew array<Byte>(1024);
			UTF8Encoding^ temp = gcnew UTF8Encoding( true );
			while ( fs->Read( b, 0, b->Length ) > 0 )
			{
				String^ strContentB = temp->GetString( b );
				pin_ptr<const wchar_t> pszContentB = PtrToStringChars(strContentB);
				Assert::IsTrue(0 == wcscmp(L"A", pszContentB), "MoveFileTransacted failed replacing");
			}
			fs->Close();

			bRet = VsiTxF::VsiCloseTransaction(hTransactionReplace);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Clean up
			File::Delete(strDirB);
		}

		[TestMethod]
		void TestVsiCopyFileTransacted()
		{
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			String^ strDir = testContextInstance->TestDeploymentDir;

			// Create file
			String^ strTestFileSrc = gcnew String(strDir);
			strTestFileSrc += "\\VsiCommUtl.dll";

			String^ strTestFileSrcCopy = gcnew String(strDir);
			strTestFileSrcCopy += "\\VsiTestTransactedSource"; //create a copy to protect source

			File::Copy(strTestFileSrc, strTestFileSrcCopy);

			String^ strTestFileDest = gcnew String(strDir);
			strTestFileDest += "\\VsiTestTransactedDest";

			pin_ptr<const wchar_t> pszTestFileSrcCopy = PtrToStringChars(strTestFileSrcCopy);
			pin_ptr<const wchar_t> pszTestFileDest = PtrToStringChars(strTestFileDest);

			BOOL bRet = VsiTxF::VsiCopyFileTransacted(
				pszTestFileSrcCopy,	// Source
				pszTestFileDest,	// Destination
				NULL,
				NULL,
				NULL,
				0,
				hTransaction);
			Assert::IsTrue(TRUE == bRet, "CopyFileTransacted failed");

			Assert::IsTrue(File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(!File::Exists(strTestFileDest), "CopyFileTransacted failed"); 

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(File::Exists(strTestFileDest), "CopyFileTransacted failed");

			bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Now test that we can roll back the transaction by closing
			// the transaction handle before committing
			HANDLE hTransactionRollBack = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionRollBack, "Create transaction failed");

			File::Delete(strTestFileDest);
			Assert::IsTrue(!File::Exists(strTestFileDest), "Delete destination file failed"); 

			bRet = VsiTxF::VsiCopyFileTransacted(
				pszTestFileSrcCopy,	// Source
				pszTestFileDest,	// Destination
				NULL,
				NULL,
				NULL,
				0,
				hTransactionRollBack);
			Assert::IsTrue(TRUE == bRet, "CopyFileTransacted failed");

			Assert::IsTrue(File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(!File::Exists(strTestFileDest), "CopyFileTransacted failed"); 

			bRet = VsiTxF::VsiCloseTransaction(hTransactionRollBack);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			Assert::IsTrue(File::Exists(strTestFileSrcCopy), "MoveFileTransacted failed");
			Assert::IsTrue(!File::Exists(strTestFileDest), "CopyFileTransacted failed");

			// Cleanup
			File::Delete(strTestFileSrcCopy);
			File::Delete(strTestFileDest);

			// Now test overwrite existing opened file
			HANDLE hTransactionOverwriteO = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionOverwriteO, "Create transaction failed");

			// Create file A.txt (content: "A")
			String^ strDirA = gcnew String(strDir);
			strDirA += "\\A.txt";
			FileStream^ fs = File::Create(strDirA);
			Assert::IsTrue(File::Exists(strDirA), "Create A.txt failed.");
			fs->Close();

			fs = File::OpenWrite(strDirA);
			array<Byte>^info = (gcnew UTF8Encoding( true ))->GetBytes( "A" );
			fs->Write( info, 0, info->Length );
			fs->Close();

			// Create file B.txt (content: "B"), open
			String^ strDirB = gcnew String(strDir);
			strDirB += "\\B.txt";
			fs = File::Create(strDirB);
			Assert::IsTrue(File::Exists(strDirB), "Create B.txt failed.");
			fs->Close();

			fs = File::OpenWrite(strDirB);
			info = (gcnew UTF8Encoding( true ))->GetBytes( "B" );
			fs->Write( info, 0, info->Length );

			// Copy file A to file B (open)
			pin_ptr<const wchar_t> pszDirA = PtrToStringChars(strDirA);
			pin_ptr<const wchar_t> pszDirB = PtrToStringChars(strDirB);

			bRet = VsiTxF::VsiCopyFileTransacted(
				pszDirA,	// Source
				pszDirB,	// Destination
				NULL,
				NULL,
				NULL,
				0, //Copy is always overwrite
				hTransactionOverwriteO);
			Assert::IsTrue(FALSE == bRet, "CopyFileTransacted succeed - unexpected");

			// Commit
			bRet = VsiTxF::VsiCommitTransaction(hTransactionOverwriteO);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(File::Exists(strDirA), "CopyFileTransacted failed");
			Assert::IsTrue(File::Exists(strDirB), "CopyFileTransacted failed");

			// Check the content of file B - should be "B"
			fs->Close();
			fs = File::OpenRead(strDirB);
			array<Byte>^b = gcnew array<Byte>(1024);
			UTF8Encoding^ temp = gcnew UTF8Encoding( true );
			while ( fs->Read( b, 0, b->Length ) > 0 )
			{
				String^ strContentB = temp->GetString( b );
				pin_ptr<const wchar_t> pszContentB = PtrToStringChars(strContentB);
				Assert::IsTrue(0 == wcscmp(L"B", pszContentB), "CopyFileTransacted failed replacing");
			}
			// Close file B
			fs->Close();

			bRet = VsiTxF::VsiCloseTransaction(hTransactionOverwriteO);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Now test overwrite existing closed file
			HANDLE hTransactionOverwriteC = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionOverwriteC, "Create transaction failed");

			// Copy file A to file B (close)
			bRet = VsiTxF::VsiCopyFileTransacted(
				pszDirA,	// Source
				pszDirB,	// Destination
				NULL,
				NULL,
				NULL,
				0,  //Copy is always overwrite
				hTransactionOverwriteC);
			Assert::IsTrue(TRUE == bRet, "CopyFileTransacted failed");

			// Commit
			bRet = VsiTxF::VsiCommitTransaction(hTransactionOverwriteC);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(File::Exists(strDirA), "CopyFileTransacted failed");
			Assert::IsTrue(File::Exists(strDirB), "CopyFileTransacted failed");

			// Check the content of file B - should be "A"
			fs = File::OpenRead(strDirB);
			b = gcnew array<Byte>(1024);
			temp = gcnew UTF8Encoding( true );
			while ( fs->Read( b, 0, b->Length ) > 0 )
			{
				String^ strContentB = temp->GetString( b );
				pin_ptr<const wchar_t> pszContentB = PtrToStringChars(strContentB);
				Assert::IsTrue(0 == wcscmp(L"A", pszContentB), "CopyFileTransacted failed replacing");
			}
			fs->Close();

			bRet = VsiTxF::VsiCloseTransaction(hTransactionOverwriteC);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Clean up
			File::Delete(strDirA);
			File::Delete(strDirB);			
		}

		[TestMethod]
		void TestVsiDeleteFileTransacted()
		{
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			String^ strDir = testContextInstance->TestDeploymentDir;

			// Create file1
			String^ strTestFileSrc1 = gcnew String(strDir);
			strTestFileSrc1 += "\\VsiCommUtl.dll";

			String^ strTestFileSrcCopy1 = gcnew String(strDir);
			strTestFileSrcCopy1 += "\\VsiTestTransactedSource1"; //create a copy to protect source

			File::Copy(strTestFileSrc1, strTestFileSrcCopy1);

			// Create file2
			String^ strTestFileSrc2 = gcnew String(strDir);
			strTestFileSrc2 += "\\3d-Mode.3img";

			String^ strTestFileSrcCopy2 = gcnew String(strDir);
			strTestFileSrcCopy2 += "\\VsiTestTransactedSource2"; //create a copy to protect source

			File::Copy(strTestFileSrc2, strTestFileSrcCopy2);

			pin_ptr<const wchar_t> pszTestFileSrcCopy1 = PtrToStringChars(strTestFileSrcCopy1);
			pin_ptr<const wchar_t> pszTestFileSrcCopy2 = PtrToStringChars(strTestFileSrcCopy2);

			//to delete multiple files: call DeleteFileTransacted multiple times with different FileName but same HANDLE
			BOOL bRet = VsiTxF::VsiDeleteFileTransacted(
				pszTestFileSrcCopy1,
				hTransaction);
			Assert::IsTrue(TRUE == bRet, "DeleteFileTransacted failed");

			bRet = VsiTxF::VsiDeleteFileTransacted(
				pszTestFileSrcCopy2,
				hTransaction);
			Assert::IsTrue(TRUE == bRet, "DeleteFileTransacted failed");

			Assert::IsTrue(File::Exists(strTestFileSrcCopy1), "DeleteFileTransacted failed"); 
			Assert::IsTrue(File::Exists(strTestFileSrcCopy2), "DeleteFileTransacted failed");

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!File::Exists(strTestFileSrcCopy1), "DeleteFileTransacted failed"); 
			Assert::IsTrue(!File::Exists(strTestFileSrcCopy2), "DeleteFileTransacted failed"); 

			bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Cannot test the situation that exception happens during delete, not easy to fool:
			// Try1: 
			// create 2 files where file 2 is open (use File::Create to make file 2 open),
			// DeleteFileTransacted file 1 returns TRUE, but DeleteFileTransacted file 2 returns FALSE,
			// file 1 and file 2 both exists before CommitTransaction,
			// CommitTransaction returns TRUE,
			// file 1 no longer exists but file 2 still exists after CommitTransaction,
			// CloseTransaction returns TRUE;
			// Try2: 
			// create 2 files where both are closed, 
			// DeleteFileTransacted file 1 returns TRUE, and DeleteFileTransacted file 2 returns TRUE,
			// file 1 and file 2 both exists before CommitTransaction,
			// attempt to open file 2 before CommitTransaction (use File::OpenRead), but system throw exception;
		}
		
		[TestMethod]
		void TestVsiRemoveDirectoryTransacted() //Remove empty directory
		{
			//Test not empty directory
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			String^ strDir = testContextInstance->TestDeploymentDir;

			// Create a test directory	
			String^ strTestDir = gcnew String(strDir);
			strTestDir += "\\TestDir";
			Directory::CreateDirectory(strTestDir);
			Assert::IsTrue(Directory::Exists(strTestDir), "Create test directory failed.");

			// Create test.txt in the test directory
			String^ strTestDirin = gcnew String(strTestDir);
			strTestDirin += "\\test.txt";
			FileStream^ fs = File::Create(strTestDirin);
			Assert::IsTrue(File::Exists(strTestDirin), "Create test.txt in the test directory failed.");
			fs->Close();

			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strTestDir);

			//attempt to delete the not empty test directory
			BOOL bRet = VsiTxF::VsiRemoveDirectoryTransacted(
				pszTestDir,
				hTransaction);
			Assert::IsTrue(FALSE == bRet, "RemoveDirectoryTransacted succeed - unexpected");

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted succeed - unexpected"); 
			Assert::IsTrue(File::Exists(strTestDirin), "RemoveDirectoryTransacted succeed - unexpected");

			//Test empty directory
			HANDLE hTransactionEmpty = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionEmpty, "Create transaction failed");

			File::Delete(strTestDirin);
			Assert::IsTrue(!File::Exists(strTestDirin), "Delete test.txt from the test directory failed.");

			//attempt to delete the empty test directory
			bRet = VsiTxF::VsiRemoveDirectoryTransacted(
				pszTestDir,
				hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "RemoveDirectoryTransacted failed");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted failed");

			bRet = VsiTxF::VsiCommitTransaction(hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!Directory::Exists(strTestDir), "RemoveDirectoryTransacted succeed - unexpected"); 

			bRet = VsiTxF::VsiCloseTransaction(hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");
		}

		[TestMethod]
		void TestVsiRemoveAllDirectoryTransacted() //Remove directory and its contents
		{
			//Test not empty directory
			HANDLE hTransaction = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransaction, "Create transaction failed");

			String^ strDir = testContextInstance->TestDeploymentDir;

			// Create a test directory	
			String^ strTestDir = gcnew String(strDir);
			strTestDir += "\\TestDir";
			Directory::CreateDirectory(strTestDir);
			Assert::IsTrue(Directory::Exists(strTestDir), "Create test directory failed.");

			// Create test.txt in the test directory
			String^ strTestDirin = gcnew String(strTestDir);
			strTestDirin += "\\test.txt";
			FileStream^ fs = File::Create(strTestDirin);
			Assert::IsTrue(File::Exists(strTestDirin), "Create test.txt in the test directory failed.");
			fs->Close();

			// Create a subtest directory under the test directory	
			String^ strSubTestDir = gcnew String(strTestDir);
			strSubTestDir += "\\SubTestDir";
			Directory::CreateDirectory(strSubTestDir);
			Assert::IsTrue(Directory::Exists(strSubTestDir), "Create subtest directory failed.");

			// Create subtest.txt in the subtest directory
			String^ strSubTestDirin = gcnew String(strSubTestDir);
			strSubTestDirin += "\\subtest.txt";
			fs = File::Create(strSubTestDirin);
			Assert::IsTrue(File::Exists(strSubTestDirin), "Create subtest.txt in the subtest directory failed.");
			fs->Close();

			pin_ptr<const wchar_t> pszTestDir = PtrToStringChars(strTestDir);

			//attempt to delete the not empty test directory
			BOOL bRet = VsiTxF::VsiRemoveAllDirectoryTransacted(
				pszTestDir,
				hTransaction);
			Assert::IsTrue(TRUE == bRet, "RemoveDirectoryTransacted failed");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted failed"); 
			Assert::IsTrue(File::Exists(strTestDirin), "RemoveDirectoryTransacted failed");
			Assert::IsTrue(Directory::Exists(strSubTestDir), "RemoveDirectoryTransacted failed"); 
			Assert::IsTrue(File::Exists(strSubTestDirin), "RemoveDirectoryTransacted failed");

			bRet = VsiTxF::VsiCommitTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!Directory::Exists(strTestDir), "RemoveDirectoryTransacted failed"); 
			Assert::IsTrue(!File::Exists(strTestDirin), "RemoveDirectoryTransacted failed");
			Assert::IsTrue(!Directory::Exists(strSubTestDir), "RemoveDirectoryTransacted failed"); 
			Assert::IsTrue(!File::Exists(strSubTestDirin), "RemoveDirectoryTransacted failed");

			bRet = VsiTxF::VsiCloseTransaction(hTransaction);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			//Test empty directory
			HANDLE hTransactionEmpty = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionEmpty, "Create transaction failed");

			// Create a test directory
			Directory::CreateDirectory(strTestDir);
			Assert::IsTrue(Directory::Exists(strTestDir), "Create test directory failed.");

			//attempt to delete the empty test directory
			bRet = VsiTxF::VsiRemoveAllDirectoryTransacted(
				pszTestDir,
				hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "RemoveDirectoryTransacted failed");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted failed");

			bRet = VsiTxF::VsiCommitTransaction(hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(!Directory::Exists(strTestDir), "RemoveDirectoryTransacted succeed - unexpected"); 

			bRet = VsiTxF::VsiCloseTransaction(hTransactionEmpty);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			//Test not empty directory with open files inside
			HANDLE hTransactionOpen = VsiTxF::VsiCreateTransaction(
				NULL,	// NULL - this means the returned handle cannot be inherited by child processes
				0,		// Reserved, must be 0
				TRANSACTION_DO_NOT_PROMOTE,		// Transaction cannot be distributed
				0,		// Reserved, must be 0
				0,		// Reserved, must be 0
				INFINITE,	// No time out
				L"Test transaction");	// User readable description
			Assert::IsTrue(INVALID_HANDLE_VALUE != hTransactionOpen, "Create transaction failed");

			// Create a test directory
			Directory::CreateDirectory(strTestDir);
			Assert::IsTrue(Directory::Exists(strTestDir), "Create test directory failed.");

			// Create test.txt in the test directory
			fs = File::Create(strTestDirin);
			Assert::IsTrue(File::Exists(strTestDirin), "Create test.txt in the test directory failed.");

			//attempt to delete the test directory
			bRet = VsiTxF::VsiRemoveAllDirectoryTransacted(
				pszTestDir,
				hTransactionOpen);
			Assert::IsTrue(FALSE == bRet, "RemoveDirectoryTransacted succeed - unexpected");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted succeed - unexpected");
			Assert::IsTrue(File::Exists(strTestDirin), "RemoveDirectoryTransacted succeed - unexpected");

			bRet = VsiTxF::VsiCommitTransaction(hTransactionOpen);
			Assert::IsTrue(TRUE == bRet, "CommitTransaction failed");

			Assert::IsTrue(Directory::Exists(strTestDir), "RemoveDirectoryTransacted succeed - unexpected");
			Assert::IsTrue(File::Exists(strTestDirin), "RemoveDirectoryTransacted succeed - unexpected");

			bRet = VsiTxF::VsiCloseTransaction(hTransactionOpen);
			Assert::IsTrue(TRUE == bRet, "Close transaction failed");

			// Clean up
			fs->Close();
			File::Delete(strTestDirin);
			Directory::Delete(strTestDir);
		}
	};
}
