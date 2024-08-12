using System;
using CDOEXM;
using System.Collections;

namespace sode
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Class1
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
		CDOEXM.IExchangeServer iExs;
		CDOEXM.IStorageGroup iSg;
		CDOEXM.MailboxStoreDB iMdb;
		CDOEXM.PublicStoreDB iPf;
		iExs = new CDOEXM.ExchangeServerClass();
		iSg = new CDOEXM.StorageGroupClass();
		iMdb = new CDOEXM.MailboxStoreDBClass();
		iPf = new CDOEXM.PublicStoreDBClass();
		string snServername = "mgnms01";
		iExs.DataSource.Open(snServername,null,ADODB.ConnectModeEnum.adModeReadWrite,
		ADODB.RecordCreateOptionsEnum.adOpenIfExists,ADODB.RecordOpenOptionsEnum.adOpenSource,null,null);
		foreach (string Sgname in (IEnumerable)iExs.StorageGroups){
			iSg.DataSource.Open(Sgname,null,ADODB.ConnectModeEnum.adModeReadWrite,
			ADODB.RecordCreateOptionsEnum.adOpenIfExists,ADODB.RecordOpenOptionsEnum.adOpenSource,null,null);
			foreach( string Mbname in (IEnumerable)iSg.MailboxStoreDBs){
				iMdb.DataSource.Open(Mbname,null,ADODB.ConnectModeEnum.adModeReadWrite,
				ADODB.RecordCreateOptionsEnum.adOpenIfExists,ADODB.RecordOpenOptionsEnum.adOpenSource,null,null);
				System.Console.WriteLine("Store Name: " + iMdb.Name);
				System.Console.WriteLine("Status: " + iMdb.Status);
				System.Console.WriteLine("Storage Quota Limit: " + iMdb.StoreQuota);
				System.Console.WriteLine("Over Quota Limit: " + iMdb.OverQuotaLimit);
				System.Console.WriteLine("HardLimit: " + iMdb.HardLimit);
				System.Console.WriteLine("");
				}
			foreach( string Pfname in (IEnumerable)iSg.PublicStoreDBs){
				iPf.DataSource.Open(Pfname,null,ADODB.ConnectModeEnum.adModeReadWrite,
				ADODB.RecordCreateOptionsEnum.adOpenIfExists,ADODB.RecordOpenOptionsEnum.adOpenSource,null,null);
				System.Console.WriteLine("Store Name: " + iPf.Name);
				System.Console.WriteLine("Status: " + iPf.Status);
				System.Console.WriteLine("");
			}
			}
		}
	}
}
