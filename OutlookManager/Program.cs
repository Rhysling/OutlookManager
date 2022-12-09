using OutlookManager.Scenarios;

namespace OutlookManager
{
	internal class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Starting...");
			Console.WriteLine();

			string sourceAccount = "robert.kummer@americanresearchcapital.com";
			string sourceFolder = "Inbox";
			string targetAccount = "ARC-Bob";
			string targetFolder = "XferIn";

			//string sourcePath = @"D:\UserData\Documents\_Sync\AmericanResearchCapital\_admin\Messages_ToCopy";
			//string completedPath = @"D:\UserData\Documents\_Sync\AmericanResearchCapital\_admin\Messages_ZZCopied";


			XferItems.MoveMail(sourceAccount, sourceFolder, targetAccount, targetFolder, 100);
			//XferItems.CopyMailFromFileFolder(sourcePath, completedPath, targetAccount, targetFolder);


			Console.WriteLine("Done");

			Console.ReadKey();
		}
	}
}