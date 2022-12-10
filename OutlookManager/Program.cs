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
			string targetAccount = "ARC-Bob";

			//string sourceFolder = "Inbox";
			//string targetFolder = "XferIn";
			//string targetFolder = "Inbox";

			string sourceFolder = "Sent Items";
			string targetFolder = "Sent";

			//string sourceFolder = "Calendar";
			//string targetFolder = "ARC RPK Calendar";

			//string sourcePath = @"D:\UserData\Documents\_Sync\AmericanResearchCapital\_admin\Messages_ToCopy";
			//string completedPath = @"D:\UserData\Documents\_Sync\AmericanResearchCapital\_admin\Messages_ZZCopied";


			XferItems.MoveMail(sourceAccount, sourceFolder, targetAccount, targetFolder, 2000);
			//XferItems.MoveAppointments(sourceAccount, sourceFolder, targetAccount, targetFolder, 99999);
			//XferItems.CopyMailFromFileFolder(sourcePath, completedPath, targetAccount, targetFolder);


			Console.WriteLine("Done");

			Console.ReadKey();
		}
	}
}