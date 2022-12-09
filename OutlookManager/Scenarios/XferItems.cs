
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using OutlookMailParser.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookManager.Scenarios
{
	public static class XferItems
	{
		public static string CopyMail()
		{
			var sb = new StringBuilder();

			//set up variables
			Outlook.Application app;
			Outlook.MAPIFolder source;
			Outlook.MAPIFolder target;

			//Outlook.Folders folders = null;

			app = new Outlook.Application();
			//source = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

			source = app.Session.Folders["robert.kummer@falah-capital.com"].Folders["Sent Items"];
			target = app.Session.Folders["FalahCapital_Archive"].Folders["Sent Items"];
			int maxCount = 100;


			List<Outlook.MailItem> targetItems;
			targetItems = OutlookInspector.GetMailItems(target, maxCount);

			var th = new HashSet<string>();

			foreach (Outlook.MailItem item in targetItems)
				th.Add(MakeMsgHash(item));


			List<Outlook.MailItem> sourceItems;
			sourceItems = OutlookInspector.GetMailItems(source, maxCount);

			int i = 0;

			foreach (Outlook.MailItem item in sourceItems)
			{
				if (!th.Contains(MakeMsgHash(item)))
				{
					Outlook.MailItem m = item.Copy();
					m.Move(target);

					sb.AppendLine($"{item.SentOn:yyMMdd-HH:mm:ss}--{item.Subject.Left(10)}");
					i += 1;
					//if (i > 5) break;
				}
			}


			ReleaseObjects(app, source, target, sourceItems);


			return sb.ToString();
		}

		public static void MoveMail(string sourceAccountName, string sourceAccountFolder, string targetAccountName, string targetAccountFolder, int maxCount = 100)
		{
			//var sb = new StringBuilder();

			//set up variables
			Outlook.Application app;
			Outlook.MAPIFolder? source;
			Outlook.MAPIFolder? target;

			//Outlook.Folders folders = null;

			app = new Outlook.Application();

			//source = app.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
			//source = app.Session.Folders["robert.kummer@falah-capital.com"].Folders["Sent Items"];
			//target = app.Session.Folders["FalahCapital_Archive"].Folders["Sent Items"];

			source = app.Session.Folders[sourceAccountName]?.Folders[sourceAccountFolder];
			target = app.Session.Folders[targetAccountName]?.Folders[targetAccountFolder];

			if (source is null)
				throw new NullReferenceException($"Source: {sourceAccountName}/{sourceAccountFolder} is null.");

			if (target is null)
				throw new NullReferenceException($"Target: {targetAccountName}/{targetAccountFolder} is null.");


			List<Outlook.MailItem> sourceItems;
			sourceItems = OutlookInspector.GetMailItems(source, maxCount);

			foreach (Outlook.MailItem item in sourceItems)
			{
				item.Move(target);
			}


			ReleaseObjects(app, source, target, sourceItems);

		}

		public static void CopyMailFromFileFolder(string sourcePath, string completedPath, string targetAccountName, string targetAccountFolder)
		{
			//set up variables
			Outlook.Application app;
			Outlook.MAPIFolder? target;

			app = new Outlook.Application();
			target = app.Session.Folders[targetAccountName]?.Folders[targetAccountFolder];

			if (target is null)
				throw new NullReferenceException($"Target: {targetAccountName}/{targetAccountFolder} is null.");


			var files = Directory.GetFiles(sourcePath).Where(a => a.EndsWith(".msg")).OrderBy(a => a).ToList();

			int i = 0;

			var items = files.Select(a => new {
				Source = a,
				Dest = a.Replace(sourcePath, completedPath),
				MI = app.Session.OpenSharedItem(a) as Outlook.MailItem
			});

			foreach (var item in items)
			{
				if (item.MI is not null)
				{
					item.MI.Move(target);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(item.MI);
					File.Move(item.Source, item.Dest);
					i += 1;
				}

				if (i > 1000) break;
			}

			GC.WaitForPendingFinalizers();
			GC.Collect();




			//foreach (var file in files)
			//{
			//	var item = app.Session.OpenSharedItem(file) as Outlook.MailItem;

			//	if (item is null)
			//		throw new NullReferenceException($"MailItem: '{file}' is null.");

			//	item.Move(target);

			//	System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
			//	GC.WaitForPendingFinalizers();
			//	GC.Collect();

			//	File.Move(file, file.Replace(sourcePath, completedPath));


			//	i += 1;
			//	if (i > 2) break;
			//}


			ReleaseObjects(app, null, target, null);

		}


		private static void ReleaseObjects(Outlook.Application? app, Outlook.MAPIFolder? source, Outlook.MAPIFolder? target, List<Outlook.MailItem>? mailItems)
		{
			//release objects

			if (mailItems != null)
			{
				foreach (var item in mailItems)
				{
					if (item is not null)
						System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
				}
				GC.WaitForPendingFinalizers();
				GC.Collect();
			}

			if (target != null)
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(target);
				GC.WaitForPendingFinalizers();
				GC.Collect();
			}

			if (source != null)
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(source);
				GC.WaitForPendingFinalizers();
				GC.Collect();
			}

			if (app != null)
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
				GC.WaitForPendingFinalizers();
				GC.Collect();
			}
		}



		private static string MakeMsgHash(Outlook.MailItem msg)
		{
			if (msg == null)
				return "";

			string hs = "";

			foreach (Outlook.Recipient r in msg.Recipients)
				hs += r.Address;

			hs += ":" + msg.SenderEmailAddress;
			//hs += "--" + msg.SentOn.ToJsTime().ToString();

			return hs;
		}
	}
}
