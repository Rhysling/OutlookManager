using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMailParser.Utilities;

public static class OutlookInspector
{
	public static Outlook.Account? FindAccountBySmtpAddress(string smtpAddress)
	{
		var App = new Outlook.Application();

		Outlook.NameSpace? ns = null;
		Outlook.Accounts? accounts = null;
		Outlook.Account? account = null;

		try
		{
			ns = App.Session;
			accounts = ns.Accounts;
			int i = 0, ic = accounts.Count;

			for (i = 1; i <= ic; i += 1)
			{
				account = accounts[i];

				if (account != null)
				{
					if (account.SmtpAddress == smtpAddress) break;

					Marshal.ReleaseComObject(account);
				}
			}
		}
		finally
		{
			if (accounts != null)
				Marshal.ReleaseComObject(accounts);

			if (ns != null)
				Marshal.ReleaseComObject(ns);
		}

		return account;
	}

	public static Outlook.MAPIFolder? GetFolder(Outlook.Account account, string folderName)
	{
		//var App = new Outlook.Application();

		Outlook.Folders? folders = null;
		Outlook.MAPIFolder? folder = null;

		try
		{
			folders = account.DeliveryStore.GetRootFolder().Folders;
			int i = 0, ic = folders.Count;

			for (i = 1; i <= ic; i += 1)
			{
				folder = folders[i];
				if (folder != null)
				{
					if (String.Compare(folder.Name, folderName, true) == 0) break;

					Marshal.ReleaseComObject(folder);
				}
			}
		}
		finally
		{
			if (folders != null)
				Marshal.ReleaseComObject(folders);
		}

		return folder;
	}


	public static List<Outlook.MailItem> GetMailItems(Outlook.MAPIFolder folder, int maxCount)
	{
		var items = new List<Outlook.MailItem>();

		int i, ic = folder.Items.Count;

		for (i = 1; i <= ic; i += 1)
		{
			if (folder.Items[i] is null) continue;

			if (folder.Items[i] is Outlook.MailItem)
				items.Add((Outlook.MailItem)folder.Items[i]);

			if (i > maxCount) break;
		}

		return items;
	}

	public static List<Outlook.AppointmentItem> GetAppointmentItems(Outlook.MAPIFolder folder, int maxCount)
	{
		var items = new List<Outlook.AppointmentItem>();

		int i, ic = folder.Items.Count;

		for (i = 1; i <= ic; i += 1)
		{
			if (folder.Items[i] is null) continue;

			if (folder.Items[i] is Outlook.AppointmentItem)
				items.Add((Outlook.AppointmentItem)folder.Items[i]);

			if (i > maxCount) break;
		}

		return items;
	}

}
