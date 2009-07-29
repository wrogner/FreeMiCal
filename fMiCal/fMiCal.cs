using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FreeMiCal
{
	class fMiCal
	{

		static void Main (string[] args)
		{
			int startIdx = 1;								// start index is 1
			int endIdx = -1;								// imply up to last record
			String profileName = "Outlook";					// use default profile name

			#region create default output filename
			StringBuilder fn = new StringBuilder ("freemical_");
			DateTime fnd = DateTime.Now;
			fn.Append (fnd.Year.ToString ("0000"));
			fn.Append (fnd.Month.ToString ("00"));
			fn.Append (fnd.Day.ToString ("00"));
			fn.Append (".ics");
			String fnOutput = fn.ToString ();
			#endregion

			#region Check arguments

			// parse command line arguments for start and end records
			foreach (String arg in args)
			{
				String a = arg.ToLower ();

				#region --help
				if (a.Equals ("--help") || a.Equals ("/h") || a.Equals ("-h") || a.Equals ("?"))
				{
					showHelp ();
					return;
				}
				#endregion

				#region --start
				if (a.StartsWith ("--start") || a.StartsWith ("-s") || a.StartsWith ("/s"))
				{
					String[] p = arg.Split (new char[] { '=', ':' });
					// assume that no 2nd p means from start
					if (p.Length > 1)
					{
						try
						{									// this is dirty, works for now
							startIdx = Int32.Parse (p[1]);
						}
						catch (Exception ex)
						{
							if (p[1].Length > 0)			// the value after = was not a number
							{
								showHelp ();
								return;
							}
						}
					}
				}
				#endregion

				#region --end
				if (a.StartsWith ("--end") || a.StartsWith ("-e") || a.StartsWith ("/e"))
				{
					String[] p = arg.Split (new char[] { '=', ':' });
					// assume that no 2nd p means to end
					if (p.Length > 1)
					{
						try
						{									// this is dirty, works for now
							endIdx = Int32.Parse (p[1]);
						}
						catch (Exception ex)
						{
							if (p[1].Length > 0)			// the value after = was not a number
							{
								showHelp ();
								return;
							}
						}
					}
				}
				#endregion

				#region --output
				if (a.StartsWith ("--output") || a.StartsWith ("-o") || a.StartsWith ("/o"))
				{
					String[] p = arg.Split (new char[] { '=', ':' }, 2);
					if (p.Length > 1)						// assume that no length specifier means from start
					{
						try
						{									// this is dirty, works for now
							fnOutput = p[1];
						}
						catch (Exception ex)
						{
							showHelp ();
							return;
						}
					}
				}
				#endregion

				#region --profile
				if (a.StartsWith ("--profile") || a.StartsWith ("-p") || a.StartsWith ("/p"))
				{
					String[] p = arg.Split (new char[] { '=', ':' }, 2);
					if (p.Length > 1)
					{
						profileName = p[1];
					}
				}
				#endregion
			}

			// clean up parameter values
			if (startIdx < 1)							// correct index < 1
				startIdx = 1;

			if (endIdx > 0 && endIdx < startIdx)		// correct lower and upper boundaries
			{
				int t = endIdx;
				endIdx = startIdx;
				startIdx = t;
			}

			// check for possible output
			try
			{
				FileInfo fi = new FileInfo (fnOutput);
				if (fi.Directory.Exists == false)
				{
					Console.WriteLine ("The directory you specified does not exist");
					return;
				}
				StreamWriter sw = new StreamWriter (fnOutput, false);
				sw.Write ("");
				sw.Flush ();
				sw.Close ();
				if (fi.Exists == true)
					fi.Delete ();
			}
			#region specific error handling
			catch (System.UnauthorizedAccessException ex)
			{
				Console.WriteLine ("You are not authorized to access this file or directory.");
				Console.WriteLine ("Contact you System administrator");
				return;
			}
			catch (System.Security.SecurityException ex)
			{
				Console.WriteLine ("You are do not have sufficient rights to access this file or directory.");
				Console.WriteLine ("Contact you System administrator");
				return;
			}
			catch (System.IO.PathTooLongException ex)
			{
				Console.WriteLine ("The path to the file is too long.");
				return;
			}
			#endregion
			// generic error handling
			catch (Exception ex)
			{
				Console.WriteLine ("Some error occured processing the output file name.");
				return;
			}
			#endregion


			try
			{
				OLCalReader olCal = new OLCalReader (profileName);
				// constructor calls Logon implicitly

				if (!olCal.LoggedOn)					// handle error in profile name
				{
					Console.WriteLine ("Could not log on to Outlook. Check the profile name");
					writeLock (-1);						// signal error to FreeMiCal
					return;
				}

				int maxRecord = olCal.Count;
				if (maxRecord == 0)
				{
					Console.WriteLine ("There are no records in your Outlook calendar!");
					return;
				}

				// this check cannot be carried out without Outlook connection. This is why it's here
				if (endIdx < 0 || endIdx > maxRecord)		// negative index not allowed, no end record set.
					endIdx = maxRecord;						// export till last event

				if (endIdx == 0)							// if user selected end record 0 then abort
				{
					Console.WriteLine ("You chose not to export any items!");
					return;
				}

				Console.WriteLine ("Start at {0}, end at {1}, output to {2}", startIdx, endIdx, fnOutput);

				ICalWriter iCal = new ICalWriter (fnOutput, true);

				for (int record = startIdx; record <= endIdx; record++)
				{
					OLCalItem olItem = new OLCalItem ((Outlook.AppointmentItem)olCal[record]);
					iCal.WriteEvent (olItem);
					writeLock (record);
				}

				iCal.CloseWriter ();

				olCal.Logoff ();
					
			}
			catch (RuntimeException re) { }

		}

		private static void showHelp ()
		{
			Console.WriteLine ("\nfMiCal\tConverts Outlook calendar items to RFC2445 iCal format.");
			Console.WriteLine ("\tVersion: 0.2.0.0, @2007 by war under OSL v3.0\n");
			Console.WriteLine ("\tFor Windows version, run FreeMiCal\n");

			Console.WriteLine ("usage:\tfmical [option|...]\n");
			Console.WriteLine (" --help\t\tprint this page. Same as ?");
			Console.WriteLine (" --start=n\tstart export with record n");
			Console.WriteLine (" --end=m\tend export after record m. Negative numbers are ignored");
			Console.WriteLine (" --output=file\tFilename to export calendar items");
			Console.WriteLine (" --profile=name\tOutlook profile name\n");

			Console.WriteLine (" parameters can be abbreviated windows (e.g. /h) or old unix style (e.g. -h)");
			Console.WriteLine (" equal signs (=) may be replaced with colons (:)");
		}

		// write current record that is written by iCalWriter to the lock file
		// used for synchronisation purposes with the GUI
		private static bool writeLock (int cur)
		{
			try
			{
				FileStream fs = File.OpenWrite ("fMiCal.run");
				fs.Seek (0, System.IO.SeekOrigin.Begin);

				Byte[] c;
				String s;
				s = cur.ToString ();
				c = System.Text.Encoding.GetEncoding (0).GetBytes (s);

				fs.Write (c, 0, s.Length);
				fs.SetLength (s.Length);
				fs.Flush ();
				fs.Close ();

				return true;
			}
			catch (Exception ex)
			{
				return false;
			}
		}
	}
}
