using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FreeMiCal
{
	class RuntimeException : Exception
	{
		public RuntimeException (String msg)
		{
			Console.WriteLine ("{0} not implemented yet", msg);
		}
	}

	public class OLCalReader
	{
		private Outlook.Application myOutlook;
		private Outlook.NameSpace myProfile;
		private Outlook.MAPIFolder myFolder;
		private String profileName = "Outlook";
		private bool loggedOn = false;

		public OLCalReader () : this ("Outlook")
		{
		}

		public OLCalReader (String profileName)
		{
			myOutlook = new Outlook.Application ();
			myProfile = myOutlook.Session;
			this.profileName = profileName;
			Logon ();
		}

		public void Logon ()
		{
			if (loggedOn == true)
				return;

			//string profile = myOutlook.DefaultProfileName;		// log on to the default profile
			string pwd = null;
			bool showDialog = false;
			bool newSession = true;
			//myProfile.Logon (profile, pwd, showDialog, newSession);
			try
			{
				myProfile.Logon (profileName, pwd, showDialog, newSession);
				loggedOn = true;
			}
			catch (Exception ex)
			{
				loggedOn = false;
			}

			if (loggedOn)
			{
				try
				{
					myFolder = myProfile.GetDefaultFolder (Outlook.OlDefaultFolders.olFolderCalendar);
				}
				catch (Exception ex)
				{
					Console.WriteLine ("OLCalReader::OLCalReader(profile): Error connecting to default calendar folder");
					loggedOn = false;
				}
			}
		}

		public void Logoff ()
		{
			if (loggedOn == true)
				myProfile.Logoff ();
		}

		public bool LoggedOn
		{
			get
			{
				return loggedOn;
			}
		}

		//public Outlook.MAPIFolder MAPIFolder
		//{
		//    get
		//    {
		//        if (loggedOn == false)
		//            Logon ();
		//        return myFolder;
		//    }
		//}

		public String DefaultProfile
		{
			get
			{
				//return myOutlook.DefaultProfileName;
				return "Outlook";
			}
		}

		public String ProfileName
		{
			get
			{
				return profileName;
			}
		}

		public int Count
		{
			get
			{
				return myFolder.Items.Count;
			}
		}

		public Outlook.AppointmentItem this[int i]
		{
			get
			{
				if (loggedOn == false)
					Logon ();

				if (i < 1) i = 1;											// correct minimum index
				if (i > myFolder.Items.Count) i = myFolder.Items.Count;		// correct maximum index
				// omit using this.Count due to performance reasons
				return (Outlook.AppointmentItem)myFolder.Items[i];
			}
		}
	}

	class OLCalItem
	{
		private Outlook.AppointmentItem calItem;
		private Outlook.RecurrencePattern pattern;

		public OLCalItem (Outlook.AppointmentItem calItem)
		{
			this.calItem = calItem;
			if (calItem.IsRecurring)
				pattern = calItem.GetRecurrencePattern ();
		}

		// VEVENT properties
		public String SUMMARY
		{
			get
			{
				return "SUMMARY:" + calItem.Subject;
			}
		}

		public String LOCATION
		{
			get
			{
				if (calItem.Location != null)
					return "LOCATION:" + calItem.Location;
				else
					return "";
			}
		}

		public String DTSTART
		{
			get
			{
				if (calItem.AllDayEvent == true)
					return "DTSTART" + allDayEvent (calItem.Start);		// mind the missing : as line is extended
				else
					return "DTSTART:" + time2ISO (calItem.Start);
			}
		}

		public String DTEND
		{
			get
			{
				if (calItem.AllDayEvent == true)
					return "DTEND" + allDayEvent (calItem.End);		// mind the missing : as line is extended
				else
					return "DTEND:" + time2ISO (calItem.End);
			}
		}

		public String DTSTAMP
		{
			get
			{
				return "DTSTAMP:" + time2ISO (calItem.CreationTime);
			}
		}

		public String DESCRIPTION
		{
			get
			{
				if (calItem.Body != null)
					return "DESCRIPTION:" + calItem.Body.Replace ("\r\n", "\\n");
				else
					return "";
			}
		}

		public String CLASS
		{
			get
			{
				String c;

				switch (calItem.Sensitivity)
				{
					case Outlook.OlSensitivity.olPrivate:
						c = "CLASS:PRIVATE";
						break;
					case Outlook.OlSensitivity.olConfidential:
						c = "CLASS:CONFIDENTIAL";
						break;
					default:
						//c = "CLASS:PUBLIC;	this is the default
						c = "";
						break;
				}

				return c;
			}
		}

		public String TRANSP
		{
			get
			{
				String t;

				switch (calItem.BusyStatus)
				{
					case Outlook.OlBusyStatus.olFree:
						t = "TRANSP:TRANSPARENT";
						break;
					case Outlook.OlBusyStatus.olTentative:
						t = "TRANSP:TENTATIVE";
						break;
					default:
						//t = "TRANSP:OPAQUE";	this is the default
						t = "";
						break;
				}

				return t;
			}
		}

		public String PRIORITY
		{
			get
			{
				String p;

				switch (calItem.Importance)
				{
					case Microsoft.Office.Interop.Outlook.OlImportance.olImportanceLow:
						p = "PRIORITY:9";
						break;
					case Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh:
						p = "PRIORITY:1";
						break;
					default:
						// p = "PRIORITY:5";	Importance = olImportanceNormal
						p = "";
						break;
				}

				return p;
			}
		}

		public String CATEGORIES
		{
			get
			{
				if (calItem.Categories != null)
					return "CATEGORIES:" + calItem.Categories.Replace("; ",",");
				else
					return "";
			}
		}

		// VALARM properties
		public bool hasAlarm ()
		{
			return calItem.ReminderSet;
		}

		public String TRIGGER
		{
			get
			{
				StringBuilder t = new StringBuilder ("TRIGGER:");
				t.Append ("-PT");		// currently only minutes are stored
				t.Append (calItem.ReminderMinutesBeforeStart);
				t.Append ("M");
				return t.ToString();
			}
		}

		public String ALARM_DESCRIPTION
		{
			get
			{
				return "DESCRIPTION:Alarm: " + calItem.Subject;
			}
		}

		// RECURRENCE properties
		public bool isRecurrent ()
		{
			return calItem.IsRecurring;
		}

		public String RRULE_FREQ
		{
			get
			{
				if (!calItem.IsRecurring)
					return "";

				StringBuilder f = new StringBuilder("RRULE:FREQ=");

				f.Append (rruleGetFrequency());
				if (!pattern.NoEndDate)
				{
													// one day added due to Exchange storing end date with time 0:00:00
													// which is actually the start of the day
					f.Append (";UNTIL=" + time2ISO (pattern.PatternEndDate.AddDays (1)));
				}

				return f.ToString ();
			}
		}

		// helper functions
		private String time2ISO (DateTime time)
		{
			StringBuilder t = new StringBuilder (time.Year.ToString ("0000"));
			t.Append (time.Month.ToString ("00"));
			t.Append (time.Day.ToString ("00"));
			t.Append ("T");
			t.Append (time.Hour.ToString ("00"));
			t.Append (time.Minute.ToString ("00"));
			t.Append (time.Second.ToString ("00"));

			return t.ToString();
		}

		private String allDayEvent (DateTime time)
		{
			StringBuilder t = new StringBuilder (";DATE=VALUE:");
			t.Append (time.Year.ToString ("0000"));
			t.Append (time.Month.ToString ("00"));
			t.Append (time.Day.ToString ("00"));

			return t.ToString ();
		}

		private String rruleGetFrequency ()
		{
			StringBuilder f = new StringBuilder ();

			switch (pattern.RecurrenceType)
			{
				case Outlook.OlRecurrenceType.olRecursDaily:
					f.Append ("DAILY");
					f.Append (";INTERVAL=" + (pattern.Interval.ToString () == "0" ? "1" : pattern.Interval.ToString ()));
																// correct an error in outlook that sets interval to 0 for
																// daily on all weekdays
					break;
				case Outlook.OlRecurrenceType.olRecursWeekly:
					f.Append ("WEEKLY");
					f.Append (";BYDAY=" + daysOfWeek(pattern.DayOfWeekMask));
					f.Append (";INTERVAL=" + (pattern.Interval.ToString () == "0" ? "1" : pattern.Interval.ToString ()));
					break;
				case Outlook.OlRecurrenceType.olRecursMonthly:
					f.Append ("MONTHLY");
					f.Append (";INTERVAL=" + (pattern.Interval.ToString () == "0" ? "1" : pattern.Interval.ToString ()));
					break;
				case Outlook.OlRecurrenceType.olRecursMonthNth:
					f.Append ("MONTHLY");
					f.Append (";BYDAY=" + pattern.Instance + daysOfWeek (pattern.DayOfWeekMask));
					f.Append (";INTERVAL=" + (pattern.Interval.ToString () == "0" ? "1" : pattern.Interval.ToString ()));
					break;
				case Outlook.OlRecurrenceType.olRecursYearly:
					f.Append ("YEARLY");
					f.Append (";INTERVAL=1");	// this catches another error in outlook interval is 12 in yearly events
					break;
				case Outlook.OlRecurrenceType.olRecursYearNth:
					f.Append ("YEARLY");
					f.Append (";BYMONTH=" + pattern.MonthOfYear);
					f.Append (";BYDAY=" + pattern.Instance + daysOfWeek (pattern.DayOfWeekMask));
					f.Append (";INTERVAL=1");	// this catches another error in outlook interval is 12 in yearly events
					break;
				default:
					// this should never happen, ERROR
					throw new RuntimeException ("OLCalItem::get_RRULE_FREQ: Recurrence not implemented");
					break;
			}
			return f.ToString ();
		}

		private String daysOfWeek (Outlook.OlDaysOfWeek dow)
		{
			StringBuilder d = new StringBuilder ();

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olMonday) == (int)Outlook.OlDaysOfWeek.olMonday)
			{
				d.Append ("MO");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olTuesday) == (int)Outlook.OlDaysOfWeek.olTuesday)
			{
				if (d.Length > 0) d.Append(",");
				d.Append ("TU");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olWednesday) == (int)Outlook.OlDaysOfWeek.olWednesday)
			{
				if (d.Length > 0) d.Append(",");
				d.Append ("WE");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olThursday) == (int)Outlook.OlDaysOfWeek.olThursday)
			{
				if (d.Length > 0) d.Append(",");
				d.Append ("TH");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olFriday) == (int)Outlook.OlDaysOfWeek.olFriday)
			{
				if (d.Length > 0) d.Append(",");
				d.Append ("FR");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olSaturday) == (int)Outlook.OlDaysOfWeek.olSaturday)
			{
				if (d.Length > 0) d.Append(",");
				d.Append ("SA");
			}

			if (((int)dow & (int)Outlook.OlDaysOfWeek.olSunday) == (int)Outlook.OlDaysOfWeek.olSunday)
			{
				if (d.Length > 0) d.Append (",");
				d.Append ("SU");
			}
			
			return d.ToString();
		}
	}
}
