using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace FreeMiCal
{
	class ICalWriter
	{
		private StreamWriter iCalFile;
		private String filename;

		public ICalWriter (String filename) : this (filename, false)	// call next constructor
		{
		}

		public ICalWriter (String filename, bool overwrite)
		{
			if (File.Exists (filename))
			{
				if (overwrite == true)
					File.Delete (filename);
				else
					return;		// user wants to preserve file, exit prematurely. not elegant, but works.
			}
			try
			{
				this.filename = filename;
				iCalFile = new StreamWriter (filename);
				iCalFile.AutoFlush = true;

				iCalFile.WriteLine ("BEGIN:VCALENDAR");
				iCalFile.WriteLine ("VERSION:2.0");
				iCalFile.WriteLine ("PRODID:-//RSB//FreeMiCal//EN");
			}
			#region specific error handling ...
			catch (DirectoryNotFoundException ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: Directory not found\n\n{0}", ex.Message);
			}
			catch (System.Security.SecurityException ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: You do not have sufficient access rights\n\n{0}", ex.Message);
			}
			catch (UnauthorizedAccessException ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: You do not have access rights\n\n{0}", ex.Message);
			}
			catch (PathTooLongException ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: The path is too long\n\n{0}", ex.Message);
			}
			catch (IOException ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: Some IO error occured\n\n{0}", ex.Message);
			}
			#endregion
			// generic error handling
			catch (Exception ex)
			{
				Console.WriteLine ("ICalWriter::ICalWriter: Some error occured:\n\n{0}\n\n{1}", ex.Message, ex.StackTrace);
			}
		}

		/// <summary>
		/// CloseWriter needs to be called in order to flush file content
		/// </summary>
		public void CloseWriter ()
		{
			iCalFile.WriteLine ("END:VCALENDAR");
			iCalFile.Close ();
		}

		public void WriteEvent (OLCalItem iCalItem)
		{
			iCalFile.WriteLine ("BEGIN:VEVENT");

			// basic VEVENT properties
			iCalFile.WriteLine (text2iCal (iCalItem.SUMMARY));
			iCalFile.WriteLine (iCalItem.DTSTART);
			iCalFile.WriteLine (iCalItem.DTEND);
			iCalFile.WriteLine (iCalItem.DTSTAMP);

			// optional VEVENT properties
			if (iCalItem.LOCATION.Length != 0)
				iCalFile.WriteLine (text2iCal (iCalItem.LOCATION));
			if (iCalItem.DESCRIPTION.Length != 0)
				iCalFile.WriteLine (text2iCal (iCalItem.DESCRIPTION));
			if (iCalItem.TRANSP.Length != 0)
				iCalFile.WriteLine (text2iCal (iCalItem.TRANSP));
			if (iCalItem.CLASS.Length != 0)
				iCalFile.WriteLine (iCalItem.CLASS);
			if (iCalItem.PRIORITY.Length != 0)
				iCalFile.WriteLine (iCalItem.PRIORITY);
			if (iCalItem.CATEGORIES.Length != 0)
				iCalFile.WriteLine (text2iCal (iCalItem.CATEGORIES));

			// write recurrency rule
			if (iCalItem.isRecurrent ())
				WriteRecurrence (iCalItem);

			// write alarm if one exists
			if (iCalItem.hasAlarm ())
				WriteAlarm (iCalItem);
			
			iCalFile.WriteLine ("END:VEVENT");
		}

		public void WriteAlarm (OLCalItem iCalItem)
		{
			if (!iCalItem.hasAlarm ())
				return;

			// this writes just the alarm time
			// duration and repeat count are not implemented yet
			iCalFile.WriteLine ("BEGIN:VALARM");
			iCalFile.WriteLine (iCalItem.TRIGGER);
			iCalFile.WriteLine (text2iCal(iCalItem.ALARM_DESCRIPTION));
			iCalFile.WriteLine ("ACTION:DISPLAY");
			iCalFile.WriteLine ("END:VALARM");
		}

		public void WriteRecurrence (OLCalItem iCalItem)
		{
			if (!iCalItem.isRecurrent ())
				return;
			iCalFile.WriteLine (text2iCal (iCalItem.RRULE_FREQ));

		}

		private String text2iCal (String text)
		{
			int MAXLINESIZE = 70;				// max 75 allowed, methode adds 3 chars to the line (cr, lf, sp)
			int curPos = 0;

			if (text.Length < MAXLINESIZE)
				return text;					// quick exit

			StringBuilder t = new StringBuilder ();

			while (curPos < text.Length)
			{
				if ((curPos + MAXLINESIZE) > text.Length)
					MAXLINESIZE = text.Length - curPos;
				t.Append (text.Substring (curPos, MAXLINESIZE));
				curPos += MAXLINESIZE;
				t.Append ("\r\n ");
			}
			t.Remove (t.Length - 3, 3);
			return t.ToString ();
		}

	}
}
