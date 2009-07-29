using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace FreeMiCal
{
	public partial class Form1 : Form
	{
		Process procfMiCal = null;							// stores process information about fMiCal
															// to facilitate process cancelation
		delegate void SetMessageCallback (String text, bool severity);

		public Form1 ()
		{
			InitializeComponent ();

			populate ();
		}

		// set default values in form
		private void populate ()
		{
			OLCalReader ol = new OLCalReader ();
			txtProfile.Text = ol.DefaultProfile;
			int maxCals = ol.Count;
			numStart.Maximum = maxCals;
			numEnd.Maximum = maxCals;
			numStart.Value = (maxCals == 0) ? 0 : 1;		// required to catch empty calendar folder
			numEnd.Value = maxCals;

			txtFile.Text = makeExportFileName ();

		}

		// create export file name
		private String makeExportFileName ()
		{
			StringBuilder fn = new StringBuilder ("freemical_");
			DateTime fnd = DateTime.Now;
			fn.Append (fnd.Year.ToString ("0000"));
			fn.Append (fnd.Month.ToString ("00"));
			fn.Append (fnd.Day.ToString ("00"));
			fn.Append (".ics");
			return fn.ToString ();
		}

		private void btnExit_Click (object sender, EventArgs e)
		{
			Application.Exit ();
		}

		private void btnExport_Click (object sender, EventArgs e)
		{

			if (btnExport.Text.StartsWith ("Fr&ee") == true)
			{
				if (numStart.Value == 0 && numEnd.Value == 0)
				{
					writeMessage ("No calendar items found in profile!", false);
					return;
				}

				btnExport.Text = "&Cancel...";
				btnExport.Refresh ();
				progressBar1.Visible = true;

				String fMiCal = "fMiCal.exe";
				StringBuilder sbArgs = new StringBuilder ("-s:");
				sbArgs.Append (numStart.Value.ToString ());
				sbArgs.Append (" -e:");
				sbArgs.Append (numEnd.Value.ToString ());

				if (txtFile.Text != "")
					sbArgs.Append (" -o:" + txtFile.Text);

				if (txtProfile.Text != "")
					sbArgs.Append (" -p:" + txtProfile.Text);

				ProcessStartInfo psi = new ProcessStartInfo (fMiCal, sbArgs.ToString ());
				psi.WindowStyle = ProcessWindowStyle.Hidden;

				deleteLock ();

				try
				{
					procfMiCal = Process.Start (psi);
				}
				catch (Exception ex)
				{
					MessageBox.Show ("Error starting 'fMiCal.exe " + sbArgs.ToString () + "'", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				try
				{
					backgroundWorker1.RunWorkerAsync ();
				}
				catch (InvalidOperationException ex)
				{
					writeMessage ("Error starting background monitoring of fMiCal", true);
				}
			}
			else
			{
				if (procfMiCal != null)
				{
					try
					{
						procfMiCal.Kill ();
					}
					catch (Exception ex)
					{
						writeMessage ("Error terminating background process fMiCal, ID=" + procfMiCal.Id, true);
					}
				}
				try
				{
					backgroundWorker1.CancelAsync ();
				}
				catch (InvalidOperationException ex)
				{
					writeMessage ("Error stopping background monitoring of fMiCal", true);
				}
			}

		}

		// thread safe error messaging
		private void writeMessage (String text, bool severity)
		{
			if (this.lblError.InvokeRequired)
			{
				SetMessageCallback d = new SetMessageCallback (writeMessage);
				this.Invoke (d, new object[] { text, severity });
			}
			else
			{
				if (severity == true)
				{
					lblError.ForeColor = System.Drawing.Color.Red;
				}
				else
				{
					lblError.ForeColor = System.Drawing.Color.Gray;
				}
				lblError.Text = text;
				lblError.Refresh ();
			}

		}

		// read lock file. used to sync fMiCal with FreeMiCal
		private int readLock ()
		{
			try
			{
				StreamReader sr = new StreamReader ("fMiCal.run");
				String s = sr.ReadLine ();
				sr.Close ();

				int cur = Int32.Parse (s);
				return cur;
			}
			catch (Exception ex)
			{
				writeMessage ("Error reading lock file", true);
				return 0;
			}

		}

		private void deleteLock ()
		{
			if (File.Exists ("fMiCal.run"))
				File.Delete ("fMiCal.run");
		}

		private void backgroundWorker1_DoWork (object sender, DoWorkEventArgs e)
		{
			BackgroundWorker wk = (BackgroundWorker)sender;


			int record;
			int currentRecord = 0;
			
			do
			{
				if (wk.CancellationPending)
				{
					e.Cancel = true;
					return;
				}

				record = readLock ();
				if (record > currentRecord)
					currentRecord = record;

				if (record < 0)
					return;

				int percentDone = currentRecord *100 / (int)numEnd.Value;
				percentDone = (percentDone > 100) ? 100 : percentDone;
				try
				{
					wk.ReportProgress (percentDone);
				}
				catch (Exception ex)
				{
					writeMessage ("background worker: exception reporting progress: " + percentDone, true);
				}

				System.Threading.Thread.Sleep (500);
			} while (record < numEnd.Value);

			wk.ReportProgress (100);
		}

		private void backgroundWorker1_ProgressChanged (object sender, ProgressChangedEventArgs e)
		{
			progressBar1.Value = e.ProgressPercentage;
			progressBar1.Refresh ();

			lblError.Text = "";
			lblError.Refresh ();
		}

		private void backgroundWorker1_RunWorkerCompleted (object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Error != null)
			{
				writeMessage (e.Error.Message, true);
			}
			else if (e.Cancelled)
			{
				writeMessage ("Export cancelled !", false);

				progressBar1.Visible = false;
				progressBar1.Refresh ();
			}
			else
			{
				int record = readLock ();
				if (record < 0)								// check if fMiCal exited with an error
				{
					writeMessage ("Error exporting calendar! Is the profile name correct?", true);

					progressBar1.Visible = false;
					progressBar1.Refresh ();
				}
				else
					writeMessage ("Export finished !", false);
			}

			// restore button and reporting setting
			btnExport.Text = "Fr&ee them...";
			btnExport.Refresh ();

			deleteLock ();									// reset synchronisation
		}

	}
}