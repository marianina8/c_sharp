/*
 * AUTHOR: Marian Montagnino
 * DATE: 01/21/2012
 * DESCRIPTION:
   ------------
	A multi-threaded console application in C# that:

	1. Reads a set of urls from the supplied file.
	2. Downloads the html content from the url.
	3. Writes the output to a separate file for each url.
	4. Logs the results for each url into a log file. (url, result code, output file, how long it took).
	5. Log statisitics about number of urls successfully downloaded, 
	   number of errors, length of time for program to run.

	The program waits until all urls have been processed / downloaded before exiting.  
 */

using System;
using System.Net;
using System.IO;
using System.Text;
using System.Threading;
using System.Collections.Generic;

namespace MoonValleyTest
{
	/*
	 *  CLASS NAME: UrlFetcher
	 *  DESCRIPTION: Fetches data from given URL
	 *
	 */
	public class UrlFetcher
	{
	    string url;
		Logger fetcherLog = new Logger();
		
		// Constructor
	    public UrlFetcher (string url)
	    {
	        this.url = url;
	    }
	
		// Fetch 
		// Input Parameters: i (index for unique filenames), filename (logfile name)
		// Output: none
	    public void Fetch(int i, string filename)
	    {
			fetcherLog.setFilename (filename);
	        using (WebClient client = new WebClient()) 
			{
				int result = 0;
				
				// Log Start Time
				DateTime start = DateTime.Now;
			    try {
					// Download data from URL
					byte[] dataBuffer = client.DownloadData(url);
					// Write the data into a file
					File.WriteAllBytes(string.Concat (string.Concat ("file",i.ToString()),".html"), dataBuffer);
		
				}
				catch (WebException webEx) {
					// Can write the error to log, or ignore and just note failure.
					result=-1;
		        }
				
				// Log End Time
				DateTime end = DateTime.Now;
				
				// Calculate Timespan
				TimeSpan ts = end.Subtract (start);
				
				// Log
				fetcherLog.WriteToLog(url,i,result,ts);
				
			}
	    }

	}

	
	/*
	 *  CLASS NAME: Logger
	 *  DESCRIPTION: Sets all relevant data for logging
	 *
	 */

	public class Logger
	{
		private string url;
		private int index;
		private int result;
		private DateTime start;
		private DateTime end;	
		private TimeSpan elapsed;
		private string filename;
		private static readonly object sync = new object();
		
		// Constructor
		public Logger()
		{	
		}
		
		// Set Logfile name
		public void setFilename(string _filename)
		{
			filename=_filename;	
		}
		
		
		// Write to log
		public void WriteToLog(string _url, int _index, int _result, System.TimeSpan elapsed)
		{
			lock(sync) // Lock to handle multiple threads accessing file.
			{
				// Write to log
				using (StreamWriter w = File.AppendText(filename))
		        {
		            w.WriteLine("\r\nLog Entry : ");
					w.WriteLine ("URL:\t"+_url);
					w.WriteLine ("Result:\t"+ ((String.Compare (_result.ToString (),"0")==0 ? "Success." : "Failure." )));
					w.WriteLine ("Elapsed Time:\t"+elapsed.ToString ());
					w.Flush();
		            w.Close();
		        }
			}
			
		}
		
		// Set and Get for Variables
		
		public void setUrl(string _url) 
		{
			url=_url;
		}
		public string getUrl() {return url;}
		public void setIndex(int _index) 
		{
			index=_index;
		}
		public int getIndex() {return index;}
		public void setResult(int _result) {
			result=_result;
		}
		public int getResult() {return result;}
		public void setStart(DateTime _start) {
			start=_start;	
		}
		public void setEnd(DateTime _end) {
			end=_end;	
			elapsed=end.Subtract(start);
		}
		public TimeSpan getElapsed() {return elapsed;}
		
	}
	
	
     /*
	 *  CLASS NAME: Main Class
	 *  DESCRIPTION: Pulls in urls from file and creates thread to download data, create unique files.
	 *
	 */
	class MainClass
	{
		
		public static void Main (string[] args)
		{
			// Can place these in Application Settings, Never hardcode filenames
			Console.WriteLine ("Please enter location of file containing test urls: ");
			string filename = Console.ReadLine();
			
			Console.WriteLine ("Please enter name of logfile: ");
			string logfile = Console.ReadLine();
			
			// Read in list of urls
			List<string> urls = new List<string>();
			using (StreamReader r = new StreamReader(filename))
			{
			    string url;
			    while ((url = r.ReadLine()) != null)
			    {
				    urls.Add(url);
			    }
			}
			
			int i=0;
			foreach(string url in urls)
			{
				// for each url in the file create a new thread to download data/save to file.
				UrlFetcher fetcher = new UrlFetcher (url);
				new Thread(() => fetcher.Fetch(i,logfile)).Start ();
				i++;
			}
			
		}

	}
}
