using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Globalization;
using EnvDTE;
using EnvDTE80;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //""

            try
            {
                // Register the IOleMessageFilter to handle any threading errors.
                // Not implemented in this example for simplicity.
                // See https://msdn.microsoft.com/en-us/library/ms228772
                //MessageFilter.Register();

                // create DTE
                string devenvPath;
                //devenvPath = @"C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe";
                devenvPath = @"C:\Program Files (x86)\Microsoft Visual Studio\2017\Common7\IDE\devenv.exe"; // change it to your actual path
                EnvDTE.DTE dte = CreateDteInstance(devenvPath);

                if (dte != null)
                {
                    // print edition
                    Console.WriteLine($"Edition: {dte.Edition}");
                    // show IDE
                    dte.MainWindow.Visible = true;
                    dte.UserControl = true;


                    string[] fileEntries = Directory.GetFiles(@"C:\Downloads\DB\DACPACs\");
                    foreach (string fileName in fileEntries)
                    {
                        createProjectsFromTemplates(dte, fileName);

                        System.Threading.Thread.Sleep(1000 * 10);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // turn off the IOleMessageFilter.
                //MessageFilter.Revoke();
            }
        }

        public static void createProjectsFromTemplates(EnvDTE.DTE dte, string filename)
        {
            try
            {
                // Create a solution with two projects in it, based on project 
                // templates.
                Solution2 soln = (Solution2)dte.Solution;
                 //= "abcd";
                string csTemplatePath;
                string csPrjPath = string.Format("C:\\Downloads\\DB\\MyDBProject\\{0}", Path.GetFileName(filename).Replace("dacpac", ""));
                // Get the project template path for a C# console project.
                // Console Application is the template name that appears in 
                // the right pane. "CSharp" is the Language(vstemplate) as seen 
                // in the registry.
                csTemplatePath = soln.GetProjectTemplate("SSDT.vstemplate", "Database");

                //System.Windows.Forms.MessageBox.Show("SQL template path: " + csTemplatePath);

                // Create a new C# console project using the template obtained 
                // above.
                Project prj = soln.AddFromTemplate(csTemplatePath, csPrjPath, Path.GetFileName(filename).Replace("dacpac", ""), false);

                string commandArg = string.Format(CultureInfo.InvariantCulture,
                    "/FileName {0} " +
                    "/GroupByType " +
                    "/GroupBySchema", filename);

                System.Threading.Thread.Sleep(1000 * 10);

                dte.ExecuteCommand("Project.SSDTImportDac", commandArg);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }

        /// <summary>Creates and returns a DTE instance of specified VS version.</summary>
        /// <param name="devenvPath">The full path to the devenv.exe.
        /// <returns>DTE instance or <see langword="null"> if not found.</see></returns>
        public static EnvDTE.DTE CreateDteInstance(string devenvPath)
        {
            EnvDTE.DTE dte = null;
            System.Diagnostics.Process proc = null;

            // start devenv
            ProcessStartInfo procStartInfo = new ProcessStartInfo();
            procStartInfo.Arguments = "-Embedding";
            procStartInfo.CreateNoWindow = true;
            procStartInfo.FileName = devenvPath;
            procStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            procStartInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(devenvPath);

            try
            {
                proc = System.Diagnostics.Process.Start(procStartInfo);
            }
            catch (Exception ex)
            {
                return null;
            }
            if (proc == null)
            {
                return null;
            }

            // get DTE
            dte = GetDTE(proc.Id, 120);

            return dte;
        }


        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <remarks>
        /// After starting devenv.exe, the DTE object is not ready. We need to try repeatedly and fail after the
        /// timeout.
        /// </remarks>
        /// <param name="processId">
        /// <param name="timeout">Timeout in seconds.
        /// <returns>
        /// Retrieved DTE object or <see langword="null"> if not found.
        /// </see></returns>
        private static EnvDTE.DTE GetDTE(int processId, int timeout)
        {
            EnvDTE.DTE res = null;
            DateTime startTime = DateTime.Now;

            while (res == null && DateTime.Now.Subtract(startTime).Seconds < timeout)
            {
                System.Threading.Thread.Sleep(1000);
                res = GetDTE(processId);
            }

            return res;
        }


        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);


        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <param name="processId">
        /// <returns>
        /// Retrieved DTE object or <see langword="null"> if not found.
        /// </see></returns>
        private static EnvDTE.DTE GetDTE(int processId)
        {
            object runningObject = null;

            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Marshal.ThrowExceptionForHR(CreateBindCtx(reserved: 0, ppbc: out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];

                    string name = null;

                    try
                    {
                        if (runningObjectMoniker != null)
                        {
                            runningObjectMoniker.GetDisplayName(bindCtx, null, out name);
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    Regex monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + processId, RegexOptions.IgnoreCase);
                    if (!string.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        break;
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return runningObject as EnvDTE.DTE;
        }

    }
}
