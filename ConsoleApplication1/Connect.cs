using System;
using TDAPIOLELib;
using System.Data;
using System.Drawing;

namespace ConsoleApplication1
{
    class Connect
    {
        private string domain;
        private string project;
        private string userName;
        private string password;
        private string url;

        public Connect(string domain, string project, string userName, string password, string url)
        {
            this.domain = domain;
            this.project = project;
            this.userName = userName;
            this.password = password;
            this.url = url;
        }

        public TDConnection connectToProject()
        {
            try
            {
                TDConnection qctd = new TDConnection();
                qctd.InitConnectionEx(url);
                qctd.ConnectProjectEx(domain, project, userName, password);
                return qctd;

            }
            catch
            {
                Console.WriteLine("Failed to connect");
                return null;
            }
        }

        public static string getCycleName(TDConnection tdc, string cycle_id)
        {
            CycleFactory cfc = (CycleFactory)tdc.CycleFactory;
            TDFilter filter = (TDFilter)cfc.Filter;
            filter["RCYC_ID"] = cycle_id;
            //TDAPIOLELib.Cycle cycle = cfc.Filter(filter.Text);

            List cycles = (List)cfc.NewList(filter.Text);
            string cycle_name;
            foreach (TDAPIOLELib.Cycle cl in cycles)
            {
                cycle_name = cl.Name;
                return cycle_name;
            }

            return null;
        }



    }
}
