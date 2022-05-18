using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;

namespace UploadDataDynamics
{
    internal class Class1
    {
        public CrmServiceClient connect()
        {

            var url = "https://orge29362b6.crm8.dynamics.com/";
            var userName = "Nehagarg@dyggdf.onmicrosoft.com";
            var password = "@N123eha";

            string conn = $@"  Url = {url}; AuthType = OAuth;
            UserName = {userName};
            Password = {password};
            AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;
            RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;
            LoginPrompt=Auto;
            RequireNewInstance = True";


            var svc = new CrmServiceClient(conn);
            return svc;
        }
    }
}
