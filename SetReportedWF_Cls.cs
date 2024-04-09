using System;
using System.Linq;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;

namespace SetReportedWF
{

    [ComVisible(true)]
    [ProgId("SetReportedWF.SetReportedWF_Cls")]
    public class SetReportedWF_Cls : IWorkflowExtension
    {

        INautilusServiceProvider sp;
        private IDataLayer dal;
        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {

                sp = Parameters["SERVICE_PROVIDER"];

                //   var Id = Parameters["RESULT_ID"].Value;
                var rs = Parameters["RECORDS"];
                var id = rs.Fields["RESULT_ID"].Value;


                var ntlsCon = Utils.GetNtlsCon(sp);

                Utils.CreateConstring(ntlsCon);
                dal = new DataLayer();
                dal.Connect();
                var result = dal.GetResultById((long)id);
             //   Logger.WriteLogFile("Result Id " + result.ResultId, false);
                Test test = dal.GetTestById((long)result.TestId);
               // Logger.WriteLogFile("test Id " + test.TEST_ID, false);

                var count = test.Results.Count;
                //Logger.WriteLogFile("test.Results.Count" + count, false);
                //foreach (Result result1 in test.Results)
                //{
                  //  Logger.WriteLogFile("result description = " + result1.DESCRIPTION, false);
                //}

                Result resultToUpdate = test.Results.FirstOrDefault(x => x.DESCRIPTION.StartsWith("THIS RESULT"));
                //Logger.WriteLogFile("resultToUpdate description = " + resultToUpdate.DESCRIPTION, false);


                if (resultToUpdate != null)
                {
                    var splited = resultToUpdate.DESCRIPTION.Split('=');
                    var value = splited[1];
                    if (value == " TRUE")
                    {
                        resultToUpdate.REPORTED = "T";
                    }

                    else if (value == " FALSE")
                    {
                        resultToUpdate.REPORTED = "F";
                    }

                    if (dal.HasChanges())
                    {
                        dal.SaveChanges();
                    }
                }

                dal.Close();
            }
            catch (Exception ex)
            {

                Logger.WriteLogFile(ex);
            }
            finally
            {
                dal.Close();

            }
        }


    }
}
