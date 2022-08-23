using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Emc.InputAccel.QuickModule.ClientScriptingInterface;
using Emc.InputAccel.ScriptEngine.Scripting;
using log4net;
using System.IO;
using System.Reflection;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace CNO.BPA.CreateMDWCover
{
    public class TaskEvents : ITaskEvents
    {
        EnvelopeDetail[] envelopeDetail = null;
        string inputSource = string.Empty;
        private log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); 
             
        public void ExecuteTask(ITaskInformation taskInfo)
        {
            try
            {
                //initialize the logger
                FileInfo fi = new FileInfo(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "CNO.BPA.CreateMDWCover.dll.config"));
                log4net.Config.XmlConfigurator.Configure(fi);   
                //initLogging();
                log.Info("Beginning the ExecuteTask method");
                //get envelope level details
                getEnvelopeDetails(taskInfo);
                log.Debug("Finished generating cover sheet and inserting as page 1 of the document");
                log.Info("Completed the ExecuteTask method");
            }
            catch (Exception ex)
            {
                log.Error("Error within the ExecuteTask method: " + ex.Message, ex);
                throw ex;
            }
        }
        private void getEnvelopeDetails(ITaskInformation taskInfo)
        {
            try
            {
                log.Debug("Preparing to loop through each envelope within the batch to create a coversheet using MDW data values");
                envelopeDetail = new EnvelopeDetail[taskInfo.Task.TaskRoot.ChildCount(3)];
                int i = 0;
                foreach (IBatchNode node in taskInfo.Task.TaskRoot.Children(3))
                {
                    envelopeDetail[i] = new EnvelopeDetail();

                    foreach (IWorkflowStep wfStep in taskInfo.Task.Batch.WorkflowSteps)
                    {
                        if (wfStep.Name.ToUpper() == "IMPORT")
                        {
                            int countValues = int.Parse(node.Values(wfStep).GetString("NumXMLValues", "0"));
                            if (countValues > 0)
                            {
                                envelopeDetail[i].xmlValues = new Dictionary<string, string>();
                                //we need to pull in all of the xml values and their tags
                                for (int x = 0; x <= (countValues - 1); x++)
                                {
                                    string name = node.Values(wfStep).GetString("Tag" + x.ToString() + "_Level3", "");
                                    string value = node.Values(wfStep).GetString("XMLValue" + x.ToString() + "_Level3", "");

                                    envelopeDetail[i].xmlValues.Add(name, value);
                                }
                                string coverSheetPath = createCoverSheet(envelopeDetail[i]);
                                if (File.Exists(coverSheetPath))
                                {
                                    insertCoverSheet(node, taskInfo, coverSheetPath);
                                }
                                else
                                {
                                    //if no file was generated, then we need to error and record the issue
                                    log.Error("Coversheet TIFF file could not be found");
                                }

                            }
                            //we need to handle the situation where this MDW uses a delimited file instead of XML
                            else
                            {
                                log.Error("No XML values exist for this batch, so no cover sheet can be generated at this time.");
                            }
                            break;
                        }
                    }

                    i++;
                }
            }
            catch (Exception ex)
            {
                log.Error("Error within the getEnvelopeDetails method: " + ex.Message, ex);
                throw ex;
            }
        }
        private string createCoverSheet(EnvelopeDetail envelopeDetail)
        {
            try
            {
                StringBuilder coversheetContents = new StringBuilder();
                log.Debug("Building the string that will become the tiff");
                foreach (KeyValuePair<string, string> content in envelopeDetail.xmlValues)
                {
                    //first grab the lenths of both values
                    int keyLength = content.Key.Length;
                    int valueLength = content.Value.Length;
                    string tempValue = string.Empty;

                    if ((keyLength + 5 + valueLength) <= 80)
                    {
                        coversheetContents.AppendLine(content.Key + "  -  " + content.Value);
                    }
                    else
                    {
                        string[] words = content.Value.Split(Convert.ToChar(" "));
                        foreach (string word in words)
                        {
                            if (keyLength + 5 + tempValue.Length >= 80)
                            {
                                coversheetContents.AppendLine(content.Key + "  -  " + tempValue + " " + word);
                                //clear the temp value for the next block
                                tempValue = string.Empty;
                            }
                            else
                            {
                                tempValue += word + " ";
                            }
                        }
                        //handle the last block of text
                        coversheetContents.AppendLine(content.Key + "  -  " + tempValue);
                    }
                }
                log.Debug("String has been assembled, calling the createTIFF method");
                TIFFBuilder tiffBuilder = new TIFFBuilder();
                string coverSheetPath = tiffBuilder.createTIFF(coversheetContents.ToString());
                log.Debug("Temp TIFF file has been created here: " + coverSheetPath);
                return coverSheetPath;
            }
            catch (Exception ex)
            {
                log.Error("Error within the createCoverSheet method: " + ex.Message, ex);
                throw ex;
            }

        }
        private void insertCoverSheet(IBatchNode node, ITaskInformation taskInfo, string FilePath)
        {
            try
            {
                if (FilePath.Length > 0)
                {
                    //we need to insert the cover sheet into the node
                    foreach (IBatchNode doc in node.Children(1))
                    {
                        IBatchNode page = doc.Insert(0);
                        foreach (IWorkflowStep wfStep in taskInfo.Task.Batch.WorkflowSteps)
                        {
                            if (wfStep.Name.ToUpper() == "STANDARD_MDF")
                            {
                                IIAValueProvider pagevalues = page.Values(wfStep);
                                pagevalues.SetFile("CurrentImgBW", FilePath);
                                break;
                            }
                        }
                    }
                    //once we complete the inserts, we can delete the temp file
                    File.Delete(FilePath);
                }
            }
            catch (UnauthorizedAccessException uaex)
            {
                log.Error("Failed deleting the temp file due to an UnauthorizedAccessException");
            }
            catch (Exception ex)
            {
                log.Error("Error within the insertCoverSheet method: " + ex.Message, ex);
                throw ex;
            }
        }
    }
}
