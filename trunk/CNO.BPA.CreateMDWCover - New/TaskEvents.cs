using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using log4net;
using System.IO;
using System.Reflection;
using Emc.InputAccel.CaptureClient;


[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace CNO.BPA.CreateMDWCover
{
    public class CodeModule : CustomCodeModule
    {
        EnvelopeDetail[] envelopeDetail = null;
        string inputSource = string.Empty;
        private log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); 
             
        public CodeModule()
        {
            
        }

        public override void ExecuteTask(IClientTask task, IBatchContext batchContext)
        {
            try
            {
                //initialize the logger
                FileInfo fi = new FileInfo(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "CNO.BPA.CreateMDWCover.dll.config"));
                log4net.Config.XmlConfigurator.Configure(fi);
                //initLogging();
                log.Info("Beginning the ExecuteTask method");
                //get envelope level details
                getEnvelopeDetails(task, batchContext);
                log.Debug("Finished generating cover sheet and inserting as page 1 of the document");
                log.Info("Completed the ExecuteTask method");
                task.CompleteTask();
            }
            catch (Exception ex)
            {
                log.Error("Error within the ExecuteTask method: " + ex.Message, ex);
                task.FailTask(FailTaskReasonCode.GenericUnrecoverableError, ex);
                throw ex;
            }
        }

        public override void StartModule(ICodeModuleStartInfo startInfo)
        {
            startInfo.ShowStatusMessage("Try1");
        }


        private void getEnvelopeDetails(IClientTask task, IBatchContext batchContext)
        {
            try
            {
                log.Debug("Preparing to loop through each envelope within the batch to create a coversheet using MDW data values");

                IBatchNode currentNode = task.BatchNode;
                IBatchNodeCollection envelopes = currentNode.GetDescendantNodes(3);
                envelopeDetail = new EnvelopeDetail[envelopes.Count];

                int i = 0;
                // for each envelope
                foreach (IBatchNode node in envelopes)
                {
                    envelopeDetail[i] = new EnvelopeDetail();

                    // New approach to access step node in .Net code module.
                    IBatchNode importEnvNode = batchContext.GetStepNode(node, "IMPORT");
                    
                    
                    int countValues = int.Parse(importEnvNode.NodeData.ValueSet.ReadString("NumXMLValues", "0"));
                    if (countValues > 0)
                    {
                        envelopeDetail[i].xmlValues = new Dictionary<string, string>();
                        //we need to pull in all of the xml values and their tags
                        for (int x = 0; x <= (countValues - 1); x++)
                        {
                            string name = importEnvNode.NodeData.ValueSet.ReadString("Tag" + x.ToString() + "_Level3", "");
                            string value = importEnvNode.NodeData.ValueSet.ReadString("XMLValue" + x.ToString() + "_Level3", "");

                            envelopeDetail[i].xmlValues.Add(name, value);
                        }
                        string coverSheetPath = createCoverSheet(envelopeDetail[i]);
                        if (File.Exists(coverSheetPath))
                        {
                            insertCoverSheet(node, batchContext, coverSheetPath);
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
                    //break;


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
        private void insertCoverSheet(IBatchNode node, IBatchContext batchContext, string FilePath)
        {
            try
            {
                if (FilePath.Length > 0)
                {
                    //we need to insert the cover sheet into the node
                    foreach (IBatchNode doc in node.GetDescendantNodes(1))
                    {
                        IBatchNode page = doc.AddNewChild(0);
                        page = batchContext.GetStepNode(page, "STANDARD_MDF");
                        byte[] file = System.IO.File.ReadAllBytes(FilePath);
                        page.NodeData.ValueSet.WriteFileData("CurrentImgBW", file, "TIF");

                        break;
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
