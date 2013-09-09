using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestApp
{
    class Program
    {
        /*
        static void Main(string[] args)
        {
            
            try
            {
                //ceate instance of dms object
                Com.Interwoven.WorkSite.iManage.IManDMS dms = new Com.Interwoven.WorkSite.iManage.ManDMSClass();
                if (dms != null)
                {
                    //add worksite server name
                    Com.Interwoven.WorkSite.iManage.IManSession session = dms.Sessions.Add("10.11.228.118"); //worksite server name
                    if (session != null)
                    {
                        //create server session
                        session.Login("mmaju", "!manage5"); //worksite username, password
                        if (session.Connected)
                        {
                            //Accesss Matter worklist and add a new workspace into it
                            Com.Interwoven.WorkSite.iManage.IManWorkArea workarea = session.WorkArea;
                            if (workarea != null)
                            {
                                Com.Interwoven.WorkSite.iManage.IManWorkspaces wss = workarea.RecentWorkspaces;
                                if (wss != null)
                                {
                                    //access preferred database
                                    if (session.PreferredDatabase != null)
                                    {
                                        //Create workspace
                                        CreateWorkspace(session.PreferredDatabase, wss);
                                    }
                                }
                            }

                            //create document

                            //logout
                            session.Logout();

                            //close dms
                            dms.Close();
                        }
                    }


                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + " - " + e.StackTrace);
            }

        }//*/


        /*
        /// <summary>
        /// Create workspace
        /// </summary>
        /// <param name="db"></param>
        private static void CreateWorkspace(Com.Interwoven.WorkSite.iManage.IManDatabase db, Com.Interwoven.WorkSite.iManage.IManWorkspaces wss)
        {
            try
            {
                Com.Interwoven.WorkSite.iManage.IManDocument doc = null;
                Com.Interwoven.WorkSite.iManage.IManDocument3 doc3 = null;

                //Create Workspace
                Com.Interwoven.WorkSite.iManage.IManWorkspace ws = db.CreateWorkspace();
                if (ws != null)
                {
                    ws.Name = "Sample Workspace - " + DateTime.Now.ToString();
                    ws.SubType = "work";
                    ws.Description = "Sample Workspace Description - " + DateTime.Now.ToString();

                    string filename = System.IO.Path.GetTempFileName();
                    Com.Interwoven.WorkSite.iManage.IManProfileUpdateResult result = ws.UpdateAllWithResults(filename);

                    //add workspace into matter worklist
                    wss.AddWorkspace(ws);

                    //Create a sub-folder inside newly created workspace
                    Com.Interwoven.WorkSite.iManage.IManDocumentFolder docfolder = ws.DocumentFolders.AddNewDocumentFolderInheriting("New Folder", "New Folder Description");
                    if (docfolder != null)
                    {
                        //Create new document
                        doc = docfolder.Database.CreateDocument();
                        if (doc != null)
                        {
                            doc.SetAttributeByID(Com.Interwoven.WorkSite.iManage.imProfileAttributeID.imProfileDescription, "TestApp");
                            doc.SetAttributeByID(Com.Interwoven.WorkSite.iManage.imProfileAttributeID.imProfileAuthor, db.Session.UserID);
                            doc.SetAttributeByID(Com.Interwoven.WorkSite.iManage.imProfileAttributeID.imProfileOperator, db.Session.UserID);
                            doc.SetAttributeByID(Com.Interwoven.WorkSite.iManage.imProfileAttributeID.imProfileType, "ANSI");
                            doc.SetAttributeByID(Com.Interwoven.WorkSite.iManage.imProfileAttributeID.imProfileClass, "DOC");

                            doc3 = doc as Com.Interwoven.WorkSite.iManage.IManDocument3;
                            if (doc3 != null)
                            {
                                //Checkin document into newly created folder
                                Com.Interwoven.WorkSite.iManage.IManCheckinResult CheckInResults = doc3.CheckInExToFolderAsNewDocumentWithResults(@"e:\TestApp.txt",
                                        Com.Interwoven.WorkSite.iManage.imCheckinOptions.imDontKeepCheckedOut,
                                        docfolder,
                                        Com.Interwoven.WorkSite.iManage.imHistEvent.imHistoryNew,
                                        "TestApp.exe", "Check in as new document", System.Environment.MachineName);
                            }
                        }
                    }


                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + e.StackTrace);
            }
        }//*/

    }
}
