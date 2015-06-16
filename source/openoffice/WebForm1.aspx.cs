using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;

using System.Security.Principal;
using System.Runtime.InteropServices;
using unoidl.com.sun.star.bridge;
using Microsoft.Win32;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.frame;
using System.Web.UI;
using System.IO.Ports;
using System.Net.NetworkInformation;
using System.Net;
using System.IO.Pipes;
namespace openoffice
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //string delfile = "ZIP_AREA_ROAD_" + DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.AddDays(-1).Day.ToString("00") + ".txt";

            string folderName = Server.MapPath("~/AreaRoadFile");

            // 取得資料夾內所有檔案
            foreach (String fname in System.IO.Directory.GetFiles(@"D:\openoffice\testfile"))
            {
                //if (fname == zip_name)
                //{
                //    continue;
                //}
                //File.Delete(Server.MapPath("~/AreaRoadFile/" + fname));
                TextBox1.Text += fname;
            }
                        
        }
        // IntPtr token = IntPtr.Zero;
        //WindowsImpersonationContext impersonatedUser = null;

        //private string _DomainName;
        //private string _strUserName;
        //private string _strPassword;

        //public openoffice(string DomainName, string strUserName, string strPassword)
        //{
        //    _DomainName = DomainName;
        //    _strUserName = strUserName;
        //    _strPassword = strPassword;

        //}

        //public void Login()
        //{
        //    try
        //    {
        //        bool result = LogonUser("Guest", "PC99122101",
        //                                "",
        //                                LogonSessionType.Network,
        //                                LogonProvider.Default,
        //                                out token);
        //        if (result)
        //        {
        //            WindowsIdentity id = new WindowsIdentity(token);
        //            impersonatedUser = id.Impersonate();
        //        }
        //        else
        //        {
        //        }
        //    }
        //    catch
        //    {
        //    }
        //    finally
        //    {
        //    }
        //}

        //public void Logout()
        //{
        //    try
        //    {
        //        if (impersonatedUser != null)
        //            impersonatedUser.Undo();
        //         Free the token
        //        if (token != IntPtr.Zero)
        //            CloseHandle(token);
        //    }
        //    catch
        //    {
        //    }
        //}

        //[DllImport("advapi32.dll", SetLastError = true)]
        //static extern bool LogonUser(
        //  string principal,
        //  string authority,
        //  string password,
        //  LogonSessionType logonType,
        //  LogonProvider logonProvider,
        //  out IntPtr token);
        //[DllImport("kernel32.dll", SetLastError = true)]
        //static extern bool CloseHandle(IntPtr handle);
        //enum LogonSessionType : uint
        //{
        //    Interactive = 2,
        //    Network,
        //    Batch,
        //    Service,
        //    NetworkCleartext = 8,
        //    NewCredentials
        //}
        //enum LogonProvider : uint
        //{
        //    Default = 0, // default for platform (use this!)
        //    WinNT35,     // sends smoke signals to authority
        //    WinNT40,     // uses NTLM
        //    WinNT50      // negotiates Kerb or NTLM
        //}
        public void start()
        {
            Process[] ps = Process.GetProcessesByName("soffice.exe");
            if (ps != null)
            {
                if (ps.Length > 0)
                    return;
                else
                {
                    Process p = new Process();
                    p.StartInfo.Arguments = "-headless -nofirststartwizard";
                    p.StartInfo.FileName = "soffice.exe";
                    p.StartInfo.CreateNoWindow = true;
                    bool result = p.Start();
                    if (result == false)
                        throw new InvalidProgramException("OpenOffice failed to start.");
                }
            }
            else
            {
                throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");
            }
        }
        public  void tes2()
        {
            try
            {
                Response.Write("line1 ");
                if (ConvertExtensionToFilterType(Path.GetExtension(AppDomain.CurrentDomain.BaseDirectory + "1.odp")) == null)
                    throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + AppDomain.CurrentDomain.BaseDirectory + "1.odp");

                start();
                Response.Write("over  StartOpenOffice() ");
                //Get a ComponentContext
                var xLocalContext = Bootstrap.bootstrap();
                //Get MultiServiceFactory
                unoidl.com.sun.star.lang.XMultiServiceFactory xRemoteFactory =
                    (unoidl.com.sun.star.lang.XMultiServiceFactory)
                        xLocalContext.getServiceManager();
                //Get a CompontLoader
                XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
                //Load the sourcefile

                XComponent xComponent = null;
                try
                {
                    xComponent = initDocument(aLoader,
                        PathConverter(AppDomain.CurrentDomain.BaseDirectory + "1.odp"), "_blank");

                    //Wait for loading
                    while (xComponent == null)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    Console.WriteLine("saveDocument  ");
                    // save/export the document
                    saveDocument(xComponent, AppDomain.CurrentDomain.BaseDirectory + "1.odp", PathConverter(AppDomain.CurrentDomain.BaseDirectory + "openOffice.ppt"));

                }
                catch
                {
                    throw;
                }
                finally
                {
                    xComponent.dispose();
                }
            }
            catch (System.Exception ex)
            {
                Response.Write(ex.ToString());
            }
        }

        public void other()
        {
            unoidl.com.sun.star.uno.XComponentContext xContext = null;

            //Microsoft.Win32.RegistryKey regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
            //        @"SOFTWARE\OpenOffice.org\UNO\InstallPath", false);

            RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
                                   @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
            if (regkey == null)
                return;

            string installPath = (string)regkey.GetValue("");
            if (installPath == null)
                return;

            string officePath = installPath + @"\soffice.exe";
            string officeParams = " -nodefault -nologo -nofirststartwizard";
            string unoParams = " -accept=pipe,name=officepipe1;urp;StarOffice.ServiceManager";

           // System.Environment.SetEnvironmentVariable(
           //     "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");

            System.Diagnostics.Process p = System.Diagnostics.Process.Start(officePath, officeParams + unoParams);

            XComponentContext xLocalContext = uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();
            XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
            XUnoUrlResolver xUrlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", xLocalContext);


            int i = 0;
            while (i < 20)
            {
                try
                {
                    xContext = (XComponentContext)xUrlResolver.resolve(
                        "uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");
                    if (xContext != null)
                        Response.Write("ok");
                        break;
                }
                catch (unoidl.com.sun.star.connection.NoConnectException ex)
                {
                    System.Threading.Thread.Sleep(100);
                    Response.Write(ex);
                }
                i++;
            }
            if (xContext == null)
                Response.Write("!ok");
                return;

            XMultiServiceFactory xMsf = (XMultiServiceFactory)xContext.getServiceManager();

            Object desktop = xMsf.createInstance("com.sun.star.frame.Desktop");
            XComponentLoader xLoader = (XComponentLoader)desktop;
        }


        protected void Button1_Click(object sender, EventArgs e)
        {
            tes2();
 //////////////          // other();

 //////////////           //return;
 //////////////          //Response.Write(System.Security.Principal.WindowsIdentity.GetCurrent().Name);
 //////////////         //  impersonatedUser
 //////////////           //Impersonate_User = new PClass.PImpersonate("網域名稱", "帳號", "密碼");
 //////////////           ////若是本機帳號則在網域的地方改成特定的機器名稱，如下
 //////////////          // PImpersonate Impersonate_User = new PImpersonate("", "test", "aa1234@@");
 //////////////         //PImpersonate Impersonate_User = new PImpersonate("PC99122101", "NETWORK SERVICE", "");
 //////////////          // PImpersonate Impersonate_User = new PImpersonate("STARGROUP", "pclu", "AIRair1590");
 //////////////          // PImpersonate Impersonate_User = new PImpersonate("STARGROUP", "pclu", "AIRair1590");
 //////////////         // PImpersonate Impersonate_User = new PImpersonate("STARGROUP", "Guest", "aa1234@@");
 //////////////         // PImpersonate Impersonate_User = new PImpersonate("STARGROUP", "Guest", "aa1234@@");
 //////////////          // Impersonate_User.loginpipe();//開始用前面設定的帳號執行程式

 //////////////           ////這邊就可以開始執行需要特定身份才能執行的程式囉

 //////////////           //Impersonate_User.Logout();//登出
            
 //////////////           if (ConvertExtensionToFilterType(Path.GetExtension(AppDomain.CurrentDomain.BaseDirectory + "1.odp")) == null)
 //////////////               throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + AppDomain.CurrentDomain.BaseDirectory + "1.odp");
 ////////////////if (resu)
 ////////////////           {
 //////////////         StartOpenOffice();

 ////////////////           //Get a ComponentContext
           
 //////////////           //    var xLocalContext = Bootstrap.bootstrap();

 ////////////////               //Get MultiServiceFactory
 //////////////            //   unoidl.com.sun.star.lang.XMultiServiceFactory xRemoteFactory =
 //////////////             //      (unoidl.com.sun.star.lang.XMultiServiceFactory)
 //////////////            //           xLocalContext.getServiceManager();
 ////////////////               //Get a CompontLoader
 ////////////////               XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
 ////////////////               //Load the sourcefile
 ////////////////               TextBox1.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
 //////////////             //RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
 //////////////             //                      @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
 //////////////                               //if (regkey == null)
 //////////////                               //{
 //////////////                               //    string a = "false";
 //////////////                               //}
 //////////////                              // var installPath = (string)regkey.GetValue("");
 //////////////                              //// var installPath = @"C:\Program Files (x86)\OpenOffice 4\program";
 //////////////                              // if (installPath == null)
 //////////////                              // {
 //////////////                              //    // string a = "false";
 //////////////                              // }
 //////////////                              // string officePath = installPath + @"\soffice.exe";
 //////////////                              // string officeParams = "-invisible -headless -nofirststartwizard";
                               
 //////////////                              // string unoParams = "";
                              

 //////////////                              // Environment.SetEnvironmentVariable(
 //////////////                              //     "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");

 //////////////                              // Process p = new Process();
 //////////////                              // p.StartInfo.Arguments = officeParams + unoParams;
 //////////////                              // p.StartInfo.FileName = officePath;
                                
 //////////////                              //// Process p = Process.Start(officePath, officeParams + unoParams);
 //////////////                              // p.StartInfo.CreateNoWindow = true;
 //////////////                              // p.StartInfo.UseShellExecute = false;
 //////////////                              // p.StartInfo.RedirectStandardOutput = true;
 //////////////                              // p.StartInfo.RedirectStandardError = true;
 //////////////                              // p.Start();
 //////////////           //RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
 //////////////           //                    @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
 //////////////           ////if (regkey == null)
 //////////////           ////{
 //////////////           ////    string a = "false";
 //////////////           ////}
 //////////////           //var installPath = (string)regkey.GetValue("");
 //////////////           //// var installPath = @"C:\Program Files (x86)\OpenOffice 4\program";
 //////////////           //if (installPath == null)
 //////////////           //{
 //////////////           //    // string a = "false";
 //////////////           //}
 //////////////           //string officePath = installPath + @"\soffice.exe";
 //////////////           //string officeParams = "-invisible -headless -nofirststartwizard";
 //////////////           //// string officeParams = "-headless -nologo -nofirststartwizard -terminate_after_init-invisible";

 //////////////           //string unoParams = "";
 //////////////           ////string unoParams = "-accept=socket;";
 //////////////           ////string unoParams = "-accept=socket,host=localhost,port=80;urp;";
 //////////////           ////string unoParams = "-accept=socket,host=localhost,port=8100;urp;";
 //////////////           ////string unoParams = "-accept='socket,host=localhost,port=8100;urp;StarOffice.Service'";
 //////////////           ////string unoParams = "-accept='socket,host=localhost,port=8100;urp;StarOffice.Service'";
 //////////////           ////soffice -invisible -headless -nofirststartwizard "-accept=socket,host=localhost,port=2002;urp;"


 //////////////           //Environment.SetEnvironmentVariable(
 //////////////           //    "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");

 //////////////           //Process p = new Process();
 //////////////           //p.StartInfo.Arguments = officeParams + unoParams;
 //////////////           //p.StartInfo.FileName = officePath;

 //////////////           //// Process p = Process.Start(officePath, officeParams + unoParams);
 //////////////           //p.StartInfo.CreateNoWindow = true;
 //////////////           //p.StartInfo.UseShellExecute = false;
 //////////////           //p.StartInfo.RedirectStandardOutput = true;
 //////////////           //p.StartInfo.RedirectStandardError = true;
 //////////////           //p.Start();
 //////////////            //Process p2 = new Process();
 //////////////            // Process[] ps = Process.GetProcessesByName("soffice.exe");
 //////////////            // if (ps != null)
 //////////////            // {
 //////////////            //     if (ps.Length > 0)
 //////////////            //         return;
 //////////////            //     else
 //////////////            //     {
 //////////////            //         //Process p = new Process();
 //////////////            //         //p = new Process();
 //////////////            //         //p.StartInfo.Arguments = "-invisible -headless -nofirststartwizard -accept=socket,host=localhost,port=8100;urp;StarOffice.ServiceManger";
 //////////////            //         p2.StartInfo.Arguments = "-headless -nologo -norestore -accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManger";
 //////////////            //         // p.StartInfo.Arguments = "-headless -nologo -norestore -accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManger";
 //////////////            //         //p.StartInfo.Arguments = "-headless -nofirststartwizard";
 //////////////            //         //RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
 //////////////            //         //               @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
 //////////////            //         //if (regkey != null)
 //////////////            //         //{
 //////////////            //         //     filename = (string)regkey.GetValue("");
 //////////////            //         //}
 //////////////            //         p2.StartInfo.FileName = @"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe";
 //////////////            //         p2.StartInfo.CreateNoWindow = true;
 //////////////            //         // p.StartInfo.UseShellExecute = false;
 //////////////            //         //p.StartInfo.RedirectStandardOutput = true;
 //////////////            //         //p.StartInfo.RedirectStandardError = true;
 //////////////            //         Response.Write("p.start");

 //////////////            //         bool result = p2.Start();
 //////////////            //         Response.Write("after p.start");
 //////////////            //         //  p.WaitForExit() ;
 //////////////            //         //  p.Close() ;
 //////////////            //         // p.Dispose();
 //////////////            //         if (result == false)
 //////////////            //             throw new InvalidProgramException("OpenOffice failed to start.");

 //////////////            //     }
 //////////////            // }

 //////////////         // string[] port = SerialPort.GetPortNames();

 //////////////             //if (p2 != null)
 //////////////             //{

 //////////////           IPGlobalProperties ipGlobal = IPGlobalProperties.GetIPGlobalProperties();

 //////////////           IPEndPoint[] endPoint = ipGlobal.GetActiveTcpListeners();
 //////////////           int num = 0;
 //////////////           //foreach (IPEndPoint iEndPoint in endPoint)
 //////////////           //{
 //////////////           //    if (iEndPoint.Address.ToString() == "127.0.0.1" && iEndPoint.Port == 8100)
 //////////////           //    {
 //////////////                   // string port = String.Format("Listening Address={0},Port={1}", "127.0.0.1", 2002);


 //////////////         //  String[] listOfPipes = System.IO.Directory.GetFiles(@"\.\localpipe\");

 //////////////                 //  num++;

 //////////////           //var asyncResult = pipeServer.BeginWaitForConnection(EndWait, this);

 //////////////           //if (asyncResult.AsyncWaitHandle.WaitOne(5000))
 //////////////           //{
 //////////////           //    pipeServer.EndWaitForConnection(asyncResult);

 //////////////           //     ...
 //////////////           //}
 //////////////           StartOpenOffice();

 //////////////           XComponentContext xLocalContext =
 //////////////               uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();
 //////////////           // String sUnoIni = "file:///C:/Program Files (x86)/OpenOffice%204/program/uno.ini";
 //////////////           //// String sUnoIni = @"C:\Program Files (x86)\OpenOffice 4\program\uno.ini";//@"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe"

 //////////////           // XComponentContext xLocalContext =
 //////////////           //     uno.util.Bootstrap.defaultBootstrap_InitialComponentContext(sUnoIni, null);
 //////////////           // XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

 //////////////           // XUnoUrlResolver xUrlResolver =
 //////////////           //     (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
 //////////////           //         "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

 //////////////           // XMultiServiceFactory multiServiceFactory =
 //////////////           //     (XMultiServiceFactory)xUrlResolver.resolve(
 //////////////           //          "uno:pipe,name=foo;urp;StarOffice.ServiceManager");

 //////////////           // NamedPipeClientStream pipeStream = new NamedPipeClientStream("testpipe");
 //////////////           //  pipeStream.Connect();

 //////////////           XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

 //////////////           XUnoUrlResolver xUrlResolver =
 //////////////               (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
 //////////////                   "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

 //////////////           //XMultiServiceFactory multiServiceFactory =
 //////////////           //   (XMultiServiceFactory)xUrlResolver.resolve(
 //////////////           //        "uno:socket,host=localhost,port=80;StarOffice.ServiceManager");

 //////////////           //XMultiServiceFactory multiServiceFactory =
 //////////////           //                      (XMultiServiceFactory)xUrlResolver.resolve(
 //////////////           //                           "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");

 //////////////           ////////XMultiServiceFactory multiServiceFactory =
 //////////////           ////////                     (XMultiServiceFactory)xUrlResolver.resolve(
 //////////////           ////////                           "uno:pipe,name=localpipe;urp;StarOffice.ServiceManager");

 //////////////           XComponentContext multiServiceFactory =
 //////////////                                      (XComponentContext)xUrlResolver.resolve(
 //////////////                                            "uno:pipe,name=testpipe;urp;StarOffice.ComponentContext");

 //////////////           XMultiServiceFactory multiServiceFactory1 = (XMultiServiceFactory)multiServiceFactory.getServiceManager();
 //////////////           //xContext = (XComponentContext)xUrlResolver.resolve(
 //////////////           //        "uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");


 //////////////           //XMultiServiceFactory multiServiceFactory =
 //////////////           //       (XMultiServiceFactory)xUrlResolver.resolve(
 //////////////           //            "uno:socket,host=localhost,port=2002;urp;StarOffice.ServiceManager");
 //////////////           XComponentLoader componentLoader = (XComponentLoader)multiServiceFactory1.createInstance("com.sun.star.frame.Desktop");
 //////////////           //XComponentLoader componentLoader = null;
 //////////////           XComponent xComponent = null;
 //////////////           try
 //////////////           {
 //////////////               xComponent = initDocument(componentLoader,
 //////////////                   PathConverter(AppDomain.CurrentDomain.BaseDirectory + "1.odp"), "_blank");
 //////////////               //Wait for loading
 //////////////               while (xComponent == null)
 //////////////               {
 //////////////                   System.Threading.Thread.Sleep(1000);
 //////////////               }

 //////////////               // save/export the document
 //////////////               saveDocument(xComponent, AppDomain.CurrentDomain.BaseDirectory + "1.odp", PathConverter(AppDomain.CurrentDomain.BaseDirectory + "openOffice.ppt"));
 //////////////               Response.Write("ok!");
 //////////////               // Impersonate_User.Logout();//登出
 //////////////           }
 //////////////           catch
 //////////////           {
 //////////////               throw;
 //////////////           }
 //////////////           finally
 //////////////           {

 //////////////               //if (xComponent != null)
 //////////////               //{
 //////////////               //    Marshal.FinalReleaseComObject(xComponent);
 //////////////               //}
 //////////////               //if (xLocalContext != null)
 //////////////               //{
 //////////////               //    //wrkBook.Close(false); //忽略尚未存檔內容，避免跳出提示卡住
 //////////////               //    Marshal.FinalReleaseComObject(xUrlResolver);
 //////////////               //}
 //////////////               //if (xUrlResolver != null)
 //////////////               //{

 //////////////               //    Marshal.FinalReleaseComObject(xUrlResolver);
 //////////////               //}
 //////////////               //if (multiServiceFactory != null)
 //////////////               //{

 //////////////               //    Marshal.FinalReleaseComObject(multiServiceFactory);
 //////////////               //}
 //////////////               //if (p != null)
 //////////////               //{
 //////////////               //    p.Close();//关闭进程
 //////////////               //    p.Dispose();//释放资源
 //////////////               //    Marshal.FinalReleaseComObject(p);
 //////////////               //}
 //////////////               //xComponent.dispose();
 //////////////               //xComponent = null;
 //////////////               //xLocalContext = null;
 //////////////               //xUrlResolver = null;
 //////////////               //multiServiceFactory = null;
 //////////////               //componentLoader = null;
 //////////////               //pipeServer.Close();
 //////////////               //p.Kill();
 //////////////               //p.WaitForExit();//阻塞等待进程结束
 //////////////               //p = null;
 //////////////               //p.Close();//关闭进程
 //////////////               //p.Dispose();//释放资源

 //////////////           }
 //////////////                   //break;
 //////////////           //    }
              
 //////////////           //    else
 //////////////           //    {
 //////////////           //    //    Response.Write("I can not connecting socket!");
 //////////////           //    }
 //////////////           //}
 //////////////           //if (num < 1)
 //////////////           //{
 //////////////           //    Response.Write("I can not connecting socket!");
 //////////////           //   // p.Kill();
 //////////////           //}
 //////////////             //}
 //////////////           //}
 //////////////          // else
 //////////////          // {
 //////////////           //    Response.Write(System.Security.Principal.WindowsIdentity.GetCurrent().Name);
 //////////////          // }

        }
        void EndWait(IAsyncResult iar)
        {
            var state = iar.AsyncState; // fetch state -> cast to desired type
            //do something when client connected
        }
        Process p = new Process();
        NamedPipeServerStream pipeServer;
        private void StartOpenOffice()
        {

            string filename = "";
            Process[] ps = Process.GetProcessesByName("soffice.exe");
            if (ps != null)
            {
                if (ps.Length > 0)
                    return;
                else
                {
                    //AnonymousPipeServerStream pipeStream = new AnonymousPipeServerStream(PipeDirection.Out, HandleInheritability.Inheritable);
                    // pipeServer = new NamedPipeServerStream("apipe", PipeDirection.Out, 10, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
                  // NamedPipeServerStream pipeStream = new NamedPipeServerStream("testpipe");
                   
                   
                   // NamedPipeServerStream pipeStream = new NamedPipeServerStream("locapipe");
                    //pipeStream
                    //Process p = new Process();
                    p = new Process();
                    //p.StartInfo.Arguments = "-invisible -headless -nofirststartwizard -accept=socket,host=localhost,port=8100;urp;StarOffice.ServiceManger";

                   // p.StartInfo.Arguments = "-headless -nologo -norestore -nofirststartwizard -accept=\"socket,host=localhost,port=8100;urp;\"";
                   // p.StartInfo.Arguments = "-headless -nologo -norestore -nofirststartwizard -accept=\"socket,host=localhost,port=2002;urp;\"StarOffice.ServiceManger";
                    //process.StartInfo.Arguments = pipeStream.GetClientHandleAsString();

                    //p.StartInfo.Arguments = "-accept=\"pipe,name=locapipe;urp;StarOffice.ComponentContext\" -nologo -headless -nofirststartwizard -invisible";
                   // p.StartInfo.Arguments = @"-nologo -headless -nofirststartwizard -invisible -accept=pipe,name=apipe ";
                   // p.StartInfo.Arguments = "-accept=\"pipe,name=localpipe;urp;StarOffice.ServiceManager\" -nologo -headless -nofirststartwizard -invisible";




                   // p.StartInfo.Arguments = "-headless -nologo -norestore -nofirststartwizard -accept=pipe,name=localpipe;urp;";



                    p.StartInfo.Arguments = "-headless -nologo -norestore -nofirststartwizard -accept=pipe,name=testpipe;urp;StarOffice.ServiceManger";
                    //p.StartInfo.Arguments = @"cd C:\Program Files (x86)\OpenOffice 4\program\soffice.exe -headless -nologo -norestore -accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManger";
                    // p.StartInfo.Arguments = "-headless -nologo -norestore -accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManger";
                    //p.StartInfo.Arguments = "-headless -nofirststartwizard";
                    //RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
                    //               @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
                    //if (regkey != null)
                    //{
                    //     filename = (string)regkey.GetValue("");
                    //}
                   // p.StartInfo.FileName = "cmd.exe";
                    p.StartInfo.FileName = @"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe";
                    p.StartInfo.CreateNoWindow = true;
                    p.StartInfo.RedirectStandardInput = true;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.StartInfo.UseShellExecute = false;
                   // p.StartInfo.LoadUserProfile = true;
                    // p.StartInfo.UseShellExecute = false;
                    //p.StartInfo.RedirectStandardOutput = true;
                    //p.StartInfo.RedirectStandardError = true;
                    Response.Write("p.start");
                   // p.StartInfo.Arguments = pipeStream.GetClientHandleAsString();
                    bool result = p.Start();
                    //pipeServer.WaitForConnection();
                    //bool a = pipeServer.IsConnected;
                    // p.WaitForExit(5000);
                    //if (p.HasExited == false)
                    //{
                    //    p.Kill();
                    //}

                    Response.Write(p.StartInfo.Arguments.ToString());
                    Response.Write("after p.start");
                    //  p.WaitForExit() ;
                    //  p.Close() ;
                    // p.Dispose();
                    if (result == false)
                        throw new InvalidProgramException("OpenOffice failed to start.");
                }
            }
            else
            {
                throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");
            }
        }


        private static XComponent initDocument(XComponentLoader aLoader, string file, string target)
        {
            PropertyValue[] openProps = new PropertyValue[1];
            openProps[0] = new PropertyValue();
            openProps[0].Name = "Hidden";
            openProps[0].Value = new uno.Any(true);


            XComponent xComponent = aLoader.loadComponentFromURL(
                file, target, 0,
                openProps);

            return xComponent;
        }


        private static void saveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            PropertyValue[] propertyValues = new PropertyValue[2];
            propertyValues = new PropertyValue[2];
            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue();
            propertyValues[1].Name = "Overwrite";
            propertyValues[1].Value = new uno.Any(true);
            //// Setting the filter name
            propertyValues[0] = new PropertyValue();
            propertyValues[0].Name = "FilterName";
            //propertyValues[0].Value = new uno.Any(ConvertExtensionToFilterType(Path.GetExtension(sourceFile)));
            propertyValues[0].Value = new uno.Any("MS PowerPoint 97 Vorlage");
            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        private static string PathConverter(string file)
        {
            if (file == null || file.Length == 0)
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return String.Format("file:///{0}", file.Replace(@"\", "/"));

        }
        //另存新檔的參數列表FilterList: https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
        public static string ConvertExtensionToFilterType(string extension)
        {
            switch (extension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";
                case ".xls":
                case ".xlsb":
                case ".ods":
                    return "calc_pdf_Export";
                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }
        
    }
}