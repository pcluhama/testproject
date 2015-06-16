using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;

using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Security;
using unoidl.com.sun.star.bridge;
using System.Net.Sockets;
using uno;
using unoidl.com.sun.star.connection;
using Microsoft.Win32;
using System.Security.Permissions;
using Microsoft.Win32.SafeHandles;
using System.Runtime.ConstrainedExecution;
using System.Security;
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
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.form.runtime;
using System.Net;
using System.IO.Pipes;
using System.Reflection;
namespace openoffice
{
    [PermissionSetAttribute(SecurityAction.Demand, Name = "FullTrust")]
    public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        private SafeTokenHandle()
            : base(true)
        {
        }

        [DllImport("kernel32.dll")]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);

        protected override bool ReleaseHandle()
        {
            return CloseHandle(handle);
        }
    }
    public class PImpersonate
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
            int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public extern static bool CloseHandle(IntPtr handle);

        // Test harness.
        // If you incorporate this code into a DLL, be sure to demand FullTrust.
        

        private string _DomainName;
        private string _strUserName;
        private string _strPassword;

        public PImpersonate(string DomainName, string strUserName, string strPassword)
        {
            _DomainName = DomainName;
            _strUserName = strUserName;
            _strPassword = strPassword;

        }

       




       //public static  XComponentContext boot()
       // { Type myType=typeof(PImpersonate);    
       //    Random rnd = new Random();
       //   //   String SOFFICE = 
       //    // myType.GetProperty( "os.name" ).startsWith( "Windows" ) ?"soffice.exe" : "soffice";
       //  String NOLOGO = "-nologo";
       //  String NODEFAULT = "-nodefault";
       //    String PIPENAME =
       //     "uno" +  (rnd.Next(10000) + 0xffff).ToString() ;
       //  String URL =
       //     "uno:pipe,name=" + PIPENAME + ";urp;StarOffice.ServiceManager";
       //  String CONNECTION =
       //     "-accept=pipe,name=" + PIPENAME + ";urp;StarOffice.ServiceManager";
        
       //  long SLEEPMILLIS = 500;
        
       // XComponentContext xContext = null;
       ////  unoidl.com.sun.star
       // //try
       // //{
       //     // create default local component context                
       // XComponentContext xLocalContext = Bootstrap.defaultBootstrap_InitialComponentContext();
       //   // Bootstrap..createInitialComponentContext(null);
       //     //createInitialComponentContext(null);
       //     // initial service manager
       //     XMultiComponentFactory xLocalServiceManager =
       //         xLocalContext.getServiceManager();

       //     // create a URL resolver
       //     var urlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
       //         "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

       //     // query for the XUnoUrlResolver interface
       //      //XUnoUrlResolver xUrlResolver =
       //      //    (XComponentContext)UnoRuntime.queryInterface(
       //      //    XUnoUrlResolver.class, urlResolver );

       //     // try to connect to office
       //     XComponentContext remoteServiceManager = null;
       //     // try {
       //     remoteServiceManager = (XComponentContext)urlResolver.resolve(URL);
       //     //}
       //     // catch (System.Exception e ) {
       //     // find office executable relative to this class's class loader
       //     //File fOffice = NativeLibraryLoader.getResource(
       //     //    Bootstrap.class.getClassLoader(), SOFFICE );

       //     //if ( fOffice != null ) {                        
       //     //     create call with arguments
       //     //    String[] cmdArray = new String[4];
       //     //    cmdArray[0] = fOffice.getPath();
       //     //    cmdArray[1] = NOLOGO;
       //     //    cmdArray[2] = NODEFAULT;
       //     //    cmdArray[3] = CONNECTION;

       //     //     start office process
       //     //    Runtime.getRuntime().exec( cmdArray );

       //     //     wait until office is started
       //     //    while ( remoteServiceManager == null ) {
       //     //        try {
       //     //             try to connect to office
       //     //            Thread.currentThread().sleep( SLEEPMILLIS );
       //     //            remoteServiceManager = xUrlResolver.resolve( URL );
       //     //        } catch ( com.sun.star.connection.NoConnectException ex ) {
       //     //             try to connect again
       //     //        }
       //     //    }
       //     //} else {
       //     //    throw new BootstrapException(
       //     //        "no office executable found!" );
       //     //}
       //     // }

       //     // XComponentContext
       //     if (remoteServiceManager != null)
       //     {
       //         //XPropertySet xPropertySet =
       //         //    (XPropertySet)UnoRuntime.queryInterface(
       //         //    XPropertySet.class, remoteServiceManager );
       //         //Object context =
       //         //    xPropertySet.getPropertyValue( "DefaultContext" );
       //         //xContext = (XComponentContext) UnoRuntime.queryInterface(
       //         //    XComponentContext.class, context);
       //     }
       //     //} catch ( java.lang.RuntimeException e ) {
       //     //    throw e;
       //     //} catch ( java.lang.Exception e ) {
       //     //    throw new BootstrapException( e );
       //     //}

       //     return xContext;

       // //}
       // }

        Process p;
        private void StartOpenOffice2()
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
                    p.StartInfo.LoadUserProfile = true;
                    // p.StartInfo.UseShellExecute = false;
                    //p.StartInfo.RedirectStandardOutput = true;
                    //p.StartInfo.RedirectStandardError = true;
                   // Response.Write("p.start");
                    // p.StartInfo.Arguments = pipeStream.GetClientHandleAsString();
                    bool result = p.Start();
                    //pipeServer.WaitForConnection();
                    //bool a = pipeServer.IsConnected;
                    // p.WaitForExit(5000);
                    //if (p.HasExited == false)
                    //{
                    //    p.Kill();
                    //}

                  // Response.Write(p.StartInfo.Arguments.ToString());
                  //  Response.Write("after p.start");
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
        public static void Display(Int32 indent, string format, params object[] param)
        {
           // Console.Write(new string(' ', indent * 2));
           // Console.WriteLine(format, param);
        }
        public void loginpipe()
        {
            bool result =false;
            SafeTokenHandle safeTokenHandle;
            try
            {
                 const int LOGON32_PROVIDER_DEFAULT = 0;
                //This parameter causes LogonUser to create a primary token.
                const int LOGON32_LOGON_INTERACTIVE = 2;

                // Call LogonUser to obtain a handle to an access token.
                result = LogonUser(_strUserName, _DomainName, _strPassword,
                    LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT,
                    out safeTokenHandle);

                //Console.WriteLine("LogonUser called.");

                if (result)
                {
                    //int ret = Marshal.GetLastWin32Error();
                    // Console.WriteLine("LogonUser failed with error code : {0}", ret);
                    //throw new System.ComponentModel.Win32Exception(ret);

                    using (safeTokenHandle)
                    {
                        //Console.WriteLine("Did LogonUser Succeed? " + (returnValue ? "Yes" : "No"));
                        // Console.WriteLine("Value of Windows NT token: " + safeTokenHandle);

                        // Check the identity.
                        //Console.WriteLine("Before impersonation: "
                        //    + WindowsIdentity.GetCurrent().Name);
                        // Use the token handle returned by LogonUser.
                        using (WindowsIdentity newId = new WindowsIdentity(safeTokenHandle.DangerousGetHandle()))
                        {
                            using (WindowsImpersonationContext impersonatedUser = newId.Impersonate())
                            {
                               // UnoInterfaceProxy
                                //StartOpenOffice2();
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

            System.Environment.SetEnvironmentVariable(
                "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");
            //System.Environment.SetEnvironmentVariable("URE_BOOTSTRAP", pathConverter(officePath + "\\fundamental.ini"));


//
            //Process pipeClient = new Process();
            //pipeClient.StartInfo.FileName = officePath;
            //AnonymousPipeServerStream pipeServer =
            //    new AnonymousPipeServerStream(PipeDirection.Out,
            //                                  HandleInheritability.Inheritable);
           // var pipe = new AnonymousPipeServerStream(PipeDirection.Out,
           //HandleInheritability.Inheritable);
            //var pipe = new AnonymousPipeServerStream(PipeDirection.In, HandleInheritability.Inheritable);

             //var pipeName = pipe.GetClientHandleAsString();

            // Pass the client process a handle to the server.
            //pipeClient.StartInfo.Arguments = pipeServer.GetClientHandleAsString();
            //string pipeName = pipeClient.StartInfo.Arguments;
            //pipeClient.StartInfo.UseShellExecute = false;
            //pipeClient.Start();


            // var startInfo = new ProcessStartInfo(officePath, officeParams + "-accept=pipe,name=" + pipeName + ";urp;StarOffice.ServiceManager");
            //var process = Process.Start(startInfo);



           // System.Diagnostics.Process pro = System.Diagnostics.Process.Start(officeParams + unoParams);
            System.Diagnostics.Process pro = System.Diagnostics.Process.Start(officePath, officeParams + unoParams);


            Int32 indent = 0;
            // Display information about the EXE assembly.
            Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
           string fulname =  a.FullName;
                                string codebase =  a.CodeBase;

            // Display the set of assemblies our assemblies reference.
                               // string[][] data = new string[100][]; 
            //Display(indent, "Referenced assemblies:");
            foreach (AssemblyName an in a.GetReferencedAssemblies())
            {
                string aname = an.Name.ToString();
                string aversion = an.Version.ToString();
                string aCultureInfoname = an.CultureInfo.Name.ToString();
                string aBitConverter = (BitConverter.ToString(an.GetPublicKeyToken()));
               // Display(indent + 1, "Name={0}, Version={1}, Culture={2}, PublicKey token={3}", an.Name, an.Version, an.CultureInfo.Name, (BitConverter.ToString(an.GetPublicKeyToken())));
               // string aname = an.Name;
              //string a =  "Name={0}, Version={1}, Culture={2}, PublicKey token={3}", an.Name.Name, an.Version, an.CultureInfo.Name, (BitConverter.ToString(an.GetPublicKeyToken()));
            } 




            XComponentContext xLocalContext = uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();


           // // create a URL resolver
           // XUnoUrlResolver xUrlResolver = UnoUrlResolver.create(xLocalContext);
           // // get remote context
           //// XComponentContext xRemoteContext = getRemoteContext(xUrlResolver);
           // Object context = xUrlResolver.resolve("-uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");




            XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();
            XUnoUrlResolver xUrlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", xLocalContext);
                                //System.Web.HttpContext.Current.Server.MapPath
                              string[] name=  xLocalServiceManager.getAvailableServiceNames();
                               //string name = "\\.\pipe\;
            int i = 0;
            while (i < 20)
            {
                try
                {
                   // xUrlResolver.resolve("uno:pipe,name=officepipe1;urp;StarOffice.ServiceManager");
                    //hFile = CreateFile("//./pipe/mypipe", GENERIC_WRITE,
                    //       FILE_SHARE_READ | FILE_SHARE_WRITE , NULL, OPEN_EXISTING,
                    //    FILE_ATTRIBUTE_NORMAL, NULL);
                    //xContext = (XComponentContext)xUrlResolver.resolve(
                    //    "uno:pipe,name= "+pipeName+ ";urp;StarOffice.ComponentContext");
                    xContext = (XComponentContext)xUrlResolver.resolve(
                        "uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");
                    if (xContext != null)
                        //Response.Write("ok");
                        break;
                }
                catch (unoidl.com.sun.star.connection.NoConnectException ex)
                {
                    System.Threading.Thread.Sleep(100);
                    //Response.Write(ex);
                }
                i++;
            }
            if (xContext == null)
                //Response.Write("!ok");
                return;
                            }
                        }
                        // Releasing the context object stops the impersonation
                        // Check the identity.
                        //Console.WriteLine("After closing the context: " + WindowsIdentity.GetCurrent().Name);
                    }
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                Console.WriteLine("Exception occurred. " + ex.Message);
            }
        }


        //public static void Main(string[] args)
        //{
        public void Login()
        {
            bool result =false;
            SafeTokenHandle safeTokenHandle;
            try
            {
                //string userName, domainName;
                // Get the user token for the specified user, domain, and password using the
                // unmanaged LogonUser method.
                // The local machine name can be used for the domain name to impersonate a user on this machine.
                //Console.Write("Enter the name of the domain on which to log on: ");
               // domainName = Console.ReadLine();

                //Console.Write("Enter the login of a user on {0} that you wish to impersonate: ", domainName);
                //userName = Console.ReadLine();

                //Console.Write("Enter the password for {0}: ", userName);

                const int LOGON32_PROVIDER_DEFAULT = 0;
                //This parameter causes LogonUser to create a primary token.
                const int LOGON32_LOGON_INTERACTIVE = 2;

                // Call LogonUser to obtain a handle to an access token.
                result = LogonUser(_strUserName, _DomainName, _strPassword,
                    LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT,
                    out safeTokenHandle);

                //Console.WriteLine("LogonUser called.");

                if (result)
                {
                    //int ret = Marshal.GetLastWin32Error();
                    // Console.WriteLine("LogonUser failed with error code : {0}", ret);
                    //throw new System.ComponentModel.Win32Exception(ret);

                    using (safeTokenHandle)
                    {
                        //Console.WriteLine("Did LogonUser Succeed? " + (returnValue ? "Yes" : "No"));
                        // Console.WriteLine("Value of Windows NT token: " + safeTokenHandle);

                        // Check the identity.
                        //Console.WriteLine("Before impersonation: "
                        //    + WindowsIdentity.GetCurrent().Name);
                        // Use the token handle returned by LogonUser.
                        using (WindowsIdentity newId = new WindowsIdentity(safeTokenHandle.DangerousGetHandle()))
                        {
                            using (WindowsImpersonationContext impersonatedUser = newId.Impersonate())
                            {
                                //result = true;
                                //StartOpenOffice();
                               // XComponentContext xcomponentcontext = Bootstrap.defaultBootstrap_InitialComponentContext();
                                // XComponentContext xcomponentcontext = Bootstrap.createInitialComponentContext(null);
 
                                  // create a connector, so that it can contact the office
                               //   XUnoUrlResolver urlResolver = UnoUrlResolver.create(xcomponentcontext);
 
                              //    Object initialObject = urlResolver.resolve(
                             //         "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
 
                                  //XMultiComponentFactory xOfficeFactory = (XMultiComponentFactory) UnoRuntime.queryInterface(
                                  //    XMultiComponentFactory.class, initialObject);
 
                                  // retrieve the component context as property (it is not yet exported from the office)
                                  // Query for the XPropertySet interface.
                                  //XPropertySet xProperySet = (XPropertySet) UnoRuntime.queryInterface( 
                                  //    XPropertySet.class, xOfficeFactory);
 
                                  //// Get the default context from the office server.
                                  //Object oDefaultContext = xProperySet.getPropertyValue("DefaultContext");
 
                                  //// Query for the interface XComponentContext.
                                  //XComponentContext xOfficeComponentContext = (XComponentContext) UnoRuntime.queryInterface(
                                  //    XComponentContext.class, oDefaultContext);
 
                                  //// now create the desktop service
                                  //// NOTE: use the office component context here!
                                  //Object oDesktop = xOfficeFactory.createInstanceWithContext(
                                  //     "com.sun.star.frame.Desktop", xOfficeComponentContext);


                                //System.Net.Sockets.Socket s = new System.Net.Sockets.Socket(
                                //AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
                                //
                                //////////String sUnoIni = "file:///C:/Program Files (x86)/OpenOffice 4/program/uno.ini";
                                //////////XComponentContext xLocalContext =
                                //////////    uno.util.Bootstrap.defaultBootstrap_InitialComponentContext(sUnoIni, null);
                                //////////XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

                                //////////XUnoUrlResolver xUrlResolver =
                                //////////    (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                                //////////        "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

                                ////////////XMultiServiceFactory multiServiceFactory =
                                ////////////    (XMultiServiceFactory)xUrlResolver.resolve(
                                ////////////         "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
                                //////////XMultiServiceFactory multiServiceFactory =
                                //////////    (XMultiServiceFactory)xUrlResolver.resolve(
                                //////////         "uno:pipe,name=some_unique_pipe_name;urp;StarOffice.ServiceManager");

                               ////// //Socket s = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
                               //////   String sUnoIni = "file:///C:/Program Files (x86)/OpenOffice 4/program/soffice.exe";

                               ////// XComponentContext localContext = Bootstrap.defaultBootstrap_InitialComponentContext();
                               ////// //XComponentContext localContext = Bootstrap.defaultBootstrap_InitialComponentContext(sUnoIni, null);
                               ////// XMultiComponentFactory xLocalServiceManager = localContext.getServiceManager();

                               ////// XUnoUrlResolver xUrlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext);
                               ////// Random rnd = new Random();
                               ////// String PIPENAME = "uno" + (rnd.Next(10000) + 0xffff).ToString();
                               ////// String URL ="uno:pipe,name=" + PIPENAME + ";urp;StarOffice.ServiceManager";
                               //////// XMultiServiceFactory multiServiceFactory = (XMultiServiceFactory)xUrlResolver.resolve(URL);
                               //////// XMultiServiceFactory multiServiceFactory = (XMultiServiceFactory)xUrlResolver.resolve("uno:pipe,host=localhost,port=8100;urp;StarOffice.ServiceManager");
                               ////// XComponentContext multiServiceFactory = (XComponentContext)xUrlResolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager"); 
                               //////// XMultiServiceFactory
                               //////// XComponentLoader componentLoader = (XComponentLoader)multiServiceFactory.createInstance("com.sun.star.frame.Desktop"); 


                                RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
                                    @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
                                //if (regkey == null)
                                //{
                                //    string a = "false";
                                //}
                                var installPath = (string)regkey.GetValue("");
                               // var installPath = @"C:\Program Files (x86)\OpenOffice 4\program";
                                if (installPath == null)
                                {
                                   // string a = "false";
                                }
                                string officePath = installPath + @"\soffice.exe";
                                string officeParams = "-invisible -headless -nofirststartwizard";
                                // string officeParams = "-headless -nologo -nofirststartwizard -terminate_after_init-invisible";

                                string unoParams = "";
                                //string unoParams = "-accept=socket;";
                                //string unoParams = "-accept=socket,host=localhost,port=80;urp;";
                                //string unoParams = "-accept=socket,host=localhost,port=8100;urp;";
                                //string unoParams = "-accept='socket,host=localhost,port=8100;urp;StarOffice.Service'";
                                //string unoParams = "-accept='socket,host=localhost,port=8100;urp;StarOffice.Service'";
                                //soffice -invisible -headless -nofirststartwizard "-accept=socket,host=localhost,port=2002;urp;"


                                Environment.SetEnvironmentVariable(
                                    "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");

                                Process p = new Process();
                                p.StartInfo.Arguments = officeParams + unoParams;
                                p.StartInfo.FileName = officePath;
                                
                               // Process p = Process.Start(officePath, officeParams + unoParams);
                                p.StartInfo.CreateNoWindow = true;
                                p.StartInfo.UseShellExecute = false;
                                p.StartInfo.RedirectStandardOutput = true;
                                p.StartInfo.RedirectStandardError = true;
                                p.Start();
                                XComponentContext xLocalContext =
                                    uno.util.Bootstrap.defaultBootstrap_InitialComponentContext();
                                XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

                                XUnoUrlResolver xUrlResolver =
                                    (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                                        "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

                                //XMultiServiceFactory multiServiceFactory =
                                //    (XMultiServiceFactory)xUrlResolver.resolve(
                                //         "uno:socket;StarOffice.ServiceManager");
                                XMultiServiceFactory multiServiceFactory =
                                   (XMultiServiceFactory)xUrlResolver.resolve(
                                        "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
                                //XMultiServiceFactory multiServiceFactory =
                                //    (XMultiServiceFactory)xUrlResolver.resolve(
                                //         "uno:socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
                                //XMultiServiceFactory multiServiceFactory =
                                //    (XMultiServiceFactory)xUrlResolver.resolve(
                                //         unoParams);
                                //XMultiServiceFactory multiServiceFactory =
                                //(XMultiServiceFactory)xUrlResolver.resolve(
                                //     "uno:pipe,name=some_unique_pipe_name;urp;StarOffice.ServiceManager");
                                XComponentLoader componentLoader = (XComponentLoader)multiServiceFactory.createInstance("com.sun.star.frame.Desktop"); 






                //////////////  bool value = false;
                //////////////             XComponent xComponent = null;
                //////////////                                //Get a ComponentContext
                //////////////    XComponentContext xContext = null;

                ////////////// RegistryKey regkey = Registry.LocalMachine.OpenSubKey(
                //////////////     @"SOFTWARE\OpenOffice\UNO\InstallPath", false);
                ////////////// if (regkey == null)
                ////////////// {
                //////////////     string a = "false";
                ////////////// }
                ////////////// var installPath = (string) regkey.GetValue("");
                ////////////// if (installPath == null)
                ////////////// {
                //////////////     string a = "false";
                ////////////// }
                ////////////// string officePath = installPath + @"\soffice.exe";
                ////////////// string officeParams = "-headless -nologo -nofirststartwizard -terminate_after_init-invisible";
                
                //////////////string  unoParams = "-accept='socket,host=localhost,port=8100;urp;StarOffice.Service'";



                //////////////Environment.SetEnvironmentVariable(
                //////////////    "URE_BOOTSTRAP", "vnd.sun.star.pathname:" + installPath + "/fundamental.ini");
                

                ////////////// Process p = Process.Start(officePath, officeParams + unoParams);
                
                ////////////// XComponentContext xLocalContext = Bootstrap.defaultBootstrap_InitialComponentContext();
                //////////////// xLocalContext = Bootstrap.bootstrap();
                ////////////// XMultiComponentFactory xLocalServiceManager = xLocalContext.getServiceManager();

                 //var xUrlResolver = (XUnoUrlResolver)xLocalServiceManager.createInstanceWithContext(
                 //                    "com.sun.star.bridge.UnoUrlResolver", xLocalContext);


                 //int i = 0;
                 //while (i < 20)
                 //{
                 //    try
                 //    {
                 //        xContext = (XComponentContext)xUrlResolver.resolve(
                 //            "uno:pipe,name=officepipe1;urp;StarOffice.ComponentContext");
                 //        if (xContext != null)
                 //            break;
                 //    }
                 //    catch (NoConnectException)
                 //    {
                 //        System.Threading.Thread.Sleep(1000);
                 //    }
                 //    i++;
                 //}
                 //if (xContext == null)
                 //{
                 //    string a = "false";
                 //}



    //             unoidl.com.sun.star.lang.XMultiComponentFactory xRemoteFactory =
   //                (unoidl.com.sun.star.lang.XMultiComponentFactory)
    //                   xLocalContext.getServiceManager();
                                
                 //var xMsf = (XMultiServiceFactory)xContext.getServiceManager();

                 //Object desktop = xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
     //            XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstanceWithContext("com.sun.star.frame.Desktop",xLocalContext);
                // Object desktop = xMsf.createInstance("com.sun.star.frame.Desktop");
     //            var xLoader = (XComponentLoader)aLoader;
                // xComponent = initDocument(xLoader,
                //                           PathConverter(inputFile), "_blank");

                // var xUrlResolver = (XUnoUrlResolver) xLocalServiceManager.createInstanceWithContext(
                //     "com.sun.star.bridge.UnoUrlResolver", xLocalContext);

                                //System.Collections.Hashtable ht = new System.Collections.Hashtable();
                                //ht.Add("SYSBINDIR", @"C:\Program Files (x86)\OpenOffice 4\program");
                                //unoidl.com.sun.star.uno.XComponentContext xLocalContext =
                                //         uno.util.Bootstrap.defaultBootstrap_InitialComponentContext(
                                //        @"C:\Program Files (x86)\OpenOffice 4\program\uno.ini", ht.GetEnumerator());

                                //unoidl.com.sun.star.bridge.XUnoUrlResolver xURLResolver =
                                //                  (unoidl.com.sun.star.bridge.XUnoUrlResolver)
                                //                          xLocalContext.getServiceManager().
                                //                                   createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                //                 xLocalContext);

                                //unoidl.com.sun.star.uno.XComponentContext xRemoteContext =
                                //         (unoidl.com.sun.star.uno.XComponentContext)xURLResolver.resolve(
                                //                  "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext");

                                //unoidl.com.sun.star.lang.XMultiServiceFactory xRemoteFactory =
                                //         (unoidl.com.sun.star.lang.XMultiServiceFactory)
                                //                  xRemoteContext.getServiceManager();



                    //           var xLocalContext = Bootstrap.bootstrap();

                                //Get MultiServiceFactory
                                //unoidl.com.sun.star.lang.XMultiServiceFactory xRemoteFactory =
                                //    (unoidl.com.sun.star.lang.XMultiServiceFactory)
                                //        xLocalServiceManager;

              //////   var xMsf = (XMultiServiceFactory)xLocalContext.getServiceManager();
               /////// var xLocalContext =  boot();
                // Object desktop = xRemoteFactory.createInstance("com.sun.star.frame.Desktop");

                                //unoidl.com.sun.star.lang.XMultiServiceFactory xRemoteFactory =
                                //  (unoidl.com.sun.star.lang.XMultiServiceFactory)
                                //     xLocalContext.getServiceManager();
                 //////               Object desktop = xMsf.createInstance("com.sun.star.frame.Desktop");
                          ////      var aLoader = (XComponentLoader)desktop;
                    //normal
                                //Get a CompontLoader
                               //XComponentLoader aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
                               // //Load the sourcefile
                               //// TextBox1.Text = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                                XComponent xComponent = null;
                                try
                                {
                                    xComponent = initDocument(componentLoader,
                                        PathConverter(AppDomain.CurrentDomain.BaseDirectory + "1.odp"), "_blank");
                                    //Wait for loading
                                    while (xComponent == null)
                                    {
                                        System.Threading.Thread.Sleep(1000);
                                    }

                                    // save/export the document
                                    saveDocument(xComponent, AppDomain.CurrentDomain.BaseDirectory + "1.odp", PathConverter(AppDomain.CurrentDomain.BaseDirectory + "openOffice.ppt"));
                                    // Impersonate_User.Logout();//登出
                                }
                                catch
                                {
                                    throw;
                                }
                                finally
                                {
                                    xComponent.dispose();
                                }


                                // Check the identity.
                               // Console.WriteLine("After impersonation: "
                               //     + WindowsIdentity.GetCurrent().Name);
                            }
                        }
                        // Releasing the context object stops the impersonation
                        // Check the identity.
                        //Console.WriteLine("After closing the context: " + WindowsIdentity.GetCurrent().Name);
                    }
                }
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                Console.WriteLine("Exception occurred. " + ex.Message);
            }
            //return result;
}
    //    public static void startService() {
    //    DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
    //    try {
    //        System.out.println("准备启动服务....");
    //        configuration.setOfficeHome(OFFICE_HOME);// 设置OpenOffice.org安装目录
    //        configuration.setPortNumbers(port); // 设置转换端口，默认为8100
    //        configuration.setTaskExecutionTimeout(1000 * 60 * 5L);// 设置任务执行超时为5分钟
    //        configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);// 设置任务队列超时为24小时

    //        officeManager = configuration.buildOfficeManager();
    //        officeManager.start(); // 启动服务
    //        System.out.println("office转换服务启动成功!");
    //    } catch (Exception ce) {
    //        System.out.println("office转换服务启动失败!详细信息:" + ce);
    //    }
    //}

    //// 关闭服务器
    //public static void stopService() {
    //    System.out.println("关闭office转换服务....");
    //    if (officeManager != null) {
    //        officeManager.stop();
    //    }
    //    System.out.println("关闭office转换成功!");
    //} 

        private static void StartOpenOffice()
        {

            string name = WindowsIdentity.GetCurrent().Name;
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
                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardError = true;
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
            ////////IntPtr token = IntPtr.Zero;
            ////////WindowsImpersonationContext impersonatedUser = null;

            ////////private string _DomainName;
            ////////private string _strUserName;
            ////////private string _strPassword;

            ////////public PImpersonate(string DomainName, string strUserName, string strPassword)
            ////////{
            ////////    _DomainName = DomainName;
            ////////    _strUserName = strUserName;
            ////////    _strPassword = strPassword;

            ////////}
            ////////const int LOGON32_PROVIDER_DEFAULT = 0;
            ////////const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
            //////////public extern static bool CloseHandle(IntPtr hToken);
            //////////public void Login()
            //////////{
            //////////    IntPtr tokenHandle = new IntPtr(0);
            //////////    tokenHandle = IntPtr.Zero;
            //////////    try
            //////////    {
            //////////        bool result = LogonUser(_strUserName, _DomainName,
            //////////                                _strPassword,
            //////////                                LOGON32_LOGON_NEW_CREDENTIALS,
            //////////                               LOGON32_PROVIDER_DEFAULT,
            //////////                                ref tokenHandle);

            //////////        //bool result = LogonUser(_strUserName, _DomainName,
            //////////        //                       _strPassword,
            //////////        //                       LogonSessionType.Network,
            //////////        //                       LogonProvider.Default,
            //////////        //                       out token);
            //////////        if (result)
            //////////        {
            //////////            WindowsIdentity id = new WindowsIdentity(tokenHandle);
            //////////            // WindowsIdentity id = new WindowsIdentity(token);
            //////////            impersonatedUser = id.Impersonate();
            //////////        }
            //////////        else
            //////////        {
            //////////        }
            //////////    }
            //////////    catch
            //////////    {
            //////////    }
            //////////    finally
            //////////    {
            //////////    }
            //////////}
            ////////public bool Login()
            ////////{
            ////////    IntPtr tokenHandle = new IntPtr(0);
            ////////    tokenHandle = IntPtr.Zero;
            ////////    bool result=false;
            ////////    try
            ////////    {
            ////////         //LogonUser(_strUserName, _DomainName,
            ////////         //                       _strPassword,
            ////////         //                       LOGON32_LOGON_NEW_CREDENTIALS,
            ////////         //                      LOGON32_PROVIDER_DEFAULT,
            ////////         //                       ref tokenHandle);
            ////////        result = LogonUser(_strUserName, _DomainName,
            ////////                               _strPassword,
            ////////                               LOGON32_LOGON_NEW_CREDENTIALS,
            ////////                              LOGON32_PROVIDER_DEFAULT,
            ////////                               ref tokenHandle);
            ////////         //bool result = LogonUser(_strUserName, _DomainName,
            ////////         //                       _strPassword,
            ////////         //                       LogonSessionType.Network,
            ////////         //                       LogonProvider.Default,
            ////////         //                       out token);
            ////////        if (result)
            ////////        {
            ////////            WindowsIdentity id = new WindowsIdentity(tokenHandle);
            ////////           // WindowsIdentity id = new WindowsIdentity(token);
            ////////            impersonatedUser = id.Impersonate();
            ////////            //return true;
            ////////        }
            ////////        //else
            ////////        //{
            ////////        //    return false; ;
            ////////        //}
            ////////    }
            ////////    catch
            ////////    {
            ////////    }
            ////////    return result;

            ////////}

            ////////public void Logout()
            ////////{
            ////////    try
            ////////    {
            ////////        if (impersonatedUser != null)
            ////////            impersonatedUser.Undo();
            ////////        // Free the token
            ////////        if (token != IntPtr.Zero)
            ////////            CloseHandle(token);
            ////////    }
            ////////    catch
            ////////    {
            ////////    }
            ////////}

            ////////[DllImport("advapi32.dll", SetLastError = true)]
            ////////static extern bool LogonUser(
            ////////  string principal,
            ////////  string authority,
            ////////  string password,
            ////////  int dwLogonType,
            ////////  int dwLogonProvider,
            ////////  ref IntPtr phToken);

            //////////static extern bool LogonUser(
            //////////  string principal,
            //////////  string authority,
            //////////  string password,
            //////////  LogonSessionType logonType,
            //////////  LogonProvider logonProvider,
            //////////  out IntPtr token);
            ////////[DllImport("kernel32.dll", SetLastError = true)]
            ////////static extern bool CloseHandle(IntPtr handle);
            ////////enum LogonSessionType : uint
            ////////{
            ////////    Interactive = 2,
            ////////    Network,
            ////////    Batch,
            ////////    Service,
            ////////    NetworkCleartext = 8,
            ////////    NewCredentials
            ////////}
            ////////enum LogonProvider : uint
            ////////{
            ////////    Default = 0, // default for platform (use this!)
            ////////    WinNT35,     // sends smoke signals to authority
            ////////    WinNT40,     // uses NTLM
            ////////    WinNT50      // negotiates Kerb or NTLM
            ////////}
        //}
    }
}