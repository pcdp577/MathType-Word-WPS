using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;

public static class MathTypeNativeBridge
{
    private const uint CLSCTX_INPROC_SERVER = 4;
    private static readonly Guid ClsidMathTypeDataObject = new Guid("0002CE03-0000-0000-C000-000000000046");
    private static readonly Guid IidIUnknown = new Guid("00000000-0000-0000-C000-000000000046");

    [DllImport("ole32.dll", ExactSpelling = true, PreserveSig = true)]
    private static extern uint CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppv);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    private static extern bool GlobalUnlock(IntPtr hMem);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    private static extern int GlobalSize(IntPtr hMem);

    private static System.Runtime.InteropServices.ComTypes.IDataObject CreateMathTypeDataObject()
    {
        object instance;
        Guid clsid = ClsidMathTypeDataObject;
        Guid iid = IidIUnknown;
        uint hr = CoCreateInstance(ref clsid, IntPtr.Zero, CLSCTX_INPROC_SERVER, ref iid, out instance);
        if (hr != 0 || instance == null)
            throw new COMException("Cannot create MathType DataObject", unchecked((int)hr));
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject =
            instance as System.Runtime.InteropServices.ComTypes.IDataObject;
        if (dataObject == null)
            throw new InvalidCastException("MathType DataObject does not expose IDataObject.");
        return dataObject;
    }

    private static FORMATETC Format(string name, TYMED tymed)
    {
        return new FORMATETC
        {
            cfFormat = (short)DataFormats.GetFormat(name).Id,
            dwAspect = DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = tymed
        };
    }

    private static byte[] ReadHGlobal(IntPtr handle)
    {
        IntPtr ptr = GlobalLock(handle);
        if (ptr == IntPtr.Zero)
            throw new InvalidOperationException("GlobalLock failed.");
        try
        {
            int size = GlobalSize(handle);
            byte[] data = new byte[size];
            Marshal.Copy(ptr, data, 0, size);
            return data;
        }
        finally
        {
            GlobalUnlock(handle);
        }
    }

    public static byte[] MathMlToMtef(string mathml)
    {
        System.Runtime.InteropServices.ComTypes.IDataObject dataObject = CreateMathTypeDataObject();
        try
        {
            FORMATETC inputFormat = Format("MathML", TYMED.TYMED_HGLOBAL);
            STGMEDIUM inputMedium = new STGMEDIUM
            {
                tymed = TYMED.TYMED_HGLOBAL,
                unionmember = Marshal.StringToHGlobalUni(mathml),
                pUnkForRelease = null
            };
            dataObject.SetData(ref inputFormat, ref inputMedium, false);
            Marshal.FreeHGlobal(inputMedium.unionmember);

            foreach (TYMED tymed in new[] { TYMED.TYMED_HGLOBAL, TYMED.TYMED_ISTORAGE })
            {
                FORMATETC outputFormat = Format("MathType EF", tymed);
                STGMEDIUM outputMedium;
                try
                {
                    dataObject.GetData(ref outputFormat, out outputMedium);
                    if (outputMedium.unionmember != IntPtr.Zero)
                        return ReadHGlobal(outputMedium.unionmember);
                }
                catch (COMException)
                {
                    // Try the next storage medium.
                }
            }
            throw new InvalidOperationException("MathType DataObject did not return MathType EF data.");
        }
        finally
        {
            Marshal.FinalReleaseComObject(dataObject);
        }
    }

    private static string JsonEscape(string text)
    {
        return text.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    public static int Main(string[] args)
    {
        string mathmlFile = "";
        string outputMtef = "";
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--mathml-file" && i + 1 < args.Length) mathmlFile = args[++i];
            else if (args[i] == "--output-mtef" && i + 1 < args.Length) outputMtef = args[++i];
        }
        if (string.IsNullOrWhiteSpace(mathmlFile) || string.IsNullOrWhiteSpace(outputMtef))
        {
            Console.Error.WriteLine("usage: mathtype_native_bridge --mathml-file <in.mathml> --output-mtef <out.mtef>");
            return 2;
        }
        try
        {
            string mathml = File.ReadAllText(mathmlFile, Encoding.UTF8);
            byte[] mtef = MathMlToMtef(mathml);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputMtef)));
            File.WriteAllBytes(outputMtef, mtef);
            Console.WriteLine("{\"ok\":true,\"output_mtef\":\"" + JsonEscape(Path.GetFullPath(outputMtef)) + "\",\"bytes\":" + mtef.Length + "}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("{\"ok\":false,\"error\":\"" + JsonEscape(ex.GetType().Name + ": " + ex.Message) + "\"}");
            return 1;
        }
    }
}
