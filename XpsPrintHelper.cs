using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;


namespace WPFPrintPages
{
    public class XpsPrintHelper
    {
        public static void Print(Stream stream, string printerName, string jobName, bool isWait
            , string outputFileName)//ファイル名
        {
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (printerName == null)
                throw new ArgumentNullException("printerName");

            IntPtr completionEvent = CreateEvent(IntPtr.Zero, true, false, null);
            if (completionEvent == IntPtr.Zero)
                throw new Win32Exception();

            try
            {
                IXpsPrintJob job;
                IXpsPrintJobStream jobStream;
                StartJob(printerName, jobName, completionEvent, out job, out jobStream, outputFileName);//ファイル名

                CopyJob(stream, job, jobStream);

                if (isWait)
                {
                    WaitForJob(completionEvent);
                    CheckJobStatus(job);
                }
            }
            finally
            {
                if (completionEvent != IntPtr.Zero)
                    CloseHandle(completionEvent);
            }
        }

        private static void StartJob(string printerName, string jobName, IntPtr completionEvent, out IXpsPrintJob job, out IXpsPrintJobStream jobStream
            , string outputFileName)//ファイル名
        {
            int result = StartXpsPrintJob
                (printerName, jobName, outputFileName //ファイル名
                , IntPtr.Zero, completionEvent
                , null, 0, out job, out jobStream, IntPtr.Zero);
            if (result != 0)
                throw new Win32Exception(result);
        }

        private static void CopyJob(Stream stream, IXpsPrintJob job, IXpsPrintJobStream jobStream)
        {
            try
            {
                byte[] buff = new byte[4096];
                while (true)
                {
                    uint read = (uint)stream.Read(buff, 0, buff.Length);
                    if (read == 0)
                        break;

                    uint written;
                    jobStream.Write(buff, read, out written);

                    if (read != written)
                        throw new Exception("Failed to copy data to the print job stream.");
                }

                // Indicate that the entire document has been copied. 
                jobStream.Close();
            }
            catch (Exception)
            {
                // Cancel the job if we had any trouble submitting it. 
                job.Cancel();
                throw;
            }
        }

        private static void WaitForJob(IntPtr completionEvent)
        {
            const int INFINITE = -1;
            switch (WaitForSingleObject(completionEvent, INFINITE))
            {
                case WAIT_RESULT.WAIT_OBJECT_0:
                    // Expected result, do nothing. 
                    break;
                case WAIT_RESULT.WAIT_FAILED:
                    throw new Win32Exception();
                default:
                    throw new Exception("Unexpected result when waiting for the print job.");
            }
        }

        private static void CheckJobStatus(IXpsPrintJob job)
        {
            XPS_JOB_STATUS jobStatus;
            job.GetJobStatus(out jobStatus);
            switch (jobStatus.completion)
            {
                case XPS_JOB_COMPLETION.XPS_JOB_COMPLETED:
                    // Expected result, do nothing. 
                    break;
                case XPS_JOB_COMPLETION.XPS_JOB_FAILED:
                    throw new Win32Exception(jobStatus.jobStatus);
                default:
                    throw new Exception("Unexpected print job status.");
            }
        }


        [DllImport("XpsPrint.dll", EntryPoint = "StartXpsPrintJob")]
        private static extern int StartXpsPrintJob(
            [MarshalAs(UnmanagedType.LPWStr)] String printerName,
            [MarshalAs(UnmanagedType.LPWStr)] String jobName,
            [MarshalAs(UnmanagedType.LPWStr)] String outputFileName, //こいつ
            IntPtr progressEvent,   // HANDLE 
            IntPtr completionEvent, // HANDLE 
            [MarshalAs(UnmanagedType.LPArray)] byte[] printablePagesOn,
            UInt32 printablePagesOnCount,
            out IXpsPrintJob xpsPrintJob,
            out IXpsPrintJobStream documentStream,
            IntPtr printTicketStream);  // This is actually "out IXpsPrintJobStream", but we don't use it and just want to pass null, hence IntPtr. 

        [DllImport("Kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);

        [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
        private static extern WAIT_RESULT WaitForSingleObject(IntPtr handle, Int32 milliseconds);

        [DllImport("Kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);
        
    }

    [Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D")]  // This is IID of ISequenatialSteam. 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IXpsPrintJobStream
    {
        // ISequentualStream methods. 
        void Read([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbWritten);
        // IXpsPrintJobStream methods. 
        void Close();
    }

    [Guid("5ab89b06-8194-425f-ab3b-d7a96e350161")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IXpsPrintJob
    {
        void Cancel();
        void GetJobStatus(out XPS_JOB_STATUS jobStatus);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct XPS_JOB_STATUS
    {
        public UInt32 jobId;
        public Int32 currentDocument;
        public Int32 currentPage;
        public Int32 currentPageTotal;
        public XPS_JOB_COMPLETION completion;
        public Int32 jobStatus; // UInt32 
    };

    enum XPS_JOB_COMPLETION
    {
        XPS_JOB_IN_PROGRESS = 0,
        XPS_JOB_COMPLETED = 1,
        XPS_JOB_CANCELLED = 2,
        XPS_JOB_FAILED = 3
    }

    enum WAIT_RESULT
    {
        WAIT_OBJECT_0 = 0,
        WAIT_ABANDONED = 0x80,
        WAIT_TIMEOUT = 0x102,
        WAIT_FAILED = -1 // 0xFFFFFFFF 
    }

}