using System;
using System.Diagnostics.SymbolStore;
using System.Runtime.InteropServices;
using System.Security.Permissions;

namespace Test
{

    /// <remarks>
    /// Adapted from http://sorin.serbans.net/blog/?p=257
    /// (which is adapted from 
    /// </remarks>
    public static class SymbolProvider
    {

        /// <remarks>
        /// We demand Unmanaged code permissions because we're reading
        /// from the file system and calling out to the Symbol Reader
        /// </remarks>
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        public static ISymbolReader GetSymbolReaderForFile(string pathModule, string searchPath)
        {
            var binder = new SymBinder();

            // Guids for imported metadata interfaces.
            var dispenserClassId = new Guid(0xe5cb7a31, 0x7512, 0x11d2, 0x89, 
                0xce, 0x00, 0x80, 0xc7, 0x92, 0xe5, 0xd8); // CLSID_CorMetaDataDispenser
            var dispenserIid = new Guid(0x809c652e, 0x7396, 0x11d2, 0x97, 0x71, 
                0x00, 0xa0, 0xc9, 0xb4, 0xd5, 0x0c); // IID_IMetaDataDispenser
            var importerIid = new Guid(0x7dac8207, 0xd3ae, 0x4c75, 0x9b, 0x67, 
                0x92, 0x80, 0x1a, 0x49, 0x7d, 0x44); // IID_IMetaDataImport

            // First create the Metadata dispenser.
            object objDispenser;
            NativeMethods.CoCreateInstance(ref dispenserClassId, null, 1, ref dispenserIid, out objDispenser);

            // Now open an Importer on the given filename. We'll end up passing this importer 
            // straight through to the Binder.
            object objImporter;
            var dispenser = (IMetaDataDispenser)objDispenser;
            dispenser.OpenScope(pathModule, 0, ref importerIid, out objImporter);

            var importerPtr = IntPtr.Zero;
            try
            {
                // This will manually AddRef the underlying object, so we need to 
                // be very careful to Release it.
                importerPtr = Marshal.GetComInterfaceForObject(objImporter, typeof(IMetadataImport));

                return binder.GetReader(importerPtr, pathModule, searchPath);
            }
            finally
            {
                if (importerPtr != IntPtr.Zero)
                {
                    Marshal.Release(importerPtr);
                }
            }
        }


        /// <remarks>
        /// We can use reflection-only load context to use reflection
        /// to query for metadata information rather than painfully
        /// import the com-classic metadata interfaces.
        /// </remarks>
        [ComVisible(true), Guid("809c652e-7396-11d2-9771-00a0c9b4d50c"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMetaDataDispenser
        {
            // We need to be able to call OpenScope, which is the 2nd vtable slot.
            // Thus we need this one placeholder here to occupy the first slot..
            void DefineScopePlaceholder();

            //STDMETHOD(OpenScope)(         // Return code.
            //  LPCWSTR     szScope,        // [in] The scope to open.
            //  DWORD       dwOpenFlags,    // [in] Open mode flags.
            //  REFIID      riid,           // [in] The interface desired.
            //  IUnknown    **ppIUnk) PURE; // [out] Return interface on success.
            void OpenScope([In, MarshalAs(UnmanagedType.LPWStr)] String szScope, 
                [In] Int32 dwOpenFlags, [In] ref Guid riid, 
                [Out, MarshalAs(UnmanagedType.IUnknown)] out Object punk);

            // Don't need any other methods.
        }


        /// <remarks>
        /// Since we're just blindly passing this interface through
        /// managed code to the Symbinder, we don't care about actually
        /// importing the specific methods. This needs to be public
        /// so that we can call Marshal.GetComInterfaceForObject() on
        /// it to get the underlying metadata pointer.
        /// </remarks>
        [ComVisible(true), Guid("7DAC8207-D3AE-4c75-9B67-92801A497D44"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMetadataImport
        {
            // Just need a single placeholder method so that it doesn't complain
            // about an empty interface.
            void Placeholder();
        }


        // PInvoke imports.
        private static class NativeMethods
        {
            [DllImport("ole32.dll")]
            public static extern int CoCreateInstance(
                [In] ref Guid rclsid,
                [In, MarshalAs(UnmanagedType.IUnknown)] Object pUnkOuter,
                [In] uint dwClsContext,
                [In] ref Guid riid,
                [Out, MarshalAs(UnmanagedType.Interface)] out Object ppv);
        }

    }

}
