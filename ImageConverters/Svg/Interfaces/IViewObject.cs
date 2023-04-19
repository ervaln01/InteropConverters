using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ImageConverters.Svg.Structs;

namespace ImageConverters.Svg.Interfaces
{
    [ComImport()]
    [Guid("0000010d-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IViewObject
    {
        int Draw(
            [MarshalAs(UnmanagedType.U4)] int dwDrawAspect,
            int lindex,
            IntPtr pvAspect,
            IntPtr ptd,
            IntPtr hdcTargetDev,
            IntPtr hdcDraw,
            ref Rectangle lprcBounds,
            IntPtr lprcWBounds,
            IntPtr pfnContinue,
            int dwContinue);

        int Freeze(uint dwDrawAspect, int lindex, IntPtr pvAspect, [Out] uint pdwFreeze);

        void GetAdvise([Out] uint pAspects, [Out] uint pAdvf, [Out] IAdviseSink ppAdvSink);

        int GetColorSet(uint dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hicTargetDev, [Out] IntPtr ppColorSet);

        void SetAdvise(uint aspects, uint ADVF, IAdviseSink pAdvSink);

        void Unfreeze(uint dwFreeze);
    }
}